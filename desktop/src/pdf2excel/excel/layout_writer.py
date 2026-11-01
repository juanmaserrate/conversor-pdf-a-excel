"""Writer del modo "Igual al PDF" (docs/LAYOUT_MODE.md §4): una `LayoutPage` -> una hoja de Excel.

Sin Qt. Reproduce la grilla de la página (anchos/altos proporcionales), combina rangos, aplica bordes y
rellenos, texto con estilos (rich text), imágenes ancladas y deja la hoja lista para imprimir en A4 ajustada
al ancho, sin líneas de cuadrícula.

Proporciones: las columnas y filas se escriben con el tamaño exacto que tienen en el PDF (nada de agrandar la
grilla "por las dudas": como la hoja se imprime ajustada al ancho, agrandar la grilla achica el texto). El
texto que no entra en su celda, medido con las métricas reales de Helvetica/Times/Courier (Calibri para el
resto), se achica **solo en esa celda** (`cell_font_scale`, hasta 0.7), de modo que el resto conserva el
tamaño del PDF.
"""

from __future__ import annotations

import math
import re
from bisect import bisect_right
from collections.abc import Callable
from datetime import date, datetime
from io import BytesIO
from typing import Any

import pymupdf
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XlImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from pdf2excel.core.layout_text import OCR_DOUBT_WARNING, doubtful_ocr_tokens
from pdf2excel.excel.naming import MAX_SHEET_NAME, RESERVED_SHEET_NAMES, clean_sheet_title
from pdf2excel.models import (
    BorderSide,
    ConversionReport,
    HAlign,
    LayoutCell,
    LayoutImage,
    LayoutPage,
    Options,
    TextRun,
)

MAX_COLS = 16384
MAX_ROWS = 1_048_576
MIN_COL_WIDTH = 0.2  # caracteres (≈ 1 pt); más finas que esto no tienen sentido
MAX_COL_WIDTH = 255.0
MAX_ROW_HEIGHT = 409.0  # puntos
MIN_ROW_HEIGHT = 1.0
MIN_FONT_PT = 6.0  # no se achica por debajo de esto (salvo que el PDF ya use menos)
MAX_TEXT_CHARS = 32000
EMU_PER_PT = 12700

FONT_SCALE_STEPS = (1.0, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7)
FIT_TOLERANCE = 1.04  # holgura al decidir si un texto "entra"
CELL_PAD_W_PT = 3.0  # margen interno horizontal total de una celda de Excel
CELL_PAD_H_PT = 1.0
LINE_FACTOR = 1.2  # alto de una línea de Excel respecto del tamaño de la fuente
CALIBRI_RATIO = 0.92  # ancho de Calibri (fuente por defecto) respecto de Helvetica
EXCEL_PT_PER_WIDTH = 5.25  # un "ancho de columna" de Excel = 7 px (Calibri 11) = 5.25 pt, relleno incluido

_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_SPLIT_WORDS_RE = re.compile(r"\S+\s*")

_SideKey = tuple[str, float, str]


# --------------------------------------------------------------------------------------- conversiones


def col_width_chars(width_pt: float) -> float:
    """Puntos -> atributo `width` de la columna. En el archivo el ancho ya incluye el relleno de 5 px: una
    unidad = 7 px (Calibri 11) = 5.25 pt, así que la conversión es lineal (sin restar el relleno)."""
    return min(MAX_COL_WIDTH, max(MIN_COL_WIDTH, round(width_pt / EXCEL_PT_PER_WIDTH, 2)))


def row_height_pt(height_pt: float) -> float:
    return min(MAX_ROW_HEIGHT, max(MIN_ROW_HEIGHT, round(height_pt, 2)))


def _rgb(color: str | None) -> str | None:
    if not color:
        return None
    c = color.lstrip("#").upper()
    return c if re.fullmatch(r"[0-9A-F]{6}", c) else None


def _clean(text: str) -> str:
    return _CONTROL_RE.sub("", ILLEGAL_CHARACTERS_RE.sub("", text))


def _border_style(side: BorderSide | None) -> Side | None:
    if side is None or side.width_pt <= 0:
        return None
    style = "thin" if side.width_pt < 1.2 else "medium" if side.width_pt < 2.5 else "thick"
    return Side(style=style, color=_rgb(side.color) or "000000")


def _sheet_title(prefix: str, page_no: int, used: set[str]) -> str:
    """`Página N` o `<prefijo> · Página N` (≤ 31); si ya existe se numera el prefijo, no la página."""
    label = f"Página {page_no}"
    taken = {u.casefold() for u in used} | {r.casefold() for r in RESERVED_SHEET_NAMES}
    clean = clean_sheet_title(prefix)
    n = 1
    while True:
        tag = clean if n == 1 else f"{clean or 'Hoja'} ({n})"
        if tag:
            room = MAX_SHEET_NAME - len(label) - 3
            name = f"{tag[:room].rstrip()} · {label}" if room >= 1 else label
            if n > 1 and len(tag) > room:
                suffix = f" ({n})"
                name = f"{tag[: max(1, room - len(suffix))].rstrip()}{suffix} · {label}"
        else:
            name = label if n == 1 else f"{label} ({n})"[:MAX_SHEET_NAME]
        if name.casefold() not in taken:
            used.add(name)
            return name
        n += 1


# --------------------------------------------------------------------------- estimación del texto


_FONT_CODES = {
    "Arial": ("helv", "hebo", "heit", "hebi"),
    "Times New Roman": ("tiro", "tibo", "tiit", "tibi"),
    "Courier New": ("cour", "cobo", "coit", "cobi"),
}
_font_cache: dict[tuple[str, bool, bool], Any] = {}


def _metric_font(run: TextRun) -> tuple[Any, float]:
    """Fuente base-14 de PyMuPDF con las mismas métricas que la de Excel y factor de ancho (Calibri ≈ 0.92 ×
    Helvetica)."""
    family = run.font if run.font in _FONT_CODES else "Arial"
    key = (family, run.bold, run.italic)
    font = _font_cache.get(key)
    if font is None:
        codes = _FONT_CODES[family]
        font = pymupdf.Font(codes[(1 if run.bold else 0) + (2 if run.italic else 0)])
        _font_cache[key] = font
    return font, (1.0 if run.font in _FONT_CODES else CALIBRI_RATIO)


def _text_width(text: str, run: TextRun, size: float) -> float:
    """Ancho en puntos de `text` con el estilo de `run` a `size` pt."""
    font, ratio = _metric_font(run)
    try:
        return float(font.text_length(text, fontsize=size)) * ratio
    except Exception:  # glifos que la fuente base no tiene
        return len(text) * size * 0.55 * ratio


def _run_width(run: TextRun, size: float) -> float:
    return _text_width(run.text, run, size)


def _run_size(run: TextRun, font_scale: float) -> float:
    """Tamaño de Excel de un run: el del PDF por la escala, a medios puntos, sin bajar de 6 pt salvo que el
    PDF ya use menos."""
    floor = min(MIN_FONT_PT, run.size_pt)
    return max(floor, round(run.size_pt * font_scale * 2) / 2)


def _wrapped_lines(line: list[TextRun], avail: float, font_scale: float) -> int:
    """Cantidad de renglones visuales que ocupa una línea lógica al envolverla en `avail` puntos."""
    n_lines = 1
    used = 0.0
    for run in line:
        size = _run_size(run, font_scale)
        for m in _SPLIT_WORDS_RE.finditer(run.text):
            w = _text_width(m.group(0).rstrip(), run, size)
            full = _text_width(m.group(0), run, size)
            if used > 0 and used + w > avail:
                n_lines += 1
                used = 0.0
            # una palabra más ancha que la celda se parte igual en Excel: cuenta los renglones extra
            if w > avail > 0:
                extra = int(w // avail)
                n_lines += extra
                used = w - extra * avail
            else:
                used += full
    return n_lines


def _text_need(
    cell: LayoutCell, width_pt: float, font_scale: float, overflow_pt: float = 0.0
) -> tuple[bool, float]:
    """`(entra_a_lo_ancho, alto_necesario_pt)` del texto de la celda con esa escala de fuente.

    `overflow_pt`: ancho libre al lado de la celda (pistas vacías) sobre el que un texto de una sola línea sin
    envolver se derrama en Excel; solo cuenta para decidir si entra a lo ancho, no para envolver."""
    if not cell.lines:
        return True, 0.0
    avail_w = max(1.0, width_pt - CELL_PAD_W_PT)
    need_h = CELL_PAD_H_PT * 2
    width_ok = True
    for line in cell.lines:
        if not line:
            continue
        size = max((_run_size(r, font_scale) for r in line), default=MIN_FONT_PT)
        if cell.wrap or len(cell.lines) > 1:
            n = _wrapped_lines(line, avail_w * FIT_TOLERANCE, font_scale)
        else:
            n = 1
            if (
                sum(_run_width(r, _run_size(r, font_scale)) for r in line)
                > (avail_w + overflow_pt) * FIT_TOLERANCE
            ):
                width_ok = False
        # un tilde o una cruz (sin letras) es un símbolo: no necesita el interlineado de un texto
        factor = LINE_FACTOR if any(ch.isalnum() for r in line for ch in r.text) else 1.0
        need_h += n * size * factor
    return width_ok, need_h


def _cell_fits(
    cell: LayoutCell, width_pt: float, height_pt: float, font_scale: float, overflow_pt: float = 0.0
) -> bool:
    width_ok, need_h = _text_need(cell, width_pt, font_scale, overflow_pt)
    return width_ok and need_h <= height_pt * FIT_TOLERANCE


def _prefix_sums(values: list[float]) -> list[float]:
    out = [0.0]
    for v in values:
        out.append(out[-1] + v)
    return out


def cell_font_scale(cell: LayoutCell, width_pt: float, height_pt: float, overflow_pt: float = 0.0) -> float:
    """Mayor escala de fuente (1.0 a 0.7) con la que el texto de la celda entra en `width_pt` × `height_pt`.
    Se decide celda por celda: las que entran con el tamaño del PDF no se tocan. `overflow_pt` es el ancho de
    pistas vacías contiguas sobre las que se puede derramar un texto de una sola línea."""
    for fs in FONT_SCALE_STEPS:
        if _cell_fits(cell, width_pt, height_pt, fs, overflow_pt):
            return fs
    return FONT_SCALE_STEPS[-1]


# --------------------------------------------------------------------------------------- estilos


class _StyleCache:
    """Evita crear miles de objetos de estilo idénticos."""

    def __init__(self) -> None:
        self._fonts: dict[Any, Font] = {}
        self._fills: dict[str, PatternFill] = {}
        self._borders: dict[Any, Border] = {}
        self._aligns: dict[Any, Alignment] = {}
        self._sides: dict[_SideKey | None, Side] = {}

    def font(self, run: TextRun, font_scale: float) -> Font:
        key = (run.bold, run.italic, _run_size(run, font_scale), _rgb(run.color), run.font)
        f = self._fonts.get(key)
        if f is None:
            color = _rgb(run.color)
            f = Font(
                name=run.font or "Calibri",
                sz=key[2],
                b=run.bold,
                i=run.italic,
                color=("FF" + color) if color else None,
            )
            self._fonts[key] = f
        return f

    def fill(self, color: str) -> PatternFill:
        f = self._fills.get(color)
        if f is None:
            f = PatternFill(fill_type="solid", start_color="FF" + color, end_color="FF" + color)
            self._fills[color] = f
        return f

    def side(self, bs: BorderSide | None) -> Side:
        s = _border_style(bs)
        if s is None:
            return Side()
        key: _SideKey = (str(s.style), 0.0, str(s.color.rgb) if s.color is not None else "")
        cached = self._sides.get(key)
        if cached is None:
            self._sides[key] = s
            cached = s
        return cached

    def border(
        self,
        left: BorderSide | None,
        top: BorderSide | None,
        right: BorderSide | None,
        bottom: BorderSide | None,
    ) -> Border | None:
        if left is None and top is None and right is None and bottom is None:
            return None
        key = (left, top, right, bottom)
        b = self._borders.get(key)
        if b is None:
            b = Border(
                left=self.side(left), top=self.side(top), right=self.side(right), bottom=self.side(bottom)
            )
            self._borders[key] = b
        return b

    def align(self, horizontal: str, vertical: str, wrap: bool) -> Alignment:
        key = (horizontal, vertical, wrap)
        a = self._aligns.get(key)
        if a is None:
            a = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)
            self._aligns[key] = a
        return a


def _same_style(a: TextRun, b: TextRun) -> bool:
    return (a.bold, a.italic, a.size_pt, a.color, a.font) == (b.bold, b.italic, b.size_pt, b.color, b.font)


def _inline_font(run: TextRun, font_scale: float) -> InlineFont:
    color = _rgb(run.color)
    return InlineFont(
        rFont=run.font or "Calibri",
        sz=_run_size(run, font_scale),
        b=run.bold,
        i=run.italic,
        color=("FF" + color) if color else None,
    )


def _text_value(cell: LayoutCell, font_scale: float) -> tuple[str | CellRichText | None, TextRun | None]:
    """Contenido de texto de la celda y, si tiene un solo estilo, el run que lo representa."""
    lines = [[r for r in line if r.text] for line in cell.lines]
    lines = [
        [TextRun(_clean(r.text), r.bold, r.italic, r.size_pt, r.color, r.font) for r in line]
        for line in lines
    ]
    flat = [r for line in lines for r in line if r.text]
    if not flat:
        return None, None
    if all(_same_style(flat[0], r) for r in flat):
        text = "\n".join("".join(r.text for r in line) for line in lines)
        return text[:MAX_TEXT_CHARS], flat[0]
    blocks: list[TextBlock] = []
    total = 0
    for i, line in enumerate(lines):
        for j, run in enumerate(line):
            text = run.text
            if j == len(line) - 1 and i < len(lines) - 1:
                text += "\n"
            total += len(text)
            if total > MAX_TEXT_CHARS:
                break
            blocks.append(TextBlock(_inline_font(run, font_scale), text))
        if total > MAX_TEXT_CHARS:
            break
        if not line and i < len(lines) - 1 and blocks:
            # línea vacía en medio: conservar el salto
            last = blocks[-1]
            blocks[-1] = TextBlock(last.font, last.text + "\n")
    return CellRichText(*blocks), None


def _write_value(ws_cell: Any, cell: LayoutCell, font_scale: float, styles: _StyleCache) -> None:
    value = cell.value
    if value is not None and not isinstance(value, str) and not isinstance(value, bool):
        ws_cell.value = value
        if cell.number_format:
            ws_cell.number_format = cell.number_format
        elif isinstance(value, (date, datetime)):
            ws_cell.number_format = "dd/mm/yyyy"
        ws_cell.font = styles.font(
            cell.lines[0][0] if cell.lines and cell.lines[0] else TextRun(""), font_scale
        )
        return
    content, single = _text_value(cell, font_scale)
    if content is None:
        return
    if isinstance(content, CellRichText):
        ws_cell.value = content
    else:
        ws_cell.value = content
        if content.startswith("="):
            ws_cell.data_type = "s"  # que Excel no lo tome por fórmula
        elif cell.number_format is None and re.fullmatch(r"[\d.,\s%$-]+", content):
            ws_cell.data_type = "s"  # parece número pero se decidió dejarlo como texto
        if single is not None:
            ws_cell.font = styles.font(single, font_scale)


# --------------------------------------------------------------------------------------- hoja


def _clip(cell: LayoutCell, n_rows: int, n_cols: int) -> tuple[int, int, int, int] | None:
    """`(r0, c0, r1, c1)` inclusivos, recortado a la grilla; None si queda fuera."""
    if cell.row < 0 or cell.col < 0 or cell.row >= n_rows or cell.col >= n_cols:
        return None
    r1 = min(cell.row + max(1, cell.row_span) - 1, n_rows - 1)
    c1 = min(cell.col + max(1, cell.col_span) - 1, n_cols - 1)
    return cell.row, cell.col, r1, c1


MIN_DONOR_ROW_PT = 1.5  # una pista de relleno no se achica por debajo de esto
DONOR_REACH = 3  # cuántas filas hacia cada lado puede recorrer para pedir alto prestado sin cruzar una línea
FAR_DONOR_REACH = (
    60  # segunda pasada: cualquier pista SIN texto, cruzando líneas (todo lo de en medio se corre junto)
)
FAR_DONOR_SHARE = 0.6  # de una pista vacía solo se presta hasta esta parte de su alto
PX_PT = 0.75  # un píxel a 96 ppp: Excel y LibreOffice redondean el alto de fila a píxeles enteros


def _occupancy(page: LayoutPage, n_rows: int, n_cols: int) -> set[tuple[int, int]]:
    occ: set[tuple[int, int]] = set()
    for cell in page.cells:
        box = _clip(cell, n_rows, n_cols)
        if box is None:
            continue
        r0, c0, r1, c1 = box
        occ.update((r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1))
    return occ


def overflow_room(
    cell: LayoutCell, box: tuple[int, int, int, int], occ: set[tuple[int, int]], xs: list[float]
) -> float:
    """Solo el ancho de `overflow_layout`."""
    return overflow_layout(cell, box, occ, xs)[0]


def overflow_layout(
    cell: LayoutCell, box: tuple[int, int, int, int], occ: set[tuple[int, int]], xs: list[float]
) -> tuple[float, HAlign]:
    """Ancho (pt) de pistas libres contiguas sobre las que se derrama el texto de una sola línea sin envolver.

    Excel dibuja el texto de una celda (alineado a la izquierda, a la derecha o centrado) sobre las celdas
    vecinas mientras estén vacías; una caja o un texto vecino lo cortan. Solo cuentan las pistas en las que
    ninguna celda del extractor ocupa NINGUNA de las filas de la celda. Devuelve `(ancho, alineación)`: un
    texto centrado que solo tiene lugar de un lado (número de página pegado al borde) se alinea hacia ese
    lado, que es hacia donde se corre apenas."""
    if cell.wrap or len(cell.lines) != 1:
        return 0.0, cell.halign
    r0, c0, r1, c1 = box
    n_cols = len(xs) - 1

    def free(c: int) -> bool:
        return all((r, c) not in occ for r in range(r0, r1 + 1))

    right = 0.0
    c = c1 + 1
    # una caja con borde no deja pasar el texto por encima de su propio borde
    while cell.border_right is None and c < n_cols and free(c):
        right += xs[c + 1] - xs[c]
        c += 1
    left = 0.0
    c = c0 - 1
    while cell.border_left is None and c >= 0 and free(c):
        left += xs[c + 1] - xs[c]
        c -= 1
    if cell.halign == "center":
        if min(left, right) <= 0.0 < max(left, right):
            return max(left, right), "right" if left > right else "left"
        return 2 * min(left, right), "center"
    if cell.halign == "right":
        return (left, "right") if left > 0.0 or right <= 0.0 else (right, "left")
    return (right, cell.halign) if right > 0.0 or left <= 0.0 else (left, "right")


def balanced_row_heights(
    page: LayoutPage, n_rows: int, n_cols: int, *, pixel_snap: bool = False
) -> list[float]:
    """Alto de cada fila, corrigiendo las pistas más bajas que el renglón que contienen.

    El extractor mide las filas de texto con el alto de la letra (8 pt para una letra de 8 pt), pero Excel
    necesita ~1.2 × el tamaño: sin esto la letra se achicaba al 70 %. Acá la fila baja le "pide prestado" alto
    a las pistas vecinas SIN texto (los espacios entre renglones), de modo que el total no cambia y los bordes
    exteriores quedan donde estaban: sin borde, el texto crece hacia los dos lados (el centro no se mueve).
    Primero se usan las pistas cercanas sin cruzar ninguna línea dibujada (un subrayado crece solo hacia
    arriba, un borde superior solo hacia abajo, con ambos no se toca). Si eso no alcanza (casillas con borde
    arriba y abajo en un renglón de 8 pt, un título de 15 pt sin aire), una segunda pasada pide el resto a la
    pista vacía más cercana aunque haya líneas en el medio: todo lo que queda entre las dos se corre unos
    puntos en bloque (nada se deforma), pero la letra conserva el tamaño del PDF. Nunca se le quita alto a una
    pista con texto. `pixel_snap=True` deja además cada alto en un múltiplo de píxel (`quantize_row_heights`).
    """
    heights = [max(0.0, h) for h in page.row_heights_pt[:n_rows]]
    xs = _prefix_sums([max(0.0, w) for w in page.col_widths_pt[:n_cols]])
    occ = _occupancy(page, n_rows, n_cols)
    multi: list[
        tuple[int, int, float]
    ] = []  # (fila_ini, fila_fin, alto_que_aun_sobra) de textos de varias filas
    singles: list[tuple[int, float]] = []
    required = [0.0] * n_rows
    line_at = [False] * (n_rows + 1)  # línea de borde en el límite entre la fila b-1 y la b
    for cell in page.cells:
        box = _clip(cell, n_rows, n_cols)
        if box is None:
            continue
        if cell.border_top is not None:
            line_at[box[0]] = True
        if cell.border_bottom is not None:
            line_at[box[2] + 1] = True
    for cell in page.cells:
        box = _clip(cell, n_rows, n_cols)
        if box is None or not cell.lines:
            continue
        r0, c0, r1, c1 = box
        room = overflow_room(cell, box, occ, xs)
        width_ok, need_h = _text_need(cell, xs[c1 + 1] - xs[c0], 1.0, room)
        if r1 > r0:
            multi.append((r0, r1, sum(heights[r0 : r1 + 1]) - need_h / FIT_TOLERANCE))
            continue
        need = need_h / FIT_TOLERANCE + 0.05  # margen para que la comparación de _cell_fits no quede justa
        required[r0] = max(required[r0], need)
        avail = max(1.0, xs[c1 + 1] - xs[c0] - CELL_PAD_W_PT) * FIT_TOLERANCE
        lines = [ln for ln in cell.lines if ln]
        wraps = sum(_wrapped_lines(ln, avail, 1.0) for ln in lines) > len(lines)
        # solo se corrigen filas "justas de alto de letra"; un texto que se envuelve en más renglones que el
        # PDF es otro problema (se resuelve achicando la letra de esa celda)
        if width_ok and not wraps:
            singles.append((r0, need))

    def donor_slack(j: int, row: int) -> float:
        room = max(0.0, heights[j] - max(required[j], MIN_DONOR_ROW_PT))
        for a, b, spare in multi:  # un texto de varias filas que cruza la donante conserva su alto
            if a <= j <= b and not a <= row <= b:
                room = min(room, max(0.0, spare))
        return room

    def apply(row: int, takes: dict[int, float]) -> float:
        nonlocal multi
        got = sum(takes.values())
        if got <= 0.01:
            return 0.0
        for j, t in takes.items():
            heights[j] -= t
            multi = [(a, b, sp - t if a <= j <= b and not a <= row <= b else sp) for a, b, sp in multi]
        heights[row] += got
        multi = [(a, b, sp + got if a <= row <= b else sp) for a, b, sp in multi]
        return got

    for r, need in singles:
        deficit = need - heights[r]
        if deficit <= 0.01:
            continue
        takes: dict[int, float] = {}
        remaining = deficit
        # a la par (el centro no se mueve) y, si un lado no alcanza o está trabado por una línea, el otro
        # completa el resto
        for share in (deficit / 2, deficit):
            for side in (-1, 1):
                want = min(share if share != deficit else remaining, remaining)
                got_side = 0.0
                for k in range(1, DONOR_REACH + 1):
                    j = r + side * k
                    if not 0 <= j < n_rows or want - got_side <= 0.01:
                        break
                    # límites que se correrían: de la fila r a la donante j
                    crossed = range(j + 1, r + 1) if side < 0 else range(r + 1, j + 1)
                    if any(line_at[b] for b in crossed):
                        break
                    t = min(donor_slack(j, r) - takes.get(j, 0.0), want - got_side)
                    if t > 0.01:
                        takes[j] = takes.get(j, 0.0) + t
                        got_side += t
                remaining -= got_side
            if remaining <= 0.01:
                break
        apply(r, takes)

    # segunda pasada: lo que falta se pide a la pista vacía más cercana, cruzando líneas si hace falta
    wanted = [(r, need) for r, need in singles]
    for _a, b, spare in list(multi):  # un texto de varias filas que no entra crece en su última fila
        if spare < -0.01:
            wanted.append((b, heights[b] - spare))
    for r, need in wanted:
        remaining = need - heights[r]
        if remaining <= 0.01:
            continue
        takes = {}
        for k in range(1, FAR_DONOR_REACH + 1):
            for side in (-1, 1):
                j = r + side * k
                if not 0 <= j < n_rows or remaining <= 0.01 or required[j] > 0.0:
                    continue
                if any(a <= j <= b and not a <= r <= b for a, b, _ in multi):
                    continue
                cap = min(donor_slack(j, r), heights[j] * FAR_DONOR_SHARE) - takes.get(j, 0.0)
                t = min(cap, remaining)
                if t > 0.01:
                    takes[j] = takes.get(j, 0.0) + t
                    remaining -= t
            if remaining <= 0.01:
                break
        apply(r, takes)
    if pixel_snap:
        # una fila cuyo texto entra (con tolerancia) no puede quedar más baja que lo que necesita al redondear
        minimums = [
            max(0.0, required[i] - 0.05) if heights[i] >= required[i] - 0.05 - 1e-6 else 0.0
            for i in range(n_rows)
        ]
        return quantize_row_heights(heights, minimums)
    return heights


def quantize_row_heights(heights: list[float], minimums: list[float] | None = None) -> list[float]:
    """Alto de cada fila como múltiplo exacto de un píxel (0.75 pt), sin perder posición acumulada.

    Excel y LibreOffice redondean el alto de cada fila a píxeles enteros (LibreOffice hacia abajo): con 45
    filas de 20 pt el contenido de abajo subía ~6 %. Se escribe cada fila en un múltiplo de 0.75 pt y el error
    se arrastra a la fila siguiente, de modo que la posición acumulada nunca se aparta más de medio píxel.
    `minimums`: alto mínimo de cada fila (el que necesita su texto), redondeado hacia arriba; las filas
    siguientes lo compensan."""
    out: list[float] = []
    target = 0.0
    placed = 0.0
    for i, h in enumerate(heights):
        target += max(0.0, h)
        floor = placed + PX_PT
        if minimums is not None and minimums[i] > 0:
            floor = max(floor, placed + math.ceil(minimums[i] / PX_PT - 1e-9) * PX_PT)
        end = max(floor, round(target / PX_PT) * PX_PT)
        out.append(end - placed)
        placed = end
    return out


def _grows_up(cell: LayoutCell) -> bool:
    return bool(cell.lines) and cell.border_bottom is not None and cell.border_top is None


def _write_cells(
    ws: Worksheet, page: LayoutPage, n_rows: int, n_cols: int, styles: _StyleCache, row_pts: list[float]
) -> list[LayoutCell]:
    xs = _prefix_sums(page.col_widths_pt[:n_cols])
    ys = _prefix_sums(row_pts)
    all_occ = _occupancy(page, n_rows, n_cols)
    occupied: set[tuple[int, int]] = set()
    placed: list[LayoutCell] = []
    for cell in page.cells:
        box = _clip(cell, n_rows, n_cols)
        if box is None:
            continue
        r0, c0, r1, c1 = box
        valign = "bottom" if _grows_up(cell) else cell.valign
        halign = cell.halign
        # un rango que pisa a otro ya escrito se reduce a su celda de arriba a la izquierda
        if any((r, c) in occupied for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)):
            r1, c1 = r0, c0
            if (r0, c0) in occupied:
                continue
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                occupied.add((r, c))
        placed.append(cell)

        if r1 > r0 or c1 > c0:
            # primero se combina y después se estilan TODAS las celdas (si no, openpyxl pierde el relleno)
            ws.merge_cells(start_row=r0 + 1, start_column=c0 + 1, end_row=r1 + 1, end_column=c1 + 1)
        fill = styles.fill(_rgb(cell.fill) or "FFFFFF") if _rgb(cell.fill) else None
        room = 0.0
        if cell.lines:
            room, spill_align = overflow_layout(cell, (r0, c0, r1, c1), all_occ, xs)
            if not _text_need(cell, xs[c1 + 1] - xs[c0], 1.0)[0]:
                halign = spill_align  # solo si el texto no entra en su celda se corre hacia donde hay lugar
        align = styles.align(halign, valign, bool(cell.wrap or len(cell.lines) > 1))
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                xc = ws.cell(row=r + 1, column=c + 1)
                border = styles.border(
                    cell.border_left if c == c0 else None,
                    cell.border_top if r == r0 else None,
                    cell.border_right if c == c1 else None,
                    cell.border_bottom if r == r1 else None,
                )
                if border is not None:
                    xc.border = border
                if fill is not None:
                    xc.fill = fill
                xc.alignment = align
        font_scale = (
            cell_font_scale(cell, xs[c1 + 1] - xs[c0], ys[r1 + 1] - ys[r0], room) if cell.lines else 1.0
        )
        _write_value(ws.cell(row=r0 + 1, column=c0 + 1), cell, font_scale, styles)
        if OCR_DOUBT_WARNING in page.warnings:
            doubtful = doubtful_ocr_tokens(cell.text)
            if doubtful:
                note = Comment(
                    "El OCR no está seguro de esta lectura ("
                    + ", ".join(f"«{t}»" for t in doubtful[:3])
                    + "). Compará con el PDF y corregila si hace falta.",
                    "Conversor PDF a Excel",
                )
                note.width, note.height = 260, 70
                ws.cell(row=r0 + 1, column=c0 + 1).comment = note
    return placed


def _track_at(prefix: list[float], value: float) -> tuple[int, float]:
    """Pista (0-based) que contiene a `value` y la distancia desde su inicio."""
    n = len(prefix) - 1
    i = max(0, min(n - 1, bisect_right(prefix, value + 0.01) - 1))
    return i, max(0.0, min(value - prefix[i], prefix[i + 1] - prefix[i]))


def _add_image(ws: Worksheet, img: LayoutImage, xs: list[float], ys: list[float]) -> None:
    """Imagen anclada por sus DOS esquinas a la grilla, así sigue a las columnas y filas en cualquier
    programa (con tamaño absoluto, un Excel con otra métrica de columnas la corre o la solapa)."""
    n_cols, n_rows = len(xs) - 1, len(ys) - 1
    if not img.data or n_rows <= 0 or n_cols <= 0:
        return
    try:
        xl = XlImage(BytesIO(img.data))
    except Exception:  # imagen ilegible: se omite sin romper la hoja
        return
    # La grilla no empieza en el borde de la página: el origen sale de la pista y el desplazamiento que
    # calculó el extractor para la esquina superior izquierda; la esquina opuesta se mide desde ahí.
    c0 = min(max(img.col, 0), n_cols - 1)
    r0 = min(max(img.row, 0), n_rows - 1)
    dx0 = min(max(img.dx_pt, 0.0), xs[c0 + 1] - xs[c0])
    dy0 = min(max(img.dy_pt, 0.0), ys[r0 + 1] - ys[r0])
    width = img.width_pt or (img.bbox[2] - img.bbox[0])
    height = img.height_pt or (img.bbox[3] - img.bbox[1])
    if width <= 0 or height <= 0:
        return
    c1, dx1 = _track_at(xs, xs[c0] + dx0 + width)
    r1, dy1 = _track_at(ys, ys[r0] + dy0 + height)
    start = AnchorMarker(col=c0, colOff=int(dx0 * EMU_PER_PT), row=r0, rowOff=int(dy0 * EMU_PER_PT))
    end = AnchorMarker(col=c1, colOff=int(dx1 * EMU_PER_PT), row=r1, rowOff=int(dy1 * EMU_PER_PT))
    xl.anchor = TwoCellAnchor(editAs="oneCell", _from=start, to=end)
    xl.width = width / 0.75  # el tamaño real lo dan las dos esquinas
    xl.height = height / 0.75
    ws.add_image(xl)


def _write_page(ws: Worksheet, page: LayoutPage, styles: _StyleCache) -> None:
    n_cols = min(len(page.col_widths_pt), MAX_COLS)
    n_rows = min(len(page.row_heights_pt), MAX_ROWS)
    col_pts = [max(0.0, w) for w in page.col_widths_pt[:n_cols]]
    row_pts = balanced_row_heights(page, n_rows, n_cols, pixel_snap=True)
    for i, w in enumerate(col_pts):
        ws.column_dimensions[get_column_letter(i + 1)].width = col_width_chars(w)
    for i, h in enumerate(row_pts):
        ws.row_dimensions[i + 1].height = row_height_pt(h)

    _write_cells(ws, page, n_rows, n_cols, styles, row_pts)
    xs, ys = _prefix_sums(col_pts), _prefix_sums(row_pts)
    for img in page.images:
        _add_image(ws, img, xs, ys)

    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if page.landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.print_options.horizontalCentered = True
    if n_cols and n_rows:
        ws.print_area = f"A1:{get_column_letter(n_cols)}{n_rows}"


def write_layout_sheets(
    wb: Workbook,
    report: ConversionReport,
    opts: Options,
    used_names: set[str],
    name_prefix: str = "",
    on_page: Callable[[int, int], None] | None = None,
) -> list[str]:
    """Agrega al libro una hoja por `LayoutPage` del reporte y devuelve los nombres creados.

    `opts.layout_images=False` omite las imágenes. `used_names` se actualiza con los nombres nuevos.
    """
    names: list[str] = []
    styles = _StyleCache()
    total = len(report.layout_pages)
    for i, page in enumerate(report.layout_pages, start=1):
        name = _sheet_title(name_prefix, page.page + 1, used_names)
        ws = wb.create_sheet(name)
        names.append(name)
        if not opts.layout_images and page.images:
            page = LayoutPage(
                page=page.page,
                width_pt=page.width_pt,
                height_pt=page.height_pt,
                col_widths_pt=page.col_widths_pt,
                row_heights_pt=page.row_heights_pt,
                cells=page.cells,
                images=[],
                hidden_texts=page.hidden_texts,
                n_words=page.n_words,
                warnings=page.warnings,
            )
        _write_page(ws, page, styles)
        if on_page is not None:
            on_page(i, total)
    return names


__all__ = ["cell_font_scale", "col_width_chars", "row_height_pt", "write_layout_sheets"]
