"""Extractor del modo "Igual al PDF": una página PDF → ``LayoutPage`` (LAYOUT_MODE.md §3).

Pasos: texto visible (sin lo tapado) → bordes, cajas y rellenos → imágenes → grilla (líneas de la página
agrupadas) → celdas (una por caja, con su texto en renglones y estilos; texto suelto en párrafos) → bordes y
rellenos por celda → valores tipados. Funciones puras, sin Qt; salida determinista.
"""

from __future__ import annotations

import itertools
import statistics
from bisect import bisect_left
from collections import Counter
from dataclasses import dataclass, replace
from typing import Any

import pymupdf

from pdf2excel.core.layout_geom import (
    Box,
    Mark,
    Rect4,
    Segment,
    cluster_values,
    extract_primitives,
    find_closed_boxes,
    find_open_boxes,
    merge_segments,
    nearest_index,
)
from pdf2excel.core.layout_images import extract_images
from pdf2excel.core.layout_raster import (
    apply_scan_styles,
    raster_primitives,
    render_page,
    scan_width_ratio,
)
from pdf2excel.core.layout_text import (
    OCR_DOUBT_WARNING,
    Block,
    Word,
    build_blocks,
    doubtful_ocr_tokens,
    group_lines,
    line_runs,
    spans_to_words,
    split_segments,
)
from pdf2excel.core.layout_typing import apply_typed_values
from pdf2excel.core.visibility import classify_spans, looks_scanned, normalize_text, span_text
from pdf2excel.models import BorderSide, LayoutCell, LayoutImage, LayoutPage, TextRun

MAX_COLS = 80
MAX_ROWS = 250
BASE_TOL = 2.0
SCAN_TOL = 4.0  # tolerancia de agrupamiento de bordes en un escaneo con OCR
TOL_STEPS = (2.0, 3.0, 4.0, 6.0, 8.0, 12.0, 16.0, 24.0, 32.0)
MAX_ROW_HEIGHT = 400.0
OVERLAP_MAX = 0.15  # una caja que se solapa más que esto con celdas ya aceptadas es un marco, no una celda
FILL_COVERAGE = 0.8
BORDER_COVERAGE = 0.6
BAND = 32.0
MIN_TEXT_SLACK = 8.0  # ancho extra (pt) que se le da a un texto suelto de un renglón si hay pistas libres
TEXT_PAD = 1.5  # relleno interno de una celda de Excel: el texto suelto se encuadra con este margen
FRAME_COLUMN_GAP = 2.5  # hueco (× tamaño) dentro de un renglón que delata dos columnas en una caja
FRAME_ROW_GAP = 3.5  # separación de líneas base (× tamaño) que delata bloques separados en una caja alta


class _BandIndex:
    """Índice por franjas horizontales de rectángulos, para consultas de solapamiento y contención."""

    def __init__(self) -> None:
        self._bands: dict[int, list[int]] = {}
        self.rects: list[Rect4] = []

    def add(self, rect: Rect4) -> int:
        idx = len(self.rects)
        self.rects.append(rect)
        for band in range(int(rect[1] // BAND), int(rect[3] // BAND) + 1):
            self._bands.setdefault(band, []).append(idx)
        return idx

    def near(self, rect: Rect4) -> list[int]:
        found: set[int] = set()
        for band in range(int(rect[1] // BAND), int(rect[3] // BAND) + 1):
            found.update(self._bands.get(band, ()))
        return sorted(found)


def _area(r: Rect4) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def _inter(a: Rect4, b: Rect4) -> float:
    return _area((max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])))


def _select_cell_boxes(boxes: list[Box]) -> list[Box]:
    """Cajas que serán celdas: las más chicas primero; las que contienen o pisan a otras son marcos."""
    accepted: list[Box] = []
    index = _BandIndex()
    for box in sorted(boxes, key=lambda b: (b.area, b.y0, b.x0)):
        area = box.area
        if area <= 0:
            continue
        limit = OVERLAP_MAX * area
        overlap = 0.0
        x0, y0, x1, y1 = box.rect
        for i in index.near(box.rect):
            ax0, ay0, ax1, ay1 = index.rects[i]
            w = min(x1, ax1) - max(x0, ax0)
            if w <= 0:
                continue
            h = min(y1, ay1) - max(y0, ay0)
            if h <= 0:
                continue
            overlap += w * h
            if overlap > limit:
                break
        if overlap <= limit:
            accepted.append(box)
            index.add(box.rect)
    return accepted


def _split_box(box: Box, segs: list[Segment]) -> list[Box]:
    """Parte una caja por las líneas que la cruzan de lado a lado (≥ 70 %): cada parte es una celda."""
    w, h = box.x1 - box.x0, box.y1 - box.y0
    for s in segs:
        parts: tuple[Box, Box] | None = None
        if (
            s.orient == "h"
            and box.y0 + 2.0 < s.pos < box.y1 - 2.0
            and min(s.b, box.x1) - max(s.a, box.x0) >= 0.7 * w
        ):
            parts = (
                Box(box.x0, box.y0, box.x1, s.pos, box.fill, box.seq),
                Box(box.x0, s.pos, box.x1, box.y1, box.fill, box.seq),
            )
        elif (
            s.orient == "v"
            and box.x0 + 2.0 < s.pos < box.x1 - 2.0
            and min(s.b, box.y1) - max(s.a, box.y0) >= 0.7 * h
        ):
            parts = (
                Box(box.x0, box.y0, s.pos, box.y1, box.fill, box.seq),
                Box(s.pos, box.y0, box.x1, box.y1, box.fill, box.seq),
            )
        if parts is not None:
            rest = [t for t in segs if t is not s]
            return _split_box(parts[0], rest) + _split_box(parts[1], rest)
    return [box]


def _assign_words(boxes: list[Box], words: list[Word]) -> tuple[dict[int, list[Word]], list[Word]]:
    """Cada palabra va a la caja (la más chica) que contiene su centro; el resto es texto suelto."""
    index = _BandIndex()
    for b in boxes:
        index.add(b.rect)
    inside: dict[int, list[Word]] = {}
    free: list[Word] = []
    for w in words:
        cx, cy = w.center
        best = -1
        best_area = 1e18
        for i in index.near((cx, cy, cx, cy)):
            r = boxes[i].rect
            if r[0] <= cx <= r[2] and r[1] <= cy <= r[3] and boxes[i].area < best_area:
                best, best_area = i, boxes[i].area
        if best >= 0:
            inside.setdefault(best, []).append(w)
        else:
            free.append(w)
    return inside, free


def _steady_scan_sizes(words: list[Word], width_ratio: float = 1.0) -> list[Word]:
    """El OCR da a cada palabra un tamaño según su caja (y suelta restos de 1 pt): se descartan los restos y
    cada renglón toma UN tamaño, el que hace que el texto con Helvetica mida lo mismo que la palabra leída
    (mediana de las palabras del renglón), así el texto no sale con letras de varios tamaños mezclados ni
    más chico que el escaneo. ``width_ratio`` corrige que la caja del OCR sea más angosta que la tinta."""
    helv = pymupdf.Font("helv")
    kept = [w for w in words if w.size >= MIN_SCAN_SIZE or any(ch.isalnum() for ch in w.text)]
    page_fits: list[float] = []
    plan: list[tuple[list[Word], float | None]] = []
    for line in group_lines(kept):
        ref = (
            statistics.median(w.size for w in line if len(w.text) >= 2)
            if any(len(w.text) >= 2 for w in line)
            else 6.0
        )
        groups: list[list[Word]] = [[line[0]]]
        for w in line[1:]:
            if w.x0 - groups[-1][-1].x1 > SEGMENT_GAP_SCAN * ref:
                groups.append([w])
            else:
                groups[-1].append(w)
        for group in groups:
            fits = [
                (w.x1 - w.x0) / unit
                for w in group
                if len(w.text) >= 3
                and not (
                    len(w.text) <= 4 and w.text.isdigit()
                )  # "120" no da un ancho confiable (dígitos sueltos)
                and (unit := float(helv.text_length(w.text, fontsize=1.0))) > 0
            ]
            raw = statistics.median(fits) * width_ratio if fits else None
            if raw is not None:
                page_fits.append(raw)
            plan.append((group, raw))
    body = statistics.median(page_fits) if page_fits else 7.0
    for group, raw in plan:
        # un tramo sin palabras de 3 letras ("4", "UN") toma el tamaño típico de la página, no el ruido
        size = raw if raw is not None else body
        med = min(40.0, max(MIN_SCAN_SIZE + 1.0, round(size * 2) / 2.0))
        for w in group:
            w.size = med
            for ch in w.chars:
                ch.style = replace(ch.style, size=med)
    _snap_scan_sizes(kept)
    return kept


_BOLD_FONT = pymupdf.Font("hebo")
_REGULAR_FONT = pymupdf.Font("helv")


def _bold_size_factor(text: str) -> float:
    """Tamaño de una palabra en negrita medido con las métricas de la regular (lo que se usa para estimar el
    tamaño de todo el OCR) respecto del real: la negrita es más ancha, así que sale inflada."""
    regular = float(_REGULAR_FONT.text_length(text, fontsize=1.0))
    bold = float(_BOLD_FONT.text_length(text, fontsize=1.0))
    return regular / bold if regular > 0 and bold > 0 else 1.0


def _scan_styles_and_sizes(words: list[Word], rgb: Any, px_per_pt: float) -> list[Word]:
    """Negrita/color desde los píxeles y, con la negrita decidida, tamaño corregido por sus métricas. Dos
    pasadas: el grosor esperado del trazo depende del tamaño, y el de una palabra en negrita sale inflado
    hasta saber que lo es."""
    base = {id(w): w.size for w in words}
    for _ in range(2):
        apply_scan_styles(words, rgb, px_per_pt)
        for w in words:
            size = base[id(w)]
            if w.chars and w.chars[0].style.bold:
                size = round(size * _bold_size_factor(w.text) * 2) / 2.0
            if size != w.size:
                w.size = size
                for i, ch in enumerate(w.chars):
                    w.chars[i] = replace(ch, style=replace(ch.style, size=size))
    _snap_scan_sizes(words)
    return words


SEGMENT_GAP_SCAN = 1.5  # hueco (× tamaño) que separa dos tramos de un renglón del OCR
SNAP_RATIO = 0.12  # dos renglones cuyo tamaño estimado difiere menos que esto son del mismo estilo


def _snap_scan_sizes(words: list[Word]) -> None:
    """El tamaño estimado de cada renglón del OCR tiene ruido de ±0.5 pt: renglones del mismo estilo
    (el cuerpo de una tabla, las etiquetas) terminaban con 6, 6.5 y 7 pt. Cada tamaño se une al más
    frecuente (por palabras) que tiene a menos de ``SNAP_RATIO``; los que no tienen pico cerca no cambian."""
    counts: Counter[float] = Counter(w.size for w in words)
    peaks = [s for s, _ in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))]
    snapped: dict[float, float] = {}
    chosen: list[float] = []
    for size in peaks:
        near = [p for p in chosen if abs(p - size) <= SNAP_RATIO * p]
        if near:
            snapped[size] = max(near, key=lambda p: counts[p])
        else:
            chosen.append(size)
            snapped[size] = size
    for w in words:
        target = snapped[w.size]
        if target != w.size:
            w.size = target
            for ch in w.chars:
                ch.style = replace(ch.style, size=target)


def _underline_cuts(line: list[Word], segments: list[Segment], index: _BandIndex) -> list[float]:
    """Posiciones x donde empieza una línea de subrayado justo debajo del renglón y a la derecha de su
    primera palabra: lo que sigue es el valor de otro campo, no la continuación de la etiqueta."""
    x0, x1 = min(w.x0 for w in line), max(w.x1 for w in line)
    y1 = max(w.y1 for w in line)
    size = max(w.size for w in line)
    cuts: list[float] = []
    for i in index.near((x0, y1 - 2.0, x1, y1 + 5.0)):
        seg = segments[i]
        if (
            seg.orient == "h"
            and y1 - 0.4 * size <= seg.pos <= y1 + 0.7 * size
            and x0 + 2.0 < seg.a < x1
            and seg.b - seg.a >= 2.0 * size
        ):
            cuts.append(seg.a - 2.0)
    return cuts


ALIGN_MIN_LINES = 3  # renglones que comparten un borde para considerarlo una columna alineada


def _alignment_columns(lines: list[list[Word]]) -> tuple[Counter[int], Counter[int]]:
    """Bordes derechos e izquierdos (a 1 pt) que se repiten en varios renglones: columnas de un listado."""
    right: Counter[int] = Counter()
    left: Counter[int] = Counter()
    for line in lines:
        rights = {round(w.x1) for w in line}
        lefts = {round(w.x0) for w in line}
        right.update(rights)
        left.update(lefts)
    return right, left


def _on_column(cols: Counter[int], x: float) -> bool:
    k = round(x)
    return max(cols[k - 1], cols[k], cols[k + 1]) >= ALIGN_MIN_LINES


def _column_cuts(line: list[Word], right_cols: Counter[int], left_cols: Counter[int]) -> list[float]:
    """Dos palabras seguidas de un renglón donde la primera termina en una columna alineada a la derecha (o
    la segunda empieza en una alineada a la izquierda) y hay un hueco visible son campos distintos: costo y
    valor pegados no deben quedar en una sola celda."""
    cuts: list[float] = []
    for left, right in itertools.pairwise(line):
        gap = right.x0 - left.x1
        size = max(left.size, right.size)
        if gap < 0.5 * size:
            continue
        if right.spaced and gap <= size:
            continue  # espacio de verdad del texto (título con letras espaciadas): es la misma frase
        if _on_column(right_cols, left.x1) or _on_column(left_cols, right.x0):
            cuts.append(right.x0 - 0.5)
    return cuts


def _is_frame(box: Box, words: list[Word]) -> bool:
    """Una caja con texto en varias columnas o en bloques lejanos entre sí no es una sola celda."""
    if not words:
        return False
    lines = group_lines(words)
    for line in lines:
        for left, right in itertools.pairwise(line):
            if right.x0 - left.x1 > FRAME_COLUMN_GAP * max(left.size, right.size):
                return True
    if len(lines) > 1 and (box.y1 - box.y0) > 40:
        for above, below in itertools.pairwise(lines):
            if below[0].oy - above[0].oy > FRAME_ROW_GAP * max(below[0].size, above[0].size):
                return True
    return False


def _peer_alignment(blk: Block, blocks: list[Block]) -> str:
    """Texto de una sola línea alineado a la derecha o al centro con otros de su columna."""
    if len(blk.segs) != 1:
        return "left"
    tol = 1.5
    same_left = same_right = same_center = 0
    for o in blocks:
        if o is blk:
            continue
        same_left += abs(o.x0 - blk.x0) <= tol
        same_right += abs(o.x1 - blk.x1) <= tol
        same_center += abs((o.x0 + o.x1) - (blk.x0 + blk.x1)) <= 2 * tol
    if same_right >= 2 and same_right > same_left:
        return "right"
    if same_left >= 1:
        return "left"
    if same_center >= 2:
        return "center"
    return "left"


def _halign(x0: float, x1: float, cx0: float, cx1: float) -> str:
    w = cx1 - cx0
    if w <= 0:
        return "left"
    gl, gr = x0 - cx0, cx1 - x1
    if gl + gr <= 0.1 * w or gl <= max(2.0, 0.04 * w):
        return "left"
    if abs(gl - gr) <= 0.15 * w:
        return "center"
    return "right" if gr < gl else "left"


def _valign(y0: float, y1: float, cy0: float, cy1: float, n_lines: int) -> str:
    h = cy1 - cy0
    gt, gb = y0 - cy0, cy1 - y1
    if h > 2.0 * (y1 - y0) or (n_lines > 1 and h > 1.4 * (y1 - y0)):
        if gt < 0.4 * gb:
            return "top"
        if gb < 0.4 * gt:
            return "bottom"
    return "center"


def _build_lines(words: list[Word]) -> list[list[TextRun]]:
    return [runs for runs in (line_runs(line) for line in group_lines(words)) if runs]


def _color_to_border(seg: Segment) -> BorderSide:
    return BorderSide(round(seg.width, 2), seg.color)


def _grow_cell(cell: LayoutCell, occ: list[list[int]], xs: list[float], min_width: float, grow: str) -> None:
    """Amplía un texto suelto que quedó justo al ancho de su celda sobre pistas libres (del lado opuesto a su
    alineación): en Excel un número que no entra se ve como "###" y un texto se achicaría."""
    idx = occ[cell.row][cell.col]
    rows = range(cell.row, cell.row + cell.row_span)
    c0, c1 = cell.col, cell.col + cell.col_span
    nc = len(xs) - 1
    while xs[c1] - xs[c0] < min_width:
        if grow in ("left", "both") and c0 > 0 and all(occ[r][c0 - 1] < 0 for r in rows):
            c0 -= 1
            for r in rows:
                occ[r][c0] = idx
        elif grow in ("right", "both") and c1 < nc and all(occ[r][c1] < 0 for r in rows):
            for r in rows:
                occ[r][c1] = idx
            c1 += 1
        else:
            break
    cell.col, cell.col_span = c0, c1 - c0


def _fit(occ: list[list[int]], r0: int, r1: int, c0: int, c1: int) -> tuple[int, int, int, int] | None:
    """Mayor rectángulo libre anclado en la primera unidad libre del rango pedido."""
    anchor = None
    for r in range(r0, r1):
        for c in range(c0, c1):
            if occ[r][c] < 0:
                anchor = (r, c)
                break
        if anchor:
            break
    if anchor is None:
        return None
    ar, ac = anchor
    ec = ac
    while ec < c1 and occ[ar][ec] < 0:
        ec += 1
    er = ar + 1
    while er < r1 and all(occ[er][c] < 0 for c in range(ac, ec)):
        er += 1
    return ar, er, ac, ec


@dataclass
class _Edge:
    width: float
    color: str


def _cluster_with_limit(raw: list[float], limit: int, start_tol: float = BASE_TOL) -> list[float]:
    lines: list[float] = []
    for tol in (t for t in TOL_STEPS if t >= start_tol):
        lines = cluster_values(raw, tol)
        if len(lines) - 1 <= limit:
            return lines
    while len(lines) - 1 > limit and len(lines) > 2:  # último recurso: se funden las pistas más finas
        gaps = [(lines[i + 1] - lines[i], i) for i in range(1, len(lines) - 2)] or [(0.0, 0)]
        _gap, i = min(gaps)
        del lines[i + 1 if i + 1 < len(lines) - 1 else i]
    return lines


def _split_tall(lines: list[float]) -> list[float]:
    """Parte en pistas iguales las filas más altas que lo que admite Excel (≈ 409 pt)."""
    out = [lines[0]]
    for nxt in lines[1:]:
        prev = out[-1]
        parts = int((nxt - prev) // MAX_ROW_HEIGHT) + 1
        for k in range(1, parts):
            out.append(prev + (nxt - prev) * k / parts)
        out.append(nxt)
    return out


def _span_indices(lines: list[float], lo: float, hi: float) -> tuple[int, int]:
    n = len(lines) - 1
    a = nearest_index(lines, lo)
    b = nearest_index(lines, hi)
    if b <= a:
        b = a + 1
    if b > n:
        b = n
        a = max(0, b - 1)
    return a, b


def extract_layout_page(
    page: Any, *, embed_images: bool = True, scanned: bool = False, typed: bool = True
) -> LayoutPage:
    """Una página → ``LayoutPage`` (grilla + celdas + imágenes). ``scanned=True``: la página es un escaneo con
    capa OCR (texto invisible = texto real; la imagen de fondo NO se incluye). ``typed=False`` deja todo como
    texto."""
    matrix = page.rotation_matrix
    page_rect = page.rect
    width, height = float(page_rect.width), float(page_rect.height)
    trace = page.get_texttrace()
    drawings = page.get_drawings()
    image_info = page.get_image_info(xrefs=True)
    is_scanned = scanned or looks_scanned(page, trace, image_info)
    hidden_flags = classify_spans(page, trace, drawings, image_info, is_scanned)
    hidden_texts = [
        normalize_text(span_text(t))
        for t, h in zip(trace, hidden_flags, strict=True)
        if h and span_text(t).strip()
    ]
    visible = [t for t, h in zip(trace, hidden_flags, strict=True) if not h and span_text(t).strip()]
    words = spans_to_words(visible, matrix)
    rgb = None
    px_per_pt = 1.0
    if is_scanned and words:
        rgb, px_per_pt = render_page(page)
        words = _steady_scan_sizes(words, scan_width_ratio(words, rgb, px_per_pt))
    prims = extract_primitives(drawings, matrix, (0.0, 0.0, width, height))
    if rgb is not None:
        # un escaneo no trae vectores: cajas, líneas y rellenos salen de los píxeles (y la negrita, el color)
        found = raster_primitives(rgb, px_per_pt, (width, height))
        prims.segments = merge_segments([*prims.segments, *found.segments])
        prims.fills.extend(found.fills)
        prims.boxes.extend(found.boxes)
        words = _scan_styles_and_sizes(words, rgb, px_per_pt)
    images = extract_images(page, image_info, matrix, (width, height), is_scanned) if embed_images else []
    result = LayoutPage(
        page=int(page.number or 0),
        width_pt=width,
        height_pt=height,
        col_widths_pt=[width],
        row_heights_pt=[height],
        hidden_texts=hidden_texts,
        n_words=len(words),
    )
    if not words and not prims.segments and not prims.boxes and not images:
        return result

    # ---------------------------------------------------------------- cajas y texto
    raw_boxes = list(prims.boxes)
    raw_boxes.extend(Box(*r) for r in find_closed_boxes(prims.segments))
    raw_boxes.extend(Box(*r) for r in find_open_boxes(prims.segments))
    cell_boxes = _select_cell_boxes(raw_boxes)
    seg_rects: list[Rect4] = [
        (s.pos, s.a, s.pos, s.b) if s.orient == "v" else (s.a, s.pos, s.b, s.pos) for s in prims.segments
    ]
    seg_index = _BandIndex()
    for r in seg_rects:
        seg_index.add(r)
    split: list[Box] = []
    for box in cell_boxes:
        near = [prims.segments[i] for i in seg_index.near(box.rect)]
        split.extend(_split_box(box, near) if near else [box])
    cell_boxes = split
    inside, free_words = _assign_words(cell_boxes, words)
    final_boxes: list[Box] = []
    box_words: list[list[Word]] = []
    for i, box in enumerate(cell_boxes):
        bw = inside.get(i, [])
        if _is_frame(box, bw):
            free_words.extend(bw)
        else:
            final_boxes.append(box)
            box_words.append(bw)
    free_lines = group_lines(free_words)
    right_cols, left_cols = _alignment_columns(free_lines)
    segs = [
        s
        for line in free_lines
        for s in split_segments(
            line,
            cuts=_underline_cuts(line, prims.segments, seg_index) + _column_cuts(line, right_cols, left_cols),
        )
    ]
    blocks = build_blocks(segs)

    # ---------------------------------------------------------------- grilla
    xs_raw: list[float] = []
    ys_raw: list[float] = []
    for b in final_boxes:
        xs_raw.extend((b.x0, b.x1))
        ys_raw.extend((b.y0, b.y1))
    for f in prims.fills:  # los bordes de toda barra/caja rellena también son líneas de la grilla
        xs_raw.extend((f.x0, f.x1))
        ys_raw.extend((f.y0, f.y1))
    for s in prims.segments:
        if s.orient == "v":
            xs_raw.append(s.pos)
            ys_raw.extend((s.a, s.b))
        else:
            ys_raw.append(s.pos)
            xs_raw.extend((s.a, s.b))
    hard = cluster_values(xs_raw, BASE_TOL)
    block_align: list[str] = []
    block_right: list[float] = []
    for blk in blocks:
        xs_raw.append(blk.x0 - TEXT_PAD)
        ys_raw.extend((blk.y0, blk.y1))
        align = blk.mode if blk.mode in ("right", "center") else _peer_alignment(blk, blocks)
        right = blk.x1 + TEXT_PAD
        if align == "left":
            k = bisect_left(hard, blk.x1 - BASE_TOL)
            size = max(s.size for s in blk.segs)
            if k < len(hard) and hard[k] - blk.x1 <= max(0.6 * (blk.x1 - blk.x0), 6.0 * size):
                right = hard[k]
            else:
                right = blk.x1 + max(2.0, 0.08 * (blk.x1 - blk.x0))
        xs_raw.append(right)
        block_align.append(align)
        block_right.append(right)
    for img in images:
        x0, y0, x1, y1 = img.bbox
        xs_raw.extend((x0, x1))
        ys_raw.extend((y0, y1))
    # la grilla cubre al menos todo el contenido; las pistas se miden desde la primera línea
    # el texto de un escaneo (OCR) viene con ruido de varios puntos: se agrupa con más tolerancia
    start_tol = SCAN_TOL if is_scanned else BASE_TOL
    xs = _cluster_with_limit(xs_raw, MAX_COLS, start_tol)
    ys = _cluster_with_limit(ys_raw, MAX_ROWS)
    if len(xs) < 2:
        xs = [xs[0], xs[0] + 1.0] if xs else [0.0, width]
    if len(ys) < 2:
        ys = [ys[0], ys[0] + 1.0] if ys else [0.0, height]
    ys = _split_tall(ys)
    if len(xs) - 1 >= MAX_COLS or len(ys) - 1 >= MAX_ROWS:
        result.warnings.append("La página tiene muchos elementos: se simplificó la grilla.")
    nc, nr = len(xs) - 1, len(ys) - 1

    # ---------------------------------------------------------------- celdas
    occ = [[-1] * nc for _ in range(nr)]
    cells: list[LayoutCell] = []

    def place(rect: Rect4, lines: list[list[TextRun]], right_x: float | None = None) -> LayoutCell | None:
        r0, r1 = _span_indices(ys, rect[1], rect[3])
        c0, c1 = _span_indices(xs, rect[0], rect[2] if right_x is None else right_x)
        fit = _fit(occ, r0, r1, c0, c1)
        if fit is None:
            return None
        fr0, fr1, fc0, fc1 = fit
        cell = LayoutCell(
            row=fr0,
            col=fc0,
            row_span=fr1 - fr0,
            col_span=fc1 - fc0,
            lines=lines,
            bbox=(xs[fc0], ys[fr0], xs[fc1], ys[fr1]),
        )
        for r in range(fr0, fr1):
            for c in range(fc0, fc1):
                occ[r][c] = len(cells)
        cells.append(cell)
        return cell

    skipped_boxes = skipped_texts = 0
    order = sorted(range(len(final_boxes)), key=lambda i: (final_boxes[i].y0, final_boxes[i].x0))
    for i in order:
        box, bw = final_boxes[i], box_words[i]
        lines = _build_lines(bw)
        cell = place(box.rect, lines)
        if cell is None:
            skipped_boxes += 1
            continue
        if bw:
            tx0, tx1 = min(w.x0 for w in bw), max(w.x1 for w in bw)
            ty0, ty1 = min(w.y0 for w in bw), max(w.y1 for w in bw)
            cell.halign = _halign(tx0, tx1, box.x0, box.x1)  # type: ignore[assignment]
            cell.valign = _valign(ty0, ty1, box.y0, box.y1, len(lines))  # type: ignore[assignment]
        cell.wrap = len(lines) > 1
    to_grow: list[tuple[LayoutCell, float, str]] = []
    order_b = sorted(range(len(blocks)), key=lambda i: (blocks[i].y0, blocks[i].x0))
    for i in order_b:
        blk = blocks[i]
        lines = _build_lines(blk.words)
        if not lines:
            continue
        align = block_align[i]
        cell = place(blk.rect, lines, block_right[i])
        if cell is None:
            skipped_texts += 1
            continue
        if len(lines) == 1:
            to_grow.append((cell, (blk.x1 - blk.x0) + MIN_TEXT_SLACK, align))
        cell.halign = align  # type: ignore[assignment]
        cell.valign = "top" if len(lines) > 1 else "center"
        cell.wrap = len(lines) > 1

    for cell, min_width, align in to_grow:
        _grow_cell(
            cell, occ, xs, min_width, "left" if align == "right" else "both" if align == "center" else "right"
        )
        cell.bbox = (xs[cell.col], ys[cell.row], xs[cell.col + cell.col_span], ys[cell.row + cell.row_span])

    if skipped_boxes:
        result.warnings.append(f"{skipped_boxes} cajas del PDF se superponían con otras y se omitieron.")
    if skipped_texts:
        result.warnings.append(f"{skipped_texts} textos del PDF no entraron en la grilla y se omitieron.")

    _place_marks(cells, prims.marks)
    if is_scanned and any(doubtful_ocr_tokens(c.text) for c in cells):
        result.warnings.append(OCR_DOUBT_WARNING)

    # ---------------------------------------------------------------- rellenos y bordes
    _apply_fills(cells, prims.fills, prims.whites)
    h_edges, v_edges = _edges_from_segments(prims.segments, xs, ys)
    col_w = [xs[k + 1] - xs[k] for k in range(nc)]
    row_h = [ys[k + 1] - ys[k] for k in range(nr)]
    _apply_borders(cells, h_edges, v_edges, col_w, row_h)
    orphans = _orphan_cells(h_edges, v_edges, occ, row_h, col_w, len(cells))
    for orphan in orphans:
        orphan.bbox = (
            xs[orphan.col],
            ys[orphan.row],
            xs[orphan.col + orphan.col_span],
            ys[orphan.row + orphan.row_span],
        )
    _apply_fills(orphans, prims.fills, prims.whites)
    cells.extend(orphans)
    for filler in _fill_only_cells(prims.fills, xs, ys, occ, len(cells)):
        filler.bbox = (
            xs[filler.col],
            ys[filler.row],
            xs[filler.col + filler.col_span],
            ys[filler.row + filler.row_span],
        )
        cells.append(filler)

    # ---------------------------------------------------------------- imágenes
    placed: list[LayoutImage] = []
    for img in images:
        img.row = nearest_track(ys, img.bbox[1])
        img.col = nearest_track(xs, img.bbox[0])
        img.dx_pt = max(0.0, img.bbox[0] - xs[img.col])
        img.dy_pt = max(0.0, img.bbox[1] - ys[img.row])
        placed.append(img)

    if typed:
        apply_typed_values(cells)
    cells.sort(key=lambda c: (c.row, c.col))
    result.col_widths_pt = col_w
    result.row_heights_pt = row_h
    result.cells = cells
    result.images = placed
    return result


def _apply_fills(cells: list[LayoutCell], colored: list[Box], whites: list[Box] | None = None) -> None:
    """Relleno de cada celda: el último relleno (por orden de dibujo) que la cubre al menos el 80 %.
    Un relleno blanco posterior a un color lo anula (campo blanco sobre un panel gris)."""
    fills = [*colored, *(whites or [])]
    index = _BandIndex()
    for f in fills:
        index.add(f.rect)
    by_seq = sorted(range(len(fills)), key=lambda k: fills[k].seq)
    for cell in cells:
        area = _area(cell.bbox)
        if area <= 0:
            continue
        near = set(index.near(cell.bbox))
        for k in by_seq:
            if k in near and _inter(cell.bbox, fills[k].rect) >= FILL_COVERAGE * area:
                cell.fill = fills[k].fill


MAX_FILL_UNITS = 40000
MIN_SCAN_SIZE = 3.0  # pt: palabras OCR más chicas que esto (y sin letras ni números) son ruido


def _place_marks(cells: list[LayoutCell], marks: list[Mark]) -> None:
    """Los tildes y cruces vectoriales de las casillas pasan a ser un carácter (✔ / ✗) en la celda que los
    contiene, así un formulario completado no parece vacío."""
    for mark in marks:
        cx, cy = (mark.rect[0] + mark.rect[2]) / 2.0, (mark.rect[1] + mark.rect[3]) / 2.0
        best: LayoutCell | None = None
        for cell in cells:
            x0, y0, x1, y1 = cell.bbox
            inside = x0 <= cx <= x1 and y0 <= cy <= y1 and not cell.lines
            if inside and (best is None or _area(cell.bbox) < _area(best.bbox)):
                best = cell
        if best is None:
            continue
        x0, y0, x1, y1 = best.bbox
        size = max(6.0, round(min(x1 - x0, y1 - y0) * 0.9 * 2) / 2)
        best.lines = [
            [TextRun(mark.glyph, True, False, size, None if mark.color == "000000" else mark.color)]
        ]
        best.halign = "center"
        best.valign = "center"
        best.wrap = False


def _fill_only_cells(
    fills: list[Box], xs: list[float], ys: list[float], occ: list[list[int]], base: int
) -> list[LayoutCell]:
    """Celdas vacías que solo llevan color: la parte de una barra o caja rellena donde no hay texto ni borde
    (p. ej. los huecos de la franja gris de un encabezado entre una columna y la otra). El relleno se reparte
    por cobertura del rectángulo, no por la presencia de texto."""
    nr, nc = len(ys) - 1, len(xs) - 1
    colors: dict[tuple[int, int], tuple[int, str]] = {}
    budget = MAX_FILL_UNITS
    for f in fills:
        c0 = max(0, bisect_left(xs, f.x0 - 0.5, 0, nc) - 1)
        r0 = max(0, bisect_left(ys, f.y0 - 0.5, 0, nr) - 1)
        for r in range(r0, nr):
            if ys[r] >= f.y1 - 0.5 or budget <= 0:
                break
            for c in range(c0, nc):
                if xs[c] >= f.x1 - 0.5:
                    break
                if occ[r][c] >= 0:
                    continue
                unit = (xs[c], ys[r], xs[c + 1], ys[r + 1])
                area = _area(unit)
                if area > 0 and _inter(unit, f.rect) >= FILL_COVERAGE * area:
                    budget -= 1
                    prev = colors.get((r, c))
                    if prev is None or f.seq >= prev[0]:
                        colors[(r, c)] = (f.seq, f.fill or "")
    out: list[LayoutCell] = []
    used: set[tuple[int, int]] = set()
    for r, c in sorted(colors):
        if (r, c) in used or not colors[(r, c)][1]:
            continue
        color = colors[(r, c)][1]
        span = 1
        while (r, c + span) in colors and colors[(r, c + span)][1] == color and (r, c + span) not in used:
            span += 1
        for k in range(span):
            used.add((r, c + k))
            occ[r][c + k] = base + len(out)
        out.append(LayoutCell(row=r, col=c, row_span=1, col_span=span, fill=color, wrap=False))
    return out


def _edges_from_segments(
    segments: list[Segment], xs: list[float], ys: list[float]
) -> tuple[dict[tuple[int, int], _Edge], dict[tuple[int, int], _Edge]]:
    """Segmentos de borde → unidades de la grilla: ``h[(línea_y, columna)]`` y ``v[(línea_x, fila)]``."""
    nc, nr = len(xs) - 1, len(ys) - 1
    h_edges: dict[tuple[int, int], _Edge] = {}
    v_edges: dict[tuple[int, int], _Edge] = {}
    for s in segments:
        if s.orient == "h":
            line = nearest_index(ys, s.pos)
            for c in range(nearest_index(xs, s.a), nearest_index(xs, s.b)):
                if 0 <= c < nc:
                    _put_edge(h_edges, (line, c), s)
        else:
            line = nearest_index(xs, s.pos)
            for r in range(nearest_index(ys, s.a), nearest_index(ys, s.b)):
                if 0 <= r < nr:
                    _put_edge(v_edges, (line, r), s)
    return h_edges, v_edges


def _apply_borders(
    cells: list[LayoutCell],
    h_edges: dict[tuple[int, int], _Edge],
    v_edges: dict[tuple[int, int], _Edge],
    col_w: list[float],
    row_h: list[float],
) -> None:
    for cell in cells:
        r, c, rs, cs = cell.row, cell.col, cell.row_span, cell.col_span
        widths = col_w[c : c + cs]
        heights = row_h[r : r + rs]
        cell.border_top = _side(h_edges, [(r, k) for k in range(c, c + cs)], widths)
        cell.border_bottom = _side(h_edges, [(r + rs, k) for k in range(c, c + cs)], widths)
        cell.border_left = _side(v_edges, [(c, k) for k in range(r, r + rs)], heights)
        cell.border_right = _side(v_edges, [(c + cs, k) for k in range(r, r + rs)], heights)


def nearest_track(lines: list[float], value: float) -> int:
    """Pista (0-based) que contiene a ``value``."""
    i = bisect_left(lines, value + 0.01) - 1
    return max(0, min(len(lines) - 2, i))


def _put_edge(edges: dict[tuple[int, int], _Edge], key: tuple[int, int], s: Segment) -> None:
    cur = edges.get(key)
    if cur is None or s.width > cur.width:
        edges[key] = _Edge(s.width, s.color)


def _side(
    edges: dict[tuple[int, int], _Edge], keys: list[tuple[int, int]], lengths: list[float]
) -> BorderSide | None:
    total = sum(lengths)
    if total <= 0:
        return None
    covered = 0.0
    seen: Counter[tuple[float, str]] = Counter()
    for key, length in zip(keys, lengths, strict=True):
        e = edges.get(key)
        if e is not None:
            covered += length
            seen[(round(e.width, 2), e.color)] += 1
    if covered < BORDER_COVERAGE * total or not seen:
        return None
    (width, color), _n = max(seen.items(), key=lambda kv: (kv[1], kv[0][0]))
    return BorderSide(width, color)


def _orphan_cells(
    h_edges: dict[tuple[int, int], _Edge],
    v_edges: dict[tuple[int, int], _Edge],
    occ: list[list[int]],
    row_h: list[float],
    col_w: list[float],
    base: int,
) -> list[LayoutCell]:
    """Celdas vacías que solo llevan borde: líneas sueltas del PDF que no tocan ninguna celda."""
    nr, nc = len(row_h), len(col_w)
    sides: dict[tuple[int, int], dict[str, BorderSide]] = {}

    def add(unit: tuple[int, int], side: str, e: _Edge) -> None:
        sides.setdefault(unit, {})[side] = BorderSide(round(e.width, 2), e.color)

    for (i, c), e in sorted(h_edges.items()):
        above = occ[i - 1][c] if 0 < i <= nr else -1
        below = occ[i][c] if i < nr else -1
        if above >= 0 or below >= 0:
            continue
        # la línea se apoya en la pista libre más baja (la que menos altera la grilla)
        use_below = i < nr and (i == 0 or row_h[i] <= row_h[i - 1])
        if use_below:
            add((i, c), "top", e)
        elif i >= 1:
            add((i - 1, c), "bottom", e)
    for (j, r), e in sorted(v_edges.items()):
        left = occ[r][j - 1] if 0 < j <= nc else -1
        right = occ[r][j] if j < nc else -1
        if left >= 0 or right >= 0:
            continue
        use_right = j < nc and (j == 0 or col_w[j] <= col_w[j - 1])
        if use_right:
            add((r, j), "left", e)
        elif j >= 1:
            add((r, j - 1), "right", e)
    out: list[LayoutCell] = []
    used: set[tuple[int, int]] = set()
    for unit in sorted(sides):
        if unit in used:
            continue
        r, c = unit
        cfg = sides[unit]
        row_span = col_span = 1
        if set(cfg) <= {"top", "bottom"}:
            while (
                (r, c + col_span) in sides
                and sides[(r, c + col_span)] == cfg
                and (r, c + col_span) not in used
            ):
                col_span += 1
        elif set(cfg) <= {"left", "right"}:
            while (
                (r + row_span, c) in sides
                and sides[(r + row_span, c)] == cfg
                and (r + row_span, c) not in used
            ):
                row_span += 1
        for rr in range(r, r + row_span):
            for cc in range(c, c + col_span):
                used.add((rr, cc))
        cell = LayoutCell(row=r, col=c, row_span=row_span, col_span=col_span, wrap=False)
        cell.border_top = cfg.get("top")
        cell.border_bottom = cfg.get("bottom")
        cell.border_left = cfg.get("left")
        cell.border_right = cfg.get("right")
        for rr in range(r, r + row_span):
            for cc in range(c, c + col_span):
                occ[rr][cc] = base + len(out)
        out.append(cell)
    return out


__all__ = ["extract_layout_page"]
