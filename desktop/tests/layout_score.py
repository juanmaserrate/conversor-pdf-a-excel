"""Medidor de FIDELIDAD de un .xlsx contra el PDF del que salió (docs/LAYOUT_MODE.md §7).

Método: se renderiza el .xlsx con LibreOffice (``soffice --headless --convert-to pdf``, con un perfil de
usuario temporal propio) sobre una COPIA donde cada hoja se fuerza a una sola página al imprimir, y se compara
contra el PDF original página por página (la hoja N se mide contra la página N):

* **recall**: fracción de las palabras visibles del PDF que aparecen en el render. "Visible" excluye el texto
  que el PDF dibuja y luego TAPA con un rectángulo blanco (chequeo propio, independiente de la app).
* **palabras sobrantes**: palabras del render que no están (o están de más) en el original, p. ej. los
  "N/A" tapados.
* **error de posición**: cada palabra se empareja por texto; las coordenadas se normalizan al bbox del
  contenido de su página (en ambos lados) y se reporta mediana y p90, en % del ancho/alto del contenido.
* **bordes**: tramos de línea de borde del PDF (rectángulos finos, líneas, cajas) vs. tramos del xlsx.
* **celdas mezcladas**: celdas del xlsx cuyo texto junta palabras que en el PDF están separadas por un borde
  o en cajas distintas del ground truth.
* hojas vs. páginas e imágenes presentes.

Uso: ``python -m tests.layout_score <pdf> <xlsx> [--png out.png] [--json]``.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from collections import Counter
from collections.abc import Sequence
from dataclasses import asdict, dataclass, field
from itertools import pairwise
from pathlib import Path
from typing import Any

import pymupdf
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText
from openpyxl.worksheet.properties import PageSetupProperties
from PIL import Image, ImageDraw

SOFFICE_TIMEOUT_S = 180
SOFFICE_ATTEMPTS = 3
PNG_DPI = 70
GROUND_TRUTH_NAME = "ground_truth_forms.json"

# umbrales iniciales (docs/LAYOUT_MODE.md §7)
MIN_RECALL = 0.98
MAX_EXTRA_WORDS = 0
MAX_POS_ERR_MEDIAN_PCT = 2.5
MIN_BORDER_RATIO = 0.5  # el Excel debe conservar al menos la mitad de los tramos de borde del PDF
MIN_BORDERS_TO_CHECK = 4

_TOKEN_RE = re.compile(r"\w+", re.UNICODE)
_THIN = 2.5  # pt: un relleno más fino que esto es un borde dibujado como rectángulo
_MIN_SEGMENT = 6.0  # pt


class LibreOfficeUnavailable(RuntimeError):
    """No hay ``soffice`` en el sistema: no se puede renderizar el .xlsx."""


class LibreOfficeError(RuntimeError):
    """LibreOffice está instalado pero no pudo convertir el archivo."""


def tokens(text: str) -> list[str]:
    """Palabras normalizadas (minúsculas, solo caracteres alfanuméricos)."""
    return _TOKEN_RE.findall(text.casefold())


# ------------------------------------------------------------------------------------------ resultado


@dataclass
class FidelityResult:
    pdf: str
    xlsx: str
    pdf_pages: int
    sheets: int
    render_pages: int
    words_expected: int
    words_found: int
    recall: float | None
    recall_any_page: float | None
    extra_words: int
    extra_examples: list[str]
    missing_examples: list[str]
    pos_matched: int
    pos_err_median_pct: float | None
    pos_err_p90_pct: float | None
    dx_median_pct: float | None
    dy_median_pct: float | None
    pdf_border_segments: int
    xlsx_border_segments: int
    xlsx_bordered_cells: int
    border_ratio: float | None
    mixed_cells: int
    mixed_examples: list[str]
    pdf_images: int
    xlsx_images: int
    covered_words_in_pdf: int
    per_page: list[dict[str, Any]] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=1)

    def problems(self) -> list[str]:
        """Umbrales iniciales incumplidos (lista vacía = fidelidad aceptable)."""
        out: list[str] = []
        if self.recall is not None and self.recall < MIN_RECALL:
            out.append(f"recall {self.recall:.3f} < {MIN_RECALL}")
        # un PDF sin capa de texto (escaneo) no tiene con qué comparar: no hay "sobrantes" medibles
        if self.words_expected > 0 and self.extra_words > MAX_EXTRA_WORDS:
            out.append(f"{self.extra_words} palabras sobrantes")
        if self.pos_err_median_pct is not None and self.pos_err_median_pct > MAX_POS_ERR_MEDIAN_PCT:
            out.append(
                f"error de posición mediano {self.pos_err_median_pct:.1f} % > {MAX_POS_ERR_MEDIAN_PCT} %"
            )
        checkable = self.border_ratio is not None and self.pdf_border_segments >= MIN_BORDERS_TO_CHECK
        if checkable and self.border_ratio is not None and self.border_ratio < MIN_BORDER_RATIO:
            out.append(
                f"faltan bordes: {self.xlsx_border_segments} tramos en el Excel para "
                f"{self.pdf_border_segments} en el PDF"
            )
        if self.mixed_cells:
            out.append(f"{self.mixed_cells} celdas mezcladas")
        if self.sheets != self.pdf_pages:
            out.append(f"{self.sheets} hojas para {self.pdf_pages} páginas")
        return out

    @property
    def passes(self) -> bool:
        return not self.problems()

    def summary(self) -> str:
        def pct(v: float | None) -> str:
            return "n/d" if v is None else f"{v:.1f} %"

        lines = [
            f"PDF: {self.pdf} ({self.pdf_pages} pág.)   "
            f"Excel: {self.xlsx} ({self.sheets} hojas, render {self.render_pages} pág.)",
            f"recall de palabras      : {'n/d' if self.recall is None else f'{self.recall:.3f}'} "
            f"({self.words_found}/{self.words_expected}; sin importar la página: "
            f"{'n/d' if self.recall_any_page is None else f'{self.recall_any_page:.3f}'})",
            f"palabras sobrantes      : {self.extra_words} {self.extra_examples[:8]}",
            f"error de posición       : mediana {pct(self.pos_err_median_pct)}  "
            f"p90 {pct(self.pos_err_p90_pct)} "
            f"({self.pos_matched} palabras emparejadas)",
            f"bordes (tramos)         : PDF {self.pdf_border_segments} vs Excel {self.xlsx_border_segments} "
            f"({self.xlsx_bordered_cells} celdas con borde)",
            f"celdas mezcladas        : {self.mixed_cells} {self.mixed_examples[:3]}",
            f"imágenes                : PDF {self.pdf_images} vs Excel {self.xlsx_images}",
            f"textos tapados en el PDF: {self.covered_words_in_pdf} palabras",
        ]
        problems = self.problems()
        lines.append(
            "veredicto               : " + ("OK" if not problems else "MAL — " + "; ".join(problems))
        )
        lines.extend(f"nota: {n}" for n in self.notes)
        return "\n".join(lines)


# ------------------------------------------------------------------------------------------ PDF


@dataclass(frozen=True)
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


def _is_white(color: Any) -> bool:
    return bool(color) and all(float(v) >= 0.97 for v in color)


def find_covered_spans(page: pymupdf.Page) -> list[tuple[str, pymupdf.Rect]]:
    """Spans de texto que un relleno BLANCO opaco dibujado DESPUÉS tapa (≥ 80 % de su área).

    Devuelve ``(texto, bbox)`` en el marco sin rotar de la página (el de ``get_texttrace``).
    """
    whites: list[tuple[int, pymupdf.Rect]] = []
    for d in page.get_drawings():
        if (
            "f" in str(d.get("type", ""))
            and _is_white(d.get("fill"))
            and float(d.get("fill_opacity") or 1.0) >= 0.99
        ):
            whites.append((int(d.get("seqno", 0)), pymupdf.Rect(d["rect"])))
    covered: list[tuple[str, pymupdf.Rect]] = []
    for span in page.get_texttrace():
        text = "".join(chr(ch[0]) for ch in span.get("chars", ()))
        bbox = pymupdf.Rect(span["bbox"])
        if not text.strip() or bbox.is_empty:
            continue
        seq = int(span.get("seqno", 0))
        area = bbox.get_area()
        for wseq, rect in whites:
            if wseq <= seq:
                continue
            inter = bbox & rect
            if not inter.is_empty and inter.get_area() >= 0.8 * area:
                covered.append((text, bbox))
                break
    return covered


def visible_words(page: pymupdf.Page, covered: list[tuple[str, pymupdf.Rect]] | None = None) -> list[Word]:
    """Palabras visibles de la página en el marco VISUAL (con ``/Rotate`` aplicado), sin las tapadas."""
    spans = find_covered_spans(page) if covered is None else covered
    rot = page.rotation_matrix
    out: list[Word] = []
    for x0, y0, x1, y1, text, *_ in page.get_text("words"):
        r = pymupdf.Rect(x0, y0, x1, y1)
        center = pymupdf.Point((x0 + x1) / 2, (y0 + y1) / 2)
        if any(center in bbox for _, bbox in spans):
            continue
        vr = r * rot
        vr.normalize()
        if not vr.intersects(page.rect):
            continue
        for tok in tokens(text):
            out.append(Word(tok, vr.x0, vr.y0, vr.x1, vr.y1))
    return out


@dataclass
class BorderRun:
    orient: str  # "h" | "v"
    pos: float
    start: float
    end: float


def _merge_runs(raw: list[BorderRun]) -> list[BorderRun]:
    merged: list[BorderRun] = []
    for orient in ("h", "v"):
        group = sorted((r for r in raw if r.orient == orient), key=lambda r: (round(r.pos / 2.0), r.start))
        cur: BorderRun | None = None
        for r in group:
            if cur and abs(r.pos - cur.pos) <= 2.0 and r.start <= cur.end + 2.0:
                cur.end = max(cur.end, r.end)
            else:
                if cur:
                    merged.append(cur)
                cur = BorderRun(orient, r.pos, r.start, r.end)
        if cur:
            merged.append(cur)
    return merged


def page_border_runs(page: pymupdf.Page) -> list[BorderRun]:
    """Tramos de borde dibujados (marco visual): líneas, bordes de rectángulos, rellenos finos, cajas."""
    rot = page.rotation_matrix
    raw: list[BorderRun] = []

    def add(p0: pymupdf.Point, p1: pymupdf.Point) -> None:
        a, b = p0 * rot, p1 * rot
        if abs(a.y - b.y) <= 1.0 and abs(a.x - b.x) >= _MIN_SEGMENT:
            raw.append(BorderRun("h", (a.y + b.y) / 2, min(a.x, b.x), max(a.x, b.x)))
        elif abs(a.x - b.x) <= 1.0 and abs(a.y - b.y) >= _MIN_SEGMENT:
            raw.append(BorderRun("v", (a.x + b.x) / 2, min(a.y, b.y), max(a.y, b.y)))

    def add_rect_edges(r: pymupdf.Rect) -> None:
        add(pymupdf.Point(r.x0, r.y0), pymupdf.Point(r.x1, r.y0))
        add(pymupdf.Point(r.x0, r.y1), pymupdf.Point(r.x1, r.y1))
        add(pymupdf.Point(r.x0, r.y0), pymupdf.Point(r.x0, r.y1))
        add(pymupdf.Point(r.x1, r.y0), pymupdf.Point(r.x1, r.y1))

    for d in page.get_drawings():
        kind = str(d.get("type", ""))
        rect = pymupdf.Rect(d["rect"])
        if "s" in kind:
            items = d.get("items", [])
            if any(it[0] == "c" for it in items) and rect.width >= 8 and rect.height >= 8:
                add_rect_edges(rect)  # caja con esquinas redondeadas
                continue
            for it in items:
                if it[0] == "l":
                    add(it[1], it[2])
                elif it[0] == "re":
                    add_rect_edges(pymupdf.Rect(it[1]))
                elif it[0] == "qu":
                    add_rect_edges(pymupdf.Quad(it[1]).rect)
        elif "f" in kind and not _is_white(d.get("fill")):
            short, long_ = sorted((rect.width, rect.height))
            if short <= _THIN and long_ >= _MIN_SEGMENT:
                if rect.width >= rect.height:
                    add(
                        pymupdf.Point(rect.x0, rect.y0 + rect.height / 2),
                        pymupdf.Point(rect.x1, rect.y0 + rect.height / 2),
                    )
                else:
                    add(
                        pymupdf.Point(rect.x0 + rect.width / 2, rect.y0),
                        pymupdf.Point(rect.x0 + rect.width / 2, rect.y1),
                    )
    return _merge_runs(raw)


def _raster_border_count(page: pymupdf.Page) -> int:
    """Tramos de borde que se ven en el render de una página sin vectores (escaneo), con el detector del
    extractor: sirve para saber si el Excel perdió las cajas aunque el PDF no tenga dibujos."""
    from pdf2excel.core.layout_raster import raster_primitives, render_page

    rgb, scale = render_page(page)
    return len(raster_primitives(rgb, scale, (float(page.rect.width), float(page.rect.height))).segments)


def _crosses(a: tuple[float, float], b: tuple[float, float], run: BorderRun) -> bool:
    """¿El segmento a-b atraviesa el tramo de borde? (extremos con 1 pt de tolerancia)."""
    (ax, ay), (bx, by) = a, b
    if run.orient == "h":
        if (ay - run.pos) * (by - run.pos) >= 0:
            return False
        x = ax + (run.pos - ay) / (by - ay) * (bx - ax)
        return run.start - 1 <= x <= run.end + 1
    if (ax - run.pos) * (bx - run.pos) >= 0:
        return False
    y = ay + (run.pos - ax) / (bx - ax) * (by - ay)
    return run.start - 1 <= y <= run.end + 1


# ------------------------------------------------------------------------------------------ xlsx


@dataclass
class SheetInfo:
    name: str
    cells: list[tuple[int, int, str]]  # (fila, columna, texto) de celdas con texto de ≥ 2 palabras
    edge_runs: int
    bordered_cells: int
    images: int
    has_content: bool


def _cell_text(value: Any) -> str:
    return "" if value is None else str(value)


def _sheet_infos(xlsx: Path) -> list[SheetInfo]:
    wb = load_workbook(xlsx, rich_text=True)
    infos: list[SheetInfo] = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        texts: list[tuple[int, int, str]] = []
        h_edges: set[tuple[int, int]] = set()  # (límite entre filas, columna)
        v_edges: set[tuple[int, int]] = set()  # (límite entre columnas, fila)
        bordered = 0
        content = False
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if value is not None and str(value).strip():
                    content = True
                    text = _cell_text(value)
                    # un número o una fecha tipados no juntan palabras de cajas distintas, y su str() no es
                    # lo que se ve (con formato) ni lo que dice el PDF: solo se miran los textos
                    if isinstance(value, (str, CellRichText)) and len(tokens(text)) >= 2:
                        texts.append((cell.row, cell.column, text))
                b = cell.border
                has = False
                if b is not None:
                    if b.top is not None and b.top.style:
                        h_edges.add((cell.row - 1, cell.column))
                        has = True
                    if b.bottom is not None and b.bottom.style:
                        h_edges.add((cell.row, cell.column))
                        has = True
                    if b.left is not None and b.left.style:
                        v_edges.add((cell.column - 1, cell.row))
                        has = True
                    if b.right is not None and b.right.style:
                        v_edges.add((cell.column, cell.row))
                        has = True
                if has:
                    bordered += 1
                    content = True
        images = len(getattr(ws, "_images", []))
        infos.append(
            SheetInfo(
                ws.title,
                texts,
                _count_runs(h_edges) + _count_runs(v_edges),
                bordered,
                images,
                content or images > 0,
            )
        )
    wb.close()
    return infos


def _count_runs(edges: set[tuple[int, int]]) -> int:
    """Cuenta tramos maximales: bordes de la misma línea (mismo índice) en posiciones consecutivas."""
    by_line: dict[int, list[int]] = {}
    for line, pos in edges:
        by_line.setdefault(line, []).append(pos)
    runs = 0
    for positions in by_line.values():
        positions.sort()
        runs += 1 + sum(1 for a, b in pairwise(positions) if b != a + 1)
    return runs


# ------------------------------------------------------------------------------------------ render


def _find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    raise LibreOfficeUnavailable("LibreOffice (soffice) no está instalado: no se puede renderizar el Excel.")


def _prepare_for_print(xlsx: Path, dest: Path) -> list[str]:
    """Copia el libro forzando 1 página por hoja (solo para medir). Devuelve notas."""
    notes: list[str] = []
    try:
        wb = load_workbook(xlsx, rich_text=True)
        for ws in wb.worksheets:
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.page_setup.paperSize = 9
            ws.print_title_rows = None
            ws.print_title_cols = None
        wb.save(dest)
        wb.close()
    except Exception as exc:  # libro raro: se mide tal cual
        shutil.copyfile(xlsx, dest)
        notes.append(
            f"no se pudo forzar 1 página por hoja ({type(exc).__name__}); se midió el libro tal cual"
        )
    return notes


def render_xlsx(xlsx: Path, workdir: Path) -> tuple[Path, list[str]]:
    """Renderiza el .xlsx a PDF con LibreOffice. Devuelve (pdf, notas)."""
    soffice = _find_soffice()
    workdir.mkdir(parents=True, exist_ok=True)
    src = workdir / "measure.xlsx"
    notes = _prepare_for_print(xlsx, src)
    out_pdf = workdir / "measure.pdf"
    last_error = ""
    for attempt in range(SOFFICE_ATTEMPTS):
        out_pdf.unlink(missing_ok=True)
        profile = Path(tempfile.mkdtemp(prefix="lo_profile_", dir=workdir))
        env = dict(os.environ, HOME=str(profile))
        cmd = [
            soffice,
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(workdir),
            str(src),
        ]
        try:
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=SOFFICE_TIMEOUT_S, env=env, check=False
            )
            last_error = (proc.stderr or proc.stdout or "").strip()[-300:]
        except subprocess.TimeoutExpired:
            last_error = f"timeout de {SOFFICE_TIMEOUT_S} s"
        finally:
            shutil.rmtree(profile, ignore_errors=True)
        if out_pdf.exists() and out_pdf.stat().st_size > 0:
            return out_pdf, notes
        time.sleep(1.0 + attempt)
    raise LibreOfficeError(
        f"LibreOffice no pudo convertir el Excel tras {SOFFICE_ATTEMPTS} intentos: {last_error}"
    )


# ------------------------------------------------------------------------------------------ comparación


def _percentile(values: list[float], q: float) -> float:
    s = sorted(values)
    k = (len(s) - 1) * q
    lo, hi = math.floor(k), math.ceil(k)
    return s[lo] + (s[hi] - s[lo]) * (k - lo)


def _bbox(words: Sequence[Word]) -> tuple[float, float, float, float]:
    x0 = min(w.x0 for w in words)
    y0 = min(w.y0 for w in words)
    x1 = max(w.x1 for w in words)
    y1 = max(w.y1 for w in words)
    return x0, y0, max(x1, x0 + 1.0), max(y1, y0 + 1.0)


def _match_positions(orig: list[Word], rend: list[Word]) -> tuple[list[tuple[float, float]], int]:
    """Empareja por texto y devuelve (dx, dy) normalizados al bbox del contenido (en fracción) por palabra."""
    by_tok_o: dict[str, list[Word]] = {}
    by_tok_r: dict[str, list[Word]] = {}
    for w in orig:
        by_tok_o.setdefault(w.text, []).append(w)
    for w in rend:
        by_tok_r.setdefault(w.text, []).append(w)
    common = [t for t in by_tok_o if t in by_tok_r]
    if not common:
        return [], 0
    anchors = [t for t in common if len(by_tok_o[t]) == 1 and len(by_tok_r[t]) == 1]
    o_anchor = (
        [by_tok_o[t][0] for t in anchors] if len(anchors) >= 3 else [w for t in common for w in by_tok_o[t]]
    )
    r_anchor = (
        [by_tok_r[t][0] for t in anchors] if len(anchors) >= 3 else [w for t in common for w in by_tok_r[t]]
    )
    ox0, oy0, ox1, oy1 = _bbox(o_anchor)
    rx0, ry0, rx1, ry1 = _bbox(r_anchor)

    def norm_o(w: Word) -> tuple[float, float]:
        return (w.cx - ox0) / (ox1 - ox0), (w.cy - oy0) / (oy1 - oy0)

    def norm_r(w: Word) -> tuple[float, float]:
        return (w.cx - rx0) / (rx1 - rx0), (w.cy - ry0) / (ry1 - ry0)

    diffs: list[tuple[float, float]] = []
    for tok in common:
        os_ = [norm_o(w) for w in by_tok_o[tok]]
        rs = [norm_r(w) for w in by_tok_r[tok]]
        pairs = sorted(
            (
                (math.hypot(o[0] - r[0], o[1] - r[1]), i, j)
                for i, o in enumerate(os_)
                for j, r in enumerate(rs)
            ),
        )
        used_o: set[int] = set()
        used_r: set[int] = set()
        for _, i, j in pairs:
            if i in used_o or j in used_r:
                continue
            used_o.add(i)
            used_r.add(j)
            diffs.append((rs[j][0] - os_[i][0], rs[j][1] - os_[i][1]))
    return diffs, len(diffs)


def _load_ground_truth(pdf: Path, ground_truth: dict[str, Any] | None) -> dict[str, Any] | None:
    if ground_truth is not None:
        return ground_truth
    candidate = pdf.parent / GROUND_TRUTH_NAME
    if candidate.exists():
        try:
            data = json.loads(candidate.read_text(encoding="utf-8"))
            entry = data.get(pdf.name)
            return entry if isinstance(entry, dict) else None
        except (OSError, ValueError):
            return None
    return None


def _gt_boxes(gt: dict[str, Any] | None, page_no: int) -> list[tuple[float, float, float, float]]:
    if not gt:
        return []
    items = gt.get("boxes", {}).get("items", [])
    return [
        tuple(b["bbox"])
        for b in items
        if b.get("page") == page_no and b.get("kind") in ("label_value", "empty", "text", "checkbox")
    ]


def _chain_cost(prev: Word, cand: Word) -> tuple[int, float]:
    """Orden de preferencia del próximo token de una celda: el que sigue en el mismo renglón (a la derecha)
    gana a cualquier otro, aunque uno de otro renglón quede apenas más cerca."""
    same_line = abs(cand.cy - prev.cy) <= 0.5 * max(prev.y1 - prev.y0, 1.0) and cand.cx >= prev.cx - 0.5
    return (0 if same_line else 1, math.hypot(cand.cx - prev.cx, cand.cy - prev.cy))


def _count_mixed(
    page_words: list[Word],
    cells: list[tuple[int, int, str]],
    runs: list[BorderRun],
    boxes: list[tuple[float, float, float, float]],
) -> list[str]:
    """Textos de celdas que juntan palabras separadas por un borde o de cajas distintas."""
    if not page_words:
        return []
    by_tok: dict[str, list[Word]] = {}
    for w in page_words:
        by_tok.setdefault(w.text, []).append(w)

    def box_id(w: Word) -> int:
        best = -1
        best_area = float("inf")
        for k, (x0, y0, x1, y1) in enumerate(boxes):
            if x0 <= w.cx <= x1 and y0 <= w.cy <= y1:
                area = (x1 - x0) * (y1 - y0)
                if area < best_area:
                    best, best_area = k, area
        return best

    mixed: list[str] = []
    for _, _, text in cells:
        toks = tokens(text)
        if any(t not in by_tok for t in toks):
            toks = [t for t in toks if t in by_tok]
            if len(toks) < 2:
                continue
        # Un texto repetido en el PDF puede salir de cualquiera de sus apariciones: la celda se marca como
        # mezclada solo si NINGUNA cadena de palabras (empezando en cada aparición) es limpia.
        chains: list[list[Word]] = []
        for start in by_tok[toks[0]][:400]:
            chain = [start]
            for t in toks[1:]:
                prev = chain[-1]
                chain.append(min(by_tok[t], key=lambda w, p=prev: _chain_cost(p, w)))
            chains.append(chain)
        if not chains:
            continue

        def chain_is_mixed(chain: list[Word]) -> bool:
            if boxes and len({box_id(w) for w in chain}) > 1:
                return True
            return any(_crosses((a.cx, a.cy), (b.cx, b.cy), r) for a, b in pairwise(chain) for r in runs)

        is_mixed = all(chain_is_mixed(c) for c in chains)
        if is_mixed:
            mixed.append(text.replace("\n", " ")[:80])
    return mixed


# ------------------------------------------------------------------------------------------ PNG


def _side_by_side(
    pdf_doc: pymupdf.Document, render_doc: pymupdf.Document, out: Path, max_pages: int = 3
) -> None:
    rows: list[Image.Image] = []
    cell_h = 640
    for i in range(min(max_pages, max(len(pdf_doc), len(render_doc)))):
        pair: list[Image.Image] = []
        for doc in (pdf_doc, render_doc):
            if i < len(doc):
                pix = doc[i].get_pixmap(dpi=PNG_DPI)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            else:
                img = Image.new("RGB", (400, cell_h), "white")
            scale = cell_h / img.height
            pair.append(img.resize((max(1, int(img.width * scale)), cell_h)))
        width = pair[0].width + pair[1].width + 30
        row = Image.new("RGB", (width, cell_h + 24), (225, 225, 225))
        draw = ImageDraw.Draw(row)
        draw.text((8, 6), f"PDF original (pag. {i + 1})", fill=(0, 0, 0))
        draw.text((pair[0].width + 38, 6), "Excel renderizado", fill=(0, 0, 0))
        row.paste(pair[0], (0, 24))
        row.paste(pair[1], (pair[0].width + 30, 24))
        rows.append(row)
    if not rows:
        return
    sheet = Image.new("RGB", (max(r.width for r in rows), sum(r.height for r in rows)), (225, 225, 225))
    y = 0
    for r in rows:
        sheet.paste(r, (0, y))
        y += r.height
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)


# ------------------------------------------------------------------------------------------ score


def score(
    pdf: Path,
    xlsx: Path,
    workdir: Path,
    png: Path | None = None,
    ground_truth: dict[str, Any] | None = None,
) -> FidelityResult:
    """Mide la fidelidad de ``xlsx`` contra ``pdf``. Lanza ``LibreOfficeUnavailable`` si no hay soffice."""
    pdf, xlsx, workdir = Path(pdf), Path(xlsx), Path(workdir)
    rendered_path, notes = render_xlsx(xlsx, workdir)
    gt = _load_ground_truth(pdf, ground_truth)
    sheets = [s for s in _sheet_infos(xlsx)]
    printed = [s for s in sheets if s.has_content]  # LibreOffice no imprime hojas vacías

    with pymupdf.open(str(pdf)) as pdf_doc, pymupdf.open(str(rendered_path)) as render_doc:
        n_pdf, n_render = len(pdf_doc), len(render_doc)
        if n_render != len(printed):
            notes.append(f"el render tiene {n_render} páginas y el libro {len(printed)} hojas con contenido")
        words_expected = words_found = covered_words = extra_total = pos_matched = 0
        extra_examples: list[str] = []
        missing_examples: list[str] = []
        diffs: list[tuple[float, float]] = []
        pdf_border_segments = pdf_images = 0
        mixed_examples: list[str] = []
        mixed_total = 0
        per_page: list[dict[str, Any]] = []
        all_orig: Counter[str] = Counter()
        all_rend: Counter[str] = Counter()
        has_text = False

        for i in range(n_pdf):
            page = pdf_doc[i]
            spans = find_covered_spans(page)
            covered_words += sum(len(tokens(t)) for t, _ in spans)
            orig = visible_words(page, spans)
            runs = page_border_runs(page)
            n_border = len(runs)
            if not orig and not runs and page.get_images():
                n_border = _raster_border_count(page)  # escaneo: las cajas son píxeles, no vectores
            pdf_border_segments += n_border
            pdf_images += len(page.get_images())
            rend = visible_words(render_doc[i]) if i < n_render else []
            has_text = has_text or bool(orig)
            c_o = Counter(w.text for w in orig)
            c_r = Counter(w.text for w in rend)
            all_orig.update(c_o)
            all_rend.update(c_r)
            found = sum(min(n, c_r[t]) for t, n in c_o.items())
            extra = c_r - c_o
            missing = c_o - c_r
            words_expected += len(orig)
            words_found += found
            extra_total += sum(extra.values())
            extra_examples.extend(sorted(extra))
            missing_examples.extend(sorted(missing))
            d, n = _match_positions(orig, rend)
            diffs.extend(d)
            pos_matched += n
            per_page.append(
                {
                    "page": i + 1,
                    "words_expected": len(orig),
                    "words_found": found,
                    "extra": sum(extra.values()),
                    "recall": (found / len(orig)) if orig else None,
                }
            )
            if i < len(printed):
                mixed = _count_mixed(orig, printed[i].cells, runs, _gt_boxes(gt, i + 1))
                mixed_total += len(mixed)
                mixed_examples.extend(mixed)
        # hojas de más: todo el texto de páginas que el PDF no tiene es sobrante
        for j in range(n_pdf, n_render):
            surplus = Counter(w.text for w in visible_words(render_doc[j]))
            extra_total += sum(surplus.values())
            extra_examples.extend(sorted(surplus))
        if png is not None:
            _side_by_side(pdf_doc, render_doc, Path(png))

    found_any = sum(min(n, all_rend[t]) for t, n in all_orig.items())
    border_runs_x = sum(s.edge_runs for s in sheets)
    errs = [math.hypot(dx, dy) * 100 for dx, dy in diffs]
    if not has_text:
        notes.append("el PDF no tiene capa de texto (escaneado): no se puede medir recall ni posición")
    return FidelityResult(
        pdf=pdf.name,
        xlsx=xlsx.name,
        pdf_pages=n_pdf,
        sheets=len(sheets),
        render_pages=n_render,
        words_expected=words_expected,
        words_found=words_found,
        recall=(words_found / words_expected) if words_expected else None,
        recall_any_page=(found_any / words_expected) if words_expected else None,
        extra_words=extra_total,
        extra_examples=extra_examples[:30],
        missing_examples=missing_examples[:30],
        pos_matched=pos_matched,
        pos_err_median_pct=statistics.median(errs) if errs else None,
        pos_err_p90_pct=_percentile(errs, 0.9) if errs else None,
        dx_median_pct=statistics.median(abs(dx) * 100 for dx, _ in diffs) if diffs else None,
        dy_median_pct=statistics.median(abs(dy) * 100 for _, dy in diffs) if diffs else None,
        pdf_border_segments=pdf_border_segments,
        xlsx_border_segments=border_runs_x,
        xlsx_bordered_cells=sum(s.bordered_cells for s in sheets),
        border_ratio=(border_runs_x / pdf_border_segments) if pdf_border_segments else None,
        mixed_cells=mixed_total,
        mixed_examples=mixed_examples[:10],
        pdf_images=pdf_images,
        xlsx_images=sum(s.images for s in sheets),
        covered_words_in_pdf=covered_words,
        per_page=per_page,
        notes=notes,
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m tests.layout_score",
        description="Mide la fidelidad de un .xlsx contra su PDF (renderiza con LibreOffice).",
    )
    parser.add_argument("pdf", type=Path)
    parser.add_argument("xlsx", type=Path)
    parser.add_argument("--png", type=Path, default=None, help="guarda 'PDF original | Excel renderizado'")
    parser.add_argument("--json", action="store_true", help="imprime el resultado como JSON")
    parser.add_argument(
        "--workdir", type=Path, default=None, help="carpeta de trabajo (por defecto, temporal)"
    )
    args = parser.parse_args(argv)
    try:
        if args.workdir is None:
            with tempfile.TemporaryDirectory(prefix="layout_score_") as tmp:
                result = score(args.pdf, args.xlsx, Path(tmp), args.png)
        else:
            result = score(args.pdf, args.xlsx, args.workdir, args.png)
    except LibreOfficeUnavailable as exc:
        print(str(exc), file=sys.stderr)
        return 2
    print(result.to_json() if args.json else result.summary())
    return 0 if result.passes else 1


if __name__ == "__main__":
    sys.exit(main())
