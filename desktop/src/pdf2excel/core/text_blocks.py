"""Texto fuera de tablas: pares clave/valor, texto libre y títulos de tabla (DESIGN.md §5.7).

Recibe ``pymupdf.Page`` y devuelve ``TextItem``. No importa Qt.
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Sequence
from typing import Any

import pymupdf

from pdf2excel.models import TextItem

Bbox = tuple[float, float, float, float]

KV_RE = re.compile(r"^([^:]{1,40}):\s*(.+)$")
_SEPARATOR_RE = re.compile(r"\s+[—–|]\s+")  # " — ", " – ", " | "
_SPACES_RE = re.compile(r"\s+")
_HAS_LETTER_RE = re.compile(r"[^\W\d_]")
_TEXT_BLOCK = 0  # block_type de page.get_text("blocks"): 0 texto, 1 imagen


def _clean(text: str) -> str:
    return _SPACES_RE.sub(" ", text.replace(" ", " ")).strip()


def _text_blocks(page: Any) -> list[tuple[Bbox, str]]:
    """Bloques de texto (sin imágenes) de la página, en orden de lectura.

    Los bbox se llevan al sistema de la página rotada (``page.rotation_matrix``) para compararlos con los
    bbox de tabla de ``find_tables``, que ya vienen rotados; con ``/Rotate 0`` la matriz es la identidad.
    """
    out: list[tuple[Bbox, str]] = []
    rotated = bool(int(getattr(page, "rotation", 0) or 0) % 360)
    matrix = page.rotation_matrix if rotated else None
    for block in page.get_text("blocks", sort=True):
        if len(block) > 6 and block[6] != _TEXT_BLOCK:
            continue
        rect = pymupdf.Rect(*(float(v) for v in block[:4]))
        if matrix is not None:
            rect = rect * matrix
        out.append(((rect.x0, rect.y0, rect.x1, rect.y1), str(block[4])))
    if matrix is not None:
        # sort=True ordenó en el sistema sin rotar: el orden de lectura es el de la página visible.
        out.sort(key=lambda item: (round(item[0][1], 1), item[0][0]))
    return out


def _overlaps_table(rect: pymupdf.Rect, table_bboxes: Sequence[Bbox], min_fraction: float = 0.5) -> bool:
    area = rect.get_area()
    if area <= 0:
        return False
    for bbox in table_bboxes:
        table_rect = pymupdf.Rect(bbox)
        if not table_rect.intersects(rect):
            continue
        if (table_rect & rect).get_area() >= min_fraction * area:
            return True
    return False


def _parse_kv(piece: str) -> tuple[str, str] | None:
    m = KV_RE.match(piece)
    if not m:
        return None
    key, value = _clean(m[1]), _clean(m[2])
    if not key or not value or not _HAS_LETTER_RE.search(key):
        return None
    return key, value


def line_items(page_no: int, line: str) -> list[TextItem]:
    """Parte una línea por `` — ``/`` – ``/`` | `` y detecta ``clave: valor`` en cada trozo."""
    items: list[TextItem] = []
    for piece in _SEPARATOR_RE.split(line):
        piece = _clean(piece)
        if not piece:
            continue
        kv = _parse_kv(piece)
        if kv is not None:
            items.append(TextItem(page=page_no, kind="kv", key=kv[0], value=kv[1]))
        else:
            items.append(TextItem(page=page_no, kind="text", key=None, value=piece))
    return items


def text_outside_tables(page: Any, table_bboxes: list[tuple]) -> list[TextItem]:
    """Texto de la página que no está dentro de una tabla, como ``TextItem`` kv/text.

    Se toman los bloques de texto cuya intersección con cualquier bbox de tabla es menor al 50 % de su
    área. Cada bloque se parte en líneas; las vacías se omiten.
    """
    page_no = int(page.number)
    bboxes: list[Bbox] = [tuple(float(v) for v in b[:4]) for b in table_bboxes]  # type: ignore[misc]
    items: list[TextItem] = []
    for bbox, text in _text_blocks(page):
        if _overlaps_table(pymupdf.Rect(bbox), bboxes):
            continue
        for line in text.split("\n"):
            if line.strip():
                items.extend(line_items(page_no, line))
    return items


def title_near(page: Any, bbox: tuple, max_gap_pt: float = 40) -> str | None:
    """Título de la tabla: bloque de una línea (≤ 60 chars, no kv) justo arriba del bbox.

    Su ``y1`` debe caer entre ``bbox[1] - max_gap_pt`` y ``bbox[1]`` y solaparse horizontalmente con la
    tabla. Si hay varios, gana el más cercano. Si no hay, ``None``.
    """
    x0, top, x1 = float(bbox[0]), float(bbox[1]), float(bbox[2])
    best: tuple[float, str] | None = None
    for (bx0, _by0, bx1, by1), text in _text_blocks(page):
        if not (top - max_gap_pt <= by1 <= top):
            continue
        if bx1 <= x0 or bx0 >= x1:
            continue
        lines = [_clean(ln) for ln in text.split("\n") if ln.strip()]
        if len(lines) != 1:
            continue
        line = lines[0]
        if len(line) > 60 or _parse_kv(line) is not None:
            continue
        if best is None or by1 > best[0]:
            best = (by1, line)
    return best[1] if best else None


def repeated_lines(pages_items: list[list[TextItem]], min_ratio: float = 0.7) -> set[str]:
    """Textos (kind ``text``) que aparecen en ≥ ``min_ratio`` de las páginas: encabezado/pie de página.

    El pipeline los emite una sola vez. Con una sola página no hay repetición posible → conjunto vacío.
    """
    n_pages = len(pages_items)
    if n_pages < 2:
        return set()
    counts: Counter[str] = Counter()
    for items in pages_items:
        counts.update({_clean(it.value) for it in items if it.kind == "text" and it.value.strip()})
    threshold = min_ratio * n_pages
    return {text for text, n in counts.items() if n >= threshold}


__all__ = ["KV_RE", "line_items", "repeated_lines", "text_outside_tables", "title_near"]
