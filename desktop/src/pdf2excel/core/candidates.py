"""Candidatos de tablas por página: ``hybrid_union`` y ``hybrid`` (DESIGN.md §5.3).

Geometría: ``page.find_tables(strategy="lines")`` de pymupdf 2.0 con el modelo de layout (``pymupdf.layout``
se autoimporta con ``import pymupdf``). Texto: palabras de ``page.get_text("words")`` asignadas por su
centro a cada celda (``cells.words_in_rect``), nunca ``Table.extract()``.

Bug conocido: ``union=True`` sobre una página cuyo layout ya se calculó devuelve 0 tablas, por eso cada
candidato se calcula sobre un ``pymupdf.Document`` fresco.

Páginas con ``/Rotate`` 90/180/270: ``find_tables`` devuelve bbox y celdas en el sistema de la página
rotada (la visible), pero ``get_text("words")`` da coordenadas sin rotar. ``words_for_page`` lleva las
palabras al espacio rotado con ``page.rotation_matrix`` para que ``words_in_rect`` encuentre el texto.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import pymupdf

from pdf2excel.core.cells import (
    clean_cell,
    drop_empty_rows,
    extend_tail_with_bottom,
    merge_ghost_columns,
    words_in_rect,
)
from pdf2excel.core.quality import table_quality
from pdf2excel.core.visibility import visible_words
from pdf2excel.models import RawTable

log = logging.getLogger(__name__)

CANDIDATE_UNION = "hybrid_union"
CANDIDATE_PLAIN = "hybrid"
MIN_ROWS = 2
MIN_COLS = 2

_FIND_TABLES_ERRORS = (TypeError, RuntimeError, ValueError)


def _find_tables(page: Any, union: bool) -> list[Any]:
    try:
        return list(page.find_tables(strategy="lines", union=union).tables)
    except _FIND_TABLES_ERRORS as exc:
        page_no = getattr(page, "number", "?")
        log.warning("find_tables(union=%s) falló en la página %s: %r", union, page_no, exc)
        return []


def words_for_page(page: Any) -> list[tuple]:
    """Palabras visibles de la página en el sistema de la página rotada (el de ``find_tables``).

    Con ``/Rotate 0`` es la lista tal cual; con 90/180/270 cada rect se multiplica por
    ``page.rotation_matrix`` (``Rect * Matrix`` normaliza x0 < x1, y0 < y1).
    """
    words = visible_words(page)  # sin el texto tapado/invisible (los "N/A" que el PDF dibuja y después tapa)
    if not int(getattr(page, "rotation", 0) or 0) % 360:
        return words
    matrix = page.rotation_matrix
    return [(*tuple(pymupdf.Rect(w[:4]) * matrix), *w[4:]) for w in words]


def _table_rows(table: Any, words: list[tuple]) -> tuple[list[list[str]], float]:
    """Filas de texto limpio de una tabla de pymupdf, con cola extendida y columnas fantasma fusionadas.

    Devuelve ``(filas, y1)``: ``y1`` es el borde inferior real (el del bbox o el del último renglón de la
    cola), para que el bbox de la ``RawTable`` cubra también los totales agregados.
    """
    rows = [[clean_cell(words_in_rect(words, c)) if c else "" for c in row.cells] for row in table.rows]
    _added, bottom = extend_tail_with_bottom(None, table, rows, words)
    rows = [[clean_cell(c) for c in r] for r in rows]  # las filas de la cola vienen crudas
    rows = drop_empty_rows(rows)
    return merge_ghost_columns(rows), bottom


def tables_from_page(page: Any, words: list[tuple], union: bool, source: str) -> list[RawTable]:
    """Tablas de una página ya abierta como `RawTable` (sin título: lo completa el pipeline)."""
    out: list[RawTable] = []
    page_no = int(page.number)
    for t in _find_tables(page, union):
        rows, bottom = _table_rows(t, words)
        if len(rows) < MIN_ROWS or not rows[0] or len(rows[0]) < MIN_COLS:
            continue
        bbox = tuple(float(v) for v in t.bbox)
        out.append(
            RawTable(
                page=page_no,
                page_end=page_no,
                bbox=(bbox[0], bbox[1], bbox[2], max(bbox[3], float(bottom))),
                rows=rows,
                source=source,
                quality=table_quality(rows),
            )
        )
    return out


def page_candidates(path: Path, page_no: int) -> dict[str, list[RawTable]]:
    """Devuelve ``{"hybrid_union": [...], "hybrid": [...]}`` para la página `page_no` (0-based) de `path`.

    Cada candidato usa un documento fresco; las palabras se toman de la misma página (rotadas al sistema de
    ``find_tables`` si la página tiene ``/Rotate``). Los errores de `find_tables` se loguean y dejan ese
    candidato vacío.
    """
    cands: dict[str, list[RawTable]] = {}
    words: list[tuple] | None = None
    for source, union in ((CANDIDATE_UNION, True), (CANDIDATE_PLAIN, False)):
        doc = pymupdf.open(path)
        try:
            page = doc[page_no]
            if words is None:
                words = words_for_page(page)
            cands[source] = tables_from_page(page, words, union, source)
        finally:
            doc.close()
    return cands


def warm_up() -> None:
    """Precalienta el modelo de layout (ONNX) con un PDF mínimo en memoria (para un hilo al iniciar)."""
    try:
        doc = pymupdf.open()
        try:
            page = doc.new_page(width=300, height=200)
            page.insert_text((40, 40), "a b", fontsize=10)
            page.draw_rect(pymupdf.Rect(30, 30, 200, 100), width=0.8)
            page.find_tables(strategy="lines")
        finally:
            doc.close()
    except Exception as exc:  # el precalentado nunca debe tirar la app
        log.warning("No se pudo precalentar el modelo de layout: %r", exc)


__all__ = [
    "CANDIDATE_PLAIN",
    "CANDIDATE_UNION",
    "page_candidates",
    "tables_from_page",
    "warm_up",
    "words_for_page",
]
