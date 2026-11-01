"""Merge de tablas que continúan en la página siguiente (DESIGN.md §5.5).

Funciones puras sobre ``RawTable``; no se modifican las tablas de entrada.
"""

from __future__ import annotations

import re
from dataclasses import replace

from pdf2excel.core.typing_ar import PCT, parse_date, parse_number
from pdf2excel.models import RawTable

_CARRIED_OVER_RE = re.compile(r"^(?:subtotal\s+)?transporte", re.IGNORECASE)
_SPACES_RE = re.compile(r"\s+")


def _norm_header(row: list[str]) -> list[str]:
    return [_SPACES_RE.sub(" ", (c or "")).strip().lower() for c in row]


def _ncols(table: RawTable) -> int:
    return len(table.rows[0]) if table.rows else 0


def _cell_is_data(cell: str) -> bool:
    text = cell.strip()
    return (
        parse_number(text, "ar") is not None
        or parse_number(text, "us") is not None
        or parse_date(text) is not None
        or bool(PCT.match(text))
    )


def _looks_like_data_row(row: list[str]) -> bool:
    """≥ 50 % de las celdas no vacías parsean como número/fecha/porcentaje → no es un encabezado."""
    nonempty = [c for c in row if c and c.strip()]
    if not nonempty:
        return False
    return sum(1 for c in nonempty if _cell_is_data(c)) >= 0.5 * len(nonempty)


def _first_nonempty(row: list[str]) -> str:
    for cell in row:
        if cell and cell.strip():
            return cell.strip()
    return ""


def _drop_carried_over(rows: list[list[str]]) -> list[list[str]]:
    """Descarta las filas "Transporte"/"Subtotal transporte" al inicio: ya están al final de la anterior."""
    i = 0
    while i < len(rows) and _CARRIED_OVER_RE.match(_first_nonempty(rows[i])):
        i += 1
    return rows[i:]


def _rows_to_append(merged: RawTable, nxt: RawTable, col_tolerance_pt: float) -> list[list[str]] | None:
    """Filas de ``nxt`` que se concatenan a ``merged``, o ``None`` si no corresponde mergear."""
    if len(merged.rows) < 2 or len(nxt.rows) < 2:
        return None
    if nxt.page != merged.page_end + 1:
        return None
    if _ncols(merged) != _ncols(nxt):
        return None
    # Regla 1: mismo encabezado → se descarta el encabezado repetido.
    if _norm_header(merged.rows[0]) == _norm_header(nxt.rows[0]):
        return nxt.rows[1:]
    # Regla 2: continuación sin encabezado, misma geometría horizontal y primera fila con datos.
    same_x = (
        abs(nxt.bbox[0] - merged.bbox[0]) <= col_tolerance_pt
        and abs(nxt.bbox[2] - merged.bbox[2]) <= col_tolerance_pt
    )
    if same_x and _looks_like_data_row(nxt.rows[0]):
        return list(nxt.rows)
    return None


def merge_multipage(tables: list[RawTable], col_tolerance_pt: float = 3.0) -> list[RawTable]:
    """Concatena tablas de páginas consecutivas que son la misma tabla cortada por el salto de página.

    La tabla ``t`` se agrega a la última mergeada ``m`` si ``t.page == m.page_end + 1``, tienen el mismo
    número de columnas y (Regla 1) los encabezados coinciden tras ``lower()`` y colapsar espacios, o
    (Regla 2) los límites x del bbox coinciden ±``col_tolerance_pt`` y la primera fila de ``t`` no parece
    encabezado. Al concatenar: ``page_end = t.page_end``, ``quality`` = media ponderada por filas, ``bbox`` y
    ``title`` de la primera. Nunca se mergea una tabla con menos de 2 filas. Devuelve tablas nuevas.
    """
    out: list[RawTable] = []
    for table in tables:
        if out:
            merged = out[-1]
            extra = _rows_to_append(merged, table, col_tolerance_pt)
            if extra is not None:
                extra = _drop_carried_over(extra)
                w_prev, w_new = len(merged.rows), len(table.rows)
                quality = (merged.quality * w_prev + table.quality * w_new) / (w_prev + w_new)
                out[-1] = replace(
                    merged,
                    page_end=table.page_end,
                    rows=merged.rows + [list(r) for r in extra],
                    quality=quality,
                )
                continue
        out.append(replace(table, rows=[list(r) for r in table.rows]))
    return out


__all__ = ["merge_multipage"]
