"""Scorer del corpus: compara las ``TypedTable`` extraídas con el ground truth (DESIGN.md §9).

El GT (``tests/corpus/ground_truth.json``) tiene por PDF una lista de tablas ``{"header": [...], "rows":
[[{"v": 350.0, "t": "float"} | {"v": "A-100", "t": "str"} | {"v": "2026-10-15", "t": "date"} | null]]}``.
Las métricas separan **contenido** (``cells_pct``: la celda dice lo mismo, aunque haya quedado como texto)
de **tipado** (``type_pct``: la celda numérica/fecha del GT llegó como número/fecha real).
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from pdf2excel.core.typing_ar import parse_date, parse_number, parse_percent
from pdf2excel.models import CellValue, TypedTable

GtCell = Mapping[str, Any] | None
TYPED_KINDS = frozenset({"float", "int", "date"})
_SUFFIX_RE = re.compile(r"_(?:scan\d*|ocr)$")
_SPACES_RE = re.compile(r"\s+")


@dataclass
class Score:
    cells_pct: float  # celdas GT (encabezado incluido) con la misma celda en la mejor fila extraída
    rows_pct: float  # filas GT con match exacto completo
    spurious_rows_pct: float  # filas extraídas (tablas elegidas) que no son la mejor de ninguna fila GT
    spurious_cols: int  # columnas extraídas de más (suma sobre las tablas GT)
    tables_detected: int
    tables_expected: int
    type_pct: float  # celdas GT float/int/date cuyo valor extraído es int/float/date (no str)


# --------------------------------------------------------------------------- normalización


def normalize(value: Any) -> str | float:
    """Forma canónica para comparar: números → ``float`` a 2 decimales, fechas → ISO, texto → minúsculas
    con espacios colapsados, ``None`` → ``""``."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).lower()
    if isinstance(value, int | float):
        return round(float(value), 2)
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return _SPACES_RE.sub(" ", str(value)).strip().lower()


def gt_value(cell: GtCell) -> CellValue:
    """Valor Python de una celda GT (``date`` para ``t == "date"``)."""
    if cell is None:
        return None
    value = cell.get("v")
    if cell.get("t") == "date" and isinstance(value, str):
        return date.fromisoformat(value)
    return value


def _lenient_numeric(text: str) -> set[str | float]:
    """Interpretaciones numéricas/fecha de un texto (AR y US): para que "350.00" como str cuente como
    contenido correcto frente a 350.0 aunque el tipado haya fallado."""
    out: set[str | float] = set()
    d = parse_date(text)
    if d is not None:
        out.add(d.isoformat())
    for locale in ("ar", "us"):
        n = parse_number(text, locale)
        if n is not None:
            out.add(round(float(n), 2))
        p = parse_percent(text, locale)
        if p is not None:
            out.add(round(float(p), 2))
    return out


def cells_match(gt_cell: GtCell, extracted: CellValue) -> bool:
    """¿La celda extraída dice lo mismo que la del GT? ``None`` equivale a ``""``; un texto que parsea al
    mismo número/fecha del GT cuenta como contenido correcto (el tipo lo mide ``type_pct``)."""
    expected = normalize(gt_value(gt_cell))
    got = normalize(extracted)
    if expected == got:
        return True
    if isinstance(extracted, str) and gt_cell is not None and gt_cell.get("t") in TYPED_KINDS:
        return expected in _lenient_numeric(extracted)
    return False


def type_ok(gt_cell: GtCell, extracted: CellValue) -> bool:
    """La celda GT es float/int/date y lo extraído es un número o fecha real (no str, no None)."""
    if gt_cell is None or gt_cell.get("t") not in TYPED_KINDS:
        return False
    if isinstance(extracted, bool) or extracted is None:
        return False
    return isinstance(extracted, int | float | date | datetime)


# --------------------------------------------------------------------------- matching


def gt_rows_with_header(gt_table: Mapping[str, Any]) -> list[list[GtCell]]:
    """Filas GT a comparar: el encabezado (como celdas ``str``) seguido de las filas de datos."""
    header = [{"v": h, "t": "str"} for h in gt_table.get("header", [])]
    return [header, *[list(r) for r in gt_table.get("rows", [])]]


def table_rows_with_header(table: TypedTable) -> list[list[CellValue]]:
    return [list(table.header), *[list(r) for r in table.rows]]


def _row_hits(gt_row: Sequence[GtCell], row: Sequence[CellValue]) -> int:
    return sum(1 for g, e in zip(gt_row, row, strict=False) if cells_match(g, e))


def _best_row(gt_row: Sequence[GtCell], rows: Sequence[Sequence[CellValue]]) -> tuple[int, int]:
    """``(índice de la mejor fila extraída, aciertos)``; ``(-1, 0)`` si no hay filas."""
    best_idx, best_hits = -1, -1
    for i, row in enumerate(rows):
        hits = _row_hits(gt_row, row)
        if hits > best_hits:
            best_idx, best_hits = i, hits
    return best_idx, max(best_hits, 0)


def _table_hits(gt_table: Mapping[str, Any], table: TypedTable) -> int:
    rows = table_rows_with_header(table)
    return sum(_best_row(g, rows)[1] for g in gt_rows_with_header(gt_table))


def match_tables(tables: Sequence[TypedTable], gt: Mapping[str, Any]) -> list[int | None]:
    """Para cada tabla GT, el índice de la tabla extraída con más celdas coincidentes (``None`` si no hay
    tablas extraídas). Dos tablas GT pueden elegir la misma tabla extraída (p. ej. si se fusionaron)."""
    chosen: list[int | None] = []
    for gt_table in gt.get("tables", []):
        best: int | None = None
        best_hits = -1
        for i, table in enumerate(tables):
            hits = _table_hits(gt_table, table)
            if hits > best_hits:
                best, best_hits = i, hits
        chosen.append(best)
    return chosen


# --------------------------------------------------------------------------- score


def _pct(num: float, den: float, empty: float = 100.0) -> float:
    return round(100.0 * num / den, 2) if den else empty


def score(tables: Sequence[TypedTable], gt: Mapping[str, Any]) -> Score:
    """Puntúa ``tables`` (las ``TypedTable`` de un ``ConversionReport``) contra la entrada GT de ese PDF."""
    gt_tables: list[Mapping[str, Any]] = list(gt.get("tables", []))
    chosen = match_tables(tables, gt)

    total_cells = hit_cells = 0
    total_rows = exact_rows = 0
    typed_total = typed_ok = 0
    spurious_cols = 0
    used_rows: dict[int, set[int]] = {}  # tabla extraída → filas que son la mejor de alguna fila GT

    for gt_table, idx in zip(gt_tables, chosen, strict=True):
        gt_rows = gt_rows_with_header(gt_table)
        ext_rows: list[list[CellValue]] = table_rows_with_header(tables[idx]) if idx is not None else []
        if idx is not None:
            ncols_ext = max((len(r) for r in ext_rows), default=0)
            ncols_gt = max((len(r) for r in gt_rows), default=0)
            spurious_cols += max(0, ncols_ext - ncols_gt)
        for r, gt_row in enumerate(gt_rows):
            total_cells += len(gt_row)
            best_idx, hits = _best_row(gt_row, ext_rows)
            hit_cells += hits
            total_rows += 1
            best_row: Sequence[CellValue] = ext_rows[best_idx] if best_idx >= 0 else []
            if best_idx >= 0:
                assert idx is not None
                used_rows.setdefault(idx, set()).add(best_idx)
                if hits == len(gt_row) and len(best_row) == len(gt_row):
                    exact_rows += 1
            if r == 0:
                continue  # el encabezado no cuenta para el tipado
            for c, gt_cell in enumerate(gt_row):
                if gt_cell is not None and gt_cell.get("t") in TYPED_KINDS:
                    typed_total += 1
                    if c < len(best_row) and type_ok(gt_cell, best_row[c]):
                        typed_ok += 1

    chosen_idx = {i for i in chosen if i is not None}
    ext_total = sum(len(table_rows_with_header(tables[i])) for i in chosen_idx)
    ext_used = sum(len(used_rows.get(i, set())) for i in chosen_idx)
    if gt_tables:
        spurious_pct = _pct(ext_total - ext_used, ext_total, empty=0.0)
    else:
        spurious_pct = 100.0 if any(t.rows or t.header for t in tables) else 0.0

    return Score(
        cells_pct=_pct(hit_cells, total_cells),
        rows_pct=_pct(exact_rows, total_rows),
        spurious_rows_pct=spurious_pct,
        spurious_cols=spurious_cols,
        tables_detected=len(tables),
        tables_expected=int(gt.get("expected_tables", len(gt_tables))),
        type_pct=_pct(typed_ok, typed_total),
    )


def mismatches(tables: Sequence[TypedTable], gt: Mapping[str, Any], limit: int = 10) -> list[str]:
    """Descripciones legibles de las primeras celdas que no coinciden (para mensajes de assert)."""
    out: list[str] = []
    for t_idx, (gt_table, idx) in enumerate(zip(gt.get("tables", []), match_tables(tables, gt), strict=True)):
        ext_rows = table_rows_with_header(tables[idx]) if idx is not None else []
        for r, gt_row in enumerate(gt_rows_with_header(gt_table)):
            best_idx, _ = _best_row(gt_row, ext_rows)
            best_row: Sequence[CellValue] = ext_rows[best_idx] if best_idx >= 0 else []
            for c, gt_cell in enumerate(gt_row):
                got = best_row[c] if c < len(best_row) else "<sin columna>"
                if c >= len(best_row) or not cells_match(gt_cell, best_row[c]):
                    out.append(f"tabla {t_idx} fila {r} col {c}: GT={gt_value(gt_cell)!r} extraído={got!r}")
                    if len(out) >= limit:
                        return out
    return out


# --------------------------------------------------------------------------- GT y reporte


def gt_name(name: str) -> str:
    """``"hard_factura_ar_scan90.pdf"`` → ``"hard_factura_ar.pdf"`` (quita ``_scan*`` / ``_ocr``)."""
    stem = name[:-4] if name.lower().endswith(".pdf") else name
    return _SUFFIX_RE.sub("", stem) + ".pdf"


def gt_for(name: str, gt: Mapping[str, Any]) -> dict[str, Any]:
    """Entrada GT para ``name`` (con o sin ``.pdf``); los ``_scan``/``_scan90``/``_scan150``/``_ocr`` usan
    la del original. ``KeyError`` si no existe."""
    return dict(gt[gt_name(name)])


def _fmt_pct(value: float) -> str:
    return f"{value:.1f}".replace(".", ",")


def format_table(
    results: Mapping[str, Score],
    meta: Mapping[str, Mapping[str, Any]] | None = None,
    header_lines: Iterable[str] = (),
) -> str:
    """Tabla markdown archivo × (celdas %, filas %, espurias %, tipos %, tablas, OCR, segundos) para el
    summary de CI. ``meta[name]`` puede traer ``ocr`` (bool) y ``seconds`` (float)."""
    lines = list(header_lines)
    lines.append("| Archivo | Celdas % | Filas % | Espurias % | Tipos % | Tablas | OCR | Segundos |")
    lines.append("|---|---:|---:|---:|---:|---:|:---:|---:|")
    for name in sorted(results):
        s = results[name]
        info = (meta or {}).get(name, {})
        ocr = info.get("ocr")
        ocr_txt = "—" if ocr is None else ("sí" if ocr else "no")
        seconds = info.get("seconds")
        sec_txt = "—" if seconds is None else f"{seconds:.1f}".replace(".", ",")
        tables_txt = f"{s.tables_detected}/{s.tables_expected}"
        lines.append(
            f"| {name} | {_fmt_pct(s.cells_pct)} | {_fmt_pct(s.rows_pct)} | {_fmt_pct(s.spurious_rows_pct)} "
            f"| {_fmt_pct(s.type_pct)} | {tables_txt} | {ocr_txt} | {sec_txt} |"
        )
    return "\n".join(lines) + "\n"


__all__ = [
    "TYPED_KINDS",
    "Score",
    "cells_match",
    "format_table",
    "gt_for",
    "gt_name",
    "gt_rows_with_header",
    "gt_value",
    "match_tables",
    "mismatches",
    "normalize",
    "score",
    "table_rows_with_header",
    "type_ok",
]
