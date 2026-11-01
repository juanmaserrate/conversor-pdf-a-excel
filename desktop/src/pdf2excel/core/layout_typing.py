"""Valores tipados conservadores para las celdas del modo "Igual al PDF" (LAYOUT_MODE.md §3.6).

Una celda de un solo renglón y un solo estilo se convierte en número o fecha solo si es inequívoca. Horas,
identificadores (N°, CUIT, teléfonos, patentes, códigos con letras), encabezados y todo lo ambiguo quedan como
texto: nunca se inventa un valor.
"""

from __future__ import annotations

import re

import pymupdf

from pdf2excel.core import typing_ar
from pdf2excel.models import CellValue, ColumnInfo, LayoutCell

_DATE_FULL_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
_AR_DECIMAL_RE = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d+$|^\d+,\d+$")
_LEADING_ZERO_RE = re.compile(r"^[-+]?0\d")
_INT_RE = re.compile(r"^\d{1,15}$")
_ALPHA_RE = re.compile(r"[^\W\d_]")
NUMBER_FIT = 1.03  # holgura de ancho para mostrar un número sin "###"
NUMBER_PAD_PT = 4.0  # relleno interno de una celda de Excel


def _single_text(cell: LayoutCell) -> str | None:
    """Texto de la celda si es un solo renglón de un solo estilo; si no, ``None``."""
    if len(cell.lines) != 1 or len(cell.lines[0]) != 1:
        return None
    text = cell.lines[0][0].text.strip()
    return text or None


def _decimals_format(text: str, currency: bool) -> str:
    clean, _neg, _cur = typing_ar.normalize_number_text(text)
    decimals = len(clean.split(",")[1]) if "," in clean else 0
    base = "#,##0" if "." in clean else "0"
    fmt = base + ("." + "0" * decimals if decimals else "")
    return f'"$" {fmt}' if currency else fmt


def _percent_format(text: str) -> str:
    """``0%`` / ``0.0%`` / ``0.00%`` según los decimales del texto (el formato no alarga lo que se ve)."""
    decimals = (
        len(re.split(r"[,.]", text.replace("%", "").strip())[1].strip()) if re.search(r"[,.]", text) else 0
    )
    return "0%" if decimals == 0 else "0." + "0" * decimals + "%"


def _standalone(text: str) -> tuple[CellValue, str | None]:
    """Valor de una celda aislada (sin columna): solo fechas completas y números con coma decimal."""
    m = _DATE_FULL_RE.match(text)
    if m:
        parsed = typing_ar.parse_date(text)
        if parsed is not None:
            return parsed, typing_ar.DATE_FORMAT
        return None, None
    if _ALPHA_RE.search(re.sub(r"U\$S|USD|ARS", "", text, flags=re.IGNORECASE)):
        return None, None
    if typing_ar.PCT.match(text):
        number = typing_ar.parse_percent(text, "ar")
        if number is not None and "." not in text:
            return number, _percent_format(text)
        return None, None
    clean, _neg, currency = typing_ar.normalize_number_text(text)
    if (currency or _AR_DECIMAL_RE.match(clean)) and typing_ar.NUM_AR.match(clean):
        if "," not in clean and "." in clean:
            return None, None  # "$ 1.234" sin decimales: ambiguo
        value = typing_ar.parse_number(text, "ar")
        if value is not None and not _LEADING_ZERO_RE.match(clean):
            return value, _decimals_format(text, currency)
    return None, None


def _fits_as_number(cell: LayoutCell, text: str) -> bool:
    """Excel muestra ``###`` si un número no entra en la celda (el texto, en cambio, se desborda): se tipa
    solo si el ancho de la celda alcanza para el texto con las métricas de Helvetica más el relleno."""
    run = cell.lines[0][0]
    font = pymupdf.Font("hebo" if run.bold else "helv")
    try:
        needed = float(font.text_length(text, fontsize=run.size_pt))
    except Exception:
        needed = len(text) * run.size_pt * 0.6
    return (cell.bbox[2] - cell.bbox[0]) >= needed * NUMBER_FIT + NUMBER_PAD_PT


def _is_data_candidate(cell: LayoutCell) -> bool:
    return _single_text(cell) is not None


def _column_runs(cells: list[LayoutCell]) -> list[list[LayoutCell]]:
    """Corridas verticales de ≥ 2 celdas contiguas con el mismo ``col``/``col_span``. Las filas de sección
    (una celda más ancha que abarca la columna) no cortan la corrida."""
    by_col: dict[tuple[str, int], list[LayoutCell]] = {}
    for c in cells:
        if _is_data_candidate(c):
            # la columna la define el borde al que se alinea el texto (el derecho en los números), no el
            # ancho de la celda: dos cifras de la misma columna pueden abarcar pistas distintas
            edge = c.col + c.col_span if c.halign == "right" else c.col
            by_col.setdefault((c.halign, edge), []).append(c)
    runs: list[list[LayoutCell]] = []
    for group in by_col.values():
        group.sort(key=lambda c: c.row)
        run = [group[0]]
        for c in group[1:]:
            if _contiguous(run[-1], c, cells):
                run.append(c)
            else:
                if len(run) >= 2:
                    runs.append(run)
                run = [c]
        if len(run) >= 2:
            runs.append(run)
    return runs


def _contiguous(prev: LayoutCell, nxt: LayoutCell, cells: list[LayoutCell]) -> bool:
    first = prev.row + prev.row_span
    if nxt.row == first:
        return True
    if nxt.row < first:
        return False
    for r in range(first, nxt.row):
        covering = [
            c for c in cells if c.row <= r < c.row + c.row_span and c.col <= prev.col < c.col + c.col_span
        ]
        if not any(c.lines for c in cells if c.row <= r < c.row + c.row_span):
            continue  # pista de separación entre renglones: no corta la columna
        extends = all(
            c.lines and (c.col < prev.col or c.col + c.col_span > prev.col + prev.col_span) for c in covering
        )
        if not covering or not extends:
            return False
    return True


def _is_bold(cell: LayoutCell) -> bool:
    return bool(cell.lines and cell.lines[0] and cell.lines[0][0].bold)


def apply_typed_values(cells: list[LayoutCell]) -> None:
    """Completa ``value`` / ``number_format`` de las celdas que se pueden tipar sin dudas."""
    for cell in cells:
        text = _single_text(cell)
        if text is None:
            continue
        value, fmt = _standalone(text)
        if value is not None and _fits_as_number(cell, text):
            cell.value, cell.number_format = value, fmt
    for run in _column_runs(cells):
        first = run[0]
        has_header = _is_bold(first) and not _is_bold(run[1])
        body = run[1:] if has_header else run
        header = (_single_text(first) or "") if has_header else ""
        texts = [_single_text(c) or "" for c in body]
        col = typing_ar.classify_column(header, texts)
        if col.kind in ("id", "text"):
            continue
        grouped = col.kind == "number" and any(_has_group_sep(t, col.locale or "ar") for t in texts)
        for cell, text in zip(body, texts, strict=True):
            if (
                isinstance(cell.value, (int, float))
                and not isinstance(cell.value, bool)
                and col.kind == "number"
            ):
                if "%" in text:
                    cell.value, cell.number_format = (
                        None,
                        None,
                    )  # "15%" en una columna numérica: texto tal cual
                    continue
                cell.number_format = _number_format(text, col, grouped)  # mismo formato en toda la columna
                continue
            if cell.value is not None or not text:
                continue
            value, fmt = _in_column(text, col, grouped)
            if value is not None and _fits_as_number(cell, text):
                cell.value, cell.number_format = value, fmt


def _has_group_sep(text: str, locale: str) -> bool:
    clean, _neg, _cur = typing_ar.normalize_number_text(text)
    return ("." if locale == "ar" else ",") in clean


def _number_format(text: str, col: ColumnInfo, grouped: bool) -> str:
    """Formato que muestra el número igual que el texto del PDF: los decimales son los de ese texto y el
    separador de miles solo si algún valor de la columna lo trae (no se inventan los que el PDF no tiene)."""
    locale = col.locale or "ar"
    clean, _neg, _cur = typing_ar.normalize_number_text(text)
    dec_sep = "," if locale == "ar" else "."
    decimals = len(clean.split(dec_sep)[1]) if dec_sep in clean else 0
    fmt = ("#,##0" if grouped else "0") + ("." + "0" * decimals if decimals else "")
    return f'"$" {fmt}' if col.number_format == typing_ar.CURRENCY_FORMAT else fmt


def _in_column(text: str, col: ColumnInfo, grouped: bool = False) -> tuple[CellValue, str | None]:
    if _LEADING_ZERO_RE.match(text.strip()):
        return None, None
    if col.kind == "date":
        parsed = typing_ar.parse_date(text)
        return (parsed, typing_ar.DATE_FORMAT) if parsed is not None else (None, None)
    if col.kind == "percent":
        number = typing_ar.parse_percent(text, col.locale or "ar")
        return (number, _percent_format(text)) if number is not None else (None, None)
    if col.kind == "number":
        if "%" in text:
            return None, None  # un "15%" suelto en una columna numérica quedaría como "0" con formato entero
        value = typing_ar.parse_number(text, col.locale or "ar")
        if value is None:
            return None, None
        return value, _number_format(text, col, grouped)
    return None, None
