"""Escritura del libro .xlsx a partir de un `ConversionReport` (DESIGN.md §6).

Una hoja por tabla con tipos reales (int/float/date/str), estilos, celdas para revisar en amarillo,
hoja "Revisión" (si hay algo que revisar), "Datos del documento" (texto fuera de tablas) y "Origen"
(oculta por defecto, para reportar problemas). Nunca se escriben fórmulas.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from pdf2excel.errors import UserError, from_exception
from pdf2excel.excel import styles
from pdf2excel.excel.layout_writer import write_layout_sheets
from pdf2excel.excel.naming import safe_sheet_name
from pdf2excel.models import (
    CellValue,
    ColumnInfo,
    ConversionReport,
    Options,
    TextItem,
    TypedTable,
)

# Hojas fijas.
SHEET_REVIEW = "Revisión"
SHEET_DOC_DATA = "Datos del documento"
SHEET_TEXT = "Texto"
SHEET_ORIGIN = "Origen"

NO_TEXT_MESSAGE = "El PDF no tiene texto extraíble"
OCR_REVIEW_MESSAGE = "Este PDF era un escaneo: los números pueden tener errores de lectura"

_STAGE_LABELS = {
    "classify": "Clasificar páginas",
    "ocr": "Reconocer texto (OCR)",
    "tables": "Buscar tablas",
    "layout": "Armar el diseño de la hoja",
    "typing": "Interpretar números y fechas",
    "excel": "Armar el Excel",
}
_PAGE_KIND_LABELS = {
    "digital": "digital",
    "scanned": "escaneada",
    "scanned_with_text": "escaneada con texto (OCR previo)",
    "garbage_text": "texto ilegible (tratada como escaneo)",
    "empty": "vacía",
}
_REVIEW_REASONS = {
    "number": "No se pudo interpretar como número",
    "date": "No se pudo interpretar como fecha",
    "percent": "No se pudo interpretar como porcentaje",
}

_RIGHT_ALIGNED_KINDS = {"number", "percent", "date"}
_SIMPLE_DATE_RE = re.compile(r"^\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})\s*$")


# --------------------------------------------------------------------------- formato AR (vista previa)


def _ar_number(value: float, decimals: int) -> str:
    """Formatea con punto de miles y coma decimal, redondeo "half up" como Excel."""
    quant = Decimal(1).scaleb(-decimals)  # 1, 0.1, 0.01...
    d = Decimal(str(value)).quantize(quant, rounding=ROUND_HALF_UP)
    text = f"{abs(d):,.{decimals}f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return f"-{text}" if d < 0 else text


def default_number_format(col: ColumnInfo, values: Iterable[CellValue] = ()) -> str | None:
    """Formato a usar cuando `col.number_format` es None, según el tipo de columna y sus valores."""
    if col.number_format:
        return col.number_format
    if col.kind == "date":
        return styles.FMT_DATE
    if col.kind == "percent":
        return styles.FMT_PCT
    if col.kind == "number":
        has_decimals = any(isinstance(v, float) and not float(v).is_integer() for v in values)
        return styles.FMT_DEC if has_decimals else styles.FMT_INT
    return None


def format_preview_value(value: CellValue, col: ColumnInfo) -> str:
    """Representación de la celda tal como la mostrará Excel en es-AR (para la vista previa de la UI).

    1234.5 con "#,##0.00" -> "1.234,50"; date -> "15/10/2026"; 0.21 con percent -> "21,00 %".
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, bool):
        return "Sí" if value else "No"
    if isinstance(value, int | float):
        fmt = default_number_format(col, (value,))
        if fmt == styles.FMT_PCT:
            return f"{_ar_number(float(value) * 100, 2)} %"
        if fmt == styles.FMT_MONEY:
            return f"$ {_ar_number(float(value), 2)}"
        if fmt == styles.FMT_INT:
            return _ar_number(float(value), 0)
        if fmt == styles.FMT_DEC:
            return _ar_number(float(value), 2)
        if isinstance(value, int):
            return _ar_number(float(value), 0)
        return _ar_number(float(value), 2)
    return str(value)


# --------------------------------------------------------------------------- helpers de celdas


def _set_value(cell: Cell, value: CellValue | datetime) -> None:
    """Escribe el valor Python tal cual; los strings nunca se interpretan como fórmula."""
    if value is None:
        return
    if isinstance(value, str):
        text = ILLEGAL_CHARACTERS_RE.sub("", value)
        cell.value = text
        if text.startswith("="):
            cell.data_type = "s"  # texto literal, nunca fórmula
        return
    cell.value = value


def _apply_style(cell: Cell, style: dict[str, Any]) -> None:
    for attr, val in style.items():
        setattr(cell, attr, val)


def _write_header_row(ws: Worksheet, row: int, labels: Iterable[str]) -> None:
    style = styles.header_style()
    for c, label in enumerate(labels, start=1):
        cell = ws.cell(row=row, column=c)
        _set_value(cell, label)
        _apply_style(cell, style)


def _display_width(text: str) -> int:
    """Largo del renglón más largo de un texto (para medir anchos)."""
    return max((len(line) for line in str(text).split("\n")), default=0)


def _set_col_width(ws: Worksheet, col_idx: int, longest: int) -> None:
    width = max(styles.MIN_COL_WIDTH, min(styles.MAX_COL_WIDTH, longest + 2))
    ws.column_dimensions[get_column_letter(col_idx)].width = width


def _simple_date(text: str) -> date | None:
    """`dd/mm/aaaa` (también con `-` o `.`) -> date; cualquier otra cosa -> None."""
    m = _SIMPLE_DATE_RE.match(text)
    if not m:
        return None
    day, month, year = (int(g) for g in m.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


# --------------------------------------------------------------------------- hojas


def _sheet_title_for(table: TypedTable, index: int, opts: Options, used: set[str]) -> str:
    title = table.name or f"Tabla {index}"
    if opts.one_sheet_per_page:
        title = f"{title} p.{table.pages[0] + 1}"
    return safe_sheet_name(title, index, used)


def _write_table(ws: Worksheet, table: TypedTable) -> None:
    n_cols = max(len(table.header), len(table.columns), max((len(r) for r in table.rows), default=0))
    if n_cols == 0:
        return
    columns = list(table.columns) + [ColumnInfo(kind="text")] * (n_cols - len(table.columns))
    header = list(table.header) + [""] * (n_cols - len(table.header))

    _write_header_row(ws, 1, header)

    border = styles.thin_border()
    total_border = styles.total_row_border()
    bold = Font(bold=True)
    fill = styles.review_fill()
    widths = [_display_width(h) for h in header]
    long_text_cols: set[int] = set()
    formats: list[str | None] = []
    for c, col in enumerate(columns):
        values = (row[c] if c < len(row) else None for row in table.rows)
        formats.append(default_number_format(col, values))

    for r, row in enumerate(table.rows):
        excel_row = r + 2
        is_total = r in table.total_rows
        for c in range(n_cols):
            value = row[c] if c < len(row) else None
            col = columns[c]
            cell = ws.cell(row=excel_row, column=c + 1)
            _set_value(cell, value)
            cell.border = total_border if is_total else border
            is_review = (r, c) in table.review_cells
            right = col.kind in _RIGHT_ALIGNED_KINDS and not is_review
            cell.alignment = styles.body_alignment(right=right)
            fmt = formats[c]
            if fmt and not isinstance(value, str):
                cell.number_format = fmt
            if is_total:
                cell.font = bold
            if is_review:
                cell.fill = fill
                cell.comment = Comment(styles.review_comment_text, styles.REVIEW_COMMENT_AUTHOR)
            shown = _display_width(format_preview_value(value, col))
            if shown > widths[c]:
                widths[c] = shown
            if isinstance(value, str) and shown > styles.MAX_COL_WIDTH:
                long_text_cols.add(c)

    for c in range(n_cols):
        _set_col_width(ws, c + 1, widths[c])
        if c in long_text_cols:
            col = columns[c]
            for r in range(len(table.rows)):
                cell = ws.cell(row=r + 2, column=c + 1)
                right = col.kind in _RIGHT_ALIGNED_KINDS and (r, c) not in table.review_cells
                cell.alignment = styles.body_alignment(right=right, wrap=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _review_reason(col: ColumnInfo) -> str:
    return _REVIEW_REASONS.get(col.kind, styles.review_comment_text)


def _write_review_sheet(wb: Workbook, report: ConversionReport, sheet_names: list[str]) -> None:
    ws = wb.create_sheet(SHEET_REVIEW)
    _write_header_row(ws, 1, ["Hoja", "Celda", "Valor", "Motivo"])
    row = 2
    if report.ocr_applied:
        cell = ws.cell(row=row, column=1)
        _set_value(cell, OCR_REVIEW_MESSAGE)
        cell.font = Font(italic=True)
        row += 1
    for table, sheet in zip(report.tables, sheet_names, strict=True):
        for r, c in sorted(table.review_cells):
            value = table.rows[r][c] if r < len(table.rows) and c < len(table.rows[r]) else None
            col = table.columns[c] if c < len(table.columns) else ColumnInfo(kind="text")
            ref = f"{get_column_letter(c + 1)}{r + 2}"
            for ci, val in enumerate([sheet, ref, value, _review_reason(col)], start=1):
                cell = ws.cell(row=row, column=ci)
                _set_value(cell, val)
                cell.alignment = styles.body_alignment(right=False)
            row += 1
    for c, width in enumerate([31, 8, 40, 48], start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A2"


def _write_text_sheet(wb: Workbook, items: list[TextItem], title: str) -> Worksheet:
    """Bloque clave/valor, fila en blanco y texto libre por página."""
    ws = wb.create_sheet(title)
    kv = [i for i in items if i.kind == "kv"]
    free = [i for i in items if i.kind != "kv"]
    row = 1
    if kv:
        _write_header_row(ws, row, ["Campo", "Valor", "Página"])
        row += 1
        for item in kv:
            _set_value(ws.cell(row=row, column=1), item.key or "")
            value_cell = ws.cell(row=row, column=2)
            parsed = _simple_date(item.value)
            if parsed is not None:
                _set_value(value_cell, parsed)
                value_cell.number_format = styles.FMT_DATE
                value_cell.alignment = styles.body_alignment(right=False)
            else:
                _set_value(value_cell, item.value)
                value_cell.alignment = styles.body_alignment(right=False, wrap=True)
            page_cell = ws.cell(row=row, column=3)
            _set_value(page_cell, item.page + 1)
            page_cell.alignment = styles.body_alignment(right=True)
            row += 1
        row += 1  # fila en blanco entre bloques
    if free or not kv:
        _write_header_row(ws, row, ["Página", "Texto"])
        row += 1
        if not free and not kv:
            _set_value(ws.cell(row=row, column=2), NO_TEXT_MESSAGE)
            row += 1
        for item in free:
            page_cell = ws.cell(row=row, column=1)
            _set_value(page_cell, item.page + 1)
            page_cell.alignment = styles.body_alignment(right=True)
            text_cell = ws.cell(row=row, column=2)
            _set_value(text_cell, item.value)
            text_cell.alignment = styles.body_alignment(right=False, wrap=True)
            row += 1
    key_width = max((len(i.key or "") for i in kv), default=0)
    ws.column_dimensions["A"].width = max(styles.MIN_COL_WIDTH, min(styles.MAX_COL_WIDTH, key_width + 2))
    ws.column_dimensions["B"].width = styles.TEXT_COL_WIDTH
    ws.column_dimensions["C"].width = styles.MIN_COL_WIDTH
    return ws


_MODE_LABELS = {"layout": "Igual al PDF", "tables": "Solo tablas de datos"}
_MAX_HIDDEN_LISTED = 200  # textos tapados que se listan en Origen (es diagnóstico, no un volcado)


def _write_layout_origin(
    ws: Worksheet, row: int, groups: Sequence[tuple[ConversionReport, list[str]]]
) -> int:
    """Hojas del modo "Igual al PDF" y textos tapados que se descartaron. Devuelve la fila siguiente."""
    _write_header_row(ws, row, ["Hoja", "Página", "Palabras", "Celdas", "Imágenes"])
    row += 1
    hidden: list[tuple[str, str]] = []
    for report, names in groups:
        for page, sheet in zip(report.layout_pages, names, strict=False):
            values: list[CellValue] = [sheet, page.page + 1, page.n_words, len(page.cells), len(page.images)]
            for ci, val in enumerate(values, start=1):
                _set_value(ws.cell(row=row, column=ci), val)
            row += 1
            hidden.extend((sheet, text) for text in page.hidden_texts)
    row += 1
    if hidden:
        _write_header_row(ws, row, ["Texto tapado o invisible (no se copió)", "Hoja"])
        row += 1
        for sheet, text in hidden[:_MAX_HIDDEN_LISTED]:
            _set_value(ws.cell(row=row, column=1), text)
            _set_value(ws.cell(row=row, column=2), sheet)
            row += 1
        row += 1
    return row


def _write_origin_sheet(
    wb: Workbook,
    report: ConversionReport,
    opts: Options,
    sheet_names: list[str],
    layout_groups: Sequence[tuple[ConversionReport, list[str]]] = (),
) -> Worksheet:
    ws = wb.create_sheet(SHEET_ORIGIN)
    bold = Font(bold=True)
    row = 1

    def kv(key: str, value: CellValue) -> None:
        nonlocal row
        k = ws.cell(row=row, column=1)
        _set_value(k, key)
        k.font = bold
        _set_value(ws.cell(row=row, column=2), value)
        row += 1

    kv("Archivo", str(report.source))
    kv("Páginas", len(report.pages))
    kv("OCR aplicado", "Sí" if report.ocr_applied else "No")
    kv("Versión", report.app_version or "")
    kv("Celdas para revisar", report.n_review)
    kv("Modo", _MODE_LABELS.get(report.mode, report.mode))
    if report.mode_reason:
        kv("Motivo del modo", report.mode_reason)
    row += 1

    if layout_groups:
        row = _write_layout_origin(ws, row, layout_groups)

    if sheet_names:
        _write_header_row(ws, row, ["Hoja", "Páginas", "Filas", "Columnas", "Fuente", "Puntaje"])
        row += 1
    for table, sheet in zip(report.tables, sheet_names, strict=False):
        first, last = table.pages
        pages = str(first + 1) if first == last else f"{first + 1}-{last + 1}"
        values: list[CellValue] = [
            sheet,
            pages,
            len(table.rows),
            len(table.header),
            table.source,
            table.quality,
        ]
        for ci, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=ci)
            _set_value(cell, val)
            if ci == 6:
                cell.number_format = "0.000"
        row += 1
    row += 1

    if report.pages:
        _write_header_row(ws, row, ["Página", "Tipo", "Caracteres", "Imagen %", "Rotación"])
        row += 1
        for page in report.pages:
            vals: list[CellValue] = [
                page.index + 1,
                _PAGE_KIND_LABELS.get(page.kind, page.kind),
                page.n_chars,
                round(page.image_coverage * 100, 1),
                page.rotation,
            ]
            for ci, val in enumerate(vals, start=1):
                _set_value(ws.cell(row=row, column=ci), val)
            row += 1
        row += 1

    if report.timings:
        _write_header_row(ws, row, ["Etapa", "Segundos"])
        row += 1
        for stage, seconds in report.timings.items():
            _set_value(ws.cell(row=row, column=1), _STAGE_LABELS.get(stage, stage))
            cell = ws.cell(row=row, column=2)
            _set_value(cell, round(float(seconds), 3))
            cell.number_format = "0.000"
            row += 1
        row += 1

    if report.warnings:
        _write_header_row(ws, row, ["Avisos"])
        row += 1
        for warning in report.warnings:
            _set_value(ws.cell(row=row, column=1), warning)
            row += 1

    for letter, width in (("A", 28), ("B", 36), ("C", 12), ("D", 12), ("E", 16), ("F", 10)):
        ws.column_dimensions[letter].width = width
    ws.sheet_state = "visible" if opts.show_origin_sheet else "hidden"
    return ws


# --------------------------------------------------------------------------- API


def build_workbook(
    report: ConversionReport,
    opts: Options,
    layout_reports: Sequence[tuple[str, ConversionReport]] = (),
) -> Workbook:
    """Arma el libro en memoria (sin guardarlo). `write_workbook` lo guarda.

    ``report.mode == "layout"``: una hoja por página (``write_layout_sheets``), no tablas/texto/revisión.
    ``layout_reports`` (lotes con un solo libro): reportes en modo layout que se suman, con su prefijo
    (nombre del PDF), a las tablas del reporte combinado.
    """
    wb = Workbook()
    default = wb.active
    if default is not None:
        wb.remove(default)
    wb.properties.creator = f"Conversor PDF a Excel {report.app_version}".rstrip()
    wb.properties.title = Path(report.source).name

    used: set[str] = set()
    sheet_names: list[str] = []
    layout_groups: list[tuple[ConversionReport, list[str]]] = []
    layout_only = report.mode == "layout" and bool(report.layout_pages)
    if layout_only:
        names = write_layout_sheets(wb, report, opts, used)
        layout_groups.append((report, names))
    else:
        for i, table in enumerate(report.tables, start=1):
            name = _sheet_title_for(table, i, opts, used)
            sheet_names.append(name)
            ws = wb.create_sheet(name)
            _write_table(ws, table)

        if any(t.review_cells for t in report.tables):
            _write_review_sheet(wb, report, sheet_names)

        if not report.tables and not layout_reports:
            # Sin tablas, el texto es lo único visible: va primero y se llama "Texto".
            _write_text_sheet(wb, report.text_items, SHEET_TEXT)
        elif opts.include_text and report.text_items:
            _write_text_sheet(wb, report.text_items, SHEET_DOC_DATA)

    for prefix, sub in layout_reports:
        layout_groups.append((sub, write_layout_sheets(wb, sub, opts, used, prefix)))

    _write_origin_sheet(wb, report, opts, sheet_names, layout_groups)
    wb.active = 0
    return wb


def write_workbook(
    report: ConversionReport,
    dst: Path,
    opts: Options,
    layout_reports: Sequence[tuple[str, ConversionReport]] = (),
) -> None:
    """Escribe el `.xlsx` en `dst` (crea la carpeta si falta). Errores de disco -> `UserError`."""
    dst = Path(dst)
    wb = build_workbook(report, opts, layout_reports)
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        wb.save(dst)
    except UserError:
        raise
    except OSError as exc:
        raise from_exception(exc, dst) from exc


__all__ = [
    "NO_TEXT_MESSAGE",
    "OCR_REVIEW_MESSAGE",
    "SHEET_DOC_DATA",
    "SHEET_ORIGIN",
    "SHEET_REVIEW",
    "SHEET_TEXT",
    "build_workbook",
    "default_number_format",
    "format_preview_value",
    "write_workbook",
]
