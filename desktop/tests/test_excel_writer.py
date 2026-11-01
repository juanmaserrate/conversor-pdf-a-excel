"""Tests del writer de Excel: relectura con openpyxl de un ConversionReport sintético (sin pymupdf)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from pdf2excel.errors import UserError
from pdf2excel.excel import styles
from pdf2excel.excel.naming import RESERVED_SHEET_NAMES, safe_sheet_name, unique_output_path
from pdf2excel.excel.writer import (
    NO_TEXT_MESSAGE,
    OCR_REVIEW_MESSAGE,
    format_preview_value,
    write_workbook,
)
from pdf2excel.models import (
    ColumnInfo,
    ConversionReport,
    Options,
    PageInfo,
    TextItem,
    TypedTable,
)

# --------------------------------------------------------------------------- fixtures


def _strip_trailing_none(values: list[object]) -> list[object]:
    while values and values[-1] is None:
        values = values[:-1]
    return values


HEADER = ["Fecha", "Cantidad", "Importe", "Total", "Descuento", "Comprobante", "Descripción"]
COLUMNS = [
    ColumnInfo(kind="date", number_format=styles.FMT_DATE, header="Fecha"),
    ColumnInfo(kind="number", locale="ar", number_format=styles.FMT_INT, header="Cantidad"),
    ColumnInfo(kind="number", locale="ar", number_format=styles.FMT_DEC, header="Importe"),
    ColumnInfo(kind="number", locale="ar", number_format=styles.FMT_MONEY, header="Total"),
    ColumnInfo(kind="percent", number_format=styles.FMT_PCT, header="Descuento"),
    ColumnInfo(kind="id", header="Comprobante"),
    ColumnInfo(kind="text", header="Descripción"),
]
ROWS: list[list[object]] = [
    [date(2026, 10, 15), 12, 1234.5, 14814.25, 0.21, "778563", "Fernet Branca 750 ml"],
    [date(2026, 10, 15), 6, 350.0, "s/d", 0.105, "0001-00004521", "Gin Beefeater"],
    [date(2026, 10, 16), 24, 980.75, 23538.0, 0.0, "778564", None],
    [date(2026, 10, 16), None, 1200.0, 1200.0, None, "A-100", "Vermut Cinzano (sin stock)"],
    [date(2026, 10, 17), 1, 15.0, 15.0, 0.5, "778565", "=SUMA(A1:A3) no es fórmula"],
    [None, None, None, 39567.0, None, None, "TOTAL"],
]


def make_table_1() -> TypedTable:
    return TypedTable(
        name="Bebidas con alcohol",
        header=HEADER,
        rows=[list(r) for r in ROWS],  # type: ignore[arg-type]
        columns=COLUMNS,
        review_cells={(1, 3)},
        total_rows={5},
        pages=(0, 1),
        source="hybrid_union",
        quality=0.93,
    )


def make_table_2() -> TypedTable:
    return TypedTable(
        name="",
        header=["Código", "Stock"],
        rows=[["X1", 24], ["X2", 0], ["X3", 112]],
        columns=[
            ColumnInfo(kind="id", header="Código"),
            ColumnInfo(kind="number", locale="ar", header="Stock"),
        ],
        pages=(1, 1),
        source="hybrid",
        quality=0.8,
    )


def make_text_items() -> list[TextItem]:
    return [
        TextItem(page=0, kind="kv", key="Cliente", value="Bar El Catorce"),
        TextItem(page=0, kind="kv", key="Fecha", value="15/10/2026"),
        TextItem(page=0, kind="kv", key="CUIT", value="20-12345678-9"),
        TextItem(
            page=1, kind="text", key=None, value="Mercadería entregada en depósito. Firma y aclaración."
        ),
    ]


def make_report(tmp_path: Path, **overrides: object) -> ConversionReport:
    data: dict[str, object] = {
        "source": tmp_path / "remito.pdf",
        "pages": [
            PageInfo(index=0, kind="digital", n_chars=1200, image_coverage=0.0),
            PageInfo(index=1, kind="digital", n_chars=800, image_coverage=0.1),
        ],
        "tables": [make_table_1(), make_table_2()],
        "text_items": make_text_items(),
        "timings": {"classify": 0.01, "tables": 0.4, "typing": 0.02, "excel": 0.1},
        "ocr_applied": False,
        "app_version": "0.1.0",
    }
    data.update(overrides)
    return ConversionReport(**data)  # type: ignore[arg-type]


@pytest.fixture
def workbook_path(tmp_path: Path) -> Path:
    dst = tmp_path / "salida" / "remito.xlsx"
    write_workbook(make_report(tmp_path), dst, Options())
    assert dst.exists()
    return dst


@pytest.fixture
def wb(workbook_path: Path) -> Workbook:
    return load_workbook(workbook_path)


# --------------------------------------------------------------------------- hojas y propiedades


def test_sheet_names_and_order(wb: Workbook) -> None:
    assert wb.sheetnames == ["Bebidas con alcohol", "Tabla 2", "Revisión", "Datos del documento", "Origen"]
    assert wb.active.title == "Bebidas con alcohol"


def test_workbook_properties(wb: Workbook) -> None:
    assert wb.properties.creator == "Conversor PDF a Excel 0.1.0"
    assert wb.properties.title == "remito.pdf"


# --------------------------------------------------------------------------- tipos reales


def test_header_row_styled(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    assert [c.value for c in ws[1]] == HEADER
    head = ws["A1"]
    assert head.font.bold is True
    assert head.font.color.rgb.endswith(styles.HEADER_FONT_COLOR)
    assert head.fill.fgColor.rgb.endswith(styles.HEADER_FILL)
    assert head.alignment.horizontal == "center" and head.alignment.wrap_text is True
    assert head.border.left.style == "thin"


def test_date_cell_is_datetime_with_format(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    assert isinstance(ws["A2"].value, datetime)
    assert ws["A2"].value == datetime(2026, 10, 15)
    assert ws["A2"].number_format == "dd/mm/yyyy"
    assert ws["A2"].alignment.horizontal == "right"


def test_numeric_cells_types_and_formats(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    # entero
    assert ws["B2"].value == 12 and isinstance(ws["B2"].value, int)
    assert ws["B2"].number_format == "#,##0"
    # importe
    assert ws["C2"].value == 1234.5 and isinstance(ws["C2"].value, float)
    assert ws["C2"].number_format == "#,##0.00"
    # moneda
    assert ws["D2"].value == 14814.25 and isinstance(ws["D2"].value, float)
    assert ws["D2"].number_format == '"$" #,##0.00'
    # porcentaje
    assert ws["E2"].value == 0.21
    assert ws["E2"].number_format == "0.00%"
    # alineación de números a la derecha, texto a la izquierda, vertical arriba
    assert ws["C2"].alignment.horizontal == "right" and ws["C2"].alignment.vertical == "top"
    assert ws["G2"].alignment.horizontal == "left"


def test_identifier_stays_string(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    assert ws["F2"].value == "778563" and isinstance(ws["F2"].value, str)
    assert ws["F3"].value == "0001-00004521"
    assert ws["F5"].value == "A-100"


def test_empty_cells_are_none(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    assert ws["G4"].value is None
    assert ws["B5"].value is None
    assert ws["A7"].value is None


def test_no_string_in_number_columns_except_review(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    table = make_table_1()
    for c, col in enumerate(table.columns):
        if col.kind not in ("number", "percent", "date"):
            continue
        for r in range(len(table.rows)):
            cell = ws.cell(row=r + 2, column=c + 1)
            if (r, c) in table.review_cells:
                assert isinstance(cell.value, str)
            else:
                assert not isinstance(cell.value, str), f"{cell.coordinate} = {cell.value!r}"


def test_review_cell_fill_and_comment(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    cell = ws["D3"]  # review_cells {(1, 3)} -> fila 3, columna D
    assert cell.value == "s/d"
    assert cell.fill.fgColor.rgb.endswith(styles.REVIEW_FILL)
    assert cell.comment is not None
    assert cell.comment.text == styles.review_comment_text
    assert cell.comment.author == "Conversor"
    # Una celda normal no tiene relleno amarillo ni comentario.
    assert ws["D2"].comment is None
    assert ws["D2"].fill.fill_type in (None, "none")


def test_total_row_bold_with_top_border(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    total = ws["G7"]
    assert total.value == "TOTAL"
    assert total.font.bold is True
    assert total.border.top.style == "medium"
    assert ws["G6"].font.bold is not True
    assert ws["G6"].border.top.style == "thin"


def test_freeze_panes_and_autofilter(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:G7"
    assert ws.auto_filter.ref == ws.dimensions
    ws2: Worksheet = wb["Tabla 2"]
    assert ws2.freeze_panes == "A2"
    assert ws2.auto_filter.ref == "A1:B4"


def test_no_formulas_ever(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    cell = ws["G6"]
    assert cell.value == "=SUMA(A1:A3) no es fórmula"
    assert cell.data_type == "s"
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for c in row:
                assert c.data_type != "f", c.coordinate


def test_column_widths(wb: Workbook) -> None:
    ws: Worksheet = wb["Bebidas con alcohol"]
    widths = {letter: ws.column_dimensions[letter].width for letter in "ABCDEFG"}
    assert widths["A"] == 12  # fechas: 10 caracteres + 2
    assert widths["B"] == len("Cantidad") + 2  # el encabezado manda cuando los datos son cortos
    assert widths["E"] == len("Descuento") + 2
    assert widths["G"] == len("Vermut Cinzano (sin stock)") + 2
    for w in widths.values():
        assert styles.MIN_COL_WIDTH <= w <= styles.MAX_COL_WIDTH


def test_long_text_wraps_and_caps_width(tmp_path: Path) -> None:
    long = "palabra " * 20
    table = TypedTable(
        name="Larga",
        header=["Texto"],
        rows=[[long.strip()]],
        columns=[ColumnInfo(kind="text")],
    )
    dst = tmp_path / "larga.xlsx"
    write_workbook(make_report(tmp_path, tables=[table], text_items=[]), dst, Options())
    ws = load_workbook(dst)["Larga"]
    assert ws.column_dimensions["A"].width == styles.MAX_COL_WIDTH
    assert ws["A2"].alignment.wrap_text is True


def test_number_column_without_format_gets_default(wb: Workbook) -> None:
    ws: Worksheet = wb["Tabla 2"]
    assert ws["B2"].value == 24 and ws["B2"].number_format == "#,##0"


# --------------------------------------------------------------------------- hoja Revisión


def test_review_sheet_contents(wb: Workbook) -> None:
    ws: Worksheet = wb["Revisión"]
    assert [c.value for c in ws[1]] == ["Hoja", "Celda", "Valor", "Motivo"]
    assert ws["A1"].font.bold is True
    assert [c.value for c in ws[2]] == [
        "Bebidas con alcohol",
        "D3",
        "s/d",
        "No se pudo interpretar como número",
    ]
    assert ws.max_row == 2


def test_review_sheet_ocr_notice_first(tmp_path: Path) -> None:
    dst = tmp_path / "ocr.xlsx"
    write_workbook(make_report(tmp_path, ocr_applied=True), dst, Options())
    ws = load_workbook(dst)["Revisión"]
    assert ws["A2"].value == OCR_REVIEW_MESSAGE
    assert ws["B3"].value == "D3"


def test_no_review_sheet_without_review_cells(tmp_path: Path) -> None:
    dst = tmp_path / "limpio.xlsx"
    write_workbook(make_report(tmp_path, tables=[make_table_2()]), dst, Options())
    names = load_workbook(dst).sheetnames
    assert "Revisión" not in names
    assert names == ["Tabla 1", "Datos del documento", "Origen"]


# --------------------------------------------------------------------------- hoja Datos del documento / Texto


def test_document_data_sheet(wb: Workbook) -> None:
    ws: Worksheet = wb["Datos del documento"]
    assert [c.value for c in ws[1]] == ["Campo", "Valor", "Página"]
    assert ws["A2"].value == "Cliente" and ws["B2"].value == "Bar El Catorce" and ws["C2"].value == 1
    # La fecha del bloque kv se tipa.
    assert ws["A3"].value == "Fecha"
    assert isinstance(ws["B3"].value, datetime) and ws["B3"].value == datetime(2026, 10, 15)
    assert ws["B3"].number_format == "dd/mm/yyyy"
    # El CUIT queda texto.
    assert ws["B4"].value == "20-12345678-9"
    # Fila en blanco y después el texto libre.
    assert ws["A5"].value is None
    assert [c.value for c in ws[6]][:2] == ["Página", "Texto"]
    assert ws["A7"].value == 2
    assert ws["B7"].value.startswith("Mercadería entregada")
    assert ws["B7"].alignment.wrap_text is True
    assert ws.column_dimensions["B"].width == styles.TEXT_COL_WIDTH


def test_include_text_false_skips_document_sheet(tmp_path: Path) -> None:
    dst = tmp_path / "sin_texto.xlsx"
    write_workbook(make_report(tmp_path), dst, Options(include_text=False))
    assert "Datos del documento" not in load_workbook(dst).sheetnames


def test_no_tables_gives_texto_first(tmp_path: Path) -> None:
    dst = tmp_path / "solo_texto.xlsx"
    write_workbook(make_report(tmp_path, tables=[]), dst, Options())
    wb = load_workbook(dst)
    assert wb.sheetnames[0] == "Texto"
    assert wb.active.title == "Texto"
    assert "Revisión" not in wb.sheetnames and "Datos del documento" not in wb.sheetnames
    ws = wb["Texto"]
    assert ws["A2"].value == "Cliente"


def test_no_tables_no_text_writes_placeholder(tmp_path: Path) -> None:
    dst = tmp_path / "vacio.xlsx"
    write_workbook(make_report(tmp_path, tables=[], text_items=[]), dst, Options())
    wb = load_workbook(dst)
    assert wb.sheetnames == ["Texto", "Origen"]
    ws = wb["Texto"]
    assert [c.value for c in ws[1]] == ["Página", "Texto"]
    assert ws["B2"].value == NO_TEXT_MESSAGE
    assert wb["Texto"].sheet_state == "visible"


# --------------------------------------------------------------------------- hoja Origen


def test_origin_sheet_hidden_by_default(wb: Workbook) -> None:
    ws: Worksheet = wb["Origen"]
    assert ws.sheet_state == "hidden"
    assert ws["A1"].value == "Archivo" and ws["B1"].value.endswith("remito.pdf")
    assert ws["A2"].value == "Páginas" and ws["B2"].value == 2
    assert ws["A3"].value == "OCR aplicado" and ws["B3"].value == "No"
    assert ws["A4"].value == "Versión" and ws["B4"].value == "0.1.0"
    rows = [_strip_trailing_none([c.value for c in row]) for row in ws.iter_rows()]
    header: list[object] = ["Hoja", "Páginas", "Filas", "Columnas", "Fuente", "Puntaje"]
    idx = rows.index(header)
    assert rows[idx + 1] == ["Bebidas con alcohol", "1-2", 6, 7, "hybrid_union", 0.93]
    assert rows[idx + 2] == ["Tabla 2", "2", 3, 2, "hybrid", 0.8]
    assert ["Etapa", "Segundos"] in rows
    assert ["Buscar tablas", 0.4] in rows


def test_origin_sheet_visible_with_option(tmp_path: Path) -> None:
    dst = tmp_path / "origen.xlsx"
    write_workbook(make_report(tmp_path), dst, Options(show_origin_sheet=True))
    assert load_workbook(dst)["Origen"].sheet_state == "visible"


# --------------------------------------------------------------------------- opciones


def test_one_sheet_per_page_suffix(tmp_path: Path) -> None:
    dst = tmp_path / "paginas.xlsx"
    write_workbook(make_report(tmp_path), dst, Options(one_sheet_per_page=True))
    names = load_workbook(dst).sheetnames
    assert names[0] == "Bebidas con alcohol p.1"
    assert names[1] == "Tabla 2 p.2"


def test_table_named_like_reserved_sheet_is_renamed(tmp_path: Path) -> None:
    table = make_table_2()
    table.name = "Origen"
    dst = tmp_path / "reservado.xlsx"
    write_workbook(make_report(tmp_path, tables=[table], text_items=[]), dst, Options())
    names = load_workbook(dst).sheetnames
    assert names == ["Origen (2)", "Origen"]


def test_write_creates_missing_folder(tmp_path: Path) -> None:
    dst = tmp_path / "a" / "b" / "c" / "x.xlsx"
    write_workbook(make_report(tmp_path), dst, Options())
    assert dst.exists()


def test_write_error_becomes_user_error(tmp_path: Path) -> None:
    blocker = tmp_path / "carpeta.xlsx"
    blocker.mkdir()  # guardar sobre un directorio falla con IsADirectoryError/PermissionError
    with pytest.raises(UserError):
        write_workbook(make_report(tmp_path), blocker, Options())


# --------------------------------------------------------------------------- naming


def test_unique_output_path(tmp_path: Path) -> None:
    src = tmp_path / "x.pdf"
    src.touch()
    assert unique_output_path(src, None) == tmp_path / "x.xlsx"
    (tmp_path / "x.xlsx").touch()
    assert unique_output_path(src, None) == tmp_path / "x (2).xlsx"
    (tmp_path / "x (2).xlsx").touch()
    assert unique_output_path(src, None) == tmp_path / "x (3).xlsx"


def test_unique_output_path_with_output_dir(tmp_path: Path) -> None:
    src = tmp_path / "docs" / "remito.pdf"
    out = tmp_path / "salida"
    assert unique_output_path(src, out) == out / "remito.xlsx"


def test_safe_sheet_name_strips_forbidden_chars() -> None:
    used: set[str] = set()
    name = safe_sheet_name("1. Bebidas: con/alcohol [2026]*?", 1, used)
    assert len(name) <= 31
    assert not any(ch in name for ch in "[]:*?/\\")
    assert name == "1. Bebidas con alcohol 2026"
    assert name in used


def test_safe_sheet_name_dedupes() -> None:
    assert safe_sheet_name(None, 3, {"Tabla 3"}) == "Tabla 3 (2)"
    used = {"Tabla 3", "Tabla 3 (2)"}
    assert safe_sheet_name(None, 3, used) == "Tabla 3 (3)"
    # Sin distinguir mayúsculas.
    assert safe_sheet_name("ventas", 1, {"VENTAS"}) == "ventas (2)"


def test_safe_sheet_name_truncates_with_suffix() -> None:
    long_title = "Listado de precios mayoristas vigentes 2026"
    used: set[str] = set()
    first = safe_sheet_name(long_title, 1, used)
    second = safe_sheet_name(long_title, 2, used)
    assert len(first) == 31 and len(second) <= 31
    assert second.endswith(" (2)") and second != first


def test_safe_sheet_name_never_reserved() -> None:
    for reserved in RESERVED_SHEET_NAMES:
        assert safe_sheet_name(reserved, 1, set()) == f"{reserved} (2)"
    assert safe_sheet_name("   ", 7, set()) == "Tabla 7"


# --------------------------------------------------------------------------- format_preview_value


@pytest.mark.parametrize(
    ("value", "col", "expected"),
    [
        (1234.5, ColumnInfo("number", "ar", "#,##0.00"), "1.234,50"),
        (-1234.5, ColumnInfo("number", "ar", "#,##0.00"), "-1.234,50"),
        (1234, ColumnInfo("number", "ar", "#,##0"), "1.234"),
        (1234.5, ColumnInfo("number", "ar", "#,##0"), "1.235"),
        (1234.5, ColumnInfo("number", "ar", '"$" #,##0.00'), "$ 1.234,50"),
        (0.21, ColumnInfo("percent", None, "0.00%"), "21,00 %"),
        (0.105, ColumnInfo("percent"), "10,50 %"),
        (date(2026, 10, 15), ColumnInfo("date", None, "dd/mm/yyyy"), "15/10/2026"),
        (datetime(2026, 1, 7, 10, 30), ColumnInfo("date"), "07/01/2026"),
        (None, ColumnInfo("number", "ar", "#,##0.00"), ""),
        ("s/d", ColumnInfo("number", "ar", "#,##0.00"), "s/d"),
        ("778563", ColumnInfo("id"), "778563"),
        (42, ColumnInfo("number"), "42"),
        (42.0, ColumnInfo("number"), "42"),
        (42.5, ColumnInfo("number"), "42,50"),
    ],
)
def test_format_preview_value(value: object, col: ColumnInfo, expected: str) -> None:
    assert format_preview_value(value, col) == expected  # type: ignore[arg-type]
