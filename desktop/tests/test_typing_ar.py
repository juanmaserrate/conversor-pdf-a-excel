"""Tests del tipado argentino por columna (core/typing_ar.py, DESIGN.md §5.6)."""

from __future__ import annotations

from datetime import date

import pytest

from pdf2excel.core import typing_ar
from pdf2excel.core.typing_ar import (
    classify_column,
    infer_column_locale,
    is_identifier_header,
    normalize_number_text,
    parse_date,
    parse_number,
    parse_percent,
    type_table,
)
from pdf2excel.models import ColumnInfo, RawTable, TypedTable

# --------------------------------------------------------------------------- parse_number


@pytest.mark.parametrize(
    ("text", "locale", "expected"),
    [
        # AR básicos
        ("1.234,56", "ar", 1234.56),
        ("$ 1.234", "ar", 1234),
        ("-62.175,43", "ar", -62175.43),
        ("1.234,56-", "ar", -1234.56),
        ("(1.234,56)", "ar", -1234.56),
        ("12,90", "ar", 12.9),
        ("1.234", "ar", 1234),
        ("1.234.567", "ar", 1234567),
        ("1.234.567,89", "ar", 1234567.89),
        ("0,5", "ar", 0.5),
        ("1234,5", "ar", 1234.5),
        ("980,00", "ar", 980.0),
        ("$ 1.234,56", "ar", 1234.56),
        ("$1.234,56", "ar", 1234.56),
        ("U$S 1.234,50", "ar", 1234.5),
        ("USD 250", "ar", 250),
        ("ARS 500", "ar", 500),
        ("ars 1.500,25", "ar", 1500.25),
        ("+12", "ar", 12),
        ("- 5", "ar", -5),
        ("−7", "ar", -7),  # signo menos unicode
        ("($ 350,00)", "ar", -350.0),
        ("$ -1.000", "ar", -1000),
        ("1 234,50", "ar", 1234.5),  # NBSP como separador de miles
        ("  42  ", "ar", 42),
        ("0", "ar", 0),
        ("24", "ar", 24),
        ("112", "ar", 112),
        # US básicos
        ("350.00", "us", 350.0),
        ("1,234.56", "us", 1234.56),
        ("1,234,567.89", "us", 1234567.89),
        ("1.234", "us", 1.234),
        ("12.5", "us", 12.5),
        ("1,234", "us", 1234),
        ("$ 1,234.00", "us", 1234.0),
        ("(1,234.56)", "us", -1234.56),
        ("1,234.56-", "us", -1234.56),
        ("42", "us", 42),
        # No parsean
        ("350.00", "ar", None),  # formato US en columna AR: queda para revisar
        ("1.234,56", "us", None),
        ("abc", "ar", None),
        ("s/d", "ar", None),
        ("", "ar", None),
        ("   ", "ar", None),
        ("-", "ar", None),
        ("$", "ar", None),
        ("()", "ar", None),
        ("21%", "ar", None),
        ("A-100", "ar", None),
        ("20-12345678-9", "ar", None),
        ("15/10/2026", "ar", None),
        ("1.2.3", "ar", None),
        ("12.34.56", "ar", None),
        ("1,23,45", "us", None),
        ("1.2345", "ar", None),
        ("1,,5", "ar", None),
    ],
)
def test_parse_number(text: str, locale: str, expected: int | float | None) -> None:
    result = parse_number(text, locale)  # type: ignore[arg-type]
    assert result == expected
    if expected is not None:
        assert type(result) is type(expected)


def test_parse_number_int_vs_float_type() -> None:
    assert isinstance(parse_number("$ 1.234", "ar"), int)
    assert isinstance(parse_number("1.234,00", "ar"), float)
    assert isinstance(parse_number("350.00", "us"), float)
    assert isinstance(parse_number("1,234", "us"), int)


# --------------------------------------------------------------------------- normalize


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("1.234,56", ("1.234,56", False, False)),
        ("$ 1.234,56", ("1.234,56", False, True)),
        ("-62.175,43", ("62.175,43", True, False)),
        ("1.234,56-", ("1.234,56", True, False)),
        ("(1.234,56)", ("1.234,56", True, False)),
        ("U$S 1.234", ("1.234", False, True)),
        ("USD 1,234.00", ("1,234.00", False, True)),
        ("ARS 5", ("5", False, True)),
        ("+5", ("5", False, False)),
        ("($ 5)", ("5", True, True)),
        ("1 234", ("1234", False, False)),
        ("", ("", False, False)),
    ],
)
def test_normalize_number_text(text: str, expected: tuple[str, bool, bool]) -> None:
    assert normalize_number_text(text) == expected


# --------------------------------------------------------------------------- parse_percent


@pytest.mark.parametrize(
    ("text", "locale", "expected"),
    [
        ("21%", "ar", 0.21),
        ("21 %", "ar", 0.21),
        ("12,5 %", "ar", 0.125),
        ("12.5%", "us", 0.125),
        ("-3%", "ar", -0.03),
        ("+10,5%", "ar", 0.105),
        ("100%", "ar", 1.0),
        ("0%", "ar", 0.0),
        ("21", "ar", None),
        ("%", "ar", None),
        ("abc%", "ar", None),
        ("1.234,5%", "ar", None),
        ("", "ar", None),
    ],
)
def test_parse_percent(text: str, locale: str, expected: float | None) -> None:
    result = parse_percent(text, locale)  # type: ignore[arg-type]
    if expected is None:
        assert result is None
    else:
        assert result == pytest.approx(expected)
        assert isinstance(result, float)


# --------------------------------------------------------------------------- parse_date


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("15/10/2026", date(2026, 10, 15)),
        ("7/1/25", date(2025, 1, 7)),
        ("01-12-99", date(1999, 12, 1)),
        ("1.2.2024", date(2024, 2, 1)),
        ("31/12/69", date(2069, 12, 31)),
        ("31/12/70", date(1970, 12, 31)),
        ("29/02/2024", date(2024, 2, 29)),
        (" 15/10/2026 ", date(2026, 10, 15)),
        ("32/13/2026", None),
        ("29/02/2023", None),
        ("00/01/2024", None),
        ("2026/10/15", None),
        ("15/10/202", None),
        ("15/10", None),
        ("15 de octubre", None),
        ("", None),
        ("1.234,56", None),
    ],
)
def test_parse_date(text: str, expected: date | None) -> None:
    assert parse_date(text) == expected


# --------------------------------------------------------------------------- infer_column_locale


@pytest.mark.parametrize(
    ("values", "expected"),
    [
        (["1.234"], "ar"),  # ambiguo solo: sin evidencia → AR (tabla de una fila con "1.500")
        (["1,234"], "ar"),
        (["1,234.56", "350.00"], "us"),
        (["1.234,56", "980,00"], "ar"),
        (["24", "0", "112", "s/d"], "ar"),  # enteros: candidatos sin voto → empate → AR
        (["s/d", "abc", ""], None),
        ([], None),
        (["15/10/2026", "16/10/2026"], None),  # fechas no son candidatos
        (["1.234,56", "350.00", "1.200,00"], "ar"),
        (["350.00", "1,234.56", "12,90"], "us"),
        (["350.00", "12,90"], "ar"),  # empate → AR
        (["1.234", "5.678"], "ar"),  # ≥ 2 ambiguos y nada más → AR (DESIGN §10)
        (["1.234", "24"], "ar"),
        (["1.234.567"], "ar"),  # dos puntos: solo AR lo explica
        (["1,234,567"], "us"),
        (["12.5"], "us"),
        (["$ 1.234,56", "U$S 980,00"], "ar"),
        (["21%", "12,5%"], "ar"),
        (["12.5%", "7.25%"], "us"),
    ],
)
def test_infer_column_locale(values: list[str], expected: str | None) -> None:
    assert infer_column_locale(values) == expected


# --------------------------------------------------------------------------- identifier headers


@pytest.mark.parametrize(
    ("header", "expected"),
    [
        ("Comprobante", True),
        ("Código", True),
        ("codigo", True),
        ("CUIT", True),
        ("CUIL", True),
        ("CBU", True),
        ("DNI", True),
        ("N°", True),
        ("Nº", True),
        ("Nro", True),
        ("Nro.", True),
        ("Referencia", True),
        ("ID", True),
        ("Artículo", True),
        ("Articulo", True),
        ("SKU", True),
        ("Factura", True),
        ("Remito", True),
        ("Legajo", True),
        ("N° de factura", True),
        ("Nro. comprobante", True),
        ("Factura N°", True),
        ("Código de artículo", True),
        ("Artículos", True),
        ("Remitos", True),
        ("Nombre", False),
        ("Importe facturado", False),  # "factura" era substring
        ("Total facturado", False),
        ("Cantidad facturada", False),
        ("Turno", False),  # terminan en "no": no son "N°"
        ("Vino", False),
        ("Destino", False),
        ("Bono", False),
        ("Abono", False),
        ("Circuito", False),  # "cuit" era substring
        ("Precio unitario", False),
        ("Importe neto", False),
        ("Identificación", False),  # "id" solo como palabra entera
        ("Cantidad", False),
        ("Importe", False),
        ("Fecha", False),
        ("Descripción", False),
        ("Ciudad", False),
        ("", False),
    ],
)
def test_is_identifier_header(header: str, expected: bool) -> None:
    assert is_identifier_header(header) is expected


# --------------------------------------------------------------------------- classify_column


def test_classify_comprobante_is_id() -> None:
    col = classify_column("Comprobante", ["778563", "778564"])
    assert col.kind == "id"
    assert col.number_format is None
    assert col.header == "Comprobante"


def test_classify_cuit_values_is_id_without_header() -> None:
    col = classify_column("Cliente", ["20-12345678-9", "30-71234567-0"])
    assert col.kind == "id"


def test_classify_cuit_80_percent() -> None:
    values = ["20-12345678-9"] * 4 + ["s/d"]
    assert classify_column("Dato", values).kind == "id"
    values = ["20-12345678-9"] * 3 + ["s/d", "abc"]
    assert classify_column("Dato", values).kind == "text"


def test_classify_same_length_digits_is_id() -> None:
    assert classify_column("Dato", ["00123", "00456", "00789"]).kind == "id"
    # Longitudes distintas: no es un código con ceros, es una cantidad.
    assert classify_column("Dato", ["12345", "123456"]).kind == "number"
    # Menos de 5 dígitos: números.
    assert classify_column("Dato", ["1234", "5678"]).kind == "number"


def test_classify_alphanumeric_code_is_text() -> None:
    col = classify_column("Cod", ["A-100", "A-101"])
    assert col.kind == "text"
    assert col.locale is None
    assert col.number_format is None


def test_classify_stock_with_sd_is_number() -> None:
    col = classify_column("Stock", ["24", "0", "112", "s/d"])
    assert col.kind == "number"
    assert col.locale == "ar"
    assert col.number_format == "#,##0"


def test_classify_currency_format() -> None:
    col = classify_column("Importe", ["$ 1.234,56", "$ 980,00"])
    assert col.kind == "number"
    assert col.locale == "ar"
    assert col.number_format == '"$" #,##0.00'


def test_classify_float_format() -> None:
    col = classify_column("Precio", ["1.234,56", "980,00", "12,90"])
    assert col.number_format == "#,##0.00"


def test_classify_us_locale() -> None:
    col = classify_column("Amount", ["1,234.56", "350.00"])
    assert col.kind == "number"
    assert col.locale == "us"
    assert col.number_format == "#,##0.00"


def test_classify_date_tolerates_loose_text() -> None:
    values = [f"{d:02d}/10/2026" for d in range(1, 11)] + ["Saldo final:"]
    col = classify_column("Fecha", values)
    assert col.kind == "date"
    assert col.number_format == "dd/mm/yyyy"


def test_classify_date_below_90_percent_is_not_date() -> None:
    values = ["15/10/2026", "16/10/2026", "Saldo final:"]
    assert classify_column("Fecha", values).kind == "text"


def test_classify_percent() -> None:
    col = classify_column("IVA", ["21%", "10,5 %", "", "27%"])
    assert col.kind == "percent"
    assert col.number_format == "0.00%"


def test_classify_dashes_count_as_empty() -> None:
    """Guiones de 'sin dato' no cuentan en el 50 %: 2 importes + 3 guiones → número AR."""
    col = classify_column("Precio", ["1.234,56", "-", "—", "-", "9.800,00"])
    assert col.kind == "number" and col.locale == "ar" and col.number_format == typing_ar.FLOAT_FORMAT
    assert classify_column("Precio", ["-", "—", "–"]).kind == "text"  # solo guiones: nada que tipar


def test_classify_single_ambiguous_value_is_number_ar() -> None:
    """Una sola fila con '1.500' se tipa igual que dos filas ('1.500', '2.000')."""
    one = classify_column("Precio", ["1.500"])
    two = classify_column("Precio", ["1.500", "2.000"])
    assert one.kind == two.kind == "number" and one.locale == two.locale == "ar"
    assert classify_column("Importe", ["12.000"]).kind == "number"


@pytest.mark.parametrize("text", ["-", "—", "–", "...", " - ", "", None])
def test_is_blank_cell(text: str | None) -> None:
    assert typing_ar.is_blank_cell(text)


@pytest.mark.parametrize("text", ["-5", "s/d", "0", "a-b", "1.234,56-"])
def test_is_blank_cell_false(text: str) -> None:
    assert not typing_ar.is_blank_cell(text)


def test_classify_empty_column_is_text() -> None:
    assert classify_column("Obs", ["", "  ", ""]).kind == "text"


def test_classify_number_below_50_percent_is_text() -> None:
    assert classify_column("Obs", ["12", "a", "b", "c"]).kind == "text"


def test_classify_returns_column_info() -> None:
    assert isinstance(classify_column("x", ["1"]), ColumnInfo)


# --------------------------------------------------------------------------- type_table


def _raw(rows: list[list[str]], **kw: object) -> RawTable:
    base: dict = {
        "page": 0,
        "page_end": 0,
        "bbox": (10.0, 10.0, 500.0, 300.0),
        "rows": rows,
        "source": "hybrid_union",
        "quality": 0.91,
        "title": None,
    }
    base.update(kw)
    return RawTable(**base)


def test_type_table_mixed_columns() -> None:
    raw = _raw(
        [
            ["Comprobante", "Fecha", "Stock", "Importe", "Desc."],
            ["778563", "15/10/2026", "24", "$ 1.234,56", "21%"],
            ["778564", "16/10/2026", "0", "$ 980,00", "10,5%"],
            ["778565", "17/10/2026", "112", "$ 1,00", "0%"],
            ["778566", "18/10/2026", "s/d", "$ 5,00", "5%"],
            ["TOTAL", "", "", "$ 2.220,56", ""],
        ],
        title="Ventas",
        page=2,
        page_end=3,
    )
    t = type_table(raw)
    assert isinstance(t, TypedTable)
    assert t.name == "Ventas"
    assert t.header == ["Comprobante", "Fecha", "Stock", "Importe", "Desc."]
    assert [c.kind for c in t.columns] == ["id", "date", "number", "number", "percent"]
    assert t.rows[0] == ["778563", date(2026, 10, 15), 24, 1234.56, 0.21]
    assert isinstance(t.rows[0][0], str)
    assert t.rows[3][2] == "s/d"
    assert t.review_cells == {(3, 2)}
    assert t.total_rows == {4}
    assert t.rows[4] == ["TOTAL", None, None, 2220.56, None]
    assert t.pages == (2, 3)
    assert t.source == "hybrid_union"
    assert t.quality == pytest.approx(0.91)
    assert t.columns[2].number_format == "#,##0"
    assert t.columns[3].number_format == '"$" #,##0.00'


def test_type_table_stock_column_three_ints_one_review() -> None:
    raw = _raw([["Stock"], ["24"], ["0"], ["112"], ["s/d"]])
    t = type_table(raw)
    assert t.columns[0].kind == "number"
    assert t.rows == [[24], [0], [112], ["s/d"]]
    assert all(isinstance(r[0], int) for r in t.rows[:3])
    assert t.review_cells == {(3, 0)}


def test_type_table_date_column_with_loose_text_in_review() -> None:
    rows = [["Fecha"]] + [[f"{d:02d}/10/2026"] for d in range(1, 11)] + [["Saldo final:"]]
    t = type_table(_raw(rows))
    assert t.columns[0].kind == "date"
    assert t.rows[0][0] == date(2026, 10, 1)
    assert t.rows[10][0] == "Saldo final:"
    assert t.review_cells == {(10, 0)}
    assert t.total_rows == {10}  # "Saldo final:" es fila de total


def test_type_table_pads_ragged_rows_and_empty_cells_are_none() -> None:
    raw = _raw([["A", "B", "C"], ["1", "2"], ["3", "", "4"]])
    t = type_table(raw)
    assert t.header == ["A", "B", "C"]
    assert t.rows == [[1, 2, None], [3, None, 4]]
    assert len(t.columns) == 3
    assert t.review_cells == set()


def test_type_table_dash_cells_are_none_without_review() -> None:
    raw = RawTable(
        page=0,
        page_end=0,
        bbox=(0, 0, 1, 1),
        rows=[
            ["Producto", "Stock", "Precio"],
            ["Coca", "12", "1.234,56"],
            ["Fernet", "-", "—"],
            ["Gin", "3", "9.800,00"],
        ],
        source="hybrid",
        quality=1.0,
    )
    typed = type_table(raw)
    assert [c.kind for c in typed.columns] == ["text", "number", "number"]
    assert typed.rows[1] == ["Fernet", None, None]
    assert typed.rows[0] == ["Coca", 12, 1234.56] and typed.rows[2] == ["Gin", 3, 9800.0]
    assert typed.review_cells == set()
    # En una columna de texto el guión se conserva tal cual (no se inventa nada).
    raw_text = RawTable(0, 0, (0, 0, 1, 1), [["Obs", "Imp"], ["-", "1,00"], ["ok", "2,00"]], "hybrid", 1.0)
    assert type_table(raw_text).rows[0] == ["-", 1.0]


def test_type_table_untyped_keeps_strings() -> None:
    raw = _raw([["Fecha", "Importe"], ["15/10/2026", "$ 1.234,56"], ["Subtotal", "$ 1.234,56"]])
    t = type_table(raw, typed_values=False)
    assert t.rows == [["15/10/2026", "$ 1.234,56"], ["Subtotal", "$ 1.234,56"]]
    assert all(c.kind == "text" for c in t.columns)
    assert [c.header for c in t.columns] == ["Fecha", "Importe"]
    assert t.review_cells == set()
    assert t.total_rows == {1}


def test_type_table_name_empty_when_no_title() -> None:
    assert type_table(_raw([["A"], ["1"]])).name == ""


def test_type_table_total_rows_patterns() -> None:
    rows = [
        ["Detalle", "Importe"],
        ["Producto", "10,00"],
        ["Subtotal", "10,00"],
        ["IVA 21%", "2,10"],
        ["Transporte", "12,10"],
        ["Saldo final", "12,10"],
        ["TOTAL GENERAL", "12,10"],
        ["Totales", "12,10"],
        ["Ivana García", "1,00"],  # no es total
        ["", "Subtotal"],  # primera celda vacía: mira la siguiente no vacía
    ]
    t = type_table(_raw(rows))
    assert t.total_rows == {1, 2, 3, 4, 5, 6, 8}


def test_type_table_empty_rows() -> None:
    t = type_table(_raw([]))
    assert t.header == []
    assert t.rows == []
    assert t.columns == []


def test_type_table_never_invents_numbers() -> None:
    raw = _raw([["Importe"], ["1.234,56"], ["350.00"], ["980,00"]])
    t = type_table(raw)
    assert t.columns[0].locale == "ar"
    assert t.rows[1][0] == "350.00"  # queda como texto, va a revisión
    assert (1, 0) in t.review_cells


def test_module_is_pure() -> None:
    import sys

    assert "pymupdf" not in getattr(typing_ar, "__dict__", {})
    assert not any(m.startswith("PySide6") for m in sys.modules if m in typing_ar.__dict__)
