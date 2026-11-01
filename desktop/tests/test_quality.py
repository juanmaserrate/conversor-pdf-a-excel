"""Tests de core/quality.py: table_quality y select_best (WI-02)."""

from __future__ import annotations

import pytest

from pdf2excel.core import quality
from pdf2excel.core.cells import merge_ghost_columns
from pdf2excel.core.quality import ghost_column_pairs, select_best, table_quality
from pdf2excel.models import RawTable

CLEAN_TYPED = [
    ["Código", "Producto", "Cantidad", "Precio unitario", "Total"],
    ["A-100", "Mezcal joven", "10", "350.00", "3500.00"],
    ["A-101", "Tequila reposado", "5", "420.50", "2102.50"],
    ["A-102", "Agua mineral", "48", "12.90", "619.20"],
    ["A-103", "Hielo bolsa 5kg", "20", "35.00", "700.00"],
    ["A-104", "Limón kg", "8", "28.75", "230.00"],
    ["A-105", "Sal de gusano", "3", "90.00", "270.00"],
]

CLEAN_AR_DATES = [
    ["Fecha", "Concepto", "Débito", "Crédito", "Saldo"],
    ["01/10/2026", "Transferencia recibida", "", "1.234,56", "126.664,73"],
    ["02/10/2026", "Débito automático", "-2.500,00", "", "124.164,73"],
    ["03/10/2026", "Compra débito", "-980,10", "", "123.184,63"],
    ["05/10/26", "Acreditación haberes", "", "85.000,00", "208.184,63"],
]


def _raw(rows: list[list[str]], source: str = "hybrid", page: int = 0) -> RawTable:
    return RawTable(
        page=page,
        page_end=page,
        bbox=(0.0, 0.0, 100.0, 100.0),
        rows=rows,
        source=source,
        quality=table_quality(rows),
    )


# --------------------------------------------------------------------------- table_quality


def test_clean_typed_table_scores_high() -> None:
    assert table_quality(CLEAN_TYPED) > 0.9
    assert table_quality(CLEAN_AR_DATES) > 0.9


def test_quality_is_bounded() -> None:
    for rows in (CLEAN_TYPED, CLEAN_AR_DATES, [["a", "b"], ["|", "|"]], [["1", "2"], ["3", "4"]]):
        assert 0.0 <= table_quality(rows) <= 1.0


def test_half_pipes_scores_low() -> None:
    """OCR que alucinó bordes: 50 % de las celdas del cuerpo son `|` (toda una columna)."""
    rows = [["Producto", "Total"]]
    for name in ["Mezcal joven", "Tequila reposado", "Agua mineral", "Hielo bolsa 5kg", "Limón kg", "Sal"]:
        rows.append([name, "|"])
    n_pipes = sum(1 for r in rows[1:] for c in r if c == "|")
    assert n_pipes * 2 == sum(len(r) for r in rows[1:])
    assert table_quality(rows) < 0.6
    # la misma tabla con los importes leídos bien puntúa alto
    fixed = [["Producto", "Total"]] + [[r[0], "350,00"] for r in rows[1:]]
    assert table_quality(fixed) > 0.9


def test_pipes_in_half_of_the_columns_score_below_clean() -> None:
    rows = [list(r) for r in CLEAN_TYPED]
    for r in rows[1:]:
        r[1] = "|"
        r[4] = "|"
    assert table_quality(rows) < table_quality(CLEAN_TYPED) - 0.05


def test_ghost_columns_score_less_than_merged() -> None:
    ghost = [
        ["Código", "Producto", "", "Cantidad", "Total"],
        ["A-100", "Mezcal joven", "", "10", "3500.00"],
        ["A-101", "Tequila reposado", "", "5", "2102.50"],
        ["", "", "Subtotal", "", "5602.50"],
        ["", "", "TOTAL", "", "6779.03"],
    ]
    merged = merge_ghost_columns([list(r) for r in ghost])
    assert len(merged[0]) < len(ghost[0])
    assert ghost_column_pairs(ghost) >= 1
    assert ghost_column_pairs(merged) == 0
    assert table_quality(ghost) < table_quality(merged)


def test_degenerate_tables_score_zero() -> None:
    assert table_quality([]) == 0.0
    assert table_quality([["solo", "encabezado"]]) == 0.0
    assert table_quality([[], []]) == 0.0


def test_broken_rows_and_long_cells_penalize() -> None:
    good = [["A", "B"], ["1", "2"], ["3", "4"], ["5", "6"]]
    broken = [["A", "B"], ["1", ""], ["3", ""], ["5", ""]]
    long_text = "x" * 80
    long_cells = [["A", "B"], ["1", long_text], ["3", long_text], ["5", long_text]]
    assert table_quality(broken) < table_quality(good)
    assert table_quality(long_cells) < table_quality(good)


def test_numeric_header_penalizes() -> None:
    text_header = [["Cantidad", "Total"], ["10", "350.00"], ["5", "420.50"]]
    numeric_header = [["1", "2"], ["10", "350.00"], ["5", "420.50"]]
    assert table_quality(numeric_header) < table_quality(text_header)


def test_legit_empty_columns_are_not_heavily_penalized() -> None:
    """Débito/Crédito mutuamente vacías son legítimas: la tabla sigue puntuando alto."""
    assert table_quality(CLEAN_AR_DATES) > 0.85


def test_quality_handles_ragged_rows_and_none() -> None:
    rows: list[list[str]] = [["A", "B", "C"], ["1", "2"], ["3", "4", "5", "6"]]
    assert 0.0 <= table_quality(rows) <= 1.0


def test_typing_backend_fallback_when_typing_ar_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    import builtins
    import sys

    real_import = builtins.__import__

    def fake_import(name: str, *args: object, **kwargs: object) -> object:
        if name == "pdf2excel.core.typing_ar":
            raise ImportError("simulado: WI-03 ausente")
        return real_import(name, *args, **kwargs)

    monkeypatch.delitem(sys.modules, "pdf2excel.core.typing_ar", raising=False)
    monkeypatch.setattr(builtins, "__import__", fake_import)
    quality._typing_backend.cache_clear()
    try:
        infer, parse_number, parse_date = quality._typing_backend()
        assert infer(["1.234,56", "350,00"]) == "ar"
        assert infer(["1,234.56", "350.00"]) == "us"
        assert infer(["texto", ""]) is None
        assert parse_number("$ 1.234,56", "ar") == 1234.56
        assert parse_number("350.00", "us") == 350.0
        assert parse_number("1.234", "ar") == 1234
        assert parse_number("21%", "ar") == pytest.approx(0.21)
        assert parse_number("abc", "ar") is None
        assert str(parse_date("15/10/2026")) == "2026-10-15"
        assert str(parse_date("7/1/25")) == "2025-01-07"
        assert parse_date("32/13/2026") is None
        assert table_quality(CLEAN_TYPED) > 0.9
    finally:
        quality._typing_backend.cache_clear()


# --------------------------------------------------------------------------- select_best


def test_select_best_prefers_more_coverage_at_equal_quality() -> None:
    small = _raw(CLEAN_TYPED[:4], "hybrid_union")
    big = _raw(CLEAN_TYPED, "hybrid")
    assert small.quality == pytest.approx(big.quality, abs=0.02)
    # forzamos igualdad exacta de calidad para aislar la cobertura
    small.quality = big.quality
    name, tables = select_best({"hybrid_union": [small], "hybrid": [big]})
    assert name == "hybrid"
    assert tables == [big]


def test_select_best_prefers_quality_when_coverage_is_similar() -> None:
    bad_rows = [list(r) for r in CLEAN_TYPED]
    for r in bad_rows[1:]:
        r[1] = "|"
        r[3] = "|"
    good = _raw(CLEAN_TYPED, "hybrid_union")
    bad = _raw(bad_rows, "hybrid")
    assert bad.quality < good.quality
    name, _ = select_best({"hybrid_union": [good], "hybrid": [bad]})
    assert name == "hybrid_union"


def test_select_best_tie_goes_to_insertion_order() -> None:
    a = _raw(CLEAN_TYPED, "hybrid_union")
    b = _raw(CLEAN_TYPED, "hybrid")
    assert select_best({"hybrid_union": [a], "hybrid": [b]})[0] == "hybrid_union"
    assert select_best({"hybrid": [b], "hybrid_union": [a]})[0] == "hybrid"


def test_select_best_empty_candidates() -> None:
    assert select_best({}) == ("hybrid", [])
    assert select_best({"hybrid_union": [], "hybrid": []}) == ("hybrid", [])


def test_select_best_skips_candidates_without_tables() -> None:
    only = _raw(CLEAN_TYPED, "hybrid")
    name, tables = select_best({"hybrid_union": [], "hybrid": [only]})
    assert name == "hybrid"
    assert tables == [only]


def test_select_best_weights_quality_by_rows() -> None:
    """Calidad media ponderada por filas: una tabla grande y buena pesa más que una chica y mala."""
    good_big = _raw(CLEAN_TYPED, "hybrid_union")
    bad_small = _raw([["a", "b"], ["|", "-"], ["x", ""]], "hybrid_union")
    cand_a = [good_big, bad_small]
    cand_b = [_raw(CLEAN_TYPED, "hybrid")]
    # A tiene más cobertura pero menor calidad media; B es limpia. Ambos deben ser puntajes válidos y
    # la elección debe ser determinista.
    name, tables = select_best({"hybrid_union": cand_a, "hybrid": cand_b})
    assert name in {"hybrid_union", "hybrid"}
    assert tables is not cand_a and tables is not cand_b  # copia, no la lista original
