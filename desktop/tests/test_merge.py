"""Tests del merge multipágina (core/merge.py, DESIGN.md §5.5)."""

from __future__ import annotations

import pytest

from pdf2excel.core.merge import merge_multipage
from pdf2excel.models import RawTable

HEADER = ["Código", "Detalle", "Cantidad", "Importe"]
BBOX = (40.0, 100.0, 550.0, 700.0)


def _table(
    page: int, rows: list[list[str]], bbox: tuple = BBOX, quality: float = 0.9, **kw: object
) -> RawTable:
    base: dict = {
        "page": page,
        "page_end": page,
        "bbox": bbox,
        "rows": rows,
        "source": "hybrid_union",
        "quality": quality,
        "title": None,
    }
    base.update(kw)
    return RawTable(**base)


def _body(start: int, n: int) -> list[list[str]]:
    return [
        [f"A-{start + i:03d}", f"Producto {start + i}", str(i + 1), f"{(i + 1) * 100},00"] for i in range(n)
    ]


# --------------------------------------------------------------------------- Regla 1


def test_rule1_same_header_merges_without_repeating_header() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)], title="Ventas")
    t1 = _table(1, [HEADER, *_body(4, 2)])
    merged = merge_multipage([t0, t1])
    assert len(merged) == 1
    m = merged[0]
    assert m.page == 0
    assert m.page_end == 1
    assert m.rows[0] == HEADER
    assert len(m.rows) == 1 + 3 + 2
    assert sum(1 for r in m.rows if r == HEADER) == 1
    assert m.rows[1:] == _body(1, 3) + _body(4, 2)
    assert m.title == "Ventas"
    assert m.bbox == BBOX


def test_rule1_header_comparison_is_case_and_space_insensitive() -> None:
    h2 = ["CÓDIGO", "detalle ", " Cantidad", "IMPORTE"]
    t0 = _table(0, [HEADER, *_body(1, 2)])
    t1 = _table(1, [h2, *_body(3, 2)], bbox=(45.0, 60.0, 560.0, 300.0))  # bbox distinto no importa
    merged = merge_multipage([t0, t1])
    assert len(merged) == 1
    assert len(merged[0].rows) == 5


def test_rule1_three_pages_chain() -> None:
    tables = [_table(p, [HEADER, *_body(p * 10, 3)]) for p in range(3)]
    merged = merge_multipage(tables)
    assert len(merged) == 1
    assert merged[0].page_end == 2
    assert len(merged[0].rows) == 10


def test_quality_is_row_weighted_mean() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)], quality=1.0)  # 4 filas
    t1 = _table(1, [HEADER, *_body(4, 1)], quality=0.5)  # 2 filas
    m = merge_multipage([t0, t1])[0]
    assert m.quality == pytest.approx((1.0 * 4 + 0.5 * 2) / 6)


# --------------------------------------------------------------------------- Regla 2


def test_rule2_continuation_without_header_is_appended_whole() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    cont_rows = _body(4, 3)  # primera fila con datos: "A-004", "Producto 4", "1", "100,00"
    t1 = _table(1, cont_rows, bbox=(41.5, 50.0, 548.0, 400.0))  # ±3 pt en x
    merged = merge_multipage([t0, t1])
    assert len(merged) == 1
    assert merged[0].page_end == 1
    assert merged[0].rows == [HEADER, *_body(1, 3), *cont_rows]


def test_rule2_requires_matching_x_bounds() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    t1 = _table(1, _body(4, 3), bbox=(60.0, 50.0, 548.0, 400.0))  # x0 corrido 20 pt
    assert len(merge_multipage([t0, t1])) == 2


def test_rule2_tolerance_parameter() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    t1 = _table(1, _body(4, 3), bbox=(46.0, 50.0, 550.0, 400.0))  # 6 pt
    assert len(merge_multipage([t0, t1])) == 2
    assert len(merge_multipage([t0, t1], col_tolerance_pt=8.0)) == 1


def test_rule2_first_row_that_looks_like_header_does_not_merge() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    other_header = ["Fecha", "Concepto", "Débito", "Crédito"]
    t1 = _table(1, [other_header, *_body(4, 3)])
    merged = merge_multipage([t0, t1])
    assert len(merged) == 2
    assert merged[1].rows[0] == other_header


def test_rule2_dates_count_as_data() -> None:
    t0 = _table(0, [["Fecha", "Concepto", "Importe"], ["15/10/2026", "Compra", "1.000,00"]])
    t1 = _table(1, [["16/10/2026", "Venta", "2.000,00"], ["17/10/2026", "Pago", "-500,00"]])
    merged = merge_multipage([t0, t1])
    assert len(merged) == 1
    assert len(merged[0].rows) == 4


# --------------------------------------------------------------------------- no se mergea


def test_non_consecutive_pages_do_not_merge() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    t2 = _table(2, [HEADER, *_body(4, 3)])
    merged = merge_multipage([t0, t2])
    assert len(merged) == 2
    assert [m.page_end for m in merged] == [0, 2]


def test_same_page_does_not_merge() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    t0b = _table(0, [HEADER, *_body(4, 3)], bbox=(40.0, 400.0, 550.0, 700.0))
    assert len(merge_multipage([t0, t0b])) == 2


def test_different_column_count_does_not_merge() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    rows3 = [["Código", "Detalle", "Importe"], ["A-004", "Producto 4", "100,00"]]
    t1 = _table(1, rows3)
    assert len(merge_multipage([t0, t1])) == 2


def test_tables_with_less_than_two_rows_never_merge() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    t1 = _table(1, [HEADER])
    assert len(merge_multipage([t0, t1])) == 2
    t0_short = _table(0, [HEADER])
    t1_full = _table(1, [HEADER, *_body(1, 2)])
    assert len(merge_multipage([t0_short, t1_full])) == 2


def test_empty_input() -> None:
    assert merge_multipage([]) == []


def test_single_table_is_returned_unchanged() -> None:
    t0 = _table(3, [HEADER, *_body(1, 2)], title="X")
    out = merge_multipage([t0])
    assert out == [t0]
    assert out[0] is not t0  # copia, no la misma instancia


# --------------------------------------------------------------------------- transporte


def test_transporte_rows_at_start_of_continuation_are_dropped() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3), ["", "Transporte", "", "600,00"]])
    t1 = _table(
        1,
        [HEADER, ["", "Subtotal transporte", "", "600,00"], ["Transporte", "", "", "600,00"], *_body(4, 2)],
    )
    m = merge_multipage([t0, t1])[0]
    assert m.rows == [HEADER, *_body(1, 3), ["", "Transporte", "", "600,00"], *_body(4, 2)]


def test_transporte_row_in_the_middle_is_kept() -> None:
    t0 = _table(0, [HEADER, *_body(1, 2)])
    t1 = _table(1, [HEADER, *_body(3, 1), ["Transporte", "", "", "1,00"], *_body(4, 1)])
    m = merge_multipage([t0, t1])[0]
    assert ["Transporte", "", "", "1,00"] in m.rows


# --------------------------------------------------------------------------- inmutabilidad


def test_inputs_are_not_mutated() -> None:
    rows0 = [HEADER, *_body(1, 3)]
    rows1 = [HEADER, *_body(4, 2)]
    t0 = _table(0, [list(r) for r in rows0])
    t1 = _table(1, [list(r) for r in rows1])
    merge_multipage([t0, t1])
    assert t0.rows == rows0
    assert t0.page_end == 0
    assert t1.rows == rows1


def test_merge_then_other_table_starts_new_group() -> None:
    t0 = _table(0, [HEADER, *_body(1, 3)])
    t1 = _table(1, [HEADER, *_body(4, 3)])
    t2 = _table(2, [["Fecha", "Concepto", "Débito", "Crédito"], ["15/10/26", "x", "1,00", ""]])
    merged = merge_multipage([t0, t1, t2])
    assert [(m.page, m.page_end) for m in merged] == [(0, 1), (2, 2)]
