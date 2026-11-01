"""Tests de core/cells.py y core/candidates.py (WI-02).

Los PDFs de prueba se generan con reportlab en `tmp_path` con la misma receta que remito_grid.pdf y
remito_nogrid.pdf del corpus.
"""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from pathlib import Path

import pymupdf
import pytest
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from pdf2excel.core import candidates
from pdf2excel.core.candidates import page_candidates, tables_from_page, warm_up, words_for_page
from pdf2excel.core.cells import (
    MAX_TAIL_ROWS,
    clean_cell,
    drop_empty_rows,
    extend_tail,
    extend_tail_with_bottom,
    merge_ghost_columns,
    words_in_rect,
)
from pdf2excel.models import RawTable

HEADER = ["Código", "Producto", "Cantidad", "Precio unitario", "Total"]
PRODUCTS = [
    ("Mezcal joven", 10, 350.0),
    ("Tequila reposado", 5, 420.5),
    ("Agua mineral", 48, 12.9),
    ("Hielo bolsa 5kg", 20, 35.0),
    ("Limón kg", 8, 28.75),
    ("Sal de gusano", 3, 90.0),
]
REMITO_ROWS = [HEADER] + [
    [f"A-{100 + i}", p, str(q), f"{pr:.2f}", f"{q * pr:.2f}"] for i, (p, q, pr) in enumerate(PRODUCTS)
]


def build_remito(path: Path, grid: bool, total_row: bool = False) -> None:
    """Remito de 7×5 como el del corpus; con `total_row` agrega una fila TOTAL sin grilla debajo."""
    data = [list(r) for r in REMITO_ROWS]
    if total_row:
        data.append(["", "", "", "TOTAL", "7421.70"])
    t = Table(data, colWidths=[60, 140, 70, 100, 80])
    style = [("FONTSIZE", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]
    if grid:
        last_grid_row = len(REMITO_ROWS) - 1
        style.append(("GRID", (0, 0), (-1, last_grid_row), 0.8, colors.black))
    if total_row:
        style.append(("TOPPADDING", (0, len(data) - 1), (-1, len(data) - 1), 10))
    t.setStyle(TableStyle(style))
    st = getSampleStyleSheet()
    SimpleDocTemplate(str(path), pagesize=A4).build(
        [
            Paragraph("Remito N° 4521", st["Title"]),
            Paragraph("Cliente: Bar El Catorce — Fecha: 15/10/2026", st["Normal"]),
            Spacer(1, 12),
            t,
        ]
    )


@pytest.fixture(scope="module")
def remito_grid(tmp_path_factory: pytest.TempPathFactory) -> Path:
    p = tmp_path_factory.mktemp("pdfs") / "remito_grid.pdf"
    build_remito(p, grid=True)
    return p


@pytest.fixture(scope="module")
def remito_nogrid(tmp_path_factory: pytest.TempPathFactory) -> Path:
    p = tmp_path_factory.mktemp("pdfs") / "remito_nogrid.pdf"
    build_remito(p, grid=False)
    return p


@pytest.fixture(scope="module")
def remito_total(tmp_path_factory: pytest.TempPathFactory) -> Path:
    p = tmp_path_factory.mktemp("pdfs") / "remito_total.pdf"
    build_remito(p, grid=True, total_row=True)
    return p


# --------------------------------------------------------------------------- clean_cell


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("| 350.00 |", "350.00"),
        ("Mezcal\njoven", "Mezcal joven"),
        ("premi-\num", "premium"),
        ("A-\n100", "A- 100"),  # guión seguido de dígito: no es corte de palabra
        ("Precio    unitario", "Precio unitario"),
        ("| ¦ !", ""),
        ("", ""),
        (None, ""),
        ("  texto  ", "texto"),
        ("1.234,56-", "1.234,56-"),
        ("Hielo bolsa\n5kg", "Hielo bolsa 5kg"),
        ("a | b", "a b"),
    ],
)
def test_clean_cell(raw: str | None, expected: str) -> None:
    assert clean_cell(raw) == expected


def test_clean_cell_is_idempotent() -> None:
    for raw in ["| 350.00 |", "Mezcal\njoven", "premi-\num", "  a   b  "]:
        once = clean_cell(raw)
        assert clean_cell(once) == once


# --------------------------------------------------------------------------- words_in_rect


def _w(x0: float, y0: float, x1: float, y1: float, text: str) -> tuple:
    return (x0, y0, x1, y1, text, 0, 0, 0)


def test_words_in_rect_center_rule_and_x_order() -> None:
    words = [
        _w(60, 10, 80, 20, "joven"),
        _w(10, 10, 50, 20, "Mezcal"),
        _w(200, 10, 230, 20, "fuera"),  # centro fuera del rect
        _w(85, 10, 120, 20, "borde"),  # centro 102.5 > x1=100 → afuera
    ]
    assert words_in_rect(words, (0, 0, 100, 30)) == "Mezcal joven"


def test_words_in_rect_groups_lines_by_center_y() -> None:
    words = [
        _w(10, 30, 40, 40, "segundo"),
        _w(10, 10, 40, 20, "primero"),
        _w(45, 11, 70, 21, "renglón"),  # 1 pt de desvío → mismo renglón
    ]
    assert words_in_rect(words, (0, 0, 100, 50)) == "primero renglón\nsegundo"


def test_words_in_rect_empty() -> None:
    assert words_in_rect([], (0, 0, 10, 10)) == ""
    assert words_in_rect([_w(50, 50, 60, 60, "x")], pymupdf.Rect(0, 0, 10, 10)) == ""


def test_words_in_rect_never_breaks_numbers(remito_grid: Path) -> None:
    """La razón de no usar Table.extract(): "350.00" debe quedar entero."""
    doc = pymupdf.open(remito_grid)
    page = doc[0]
    words = page.get_text("words")
    table = page.find_tables(strategy="lines").tables[0]
    assert "350.00" in [clean_cell(words_in_rect(words, c)) for c in table.rows[1].cells if c]
    doc.close()


# --------------------------------------------------------------------------- ghost columns / empty rows


def test_merge_ghost_columns_acceptance() -> None:
    assert merge_ghost_columns([["A", ""], ["", "1"], ["B", ""]]) == [["A"], ["1"], ["B"]]


def test_merge_ghost_columns_keeps_legit_exclusive_columns() -> None:
    # Débito/Crédito mutuamente vacías con > 2 celdas cada una: NO se unen (caso extracto bancario)
    rows = [["Débito", "Crédito"], ["-1", ""], ["", "2"], ["-3", ""], ["", "4"], ["-5", ""], ["", "6"]]
    assert merge_ghost_columns(rows) == rows


def test_merge_ghost_columns_pads_and_drops_empty_columns() -> None:
    rows = [["a", "", "x"], ["b"], ["c", "", "z"]]
    assert merge_ghost_columns(rows) == [["a", "x"], ["b", ""], ["c", "z"]]


def test_merge_ghost_columns_empty_input() -> None:
    assert merge_ghost_columns([]) == []


def test_drop_empty_rows() -> None:
    assert drop_empty_rows([["a", ""], ["", "  "], [], ["", "b"]]) == [["a", ""], ["", "b"]]


# --------------------------------------------------------------------------- extend_tail


@dataclass
class _FakeRow:
    cells: list[tuple[float, float, float, float] | None]


@dataclass
class _FakeTable:
    bbox: tuple[float, float, float, float]
    rows: list[_FakeRow]

    @property
    def col_count(self) -> int:
        return len(self.rows[0].cells)


def _fake_table() -> _FakeTable:
    # 3 columnas (x: 0-100, 100-200, 200-300), 2 filas (y: 0-20, 20-40)
    cols = [(0.0, 100.0), (100.0, 200.0), (200.0, 300.0)]
    rows = [_FakeRow([(x0, y0, x1, y0 + 20) for x0, x1 in cols]) for y0 in (0.0, 20.0)]
    return _FakeTable(bbox=(0.0, 0.0, 300.0, 40.0), rows=rows)


def test_extend_tail_adds_total_row_10pt_below() -> None:
    table = _fake_table()
    rows = [["Producto", "Cant.", "Total"], ["Mezcal", "10", "350.00"]]
    words = [
        _w(10, 5, 60, 15, "Producto"),
        _w(110, 5, 150, 15, "Cant."),
        _w(210, 5, 250, 15, "Total"),
        _w(10, 25, 60, 35, "Mezcal"),
        _w(110, 25, 130, 35, "10"),
        _w(210, 25, 260, 35, "350.00"),
        # fila TOTAL dibujada 10 pt debajo del bbox, sin grilla
        _w(110, 50, 150, 60, "TOTAL"),
        _w(210, 50, 260, 60, "350.00"),
    ]
    added = extend_tail(None, table, rows, words)
    assert added == 1
    assert rows[-1] == ["", "TOTAL", "350.00"]


def test_extend_tail_chains_several_rows_and_stops_at_gap() -> None:
    table = _fake_table()
    rows: list[list[str]] = [["h1", "h2", "h3"], ["a", "b", "c"]]
    words = [
        _w(110, 45, 150, 55, "Subtotal"),
        _w(210, 45, 260, 55, "100,00"),
        _w(110, 60, 150, 70, "IVA"),
        _w(210, 60, 260, 70, "21,00"),
        _w(110, 75, 150, 85, "TOTAL"),
        _w(210, 75, 260, 85, "121,00"),
        _w(110, 150, 150, 160, "Lejos"),  # gap > 20 pt: no entra
        _w(210, 150, 260, 160, "999"),
    ]
    assert extend_tail(None, table, rows, words) == 3
    assert [r[1] for r in rows[-3:]] == ["Subtotal", "IVA", "TOTAL"]


def test_extend_tail_rejects_lines_outside_columns_or_with_one_cell() -> None:
    table = _fake_table()
    rows: list[list[str]] = [["h1", "h2", "h3"], ["a", "b", "c"]]
    # una palabra fuera de los rangos x → se corta
    words = [_w(110, 45, 150, 55, "TOTAL"), _w(400, 45, 450, 55, "afuera")]
    assert extend_tail(None, table, rows, words) == 0
    # una sola celda llena (texto suelto) → se corta
    words = [_w(10, 45, 150, 55, "Observaciones: entregado")]
    assert extend_tail(None, table, rows, words) == 0
    assert len(rows) == 2


def _numeric_body_rows() -> list[list[str]]:
    return [["Producto", "Cant.", "Total"], ["Mezcal", "10", "350.00"]]


def test_extend_tail_terminates_with_subpoint_words() -> None:
    """Palabras de alto < 1 pt justo debajo del bbox (OCR, marca de agua): antes era un loop infinito."""
    table = _fake_table()
    rows = _numeric_body_rows()
    words = [_w(110, 41.0, 150, 41.6, "TOTAL"), _w(210, 41.0, 260, 41.6, "6.234,56")]
    result: list[int] = []
    worker = threading.Thread(target=lambda: result.append(extend_tail(None, table, rows, words)))
    worker.daemon = True
    worker.start()
    worker.join(10)
    assert not worker.is_alive(), "extend_tail no terminó en 10 s"
    assert result == [1] and rows[-1] == ["", "TOTAL", "6.234,56"]
    assert len(rows) == 3  # el mismo renglón no se agrega dos veces


def test_extend_tail_never_exceeds_max_rows() -> None:
    table = _fake_table()
    rows = _numeric_body_rows()
    # 200 renglones de "TOTAL n" uno debajo del otro, cada 10 pt: solo entran MAX_TAIL_ROWS.
    words = []
    for k in range(200):
        y = 45 + 10 * k
        words += [_w(110, y, 150, y + 8, "Subtotal"), _w(210, y, 260, y + 8, f"{k},00")]
    assert extend_tail(None, table, rows, words) == MAX_TAIL_ROWS
    assert len(rows) == 2 + MAX_TAIL_ROWS


def test_extend_tail_rejects_kv_lines_and_prose() -> None:
    """'Vencimiento: 15/11/2026 | Forma de pago: contado' trae una fecha pero es prosa, no una fila."""
    table = _fake_table()
    rows = _numeric_body_rows()
    words = [
        _w(110, 50, 150, 60, "TOTAL"),
        _w(210, 50, 260, 60, "350.00"),
        _w(10, 65, 50, 75, "Vencimiento:"),
        _w(55, 65, 95, 75, "15/11/2026"),
        _w(210, 65, 240, 75, "Forma"),
        _w(242, 65, 250, 75, "de"),
        _w(252, 65, 270, 75, "pago:"),
        _w(272, 65, 299, 75, "contado"),
        _w(10, 80, 60, 90, "Observaciones:"),
        _w(62, 80, 95, 90, "entregar"),
        _w(110, 80, 140, 90, "2"),
    ]
    assert extend_tail(None, table, rows, words) == 1
    assert rows[-1] == ["", "TOTAL", "350.00"]


def test_extend_tail_value_must_be_in_a_numeric_column() -> None:
    """Un número en la columna de texto (Producto) no convierte prosa en fila de totales."""
    table = _fake_table()
    rows = _numeric_body_rows()
    words = [_w(10, 50, 60, 60, "15/11/2026"), _w(110, 50, 190, 60, "entregar en depósito")]
    assert extend_tail(None, table, rows, words) == 0
    # ...pero si el cuerpo no tiene ninguna columna numérica vale cualquier columna (tabla de texto).
    rows_text = [["h1", "h2", "h3"], ["a", "b", "c"]]
    words = [_w(10, 50, 60, 60, "TOTAL"), _w(110, 50, 150, 60, "350.00")]
    assert extend_tail(None, table, rows_text, words) == 1


def test_extend_tail_with_bottom_reports_last_row_y1() -> None:
    table = _fake_table()
    rows = _numeric_body_rows()
    words = [_w(110, 50, 150, 60, "TOTAL"), _w(210, 50, 260, 61.5, "350.00")]
    added, bottom = extend_tail_with_bottom(None, table, rows, words)
    assert added == 1 and bottom == pytest.approx(61.5)
    added, bottom = extend_tail_with_bottom(None, _fake_table(), _numeric_body_rows(), [])
    assert added == 0 and bottom == pytest.approx(40.0)


def test_extend_tail_on_real_pdf_recovers_total(remito_total: Path) -> None:
    """TOTAL sin grilla debajo de la tabla: termina en las filas de la tabla (vía extend_tail o layout)."""
    doc = pymupdf.open(remito_total)
    page = doc[0]
    words = page.get_text("words")
    table = page.find_tables(strategy="lines", union=False).tables[0]
    rows = [[clean_cell(words_in_rect(words, c)) if c else "" for c in row.cells] for row in table.rows]
    had_total = any("TOTAL" in r for r in rows)
    added = extend_tail(page, table, rows, words)
    assert had_total or added >= 1
    rows = merge_ghost_columns(drop_empty_rows([[clean_cell(c) for c in r] for r in rows]))
    assert any("TOTAL" in r and "7421.70" in r for r in rows)
    doc.close()


# --------------------------------------------------------------------------- page_candidates


@pytest.mark.parametrize("fixture", ["remito_grid", "remito_nogrid"])
def test_page_candidates_remito_7x5(fixture: str, request: pytest.FixtureRequest) -> None:
    path: Path = request.getfixturevalue(fixture)
    t0 = time.perf_counter()
    cands = page_candidates(path, 0)
    elapsed = time.perf_counter() - t0

    assert list(cands) == ["hybrid_union", "hybrid"]
    for source, tables in cands.items():
        assert len(tables) == 1, source
        t = tables[0]
        assert isinstance(t, RawTable)
        assert t.source == source
        assert t.page == 0 and t.page_end == 0
        assert len(t.rows) == 7 and all(len(r) == 5 for r in t.rows)
        assert t.rows[0] == HEADER
        assert t.rows[1] == ["A-100", "Mezcal joven", "10", "350.00", "3500.00"]
        assert t.rows == REMITO_ROWS
        assert 0.0 <= t.quality <= 1.0
        assert t.quality > 0.9
        assert len(t.bbox) == 4 and t.bbox[0] < t.bbox[2] and t.bbox[1] < t.bbox[3]
        assert t.title is None
    assert elapsed < 2.0, f"page_candidates tardó {elapsed:.2f} s"


ROT_TABLE = [
    ["Producto", "Cantidad", "Importe"],
    ["Coca", "2", "1.234,56"],
    ["Fernet", "1", "5.000,00"],
    ["Agua", "6", "600,00"],
]


def _upright_in_viewer(c: canvas.Canvas, rotation: int, width: float, height: float) -> None:
    """Transforma el canvas para que, con ``/Rotate rotation``, el contenido se vea derecho en el visor."""
    if rotation == 90:
        c.translate(width, 0)
        c.rotate(90)
    elif rotation == 180:
        c.translate(width, height)
        c.rotate(180)
    elif rotation == 270:
        c.translate(0, height)
        c.rotate(270)


def _rotated_table_pdf(path: Path, rotation: int) -> Path:
    """Reporte apaisado "exportado por un sistema": contenido girado en el mediabox + /Rotate."""
    width, height = landscape(A4) if rotation in (90, 270) else A4
    c = canvas.Canvas(str(path), pagesize=(width, height), invariant=1)
    _upright_in_viewer(c, rotation, width, height)
    x0, y_top = 50.0, 842.0 - 100.0
    cols = [x0, x0 + 150, x0 + 300, x0 + 450]
    rows_y = [y_top - 20 * i for i in range(len(ROT_TABLE) + 1)]
    c.setFont("Helvetica", 9)
    for ri, row in enumerate(ROT_TABLE):
        for ci, cell in enumerate(row):
            c.drawString(cols[ci] + 4, rows_y[ri] - 14, cell)
    for yy in rows_y:
        c.line(cols[0], yy, cols[-1], yy)
    for xx in cols:
        c.line(xx, rows_y[0], xx, rows_y[-1])
    c.showPage()
    c.save()
    if rotation:
        doc = pymupdf.open(path)
        doc[0].set_rotation(rotation)
        doc.save(path, incremental=True, encryption=pymupdf.PDF_ENCRYPT_KEEP)
        doc.close()
    return path


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_page_candidates_rotated_page_reads_the_table(tmp_path: Path, rotation: int) -> None:
    """Con /Rotate, find_tables da coordenadas rotadas y get_text("words") no: antes daba 0 tablas."""
    pdf = _rotated_table_pdf(tmp_path / f"rot{rotation}.pdf", rotation)
    doc = pymupdf.open(pdf)
    assert doc[0].rotation == rotation
    page_rect = doc[0].rect  # rect de la página rotada (la que ve el usuario)
    doc.close()

    cands = page_candidates(pdf, 0)
    best = max(cands.values(), key=len)
    assert best, cands
    assert best[0].rows == ROT_TABLE
    x0, y0, x1, y1 = best[0].bbox
    assert page_rect.x0 - 1 <= x0 < x1 <= page_rect.x1 + 1
    assert page_rect.y0 - 1 <= y0 < y1 <= page_rect.y1 + 1


def test_words_for_page_identity_without_rotation(remito_grid: Path) -> None:
    doc = pymupdf.open(remito_grid)
    page = doc[0]
    assert words_for_page(page) == list(page.get_text("words"))
    page.set_rotation(90)
    rotated = words_for_page(page)
    assert len(rotated) == len(page.get_text("words"))
    assert all(w[0] <= w[2] and w[1] <= w[3] for w in rotated)  # rects normalizados
    assert [w[4] for w in rotated] == [w[4] for w in page.get_text("words")]
    doc.close()


def test_table_bbox_covers_tail_rows(remito_total: Path) -> None:
    """El bbox de la RawTable se extiende con la cola para que text_outside_tables no repita el TOTAL."""
    from pdf2excel.core.text_blocks import text_outside_tables

    cands = page_candidates(remito_total, 0)
    doc = pymupdf.open(remito_total)
    page = doc[0]
    for tables in cands.values():
        for table in tables:
            if any("TOTAL" in r for r in table.rows):
                items = text_outside_tables(page, [table.bbox])
                assert not any("7421" in it.value.replace(".", "") for it in items), items
    doc.close()


def test_page_candidates_image_only_page_has_no_tables(tmp_path: Path, remito_grid: Path) -> None:
    """Un escaneo sin capa de texto no tiene palabras → 0 tablas (necesita OCR antes)."""
    src = pymupdf.open(remito_grid)
    pix = src[0].get_pixmap(dpi=100, colorspace=pymupdf.csGRAY)
    out = pymupdf.open()
    page = out.new_page(width=src[0].rect.width, height=src[0].rect.height)
    page.insert_image(page.rect, stream=pix.tobytes("png"))
    scan = tmp_path / "scan.pdf"
    out.save(scan)
    out.close()
    src.close()
    cands = page_candidates(scan, 0)
    assert cands == {"hybrid_union": [], "hybrid": []}


def test_page_candidates_closes_documents(remito_grid: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    opened: list[pymupdf.Document] = []
    real_open = pymupdf.open

    def tracking_open(*args: object, **kwargs: object) -> pymupdf.Document:
        doc = real_open(*args, **kwargs)
        opened.append(doc)
        return doc

    monkeypatch.setattr(candidates.pymupdf, "open", tracking_open)
    page_candidates(remito_grid, 0)
    assert len(opened) == 2
    assert all(d.is_closed for d in opened)


def test_tables_from_page_handles_find_tables_errors(remito_grid: Path) -> None:
    class BrokenPage:
        number = 3

        def find_tables(self, **_: object) -> object:
            raise RuntimeError("layout roto")

    assert tables_from_page(BrokenPage(), [], union=True, source="hybrid_union") == []


def test_tables_from_page_discards_tiny_tables(monkeypatch: pytest.MonkeyPatch) -> None:
    """Tablas con < 2 filas o < 2 columnas no son tablas."""

    class Row:
        def __init__(self, cells: list[tuple | None]) -> None:
            self.cells = cells

    class Tbl:
        def __init__(self, rows: list[Row], bbox: tuple) -> None:
            self.rows = rows
            self.bbox = bbox
            self.col_count = len(rows[0].cells)

    class Finder:
        def __init__(self, tables: list[Tbl]) -> None:
            self.tables = tables

    one_col = Tbl([Row([(0, 0, 50, 10)]), Row([(0, 10, 50, 20)])], (0, 0, 50, 20))
    one_row = Tbl([Row([(0, 100, 50, 110), (50, 100, 100, 110)])], (0, 100, 100, 110))  # lejos de la cola
    good = Tbl(
        [Row([(0, 0, 50, 10), (50, 0, 100, 10)]), Row([(0, 10, 50, 20), (50, 10, 100, 20)])],
        (0, 0, 100, 20),
    )

    class FakePage:
        number = 0

        def find_tables(self, **_: object) -> Finder:
            return Finder([one_col, one_row, good])

    words = [
        _w(5, 2, 20, 8, "a"),
        _w(55, 2, 70, 8, "b"),
        _w(5, 12, 20, 18, "1"),
        _w(55, 12, 70, 18, "2"),
        _w(5, 102, 20, 108, "h1"),
        _w(55, 102, 70, 108, "h2"),
    ]
    out = tables_from_page(FakePage(), words, union=False, source="hybrid")
    assert len(out) == 1
    assert out[0].rows == [["a", "b"], ["1", "2"]]
    assert out[0].bbox == (0.0, 0.0, 100.0, 20.0)


def test_warm_up_runs_quietly() -> None:
    warm_up()  # no debe lanzar


def test_core_does_not_import_qt() -> None:
    """core no importa Qt: se verifica en un intérprete aparte (pytest-qt ya carga PySide6 en este)."""
    import subprocess
    import sys

    code = (
        "import sys, pdf2excel.core.candidates, pdf2excel.core.cells, pdf2excel.core.quality; "
        "bad = [m for m in sys.modules if m.startswith(('PySide6', 'shiboken6'))]; "
        "sys.exit(1 if bad else 0)"
    )
    assert subprocess.run([sys.executable, "-c", code], check=False).returncode == 0
