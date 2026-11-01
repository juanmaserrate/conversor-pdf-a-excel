"""Tests de texto fuera de tablas y títulos (core/text_blocks.py, DESIGN.md §5.7).

Los PDFs se generan con reportlab en ``tmp_path``; la geometría de la tabla se conoce de antemano
(reportlab mide desde abajo; pymupdf desde arriba: ``y_mupdf = alto - y_reportlab``).
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pymupdf
import pytest
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas

from pdf2excel.core.text_blocks import line_items, repeated_lines, text_outside_tables, title_near
from pdf2excel.models import TextItem

W, H = A4
TABLE_TOP, TABLE_LEFT, TABLE_RIGHT, ROW_H, N_ROWS = 700.0, 50.0, 350.0, 20.0, 3
TABLE_BOTTOM = TABLE_TOP - ROW_H * N_ROWS
TABLE_BBOX = (TABLE_LEFT, H - TABLE_TOP, TABLE_RIGHT, H - TABLE_BOTTOM)
HEADER_LINE = "Cliente: Bar El Catorce — Fecha: 15/10/2026"
TITLE = "1. Bebidas con alcohol"


def _draw_table(c: canvas.Canvas) -> None:
    xs = [TABLE_LEFT, 150.0, 250.0, TABLE_RIGHT]
    for i in range(N_ROWS + 1):
        y = TABLE_TOP - i * ROW_H
        c.line(TABLE_LEFT, y, TABLE_RIGHT, y)
    for x in xs:
        c.line(x, TABLE_TOP, x, TABLE_BOTTOM)
    c.setFont("Helvetica", 9)
    cells = [
        ["Producto", "Cant.", "Precio"],
        ["Fernet 750", "12", "$ 9.800,00"],
        ["Gin 700", "6", "$ 15.200,00"],
    ]
    for r, row in enumerate(cells):
        for k, text in enumerate(row):
            c.drawString(xs[k] + 3, TABLE_TOP - r * ROW_H - 14, text)


def _make_pdf(
    path: Path, title_gap_pt: float, footer: str | None = None, n_pages: int = 1, obs_y: float = 770
) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    for p in range(n_pages):
        c.setFont("Helvetica", 10)
        c.drawString(50, 790, HEADER_LINE)
        c.drawString(50, obs_y, "Observaciones generales del remito")
        # Título: baseline a title_gap_pt sobre el borde superior de la tabla.
        c.drawString(50, TABLE_TOP + title_gap_pt, TITLE)
        _draw_table(c)
        if footer:
            c.drawString(50, 40, footer)
            c.drawString(400, 40, f"Página {p + 1} de {n_pages}")
        c.showPage()
    c.save()


@pytest.fixture
def page_close(tmp_path: Path) -> Iterator[pymupdf.Page]:
    path = tmp_path / "close.pdf"
    _make_pdf(path, title_gap_pt=12)
    doc = pymupdf.open(path)
    yield doc[0]
    doc.close()


@pytest.fixture
def page_far(tmp_path: Path) -> Iterator[pymupdf.Page]:
    path = tmp_path / "far.pdf"
    _make_pdf(path, title_gap_pt=80, obs_y=812)  # lejos del título para que no se fusionen
    doc = pymupdf.open(path)
    yield doc[0]
    doc.close()


# --------------------------------------------------------------------------- text_outside_tables


def test_kv_pairs_from_header_line(page_close: pymupdf.Page) -> None:
    items = text_outside_tables(page_close, [TABLE_BBOX])
    kvs = [it for it in items if it.kind == "kv"]
    assert [(it.key, it.value) for it in kvs] == [("Cliente", "Bar El Catorce"), ("Fecha", "15/10/2026")]
    assert all(it.page == 0 for it in items)
    assert all(isinstance(it, TextItem) for it in items)


def test_table_cells_are_excluded(page_close: pymupdf.Page) -> None:
    items = text_outside_tables(page_close, [TABLE_BBOX])
    values = [it.value for it in items]
    assert not any("Fernet" in v or "Producto" in v or "9.800" in v for v in values)
    # El texto libre y el título sí están.
    assert "Observaciones generales del remito" in values
    assert TITLE in values


def test_without_table_bboxes_everything_is_text(page_close: pymupdf.Page) -> None:
    items = text_outside_tables(page_close, [])
    values = [it.value for it in items]
    assert any("Fernet 750" in v for v in values)
    assert len([it for it in items if it.kind == "kv"]) == 2


def test_text_items_have_no_empty_lines(page_close: pymupdf.Page) -> None:
    items = text_outside_tables(page_close, [TABLE_BBOX])
    assert all(it.value.strip() for it in items)
    assert all(it.key is None for it in items if it.kind == "text")
    assert all(it.key for it in items if it.kind == "kv")


def test_items_in_reading_order(page_close: pymupdf.Page) -> None:
    items = text_outside_tables(page_close, [TABLE_BBOX])
    values = [it.value for it in items]
    assert (
        values.index("Bar El Catorce")
        < values.index("Observaciones generales del remito")
        < values.index(TITLE)
    )


# --------------------------------------------------------------------------- line_items (puro)


@pytest.mark.parametrize(
    ("line", "expected"),
    [
        (
            "Cliente: Bar El Catorce — Fecha: 15/10/2026",
            [("kv", "Cliente", "Bar El Catorce"), ("kv", "Fecha", "15/10/2026")],
        ),
        (
            "CUIT: 20-12345678-9 | Cond. IVA: Responsable Inscripto",
            [("kv", "CUIT", "20-12345678-9"), ("kv", "Cond. IVA", "Responsable Inscripto")],
        ),
        (
            "Remito N° 0001-00004521 – Original",
            [("text", None, "Remito N° 0001-00004521"), ("text", None, "Original")],
        ),
        ("Observaciones", [("text", None, "Observaciones")]),
        ("Cliente:", [("text", None, "Cliente:")]),
        ("10:30", [("text", None, "10:30")]),  # clave sin letras no es kv
        ("Nota: ver abajo — importante", [("kv", "Nota", "ver abajo"), ("text", None, "importante")]),
        ("", []),
        ("   ", []),
    ],
)
def test_line_items(line: str, expected: list[tuple]) -> None:
    items = line_items(3, line)
    assert [(it.kind, it.key, it.value) for it in items] == expected
    assert all(it.page == 3 for it in items)


def test_line_items_long_key_is_not_kv() -> None:
    items = line_items(0, "x" * 41 + ": valor")
    assert items[0].kind == "text"


# --------------------------------------------------------------------------- title_near


def test_title_near_finds_title_12pt_above(page_close: pymupdf.Page) -> None:
    assert title_near(page_close, TABLE_BBOX) == TITLE


def test_title_near_none_when_80pt_above(page_far: pymupdf.Page) -> None:
    assert title_near(page_far, TABLE_BBOX) is None


def test_title_near_respects_max_gap(page_far: pymupdf.Page) -> None:
    assert title_near(page_far, TABLE_BBOX, max_gap_pt=100) == TITLE


def test_title_near_ignores_kv_and_blocks_elsewhere(page_close: pymupdf.Page) -> None:
    # Un bbox a la derecha de la página no se solapa horizontalmente con ningún bloque.
    far_right = (400.0, TABLE_BBOX[1], 580.0, TABLE_BBOX[3])
    assert title_near(page_close, far_right) is None
    # Bbox cuya parte superior está justo debajo de la línea kv: la línea kv no cuenta como título.
    doc_header_y1 = H - 790 + 3
    under_kv = (50.0, doc_header_y1 + 5, 350.0, doc_header_y1 + 60)
    assert title_near(page_close, under_kv) is None


def test_title_near_ignores_long_titles(tmp_path: Path) -> None:
    path = tmp_path / "long.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("Helvetica", 10)
    c.drawString(
        50, TABLE_TOP + 12, "Un título extremadamente largo que supera los sesenta caracteres permitidos"
    )
    _draw_table(c)
    c.save()
    doc = pymupdf.open(path)
    try:
        assert title_near(doc[0], TABLE_BBOX) is None
    finally:
        doc.close()


# --------------------------------------------------------------------------- repeated_lines


def _items(page: int, *texts: str) -> list[TextItem]:
    return [TextItem(page=page, kind="text", key=None, value=t) for t in texts]


def test_repeated_lines_detects_header_and_footer() -> None:
    pages = [
        _items(0, "DISTRIBUIDORA SUR S.A.", "Página 1 de 3", "Lista de precios"),
        _items(1, "DISTRIBUIDORA SUR S.A.", "Página 2 de 3"),
        _items(2, "DISTRIBUIDORA SUR S.A.", "Página 3 de 3", "Fin"),
    ]
    assert repeated_lines(pages) == {"DISTRIBUIDORA SUR S.A."}


def test_repeated_lines_ratio() -> None:
    pages = [_items(0, "A", "B"), _items(1, "A", "B"), _items(2, "A"), _items(3, "A")]
    assert repeated_lines(pages) == {"A"}
    assert repeated_lines(pages, min_ratio=0.5) == {"A", "B"}


def test_repeated_lines_ignores_kv_and_single_page() -> None:
    kv = TextItem(page=0, kind="kv", key="Cliente", value="X")
    pages = [[kv, *_items(0, "Pie")], [kv, *_items(1, "Pie")]]
    assert repeated_lines(pages) == {"Pie"}
    assert repeated_lines([_items(0, "Pie")]) == set()
    assert repeated_lines([]) == set()


def test_repeated_lines_on_real_pdf(tmp_path: Path) -> None:
    path = tmp_path / "multi.pdf"
    _make_pdf(path, title_gap_pt=12, footer="DISTRIBUIDORA SUR S.A. - Lista de precios", n_pages=3)
    doc = pymupdf.open(path)
    try:
        pages_items = [text_outside_tables(doc[i], [TABLE_BBOX]) for i in range(doc.page_count)]
    finally:
        doc.close()
    repeated = repeated_lines(pages_items)
    assert "DISTRIBUIDORA SUR S.A. - Lista de precios" in repeated
    assert "Observaciones generales del remito" in repeated
    assert "Página 1 de 3" not in repeated


# --------------------------------------------------------------------------- páginas con /Rotate


def _make_rotated_pdf(path: Path, rotation: int) -> None:
    """Mismo contenido que ``_make_pdf(title_gap_pt=12)`` pero girado en el mediabox + ``/Rotate``: se ve
    derecho en el visor (reporte apaisado exportado por un sistema)."""
    width, height = landscape(A4) if rotation in (90, 270) else A4
    c = canvas.Canvas(str(path), pagesize=(width, height))
    if rotation == 90:
        c.translate(width, 0)
        c.rotate(90)
    elif rotation == 180:
        c.translate(width, height)
        c.rotate(180)
    elif rotation == 270:
        c.translate(0, height)
        c.rotate(270)
    c.setFont("Helvetica", 10)
    c.drawString(50, 790, HEADER_LINE)
    c.drawString(50, 770, "Observaciones generales del remito")
    c.drawString(50, TABLE_TOP + 12, TITLE)
    _draw_table(c)
    c.showPage()
    c.save()
    doc = pymupdf.open(path)
    doc[0].set_rotation(rotation)
    doc.save(path, incremental=True, encryption=pymupdf.PDF_ENCRYPT_KEEP)
    doc.close()


@pytest.mark.parametrize("rotation", [90, 180, 270])
def test_rotated_page_blocks_match_find_tables_space(tmp_path: Path, rotation: int) -> None:
    """Con /Rotate los bbox de los bloques se llevan al sistema de find_tables: la tabla sigue excluida."""
    plain_path = tmp_path / "plain.pdf"
    _make_pdf(plain_path, title_gap_pt=12)
    doc = pymupdf.open(plain_path)
    page = doc[0]
    plain = text_outside_tables(page, [TABLE_BBOX])
    assert title_near(page, TABLE_BBOX) == TITLE
    doc.close()

    rotated_path = tmp_path / f"rot{rotation}.pdf"
    _make_rotated_pdf(rotated_path, rotation)
    doc = pymupdf.open(rotated_path)
    page = doc[0]
    assert page.rotation == rotation
    tables = page.find_tables(strategy="lines", union=False).tables
    assert tables, "find_tables no encontró la tabla en la página rotada"
    bbox = tuple(tables[0].bbox)
    items = text_outside_tables(page, [bbox])
    assert [(it.kind, it.key, it.value) for it in items] == [(it.kind, it.key, it.value) for it in plain]
    assert not any("Fernet" in it.value or "9.800" in it.value for it in items)
    assert title_near(page, bbox) == TITLE
    doc.close()
