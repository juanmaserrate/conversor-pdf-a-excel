"""Modo "Igual al PDF" sobre un ESCANEO: cajas, rellenos, negrita, tamaños y columnas salen de los píxeles.

El escaneo se arma sin tesseract: se dibuja un formulario vectorial ficticio, se rasteriza y se lo vuelve a
poner como imagen con una capa de texto invisible (lo que deja el OCR de la app).
"""

from __future__ import annotations

import pymupdf

from pdf2excel.core.layout import extract_layout_page
from pdf2excel.core.layout_raster import raster_primitives, render_page, scan_width_ratio
from pdf2excel.core.layout_text import spans_to_words

OCR_BOX = 0.9
COLS = (30.0, 90.0, 300.0, 350.0, 410.0, 490.0, 565.0)
HEADERS = ("Codigo", "Detalle", "Cant.", "Unidad", "Precio", "Importe")
ROWS = (
    ("X-100", "Tornillo hexagonal 6mm", "120", "UN", "12,50", "1.500,00"),
    ("X-200", "Arandela plana", "4", "RL", "3,20", "12,80"),
    ("X-300", "Cinta aisladora negra", "24", "UN", "1.250,00", "30.000,00"),
    ("X-400", "Pintura latex interior 20L", "3", "LT", "74.500,00", "223.500,00"),
)


def _vector_form() -> pymupdf.Document:
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=300)
    top, row_h = 40.0, 18.0
    page.draw_rect(pymupdf.Rect(COLS[0], top, COLS[-1], top + row_h), color=None, fill=(0.9, 0.9, 0.9))
    for i in range(len(ROWS) + 2):
        page.draw_line(
            (COLS[0], top + i * row_h), (COLS[-1], top + i * row_h), color=(0.6, 0.6, 0.6), width=0.8
        )
    for x in COLS:
        page.draw_line((x, top), (x, top + (len(ROWS) + 1) * row_h), color=(0.6, 0.6, 0.6), width=0.8)
    for j, text in enumerate(HEADERS):
        page.insert_text((COLS[j] + 4, top + 12.5), text, fontsize=7.5, fontname="hebo")
    for i, row in enumerate(ROWS):
        y = top + (i + 1) * row_h + 12.5
        for j, text in enumerate(row):
            if j in (2, 4, 5):  # cifras alineadas a la derecha
                width = pymupdf.get_text_length(text, fontname="helv", fontsize=7.5)
                page.insert_text((COLS[j + 1] - 4 - width, y), text, fontsize=7.5, fontname="helv")
            else:
                page.insert_text((COLS[j] + 4, y), text, fontsize=7.5, fontname="helv")
    page.insert_text((30, 24), "Remito de prueba", fontsize=14, fontname="hebo")
    return doc


def _scanned_form() -> tuple[pymupdf.Document, pymupdf.Page]:
    """Imagen de la página + texto invisible en las mismas posiciones, con una caja de OCR ~10 % angosta."""
    src = _vector_form()
    sp = src[0]
    words = sp.get_text("words")
    pix = sp.get_pixmap(dpi=170, colorspace=pymupdf.csGRAY, alpha=False)
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=300)
    page.insert_image(page.rect, pixmap=pix)
    for x0, y0, x1, y1, text, *_ in words:
        # como el OCR de la app: el texto invisible mide la caja leída, ~10 % más angosta que la tinta
        size = OCR_BOX * (x1 - x0) / max(1.0, pymupdf.get_text_length(text, fontname="helv", fontsize=1.0))
        page.insert_text((x0, y1 - 0.22 * (y1 - y0)), text, fontsize=size, fontname="helv", render_mode=3)
    src.close()
    return doc, page


def test_raster_finds_table_lines_and_header_fill() -> None:
    doc, page = _scanned_form()
    rgb, scale = render_page(page)
    prims = raster_primitives(rgb, scale, (595.0, 300.0))
    horizontals = [s for s in prims.segments if s.orient == "h"]
    verticals = [s for s in prims.segments if s.orient == "v"]
    assert len(horizontals) == len(ROWS) + 2
    assert len(verticals) == len(COLS)
    assert all(abs(s.a - COLS[0]) < 2 and abs(s.b - COLS[-1]) < 2 for s in horizontals)
    assert len(prims.fills) == 1  # la banda gris del encabezado (y no las líneas grises)
    fill = prims.fills[0]
    assert abs(fill.y0 - 40.0) < 2 and abs(fill.y1 - 58.0) < 2 and fill.fill is not None
    doc.close()


def test_scan_width_ratio_never_shrinks_the_text() -> None:
    doc, page = _scanned_form()
    rgb, scale = render_page(page)
    words = spans_to_words(page.get_texttrace(), page.rotation_matrix)
    assert 1.0 <= scan_width_ratio(words, rgb, scale) <= 1.3
    doc.close()


def test_scanned_table_keeps_cells_columns_bold_and_sizes() -> None:
    doc, page = _scanned_form()
    lp = extract_layout_page(page, scanned=True)
    by_text = {c.text: c for c in lp.cells if c.text}
    # 1) cada celda de la tabla dibujada tiene sus bordes: 4 lados
    bordered = [c for c in lp.cells if c.border_left and c.border_right and c.border_top and c.border_bottom]
    assert len(bordered) >= (len(ROWS) + 1) * len(HEADERS)
    # 2) el encabezado va en negrita y con el gris de fondo; el cuerpo no
    for text in HEADERS:
        cell = by_text[text]
        assert cell.lines[0][0].bold and cell.fill is not None
        assert 7.0 <= cell.lines[0][0].size_pt <= 8.5  # 7.5 pt de verdad: la negrita no infla el tamaño
    assert not by_text["Tornillo hexagonal 6mm"].lines[0][0].bold
    assert by_text["Tornillo hexagonal 6mm"].fill is None
    # 3) los valores de una misma columna caen en la misma columna de la hoja, un renglón por ítem
    for j in (2, 3, 4, 5):
        cols = {by_text[row[j]].col for row in ROWS}
        assert len(cols) == 1, (j, cols)
    rows = [by_text[row[0]].row for row in ROWS]
    assert rows == sorted(set(rows)) and all(by_text[row[1]].row == by_text[row[0]].row for row in ROWS)
    # 4) un solo tamaño para el cuerpo de la tabla (el OCR da ±1 pt de ruido) y no más chico que el real
    sizes = {run.size_pt for row in ROWS for t in row for run in by_text[t].lines[0]}
    assert len(sizes) == 1
    assert next(iter(sizes)) >= 7.0  # el texto real mide 7.5 pt
    # 5) los importes quedan como números en toda la columna
    assert all(isinstance(by_text[row[5]].value, float) for row in ROWS)
    doc.close()


def test_scanned_page_without_boxes_adds_no_phantom_borders() -> None:
    doc = pymupdf.open()
    src = doc.new_page(width=300, height=120)
    src.insert_text((30, 40), "Texto corrido sin cajas ni líneas", fontsize=9, fontname="helv")
    src.insert_text((30, 60), "Segundo renglón del documento", fontsize=9, fontname="helv")
    words = src.get_text("words")
    pix = src.get_pixmap(dpi=170, colorspace=pymupdf.csGRAY, alpha=False)
    page = pymupdf.open().new_page(width=300, height=120)
    page.insert_image(page.rect, pixmap=pix)
    for x0, y0, _x1, y1, text, *_ in words:
        page.insert_text((x0, y1 - 2), text, fontsize=(y1 - y0) / 1.15, fontname="helv", render_mode=3)
    rgb, scale = render_page(page)
    prims = raster_primitives(rgb, scale, (300.0, 120.0))
    assert prims.segments == [] and prims.fills == []
    doc.close()
