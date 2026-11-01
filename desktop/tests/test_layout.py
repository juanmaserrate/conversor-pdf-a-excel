"""Tests de core/layout.py: extractor del modo "Igual al PDF" (PDFs sintéticos con datos ficticios)."""

from __future__ import annotations

import time
from datetime import date

import pymupdf

from pdf2excel.core.layout import extract_layout_page
from pdf2excel.models import LayoutCell, LayoutPage

GRAY_LINE = (0.6, 0.6, 0.6)
WHITE = (1, 1, 1)


def _new(width: float = 400, height: float = 300) -> tuple[pymupdf.Document, pymupdf.Page]:
    doc = pymupdf.open()
    return doc, doc.new_page(width=width, height=height)


def _frame(page: pymupdf.Page, rect: tuple[float, float, float, float], fill: tuple | None = None) -> None:
    page.draw_rect(pymupdf.Rect(*rect), color=GRAY_LINE, fill=fill, width=0.75)


def _thin_rect_border(page: pymupdf.Page, rect: tuple[float, float, float, float], t: float = 0.75) -> None:
    """Borde dibujado como cuatro rectángulos de relleno finos (como hace el navegador al imprimir)."""
    x0, y0, x1, y1 = rect
    for r in ((x0, y0, x1, y0 + t), (x0, y1 - t, x1, y1), (x0, y0, x0 + t, y1), (x1 - t, y0, x1, y1)):
        page.draw_rect(pymupdf.Rect(*r), color=None, fill=GRAY_LINE)


def _cell_with(lp: LayoutPage, text: str) -> LayoutCell:
    found = [c for c in lp.cells if text in c.text]
    assert len(found) == 1, f"{text!r}: {len(found)} celdas"
    return found[0]


def _all_text(lp: LayoutPage) -> str:
    return "\n".join(c.text for c in lp.cells)


def _label_box(page: pymupdf.Page, rect: tuple[float, float, float, float], label: str, value: str) -> None:
    _frame(page, rect)
    x0, y0 = rect[0], rect[1]
    page.insert_text((x0 + 4, y0 + 9), label, fontsize=6, fontname="helv", color=(0.5, 0.5, 0.5))
    if value:
        page.insert_text((x0 + 4, y0 + 22), value, fontsize=9, fontname="hebo", color=(0.1, 0.1, 0.3))


# --- cajas y texto


def test_caja_etiqueta_valor_es_una_celda_de_dos_lineas_con_estilos_distintos() -> None:
    doc, page = _new()
    _label_box(page, (20, 20, 180, 50), "TRANSPORTISTA", "PABLO MENDEZ")
    lp = extract_layout_page(page)
    cell = _cell_with(lp, "TRANSPORTISTA")
    assert cell.text == "TRANSPORTISTA\nPABLO MENDEZ"
    assert len(cell.lines) == 2
    label, value = cell.lines[0][0], cell.lines[1][0]
    assert not label.bold and value.bold
    assert label.size_pt < value.size_pt
    assert label.color != value.color
    assert cell.wrap is True
    assert (cell.row_span, cell.col_span) == (1, 1)
    assert [c for c in lp.cells if c.text] == [cell]
    doc.close()


def test_caja_vacia_conserva_sus_bordes() -> None:
    doc, page = _new()
    _label_box(page, (20, 20, 180, 50), "PATENTE", "")
    _frame(page, (200, 20, 300, 50))
    lp = extract_layout_page(page)
    empty = [c for c in lp.cells if not c.lines]
    assert len(empty) == 1
    for side in (empty[0].border_left, empty[0].border_top, empty[0].border_right, empty[0].border_bottom):
        assert side is not None and 0.5 <= side.width_pt <= 1.0
    labelled = _cell_with(lp, "PATENTE")
    assert len(labelled.lines) == 1
    assert labelled.border_top is not None and labelled.border_bottom is not None
    doc.close()


def test_texto_tapado_va_a_hidden_texts_y_a_ninguna_celda() -> None:
    doc, page = _new()
    _frame(page, (20, 20, 120, 45))
    _frame(page, (120, 20, 220, 45))
    page.insert_text((26, 36), "N/A", fontsize=9)
    page.draw_rect(pymupdf.Rect(21, 21, 119, 44), color=None, fill=WHITE)
    page.insert_text((126, 36), "Sí", fontsize=9)
    lp = extract_layout_page(page)
    assert lp.hidden_texts == ["N/A"]
    assert "N/A" not in _all_text(lp)
    assert lp.n_words == 1
    assert _cell_with(lp, "Sí").text == "Sí"
    doc.close()


def test_borde_desde_rectangulos_finos_y_relleno_de_encabezado() -> None:
    doc, page = _new()
    page.draw_rect(pymupdf.Rect(20, 20, 220, 40), color=None, fill=(0.8, 0.8, 1.0))
    for x0, x1 in ((20, 120), (120, 220)):
        _thin_rect_border(page, (x0, 20, x1, 40))
        _thin_rect_border(page, (x0, 40, x1, 60))
    page.insert_text((26, 34), "Producto", fontsize=9, fontname="hebo")
    page.insert_text((126, 34), "Cantidad", fontsize=9, fontname="hebo")
    page.insert_text((26, 54), "Tornillo", fontsize=9)
    page.insert_text((126, 54), "Tuerca", fontsize=9)
    lp = extract_layout_page(page)
    head = _cell_with(lp, "Producto")
    body = _cell_with(lp, "Tornillo")
    assert head.fill == "CCCCFF"
    assert body.fill is None
    for cell in (head, body):
        for side in (cell.border_left, cell.border_top, cell.border_right, cell.border_bottom):
            assert side is not None and abs(side.width_pt - 0.75) < 0.2
            assert side.color == "999999"
    doc.close()


def test_fila_de_seccion_combina_todas_las_columnas() -> None:
    doc, page = _new()
    xs = [20, 60, 140, 220]
    for i in range(3):
        _thin_rect_border(page, (xs[i], 20, xs[i + 1], 40))  # encabezado: 3 celdas
    _thin_rect_border(page, (20, 40, 220, 60))  # sección: una sola caja ancha
    for i in range(3):
        _thin_rect_border(page, (xs[i], 60, xs[i + 1], 80))  # fila de datos
    page.insert_text((24, 34), "N", fontsize=9)
    page.insert_text((64, 34), "Destino", fontsize=9)
    page.insert_text((144, 34), "Hora", fontsize=9)
    page.insert_text((24, 54), "R7 · Zona norte", fontsize=9, fontname="hebo")
    page.insert_text((24, 74), "1", fontsize=9)
    page.insert_text((64, 74), "Kiosco", fontsize=9)
    page.insert_text((144, 74), "08:15", fontsize=9)
    lp = extract_layout_page(page)
    section = _cell_with(lp, "R7")
    header = _cell_with(lp, "Destino")
    assert section.col_span == 3 * 1 or section.col_span >= 3
    assert section.col == min(c.col for c in lp.cells)
    assert section.col_span == sum(c.col_span for c in lp.cells if c.row == header.row)
    assert section.row_span == 1 and section.row == header.row + 1
    doc.close()


def test_tabla_conserva_filas_y_columnas() -> None:
    doc, page = _new()
    xs = [20 + 40 * i for i in range(6)]
    for r in range(4):
        for c in range(5):
            _thin_rect_border(page, (xs[c], 20 + 20 * r, xs[c + 1], 40 + 20 * r))
    lp = extract_layout_page(page)
    assert len(lp.cells) == 20
    assert {c.row for c in lp.cells} == {0, 1, 2, 3}
    assert {c.col for c in lp.cells} == {0, 1, 2, 3, 4}
    assert all(not c.lines for c in lp.cells)
    assert all(abs(w - 40) < 2 for w in lp.col_widths_pt)
    doc.close()


def test_caja_con_linea_interna_se_parte_en_dos_celdas() -> None:
    doc, page = _new()
    _frame(page, (20, 20, 160, 60))
    page.draw_line((24, 50), (156, 50), color=GRAY_LINE, width=0.75)
    page.insert_text((26, 32), "KM SALIDA", fontsize=7, fontname="hebo")
    lp = extract_layout_page(page)
    assert len(lp.cells) == 2
    top, bottom = sorted(lp.cells, key=lambda c: c.row)
    assert top.text == "KM SALIDA" and bottom.text == ""
    assert top.border_bottom is not None and bottom.border_top is not None
    doc.close()


# --- texto suelto


def test_texto_libre_se_agrupa_en_parrafos_y_conserva_su_posicion() -> None:
    doc, page = _new()
    page.insert_text((30, 40), "TITULO DEL INFORME", fontsize=14, fontname="hebo")
    page.insert_text((30, 70), "Primera linea del parrafo", fontsize=9)
    page.insert_text((30, 81), "Segunda linea del parrafo", fontsize=9)
    page.insert_text((30, 140), "Otro bloque lejano", fontsize=9)
    page.insert_text((250, 70), "Columna derecha", fontsize=9)
    lp = extract_layout_page(page)
    para = _cell_with(lp, "Primera")
    assert para.text == "Primera linea del parrafo\nSegunda linea del parrafo"
    assert para.valign == "top"
    right = _cell_with(lp, "Columna derecha")
    assert right.col > para.col + para.col_span - 1
    assert _cell_with(lp, "TITULO").lines[0][0].size_pt >= 13
    assert _cell_with(lp, "lejano").row > para.row
    assert not any(c.border_top or c.border_left for c in lp.cells)  # sin bordes dibujados
    doc.close()


def test_pagina_con_solo_texto_alineado_no_inventa_bordes_ni_columnas_de_mas() -> None:
    doc, page = _new(400, 400)
    for i in range(12):
        y = 40 + 14 * i
        page.insert_text((30, y), f"AR-{1000 + i}", fontsize=9)
        page.insert_text((120, y), f"Articulo {i}", fontsize=9)
        value = f"{1000 * (i + 1):,}".replace(",", ".") + ",00"
        page.insert_text((380 - pymupdf.get_text_length(value, fontsize=9), y), value, fontsize=9)
    lp = extract_layout_page(page)
    assert len(lp.col_widths_pt) <= 12
    amounts = [c for c in lp.cells if c.text.endswith(",00")]
    assert len(amounts) == 12
    assert {c.halign for c in amounts} == {"right"}
    assert len({c.col + c.col_span for c in amounts}) == 1  # mismo borde derecho
    doc.close()


def test_alineaciones_dentro_de_cajas() -> None:
    doc, page = _new()
    for i, (text, expected) in enumerate((("izq", "left"), ("centro", "center"), ("der", "right"))):
        rect = (20, 20 + 30 * i, 220, 45 + 30 * i)
        _frame(page, rect)
        width = pymupdf.get_text_length(text, fontsize=9)
        if expected == "left":
            x = rect[0] + 5
        elif expected == "center":
            x = (rect[0] + rect[2] - width) / 2
        else:
            x = rect[2] - 5 - width
        page.insert_text((x, rect[1] + 16), text, fontsize=9)
    lp = extract_layout_page(page)
    assert _cell_with(lp, "izq").halign == "left"
    assert _cell_with(lp, "centro").halign == "center"
    assert _cell_with(lp, "der").halign == "right"
    assert _cell_with(lp, "izq").valign == "center"
    doc.close()


def test_caja_con_texto_arriba_a_la_izquierda_usa_valign_top() -> None:
    doc, page = _new()
    _frame(page, (20, 20, 220, 120))
    page.insert_text((26, 34), "Observaciones", fontsize=8)
    page.insert_text((26, 45), "Llamar antes de entregar", fontsize=8)
    lp = extract_layout_page(page)
    assert _cell_with(lp, "Observaciones").valign == "top"
    doc.close()


# --- valores tipados


def _typed_page(values: list[str]) -> LayoutPage:
    doc, page = _new(400, 400)
    for i, text in enumerate(values):
        _frame(page, (20, 20 + 22 * i, 200, 42 + 22 * i))
        page.insert_text((26, 35 + 22 * i), text, fontsize=9)
    lp = extract_layout_page(page)
    doc.close()
    return lp


def test_tipado_conservador_de_celdas_aisladas() -> None:
    doc, page = _new(400, 400)
    samples = [
        "15/03/2025",
        "08:30",
        "N° 12345",
        "AB 123 CD",
        "1.234",
        "1.234,56",
        "$ 1.234,56",
        "12,5 %",
        "007",
    ]
    for i, text in enumerate(samples):
        _frame(page, (20, 20 + 30 * i, 200, 45 + 30 * i))
        page.insert_text((26, 37 + 30 * i), text, fontsize=9)
        if i % 2:  # separa las cajas para que ninguna forme una columna de datos
            page.insert_text((250, 37 + 30 * i), "x", fontsize=9)
    lp = extract_layout_page(page)
    by = {c.text: c for c in lp.cells if c.text in samples}
    assert by["15/03/2025"].value == date(2025, 3, 15)
    assert by["15/03/2025"].number_format == "dd/mm/yyyy"
    assert by["1.234,56"].value == 1234.56
    assert by["$ 1.234,56"].value == 1234.56
    assert by["12,5 %"].value == 0.125
    for text in ("08:30", "N° 12345", "AB 123 CD", "1.234", "007"):
        assert by[text].value is None, text
    doc.close()


def test_columna_de_enteros_se_tipa_como_numeros_y_el_encabezado_queda_texto() -> None:
    doc, page = _new(400, 400)
    page.draw_rect(pymupdf.Rect(20, 20, 120, 40), color=GRAY_LINE, width=0.75)
    page.insert_text((26, 34), "Bultos", fontsize=9, fontname="hebo")
    for i, text in enumerate(["8", "18", "12", "16"]):
        page.draw_rect(pymupdf.Rect(20, 40 + 20 * i, 120, 60 + 20 * i), color=GRAY_LINE, width=0.75)
        page.insert_text((26, 54 + 20 * i), text, fontsize=9)
    lp = extract_layout_page(page)
    assert _cell_with(lp, "Bultos").value is None
    values = sorted(c.value for c in lp.cells if isinstance(c.value, int))
    assert values == [8, 12, 16, 18]
    doc.close()


def test_typed_false_deja_todo_como_texto() -> None:
    doc, page = _new()
    _frame(page, (20, 20, 200, 42))
    page.insert_text((26, 35), "15/03/2025", fontsize=9)
    lp = extract_layout_page(page, typed=False)
    assert all(c.value is None for c in lp.cells)
    doc.close()


# --- imágenes y casos especiales


def _picture(w: int = 80, h: int = 40, value: int = 120) -> pymupdf.Pixmap:
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, w, h), False)
    pix.clear_with(value)
    return pix


def test_imagen_anclada_arriba_a_la_izquierda() -> None:
    doc, page = _new()
    page.insert_image(pymupdf.Rect(25, 25, 85, 55), pixmap=_picture())
    page.insert_text((100, 40), "Empresa Ficticia S.A.", fontsize=11)
    page.insert_image(pymupdf.Rect(1, 1, 2, 2), pixmap=_picture())  # demasiado chica: se ignora
    lp = extract_layout_page(page)
    assert len(lp.images) == 1
    img = lp.images[0]
    assert img.format in ("jpeg", "png") and len(img.data) > 100
    assert img.bbox == (25.0, 25.0, 85.0, 55.0)
    assert (img.row, img.col) == (0, 0)
    assert abs(img.width_pt - 60) < 0.5 and abs(img.height_pt - 30) < 0.5
    top_left = sum(lp.col_widths_pt[: img.col]) + img.dx_pt
    assert top_left >= 0
    assert extract_layout_page(page, embed_images=False).images == []
    doc.close()


def test_pagina_vacia() -> None:
    doc, page = _new(300, 200)
    lp = extract_layout_page(page)
    assert lp.cells == [] and lp.images == [] and lp.n_words == 0
    assert (lp.width_pt, lp.height_pt) == (300, 200)
    assert lp.col_widths_pt and lp.row_heights_pt
    doc.close()


def test_pagina_rotada_trabaja_en_el_marco_visual() -> None:
    doc, page = _new(400, 300)
    _label_box(page, (20, 20, 180, 50), "CHOFER", "LUIS PEREZ")
    page.insert_text((30, 120), "Pie", fontsize=9)
    page.set_rotation(90)
    lp = extract_layout_page(page)
    assert (lp.width_pt, lp.height_pt) == (300, 400)
    cell = _cell_with(lp, "CHOFER")
    assert cell.text == "CHOFER\nLUIS PEREZ" or cell.text.count("CHOFER") == 1
    x0, y0, x1, y1 = cell.bbox
    assert 0 <= x0 < x1 <= 300 and 0 <= y0 < y1 <= 400
    doc.close()


def test_escaneo_con_capa_ocr_cuenta_el_texto_invisible_y_omite_la_imagen_de_fondo() -> None:
    doc, page = _new(300, 400)
    page.insert_image(page.rect, pixmap=_picture(300, 400, 230))
    page.insert_text((40, 60), "Factura 0001", fontsize=10, render_mode=3)
    page.insert_text((40, 90), "Total 1.500,00", fontsize=10, render_mode=3)
    lp = extract_layout_page(page, scanned=True)
    assert lp.images == []
    assert lp.n_words == 4
    assert "Factura 0001" in _all_text(lp)
    plain = extract_layout_page(page, scanned=False)  # el escaneo se detecta solo
    assert plain.n_words == 4 and plain.images == []
    doc.close()


def test_texto_invisible_en_pagina_normal_se_descarta() -> None:
    doc, page = _new()
    page.insert_text((40, 60), "Escondido", fontsize=10, render_mode=3)
    page.insert_text((40, 90), "Visible", fontsize=10)
    lp = extract_layout_page(page)
    assert lp.hidden_texts == ["Escondido"]
    assert _all_text(lp) == "Visible"
    doc.close()


# --- límites y determinismo


def test_limites_de_grilla_y_rendimiento_con_miles_de_dibujos() -> None:
    doc, page = _new(595, 842)
    shape = page.new_shape()
    n = 0
    for r in range(40):
        for c in range(32):
            x, y = 20 + c * 17, 20 + r * 19
            for rect in ((x, y, x + 17, y + 0.75), (x, y + 19, x + 17, y + 19.75), (x, y, x + 0.75, y + 19)):
                shape.draw_rect(pymupdf.Rect(*rect))
                shape.finish(color=None, fill=GRAY_LINE)
                n += 1
    for k in range(1500):  # líneas verticales con x todas distintas
        x = 10 + k * 0.37
        shape.draw_line((x, 800), (x, 830))
        shape.finish(color=(0, 0, 0), width=0.3)
        n += 1
    shape.commit()
    assert n >= 5000
    started = time.perf_counter()
    lp = extract_layout_page(page)
    elapsed = time.perf_counter() - started
    assert len(lp.col_widths_pt) <= 80
    assert len(lp.row_heights_pt) <= 250
    assert elapsed < 2.0
    doc.close()


def test_filas_altas_se_parten_para_entrar_en_excel() -> None:
    doc, page = _new(400, 1500)
    page.insert_text((30, 40), "Arriba", fontsize=9)
    page.insert_text((30, 1450), "Abajo", fontsize=9)
    lp = extract_layout_page(page)
    assert max(lp.row_heights_pt) <= 409
    assert _cell_with(lp, "Abajo").row > _cell_with(lp, "Arriba").row
    doc.close()


def test_es_determinista() -> None:
    doc, page = _new()
    _label_box(page, (20, 20, 180, 50), "UNIDAD", "UT-14")
    page.insert_text((30, 120), "Texto suelto", fontsize=9)
    page.insert_image(pymupdf.Rect(200, 20, 260, 50), pixmap=_picture())
    assert extract_layout_page(page) == extract_layout_page(page)
    doc.close()


# --- ronda de fidelidad visual (R1): líneas finas, tildes, rellenos continuos, campos con subrayado, columnas


def test_linea_gruesa_dibujada_como_relleno_de_trazos_no_se_pierde() -> None:
    """Una regla de 2 pt hecha con un camino relleno (no con "re") sigue siendo un borde (p. ej. la línea azul
    bajo un título), también en una página rotada."""
    for rotate in (0, 90):
        doc, page = _new(420, 300)
        shape = page.new_shape()
        shape.draw_polyline([(30, 60), (200, 60), (390, 60), (390, 62), (30, 62)])  # con "l", no "re"
        shape.finish(color=None, fill=(0.1, 0.2, 0.5), closePath=True)
        shape.commit()
        page.insert_text((30, 50), "TITULO DEL INFORME", fontsize=14, fontname="hebo")
        page.insert_text((30, 90), "Detalle", fontsize=9, fontname="helv")
        if rotate:
            page.set_rotation(rotate)
        lp = extract_layout_page(page)
        bordered = [c for c in lp.cells if c.border_top or c.border_bottom or c.border_left or c.border_right]
        assert bordered, f"rotación {rotate}: se perdió la regla"
        # en una página rotada la regla es vertical en el marco del PDF: se mide el largo de su borde
        long_edge = max(
            max(
                (c.bbox[2] - c.bbox[0]) if (c.border_top or c.border_bottom) else 0.0,
                (c.bbox[3] - c.bbox[1]) if (c.border_left or c.border_right) else 0.0,
            )
            for c in bordered
        )
        assert long_edge > 300, f"rotación {rotate}: la regla quedó en {long_edge:.0f} pt"
        doc.close()


def test_lados_de_una_caja_no_se_pierden_al_unir_trazos_de_distinto_grosor() -> None:
    """Dos trazos verticales sobre el mismo x (uno 0.8 pt y otro 0.6 pt) y separados no se fusionan."""
    doc, page = _new()
    page.draw_rect(pymupdf.Rect(30, 30, 300, 130), color=(0, 0, 0), width=0.8)
    page.draw_rect(pymupdf.Rect(30, 140, 300, 190), color=(0, 0, 0), width=0.6)
    page.insert_text((60, 60), "Emisor de ejemplo", fontsize=9, fontname="helv")
    page.insert_text((60, 160), "Cliente de ejemplo", fontsize=9, fontname="helv")
    lp = extract_layout_page(page)
    left_edges = sorted({round(c.bbox[1]) for c in lp.cells if c.border_left is not None and c.bbox[0] < 40})
    assert left_edges and left_edges[0] <= 40  # el lado izquierdo del recuadro de arriba existe
    assert any(c.border_left is not None and c.bbox[1] < 100 for c in lp.cells)
    doc.close()


def test_tilde_vectorial_de_una_casilla_pasa_a_caracter() -> None:
    doc, page = _new()
    for k, checked in enumerate((True, False)):
        y = 40 + k * 30
        page.draw_rect(pymupdf.Rect(40, y, 49, y + 9), color=(0, 0, 0), width=0.7)
        if checked:
            shape = page.new_shape()
            shape.draw_polyline([(42, y + 4.5), (44.5, y + 7), (48, y + 2)])
            shape.finish(color=(0, 0, 0), width=1.0, closePath=False)
            shape.commit()
        page.insert_text((60, y + 8), f"Opción {k + 1}", fontsize=8, fontname="helv")
    lp = extract_layout_page(page)
    ticks = [c for c in lp.cells if c.text in ("✔", "✗")]
    assert len(ticks) == 1 and ticks[0].text == "✔"
    assert ticks[0].halign == "center" and ticks[0].border_left is not None
    assert (
        sum(1 for c in lp.cells if not c.lines and c.border_left and c.col_span == 1) >= 1
    )  # la casilla vacía
    doc.close()


def test_barra_de_color_continua_aunque_haya_huecos_sin_texto() -> None:
    doc, page = _new(420, 200)
    page.draw_rect(pymupdf.Rect(20, 40, 400, 58), color=None, fill=(0.85, 0.85, 0.85))
    page.insert_text((24, 52), "Código", fontsize=8, fontname="hebo")
    page.insert_text((250, 52), "Cantidad", fontsize=8, fontname="hebo")
    page.insert_text((24, 90), "S-001", fontsize=8, fontname="helv")
    lp = extract_layout_page(page)
    filled = [c for c in lp.cells if c.fill == "D9D9D9"]
    covered = sum(c.bbox[2] - c.bbox[0] for c in filled if c.bbox[1] < 50 < c.bbox[3])
    assert covered >= 0.95 * 380  # de punta a punta: también donde no hay texto
    assert not any(c.fill for c in lp.cells if c.text == "S-001")
    doc.close()


def test_valor_sobre_subrayado_es_otra_celda_que_la_etiqueta() -> None:
    doc, page = _new(420, 120)
    page.insert_text((20, 40), "Apellido y nombres:", fontsize=8, fontname="hebo")
    page.insert_text((100, 40), "Quiroga, Marcela", fontsize=8, fontname="helv", color=(0.1, 0.2, 0.6))
    page.draw_line((96, 43), (260, 43), color=(0.5, 0.5, 0.5), width=0.6)
    lp = extract_layout_page(page)
    label = _cell_with(lp, "Apellido")
    value = _cell_with(lp, "Quiroga")
    assert "Quiroga" not in label.text and "Apellido" not in value.text
    assert value.col > label.col or value.bbox[0] > label.bbox[0]
    assert value.bbox[0] >= 95  # el valor arranca donde empieza la línea, no pegado a la etiqueta
    doc.close()


def test_costo_y_valor_pegados_quedan_en_columnas_distintas() -> None:
    """Dos columnas numéricas alineadas a la derecha con poco hueco no se juntan en una celda."""
    doc, page = _new(420, 200)
    rows = (("8.067,50", "1.952.335,00"), ("6.108,00", "1.600.296,00"), ("5.568,90", "1.314.260,40"))
    for i, (cost, value) in enumerate(rows):
        y = 40 + i * 14
        page.insert_text((20, y), f"AR-10{i}", fontsize=8, fontname="helv")
        for text, right in ((cost, 300.0), (value, 380.0)):
            width = pymupdf.get_text_length(text, fontname="helv", fontsize=8)
            page.insert_text((right - width, y), text, fontsize=8, fontname="helv")
    lp = extract_layout_page(page)
    for cost, value in rows:
        assert _cell_with(lp, cost).text == cost
        assert _cell_with(lp, value).text == value
        assert _cell_with(lp, cost).col != _cell_with(lp, value).col
    doc.close()


def test_numeros_tipados_no_inventan_separador_de_miles() -> None:
    """Un listado "3500.00 / 2102.50" (sin miles) se muestra igual: formato "0.00", no "#,##0.00"."""
    doc, page = _new(300, 200)
    for i, (qty, price) in enumerate((("3", "3500.00"), ("2", "2102.50"), ("5", "420.50"))):
        y = 40 + i * 16
        page.insert_text((20, y), qty, fontsize=9, fontname="helv")
        page.insert_text((120, y), price, fontsize=9, fontname="helv")
    page.insert_text((120, 24), "Total", fontsize=9, fontname="hebo")
    page.insert_text((250, 24), "Notas", fontsize=9, fontname="hebo")  # deja pistas libres a la derecha
    lp = extract_layout_page(page)
    priced = [c for c in lp.cells if c.value is not None and isinstance(c.value, float)]
    assert len(priced) == 3
    assert all(c.number_format == "0.00" for c in priced)
    doc.close()


def test_formato_de_miles_se_usa_en_toda_la_columna_si_alguno_lo_trae() -> None:
    doc, page = _new(300, 200)
    page.insert_text((120, 24), "Importe", fontsize=9, fontname="hebo")
    page.insert_text((250, 24), "Notas", fontsize=9, fontname="hebo")  # deja pistas libres a la derecha
    for i, text in enumerate(("1.234,50", "98,00", "12,25")):
        page.insert_text((120, 40 + i * 16), text, fontsize=9, fontname="helv")
    lp = extract_layout_page(page)
    fmts = {c.number_format for c in lp.cells if isinstance(c.value, float)}
    assert fmts == {"#,##0.00"}  # el mismo formato en toda la columna (los menores a 1000 se ven igual)
    doc.close()


def test_texto_de_escaneo_toma_un_tamano_por_renglon_y_sin_restos() -> None:
    from pdf2excel.core.layout import _steady_scan_sizes
    from pdf2excel.core.layout_text import Char, Style, Word

    def word(text: str, x0: float, size: float, y: float = 100.0) -> Word:
        style = Style(size=size)
        width = pymupdf.get_text_length(text, fontname="helv", fontsize=9.0)
        chars = [Char(x0, y - 8, x0 + width, y, y, text, style)]
        return Word(chars=chars, x0=x0, y0=y - 8, x1=x0 + width, y1=y, oy=y, size=size)

    words = [
        word("Tornillo", 20, 5.0),
        word("hexagonal", 80, 7.5),
        word("6mm", 150, 6.0),
        word("-", 190, 1.0),
    ]
    out = _steady_scan_sizes(words)
    assert [w.text for w in out] == ["Tornillo", "hexagonal", "6mm"]  # el resto "-" de 1 pt desaparece
    assert len({w.size for w in out}) == 1 and abs(out[0].size - 9.0) <= 1.0
    assert all(ch.style.size == out[0].size for w in out for ch in w.chars)


def test_tilde_hecho_con_dos_lineas_sueltas_es_un_tilde_y_una_equis_es_cruz() -> None:
    doc, page = _new()
    page.draw_rect(pymupdf.Rect(40, 40, 49, 49), color=(0, 0, 0), width=0.7)
    page.draw_line((41.5, 44.5), (43.8, 47.5), color=(0, 0, 0), width=1.0)  # dos dibujos de una línea
    page.draw_line((43.8, 47.5), (48.0, 41.5), color=(0, 0, 0), width=1.0)
    page.draw_rect(pymupdf.Rect(40, 80, 49, 89), color=(0, 0, 0), width=0.7)
    page.draw_line((41.5, 81.5), (47.5, 87.5), color=(0, 0, 0), width=1.0)
    page.draw_line((47.5, 81.5), (41.5, 87.5), color=(0, 0, 0), width=1.0)
    page.insert_text((60, 48), "Sí", fontsize=8, fontname="helv")
    page.insert_text((60, 88), "No", fontsize=8, fontname="helv")
    lp = extract_layout_page(page)
    assert sorted(c.text for c in lp.cells if c.text in ("✔", "✗")) == ["✔", "✗"]
    doc.close()


def test_titulo_con_letras_espaciadas_no_se_parte_en_dos_celdas(tmp_path) -> None:  # type: ignore[no-untyped-def]
    """Un título con espaciado entre letras tiene huecos grandes entre palabras: si la
    segunda palabra cae justo en una columna alineada de otros renglones, no debe cortarse ahí."""
    from reportlab.pdfgen import canvas

    path = tmp_path / "titulo.pdf"
    c = canvas.Canvas(str(path), pagesize=(400, 300))
    font, size, spacing, x0 = "Helvetica-Bold", 13.5, 2.2, 40.0
    text = "HOJA DE RUTA"
    t = c.beginText(x0, 250)
    t.setFont(font, size)
    t.setCharSpace(spacing)
    t.textOut(text)
    c.drawText(t)
    # la palabra "RUTA" empieza donde otros tres renglones alinean su borde izquierdo
    ruta_x = x0 + c.stringWidth("HOJA DE ", font, size) + spacing * len("HOJA DE ")
    for i, line in enumerate(("Reparto norte", "Zona sur", "Base central")):
        c.setFont("Helvetica", 8)
        c.drawString(ruta_x, 200 - i * 14, line)
    c.save()
    with pymupdf.open(str(path)) as doc:
        lp = extract_layout_page(doc[0])
    titles = [c for c in lp.cells if "HOJA" in c.text]
    assert len(titles) == 1 and titles[0].text == "HOJA DE RUTA"


def test_numero_corto_de_escaneo_no_define_el_tamano_de_su_tramo() -> None:
    # "120" mide menos de lo que dice su caja (dígitos sueltos): la columna no puede salir más chica
    from pdf2excel.core.layout import _steady_scan_sizes
    from pdf2excel.core.layout_text import Char, Style, Word

    def word(text: str, x0: float, width: float, y: float) -> Word:
        chars = [Char(x0, y - 8, x0 + width, y, y, text, Style(size=7.0))]
        return Word(chars=chars, x0=x0, y0=y - 8, x1=x0 + width, y1=y, oy=y, size=7.0)

    helv = pymupdf.Font("helv")
    desc = "Tornillo hexagonal"
    full = float(helv.text_length(desc, fontsize=7.0))
    words = [word(desc, 20, full, 100.0), word("120", 200, 3 * 0.556 * 5.5, 100.0)]
    words += [word("Arandela plana", 20, float(helv.text_length("Arandela plana", fontsize=7.0)), 120.0)]
    words += [word("4", 200, 0.556 * 7.0, 120.0)]
    out = {w.text: w.size for w in _steady_scan_sizes(words)}
    assert out["120"] == out["4"] == 7.0


def test_ocr_dudoso_por_mayusculas_dentro_de_la_palabra() -> None:
    from pdf2excel.core.layout_text import doubtful_ocr_tokens

    assert doubtful_ocr_tokens("curT 30-70123456-7") == ["curT"]
    assert doubtful_ocr_tokens("MaRtIn Perez") == ["MaRtIn"]
    for fine in ("CUIT", "McDonald", "DeLuca", "Cuenta corriente", "iva21", "N° PEDIDO", "a1B"):
        assert doubtful_ocr_tokens(fine) == []


def test_escaneo_con_lecturas_dudosas_deja_el_aviso_en_la_pagina() -> None:
    from pdf2excel.core.layout_text import OCR_DOUBT_WARNING

    doc, page = _new()
    page.draw_rect(pymupdf.Rect(40, 40, 200, 70), color=(0, 0, 0), width=0.7)
    page.insert_text((46, 60), "curT 30-70123456-7", fontsize=9, fontname="helv", render_mode=3)
    page.insert_image(
        pymupdf.Rect(0, 0, 300, 300), pixmap=pymupdf.Pixmap(pymupdf.csGRAY, pymupdf.IRect(0, 0, 30, 30), 1)
    )
    lp = extract_layout_page(page, scanned=True)
    assert OCR_DOUBT_WARNING in lp.warnings
    doc.close()


def test_imagen_de_pagina_girada_no_sale_en_blanco() -> None:
    import io

    from PIL import Image

    for rotation in (0, 90, 180, 270):
        doc, page = _new(400, 300)
        pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 100, 50), False)
        pix.clear_with(0)
        pix.set_rect(pix.irect, (200, 0, 0))
        page.insert_image(pymupdf.Rect(50, 50, 250, 150), pixmap=pix)
        page.set_rotation(rotation)
        lp = extract_layout_page(page)
        assert len(lp.images) == 1
        shot = Image.open(io.BytesIO(lp.images[0].data)).convert("RGB")
        r, g, b = shot.getpixel((shot.width // 2, shot.height // 2))
        assert r > 150 and g < 60 and b < 60, rotation
        doc.close()


def test_texto_recortado_por_un_clip_no_aparece() -> None:
    doc, page = _new(400, 300)
    page.insert_text((50, 100), "VISIBLE", fontsize=11)
    page.insert_text((50, 150), "SEED", fontsize=11)
    page.clean_contents()
    xref = page.get_contents()[0]
    stream = doc.xref_stream(xref)
    stream += b"\nq 0 0 10 10 re W n BT /helv 12 Tf 100 200 Td (CLIPPEDOUT) Tj ET Q\n"
    doc.update_stream(xref, stream)
    assert "CLIPPEDOUT" not in page.get_text()
    lp = extract_layout_page(page)
    assert "CLIPPEDOUT" not in _all_text(lp)
    assert "VISIBLE" in _all_text(lp)
    doc.close()


def test_celda_blanca_sobre_panel_gris_queda_blanca() -> None:
    doc, page = _new(400, 300)
    page.draw_rect(pymupdf.Rect(50, 50, 350, 150), fill=(0.8, 0.8, 0.8), color=None)
    page.draw_rect(pymupdf.Rect(60, 60, 160, 90), fill=(1, 1, 1), color=(0, 0, 0), width=0.75)
    page.insert_text((65, 80), "white cell", fontsize=9)
    page.draw_rect(pymupdf.Rect(180, 60, 280, 90), fill=None, color=(0, 0, 0), width=0.75)
    page.insert_text((185, 80), "gray cell", fontsize=9)
    lp = extract_layout_page(page)
    assert _cell_with(lp, "white cell").fill is None
    assert _cell_with(lp, "gray cell").fill == "CCCCCC"
    doc.close()


def test_porcentaje_suelto_en_columna_numerica_queda_como_texto() -> None:
    lp = _typed_page(["Dato", "1.234,50", "3,5", "15%", "0,5"])
    cell = _cell_with(lp, "15%")
    assert cell.value is None or cell.number_format != "#,##0"
    assert not (isinstance(cell.value, float) and cell.value == 0.15 and cell.number_format == "#,##0")


def test_texto_vertical_conserva_los_espacios() -> None:
    doc, page = _new(400, 300)
    page.insert_text((100, 250), "VERTICAL TEXT", fontsize=10, rotate=90)
    lp = extract_layout_page(page)
    assert "VERTICAL TEXT" in _all_text(lp)
    doc.close()
