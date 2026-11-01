"""Tests de core/visibility.py: qué texto de una página se ve de verdad (PDFs sintéticos, datos ficticios)."""

from __future__ import annotations

import pymupdf
import pytest

from pdf2excel.core.visibility import hidden_span_rects, visible_words

WHITE = (1, 1, 1)


def _page() -> tuple[pymupdf.Document, pymupdf.Page]:
    doc = pymupdf.open()
    return doc, doc.new_page(width=300, height=200)


def _texts(words: list[tuple]) -> list[str]:
    return [w[4] for w in words]


def test_texto_tapado_por_rect_blanco_posterior_se_oculta() -> None:
    doc, page = _page()
    page.insert_text((40, 60), "N/A", fontsize=10)
    page.insert_text((40, 100), "Visible", fontsize=10)
    page.draw_rect(pymupdf.Rect(36, 48, 70, 66), color=None, fill=WHITE)  # tapa "N/A" después
    hidden = hidden_span_rects(page)
    assert len(hidden) == 1
    assert hidden[0].y0 < 58 < hidden[0].y1
    assert _texts(visible_words(page)) == ["Visible"]
    assert "N/A" in _texts(page.get_text("words"))
    doc.close()


def test_rect_blanco_anterior_es_fondo_y_no_oculta() -> None:
    doc, page = _page()
    page.draw_rect(pymupdf.Rect(30, 40, 200, 80), color=None, fill=WHITE)  # fondo, antes del texto
    page.draw_rect(pymupdf.Rect(30, 90, 200, 120), color=None, fill=(0.9, 0.9, 0.5))
    page.insert_text((40, 60), "Fondo blanco", fontsize=10)
    page.insert_text((40, 110), "Fondo amarillo", fontsize=10)
    assert hidden_span_rects(page) == []
    assert _texts(visible_words(page)) == _texts(page.get_text("words"))
    doc.close()


def test_relleno_de_color_posterior_tambien_tapa() -> None:
    doc, page = _page()
    page.insert_text((40, 60), "Tapado", fontsize=10)
    page.draw_rect(pymupdf.Rect(30, 48, 120, 70), color=None, fill=(0.2, 0.2, 0.8))
    assert len(hidden_span_rects(page)) == 1
    assert visible_words(page) == []
    doc.close()


def test_cobertura_parcial_o_translucida_no_oculta() -> None:
    doc, page = _page()
    page.insert_text((40, 60), "Parcial", fontsize=10)
    page.insert_text((40, 110), "Translucido", fontsize=10)
    page.draw_rect(pymupdf.Rect(30, 55, 52, 70), color=None, fill=WHITE)  # tapa solo una parte
    page.draw_rect(pymupdf.Rect(30, 98, 140, 118), color=None, fill=WHITE, fill_opacity=0.5)
    assert hidden_span_rects(page) == []
    assert len(visible_words(page)) == 2
    doc.close()


def test_texto_fuera_de_pagina_y_tamano_cero_se_ocultan() -> None:
    doc, page = _page()
    page.insert_text((900, 60), "Afuera", fontsize=10)
    page.insert_text((40, 60), "Diminuto", fontsize=0.2)
    page.insert_text((40, 100), "Normal", fontsize=10)
    assert len(hidden_span_rects(page)) == 2
    assert _texts(visible_words(page)) == ["Normal"]
    doc.close()


def test_texto_del_color_del_fondo_se_oculta_pero_no_sobre_relleno_oscuro() -> None:
    doc, page = _page()
    page.insert_text((40, 40), "Blanco sobre blanco", fontsize=10, color=WHITE)
    page.draw_rect(pymupdf.Rect(30, 60, 200, 90), color=None, fill=(0.1, 0.1, 0.4))
    page.insert_text((40, 80), "Blanco sobre azul", fontsize=10, color=WHITE)
    hidden = hidden_span_rects(page)
    assert len(hidden) == 1
    assert hidden[0].y1 < 50
    assert _texts(visible_words(page)) == ["Blanco", "sobre", "azul"]
    doc.close()


def test_texto_invisible_solo_cuenta_en_escaneos_con_capa_ocr() -> None:
    doc, page = _page()
    page.insert_text((40, 60), "Capa oculta", fontsize=10, render_mode=3)
    assert len(hidden_span_rects(page)) == 1
    assert visible_words(page) == []
    assert hidden_span_rects(page, scanned=True) == []
    assert _texts(visible_words(page, scanned=True)) == ["Capa", "oculta"]
    doc.close()


def test_escaneo_se_detecta_por_imagen_de_fondo_mas_texto_invisible() -> None:
    doc, page = _page()
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 60, 40), False)
    pix.clear_with(200)
    page.insert_image(page.rect, pixmap=pix)
    page.insert_text((40, 60), "Texto OCR", fontsize=10, render_mode=3)
    assert hidden_span_rects(page) == []  # scanned=None detecta el escaneo
    assert _texts(visible_words(page)) == ["Texto", "OCR"]
    doc.close()


def test_texto_tapado_y_reescrito_encima_deja_solo_el_de_arriba() -> None:
    doc, page = _page()
    page.insert_text((40, 60), "viejo", fontsize=10)
    page.draw_rect(pymupdf.Rect(30, 48, 90, 68), color=None, fill=WHITE)
    page.insert_text((40, 60), "nuevo", fontsize=10)
    assert _texts(visible_words(page)) == ["nuevo"]
    doc.close()


def test_visible_words_coincide_con_get_text_menos_las_ocultas() -> None:
    doc, page = _page()
    for i, word in enumerate(["uno", "dos", "tres", "cuatro"]):
        page.insert_text((20, 30 + 30 * i), word, fontsize=10)
    page.draw_rect(pymupdf.Rect(15, 48, 60, 66), color=None, fill=WHITE)  # tapa "dos"
    all_words = page.get_text("words")
    kept = visible_words(page)
    assert len(kept) == len(all_words) - 1
    assert set(map(tuple, kept)) == {tuple(w) for w in all_words if w[4] != "dos"}
    clip = pymupdf.Rect(0, 0, 300, 70)
    assert _texts(visible_words(page, clip=clip)) == ["uno"]
    doc.close()


@pytest.mark.parametrize("scanned", [None, False, True])
def test_pagina_sin_texto_ni_dibujos(scanned: bool | None) -> None:
    doc, page = _page()
    assert hidden_span_rects(page, scanned) == []
    assert visible_words(page, scanned=scanned) == []
    doc.close()


def test_pagina_rotada_usa_el_marco_de_get_text() -> None:
    doc, page = _page()
    page.insert_text((40, 60), "Gira", fontsize=10)
    page.insert_text((40, 100), "Tapado", fontsize=10)
    page.draw_rect(pymupdf.Rect(30, 88, 100, 106), color=None, fill=WHITE)
    page.set_rotation(90)
    assert len(hidden_span_rects(page)) == 1
    assert _texts(visible_words(page)) == ["Gira"]
    doc.close()


def test_texto_blanco_sobre_banda_oscura_translucida_no_se_oculta() -> None:
    doc, page = _page()
    page.draw_rect(pymupdf.Rect(50, 50, 300, 100), fill=(0, 0, 0.3), color=None, fill_opacity=0.8)
    page.insert_text((60, 80), "WhiteOnAlpha", fontsize=12, color=(1, 1, 1))
    assert "WhiteOnAlpha" in _texts(visible_words(page))
    doc.close()


def test_palabra_partida_en_spans_de_un_glifo_no_se_oculta() -> None:
    """Regresión: "1/2" del pie impreso por un navegador son 3 spans ("1", "/", "2") que MuPDF une en una
    palabra; el filtro de recortes descartaba el "1" y el "2" (el centro de la palabra caía en la "/")."""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    for i, ch in enumerate("1/2"):
        page.insert_text((560.0 + i * 4.5, 826.0), ch, fontsize=8)
    assert hidden_span_rects(page) == []
    assert "".join(_texts(visible_words(page))) == "1/2"


def test_span_largo_de_varias_palabras_no_se_oculta() -> None:
    """Regresión: un subtítulo de varias palabras en un solo span debe seguir visible (cada palabra cubre
    una fracción chica del span, así que no alcanza con medir superposición contra una sola palabra)."""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((70.0, 90.0), "VIAJE N 1234 REPARTO R7 FECHA 01/01/2030", fontsize=9)
    assert hidden_span_rects(page) == []
    assert len(visible_words(page)) == 7
