"""Generador de formularios y medidor de fidelidad (``tests/layout_score.py``)."""

from __future__ import annotations

import time
from collections import Counter
from pathlib import Path
from typing import Any

import pymupdf
import pytest
from openpyxl import Workbook
from openpyxl.styles import Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from tests.corpus.generate_forms import generate_forms
from tests.forms_fixtures import forms_corpus, forms_ground_truth  # noqa: F401
from tests.layout_score import (
    FidelityResult,
    LibreOfficeUnavailable,
    find_covered_spans,
    score,
    tokens,
    visible_words,
)

EXPECTED = {
    "hoja_ruta",
    "remito_cajas",
    "factura_bloques",
    "form_dos_columnas",
    "tabla_seccion_combinada",
    "pagina_rotada",
    "solo_texto_alineado",
    "multipagina_listado",
    "escaneado_formulario",
}


# ------------------------------------------------------------------------------------------ generador


def test_generates_all_forms(forms_corpus: dict[str, Path]) -> None:  # noqa: F811
    assert len(forms_corpus) >= 8
    assert set(forms_corpus) >= EXPECTED
    for path in forms_corpus.values():
        assert path.exists()
        with pymupdf.open(str(path)) as doc:
            assert len(doc) >= 1


def test_generator_is_deterministic_and_fast(tmp_path: Path) -> None:
    t0 = time.perf_counter()
    first = generate_forms(tmp_path / "a")
    elapsed = time.perf_counter() - t0
    second = generate_forms(tmp_path / "b")
    assert elapsed < 60
    for name in first:
        assert first[name].read_bytes() == second[name].read_bytes(), name
    assert (tmp_path / "a" / "ground_truth_forms.json").read_bytes() == (
        tmp_path / "b" / "ground_truth_forms.json"
    ).read_bytes()


def test_hoja_ruta_has_exactly_two_covered_spans(forms_corpus: dict[str, Path]) -> None:  # noqa: F811
    covered: list[str] = []
    with pymupdf.open(str(forms_corpus["hoja_ruta"])) as doc:
        for page in doc:
            whites = [
                d
                for d in page.get_drawings()
                if d.get("fill") == (1.0, 1.0, 1.0) and "f" in d.get("type", "")
            ]
            for span in page.get_texttrace():
                bbox = pymupdf.Rect(span["bbox"])
                text = "".join(chr(ch[0]) for ch in span["chars"])
                if any(
                    w["seqno"] > span["seqno"]
                    and (bbox & pymupdf.Rect(w["rect"])).get_area() >= 0.8 * bbox.get_area()
                    for w in whites
                ):
                    covered.append(text)
    assert covered == ["N/A", "N/A"]
    with pymupdf.open(str(forms_corpus["hoja_ruta"])) as doc:
        assert [t for page in doc for t, _ in find_covered_spans(page)] == covered


def test_hoja_ruta_structure(forms_corpus: dict[str, Path], forms_ground_truth: dict[str, Any]) -> None:  # noqa: F811
    gt = forms_ground_truth["hoja_ruta.pdf"]
    assert gt["pages"] == 2
    assert gt["covered_words"] == ["N/A", "N/A"]
    assert "N/A" not in gt["visible_words"]
    boxes = gt["boxes"]
    assert boxes["label_value"] >= 7  # 2 "viaje" + 5 cajas etiqueta/valor con dato
    kinds = Counter(b["kind"] for b in boxes["items"])
    assert kinds["empty"] >= 5  # 2 KM + vehículo + patente + temperatura + observaciones
    with pymupdf.open(str(forms_corpus["hoja_ruta"])) as doc:
        assert len(doc) == 2
        assert len(doc[1].get_images()) >= 2  # logo + mapa
        assert len(doc[0].get_drawings()) > 300  # bordes como rectángulos finos


def test_rotated_page(forms_corpus: dict[str, Path]) -> None:  # noqa: F811
    with pymupdf.open(str(forms_corpus["pagina_rotada"])) as doc:
        assert doc[0].rotation == 90
        assert doc[0].rect.width > doc[0].rect.height  # visualmente apaisada


def test_scanned_form_has_no_text(forms_corpus: dict[str, Path]) -> None:  # noqa: F811
    with pymupdf.open(str(forms_corpus["escaneado_formulario"])) as doc:
        assert doc[0].get_text().strip() == ""
        assert len(doc[0].get_images()) == 1


def test_ground_truth_matches_visible_text(
    forms_corpus: dict[str, Path],  # noqa: F811
    forms_ground_truth: dict[str, Any],  # noqa: F811
) -> None:
    """Las palabras visibles del ground truth coinciden con lo que se extrae del PDF (sin las tapadas)."""
    for name, path in forms_corpus.items():
        gt = forms_ground_truth[f"{name}.pdf"]
        if gt.get("scanned"):
            continue
        expected = Counter(t for w in gt["visible_words"] for t in tokens(w))
        with pymupdf.open(str(path)) as doc:
            assert len(doc) == gt["pages"], name
            found = Counter(w.text for page in doc for w in visible_words(page))
        assert found == expected, (name, (found - expected), (expected - found))


# ------------------------------------------------------------------------------------------ medidor


def _make_grid_pdf(path: Path, rows: list[list[str]]) -> None:
    c = canvas.Canvas(str(path), pagesize=A4, invariant=1)
    c.setFont("Helvetica", 10)
    for r, row in enumerate(rows):
        for k, word in enumerate(row):
            c.drawString(50 + k * 150, 780 - r * 20, word)
    c.save()


def _good_xlsx(path: Path, rows: list[list[str]]) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    thin = Side(style="thin")
    for r, row in enumerate(rows, start=1):
        ws.row_dimensions[r].height = 20
        for k, word in enumerate(row, start=1):
            cell = ws.cell(r, k, word)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for letter in "ABC":
        ws.column_dimensions[letter].width = 150 / 5.25
    wb.save(path)


_ROWS = [
    ["Cliente", "Fecha", "Importe"],
    ["Alfa", "01/02/2025", "1500"],
    ["Beta", "03/02/2025", "2300"],
    ["Gamma", "05/02/2025", "870"],
]


def _score_or_skip(pdf: Path, xlsx: Path, workdir: Path, **kw: Any) -> FidelityResult:
    try:
        return score(pdf, xlsx, workdir, **kw)
    except LibreOfficeUnavailable:
        pytest.skip("LibreOffice no está instalado")


@pytest.mark.slow
def test_score_good_grid(tmp_path: Path) -> None:
    pdf, xlsx = tmp_path / "grid.pdf", tmp_path / "grid.xlsx"
    _make_grid_pdf(pdf, _ROWS)
    _good_xlsx(xlsx, _ROWS)
    png = tmp_path / "cmp.png"
    result = _score_or_skip(pdf, xlsx, tmp_path / "w", png=png)
    assert result.recall is not None and result.recall >= 0.98
    assert result.extra_words == 0
    assert result.mixed_cells == 0
    assert result.pos_err_median_pct is not None and result.pos_err_median_pct < 6
    assert result.sheets == result.pdf_pages == 1
    assert png.exists() and png.stat().st_size > 0
    assert FidelityResult(**result.to_dict()) == result  # serializable


@pytest.mark.slow
def test_score_bad_grid(tmp_path: Path) -> None:
    pdf, xlsx = tmp_path / "grid.pdf", tmp_path / "bad.xlsx"
    _make_grid_pdf(pdf, _ROWS)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "Cliente Fecha"  # junta dos palabras de columnas distintas
    ws["A2"] = "fantasma"  # palabra que no existe en el PDF
    ws["A30"] = "Alfa"
    wb.create_sheet("Otra")["A1"] = "Importe"
    wb.save(xlsx)
    result = _score_or_skip(pdf, xlsx, tmp_path / "w")
    assert result.recall is not None and result.recall < 0.5
    assert result.extra_words >= 1
    assert result.sheets == 2 and result.pdf_pages == 1
    assert not result.passes


@pytest.mark.slow
def test_score_hoja_ruta_covered_text_is_extra(tmp_path: Path, forms_corpus: dict[str, Path]) -> None:  # noqa: F811
    """Un Excel que copia el texto TAPADO ("N/A") suma palabras sobrantes."""
    pdf = forms_corpus["hoja_ruta"]
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    with pymupdf.open(str(pdf)) as doc:
        words = [w[4] for w in doc[0].get_text("words")]  # incluye los N/A tapados
    for i, w in enumerate(words):
        ws.cell(1 + i // 10, 1 + i % 10, w)
    xlsx = tmp_path / "all_words.xlsx"
    wb.save(xlsx)
    result = _score_or_skip(pdf, xlsx, tmp_path / "w")
    assert result.covered_words_in_pdf == 4  # "N/A" x 2 = tokens n, a
    assert {"n", "a"} <= set(result.extra_examples)


def _result(**over: Any) -> FidelityResult:
    base: dict[str, Any] = {
        "pdf": "a.pdf", "xlsx": "a.xlsx", "pdf_pages": 1, "sheets": 1, "render_pages": 1,
        "words_expected": 0, "words_found": 0, "recall": None, "recall_any_page": None,
        "extra_words": 0, "extra_examples": [], "missing_examples": [], "pos_matched": 0,
        "pos_err_median_pct": None, "pos_err_p90_pct": None, "dx_median_pct": None, "dy_median_pct": None,
        "pdf_border_segments": 0, "xlsx_border_segments": 0, "xlsx_bordered_cells": 0, "border_ratio": None,
        "mixed_cells": 0, "mixed_examples": [], "pdf_images": 0, "xlsx_images": 0, "covered_words_in_pdf": 0,
    }  # fmt: skip
    base.update(over)
    return FidelityResult(**base)


def test_missing_borders_are_a_problem_even_without_text_layer() -> None:
    """Un escaneo con cajas cuyo Excel no tiene ninguna ya no da "OK" por comparar 0 contra 0."""
    lost = _result(pdf_border_segments=40, xlsx_border_segments=0, border_ratio=0.0)
    assert any("faltan bordes" in p for p in lost.problems()) and not lost.passes
    kept = _result(pdf_border_segments=40, xlsx_border_segments=52, border_ratio=1.3)
    assert kept.passes
    tiny = _result(pdf_border_segments=2, xlsx_border_segments=0, border_ratio=0.0)
    assert tiny.passes  # con 2 líneas sueltas no vale la pena exigir


def test_raster_border_count_on_a_scanned_page(tmp_path: Path) -> None:
    from tests.layout_score import _raster_border_count

    src = pymupdf.open()
    sp = src.new_page(width=300, height=200)
    sp.draw_rect(pymupdf.Rect(30, 30, 270, 120), color=(0.3, 0.3, 0.3), width=1.0)
    sp.draw_line((30, 75), (270, 75), color=(0.3, 0.3, 0.3), width=1.0)
    pix = sp.get_pixmap(dpi=150, colorspace=pymupdf.csGRAY, alpha=False)
    page = pymupdf.open().new_page(width=300, height=200)
    page.insert_image(page.rect, pixmap=pix)
    assert _raster_border_count(page) == 5  # 3 horizontales y 2 verticales
