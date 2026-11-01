"""Integración del modo "Igual al PDF" en el pipeline, el writer y la CLI (LAYOUT_MODE.md §5 y §6)."""

# ruff: noqa: F811  (las fixtures importadas se "redefinen" como argumentos de test)
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any

import openpyxl
import pymupdf
import pytest
from openpyxl.worksheet.worksheet import Worksheet
from tests.forms_fixtures import forms_corpus, forms_ground_truth  # noqa: F401
from tests.layout_score import LibreOfficeUnavailable, score

from pdf2excel.core import pipeline
from pdf2excel.core.ocr import ocr_available
from pdf2excel.errors import UserError
from pdf2excel.excel.writer import SHEET_DOC_DATA, SHEET_ORIGIN, build_workbook
from pdf2excel.models import ConversionReport, Options

ENV = {**os.environ, "QT_QPA_PLATFORM": "offscreen"}


def _opts(tmp_path: Path, **kw: Any) -> Options:
    return Options(output_dir=tmp_path / "out", open_when_done=False, **kw)


def _all_text(wb: openpyxl.Workbook) -> list[str]:
    out: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            out.extend(str(v) for v in row if v is not None)
    return out


# --------------------------------------------------------------------------- texto tapado en "Solo tablas"


def test_covered_text_never_shows_in_tables_mode(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    """El PDF dibuja dos "N/A" y los tapa con un rectángulo blanco: no deben aparecer en ningún lado."""
    report = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path, layout="tables"))
    assert report.mode == "tables"
    assert report.tables
    cells = [str(c) for t in report.tables for row in (t.header, *t.rows) for c in row if c is not None]
    texts = [it.value for it in report.text_items] + [it.key or "" for it in report.text_items]
    assert not [x for x in cells + texts if "N/A" in x]
    wb = openpyxl.load_workbook(pipeline.export(report, _opts(tmp_path, layout="tables")))
    assert not [x for x in _all_text(wb) if "N/A" in x]
    assert SHEET_DOC_DATA in wb.sheetnames  # el resto del texto sí está


def test_hidden_text_is_recorded_in_layout_mode(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    report = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path, layout="layout"))
    assert sum(p.hidden_texts.count("N/A") for p in report.layout_pages) == 2
    dst = pipeline.export(report, _opts(tmp_path, layout="layout", show_origin_sheet=True))
    wb = openpyxl.load_workbook(dst)
    origin = wb[SHEET_ORIGIN]
    values = [str(v) for row in origin.iter_rows(values_only=True) for v in row if v is not None]
    assert "Igual al PDF" in values  # modo
    assert "N/A" in values  # textos tapados descartados (diagnóstico)
    assert "Página 1" in values
    for ws in wb.worksheets:
        if ws.title != SHEET_ORIGIN:
            assert not [v for row in ws.iter_rows(values_only=True) for v in row if v and "N/A" in str(v)]


# --------------------------------------------------------------------------- modo layout


def test_layout_export_one_sheet_per_page(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    report = pipeline.convert(forms_corpus["hoja_ruta"], _opts(tmp_path))
    assert report.mode == report.suggested_mode == "layout"
    assert report.output_path is not None
    wb = openpyxl.load_workbook(report.output_path)
    visible = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]
    assert [ws.title for ws in visible] == ["Página 1", "Página 2"]
    assert wb[SHEET_ORIGIN].sheet_state == "hidden"
    assert "Revisión" not in wb.sheetnames and SHEET_DOC_DATA not in wb.sheetnames
    assert {"layout"} <= set(report.timings)
    first: Worksheet = visible[0]
    assert first.sheet_view.showGridLines is False


def test_mode_can_be_switched_before_export(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    report = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path))
    assert report.mode == "layout"
    report.mode = "tables"  # lo que hace el selector de la UI
    wb = build_workbook(report, _opts(tmp_path))
    assert not any(ws.title.startswith("Página") for ws in wb.worksheets)
    assert len(report.layout_pages) == 2  # el diseño sigue disponible para volver a "layout"
    report.mode = "layout"
    wb2 = build_workbook(report, _opts(tmp_path))
    assert [ws.title for ws in wb2.worksheets if ws.sheet_state == "visible"] == ["Página 1", "Página 2"]


def test_no_images_option(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    with_images = build_workbook(
        pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path)), _opts(tmp_path)
    )
    without = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path, layout_images=False))
    assert all(not p.images for p in without.layout_pages)
    plain = build_workbook(without, _opts(tmp_path, layout_images=False))
    assert sum(len(ws._images) for ws in with_images.worksheets) == 3  # type: ignore[attr-defined]
    assert sum(len(ws._images) for ws in plain.worksheets) == 0  # type: ignore[attr-defined]


def test_layout_cancel_during_layout_stage(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    cancel = threading.Event()

    def progress(_c: int, _t: int, stage: str) -> None:
        if stage == "layout":
            cancel.set()

    with pytest.raises(UserError) as info:
        pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path), progress=progress, cancel=cancel)
    assert info.value.code == "cancelled"


def test_long_documents_skip_layout_unless_forced(tmp_path: Path) -> None:
    pdf = tmp_path / "largo.pdf"
    doc = pymupdf.open()
    for i in range(pipeline.LAYOUT_MAX_PAGES + 1):
        page = doc.new_page()
        page.insert_text((72, 100), f"Linea de la pagina {i + 1}", fontsize=11)
    doc.save(pdf)
    doc.close()
    auto = pipeline.analyze(pdf, _opts(tmp_path))
    assert auto.layout_pages == []
    assert auto.suggested_mode == auto.mode == "tables"
    forced = pipeline.analyze(pdf, _opts(tmp_path, layout="layout"))
    assert len(forced.layout_pages) == pipeline.LAYOUT_MAX_PAGES + 1
    assert forced.mode == "layout"


def test_layout_forced_without_pages_falls_back_to_tables(tmp_path: Path) -> None:
    pdf = tmp_path / "vacio.pdf"
    doc = pymupdf.open()
    doc.new_page()
    doc.save(pdf)
    doc.close()
    report = pipeline.analyze(pdf, _opts(tmp_path, layout="layout"))
    assert report.mode == "tables"
    assert report.mode_reason


@pytest.mark.slow
def test_scanned_form_with_ocr_gets_layout(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    if not ocr_available():
        pytest.skip("tesseract no disponible")
    report = pipeline.analyze(forms_corpus["escaneado_formulario"], _opts(tmp_path, layout="layout"))
    assert report.ocr_applied
    assert report.pages[0].kind == "scanned_with_text"
    assert report.layout_pages
    assert report.layout_pages[0].cells  # el texto del OCR se volvió celdas (scanned=True)
    assert report.layout_pages[0].images == []  # la imagen de fondo del escaneo no se incluye
    auto = pipeline.analyze(forms_corpus["escaneado_formulario"], _opts(tmp_path))
    assert auto.suggested_mode == "tables"


# --------------------------------------------------------------------------- lotes


def test_export_batch_mixes_modes(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    opts = _opts(tmp_path, single_workbook=True)
    form = pipeline.analyze(forms_corpus["hoja_ruta"], opts)
    listing = pipeline.analyze(forms_corpus["multipagina_listado"], opts)
    assert (form.mode, listing.mode) == ("layout", "tables")
    dst = pipeline.export_batch([form, listing], opts)
    wb = openpyxl.load_workbook(dst)
    titles = [ws.title for ws in wb.worksheets]
    assert any("hoja_ruta" in t and "Página 1" in t for t in titles)
    assert any(t.startswith("multipagina_listado") and "Página" not in t for t in titles)
    assert len(set(titles)) == len(titles)
    assert all(len(t) <= 31 for t in titles)
    assert form.output_path == listing.output_path == dst


def test_export_batch_all_layout(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    opts = _opts(tmp_path, single_workbook=True)
    reports = [pipeline.analyze(forms_corpus[n], opts) for n in ("hoja_ruta", "remito_cajas")]
    dst = pipeline.export_batch(reports, opts)
    wb = openpyxl.load_workbook(dst)
    visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    assert len(visible) == 3  # 2 + 1 páginas
    assert SHEET_DOC_DATA not in wb.sheetnames and "Texto" not in wb.sheetnames


# --------------------------------------------------------------------------- fidelidad (LibreOffice)


VECTOR_FORMS = [
    "hoja_ruta",
    "remito_cajas",
    "factura_bloques",
    "form_dos_columnas",
    "tabla_seccion_combinada",
    "pagina_rotada",
    "solo_texto_alineado",
]


@pytest.mark.slow
@pytest.mark.parametrize("name", VECTOR_FORMS)
def test_layout_fidelity_end_to_end(name: str, forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    """Umbrales ideales completos del medidor (recall, sin sobrantes, posición, sin celdas mezcladas)."""
    report = pipeline.convert(forms_corpus[name], _opts(tmp_path, layout="layout"))
    assert report.mode == "layout"
    assert report.output_path is not None
    try:
        result = score(forms_corpus[name], report.output_path, tmp_path / "work")
    except LibreOfficeUnavailable:
        pytest.skip("LibreOffice no disponible")
    assert result.passes, result.summary()
    assert not [m for m in result.mixed_examples if "TRANSPORTISTA" in m or "AUXILIAR" in m]
    assert (result.pos_err_median_pct or 0.0) <= 2.0  # margen sobre el umbral de 2.5 %


@pytest.mark.slow
def test_layout_fidelity_images_and_borders_keep_up(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    report = pipeline.convert(forms_corpus["hoja_ruta"], _opts(tmp_path, layout="layout"))
    assert report.output_path is not None
    try:
        result = score(forms_corpus["hoja_ruta"], report.output_path, tmp_path / "work")
    except LibreOfficeUnavailable:
        pytest.skip("LibreOffice no disponible")
    assert result.xlsx_images == result.pdf_images == 3
    assert result.xlsx_border_segments >= result.pdf_border_segments  # no se pierde ninguna caja


@pytest.mark.slow
def test_form_checkboxes_survive_in_the_excel(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    report = pipeline.convert(forms_corpus["form_dos_columnas"], _opts(tmp_path, layout="layout"))
    assert report.output_path is not None
    ws = openpyxl.load_workbook(report.output_path).worksheets[0]
    ticks = [c.value for row in ws.iter_rows() for c in row if c.value == "✔"]
    assert len(ticks) == 5  # las casillas tildadas del PDF (4 filas de opciones + autorizaciones)


@pytest.mark.slow
def test_aligned_numbers_become_numbers_not_glued_text(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    report = pipeline.convert(forms_corpus["solo_texto_alineado"], _opts(tmp_path, layout="layout"))
    assert report.output_path is not None
    ws = openpyxl.load_workbook(report.output_path).worksheets[0]
    texts = [str(c.value) for row in ws.iter_rows() for c in row if c.value is not None]
    assert not [
        t for t in texts if len(t.split()) == 2 and all("," in x for x in t.split())
    ]  # "8.067,50 1.952…"
    assert sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, (int, float))) >= 40


@pytest.mark.slow
def test_scanned_layout_explains_its_limits(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    if not ocr_available():
        pytest.skip("tesseract no disponible")
    report = pipeline.analyze(forms_corpus["escaneado_formulario"], _opts(tmp_path, layout="layout"))
    assert report.mode == "layout"
    assert "escaneo" in report.mode_reason and "OCR" in report.mode_reason


# --------------------------------------------------------------------------- CLI


def _run(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-m", "pdf2excel", "--cli", *args],
        capture_output=True,
        text=True,
        env=ENV,
        timeout=300,
        check=False,
    )


def test_cli_layout_flags(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    pdf = forms_corpus["hoja_ruta"]
    cases = [
        (["--layout", "auto"], "layout"),
        (["--solo-tablas"], "tables"),
        (["--layout", "tables"], "tables"),
        (["--igual-al-pdf", "--no-images"], "layout"),
    ]
    for i, (flags, expected) in enumerate(cases):
        out = tmp_path / f"o{i}"
        proc = _run(str(pdf), "-o", str(out), "--json", *flags)
        assert proc.returncode == 0, proc.stderr
        data = json.loads(proc.stdout)[0]
        assert data["mode"] == expected
        assert data["mode_reason"]
        wb = openpyxl.load_workbook(data["output"])
        has_pages = any(ws.title.startswith("Página") for ws in wb.worksheets)
        assert has_pages == (expected == "layout")
        if "--no-images" in flags:
            assert sum(len(ws._images) for ws in wb.worksheets) == 0  # type: ignore[attr-defined]


def test_cli_selftest_still_checks_tables(forms_corpus: dict[str, Path]) -> None:
    proc = subprocess.run(
        [sys.executable, "-m", "pdf2excel", "--selftest", str(forms_corpus["remito_cajas"])],
        capture_output=True,
        text=True,
        env=ENV,
        timeout=300,
        check=False,
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "selftest ok" in proc.stdout


def test_report_defaults_for_unit_construction() -> None:
    report = ConversionReport(source=Path("x.pdf"), pages=[], tables=[], text_items=[])
    assert (report.mode, report.suggested_mode, report.layout_pages) == ("tables", "tables", [])
