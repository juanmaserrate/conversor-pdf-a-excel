"""``core/mode.py``: elección "Igual al PDF" / "Solo tablas" calibrada con los dos corpus sintéticos."""

# ruff: noqa: F811  (las fixtures importadas se "redefinen" como argumentos de test)
from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
from tests.forms_fixtures import forms_corpus, forms_ground_truth  # noqa: F401

from pdf2excel.core import pipeline
from pdf2excel.core.mode import (
    MIN_OUTSIDE_BORDERED,
    REASON_TABLES_NONE,
    ModeFeatures,
    decide,
    mode_features,
    suggest_mode,
)
from pdf2excel.models import BorderSide, LayoutCell, LayoutPage, Options, TypedTable

# Todo lo que dice la regla sobre los formularios coincide con ``expected_modes`` del ground truth.
FORMS_LAYOUT = (
    "hoja_ruta",
    "remito_cajas",
    "factura_bloques",
    "form_dos_columnas",
    "tabla_seccion_combinada",  # decisión documentada en core/mode.py: bandas de color → layout
    "pagina_rotada",
)
FORMS_TABLES = ("solo_texto_alineado", "multipagina_listado", "escaneado_formulario")

# Listados, extractos y comprobantes de tabla dominante: siguen en "Solo tablas".
CORPUS_TABLES = (
    "hard_extracto_banco",
    "cont_sin_header",
    "hard_lista_mixta",
    "negativos_banco",
    "hard_factura_ar",
    "header_dos_filas",
    "apaisado",
    "numeros_us",
    "porcentajes",
    "remito_grid",
    "remito_nogrid",
    "sin_tablas",
)


def _opts(tmp_path: Path, **kw: Any) -> Options:
    return Options(output_dir=tmp_path / "out", open_when_done=False, **kw)


# --------------------------------------------------------------------------- corpus real


@pytest.mark.parametrize("name", FORMS_LAYOUT + FORMS_TABLES)
def test_forms_corpus_modes(
    name: str,
    forms_corpus: dict[str, Path],
    forms_ground_truth: dict[str, Any],
    tmp_path: Path,
) -> None:
    pytest.importorskip("pymupdf")
    expected = forms_ground_truth[f"{name}.pdf"]["expected_modes"]["mode"]
    report = pipeline.analyze(forms_corpus[name], _opts(tmp_path))
    assert report.suggested_mode == expected, report.mode_reason
    assert report.mode == expected
    assert report.mode_reason  # frase para la UI


@pytest.mark.parametrize("name", CORPUS_TABLES)
def test_list_corpus_stays_in_tables(name: str, corpus_dir: Path, tmp_path: Path) -> None:
    report = pipeline.analyze(corpus_dir / f"{name}.pdf", _opts(tmp_path))
    assert report.suggested_mode == "tables", report.mode_reason
    assert report.mode == "tables"


def test_reason_phrases_in_spanish(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    form = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path))
    listing = pipeline.analyze(forms_corpus["multipagina_listado"], _opts(tmp_path))
    assert "conviene respetar el diseño" in form.mode_reason
    assert "listado" in listing.mode_reason
    assert form.mode_reason != listing.mode_reason


def test_forced_modes_override_suggestion(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    forced_tables = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path, layout="tables"))
    assert forced_tables.mode == "tables"
    assert forced_tables.layout_pages == []  # "Solo tablas" explícito no paga el costo del diseño
    assert "layout" not in forced_tables.timings

    forced_layout = pipeline.analyze(forms_corpus["multipagina_listado"], _opts(tmp_path, layout="layout"))
    assert forced_layout.mode == "layout"
    assert len(forced_layout.layout_pages) == 3


def test_form_features(forms_corpus: dict[str, Path], tmp_path: Path) -> None:
    # La regla se ve en los números: la hoja de ruta tiene muchas celdas con borde fuera de las tablas.
    report = pipeline.analyze(forms_corpus["hoja_ruta"], _opts(tmp_path))
    assert report.layout_pages
    features = mode_features(report)  # sin bboxes de tabla: todo lo que tiene borde cuenta "afuera"
    assert features.outside_bordered >= MIN_OUTSIDE_BORDERED
    assert features.n_pages == 2


# --------------------------------------------------------------------------- regla sobre números


def _features(**kw: Any) -> ModeFeatures:
    base: dict[str, Any] = {
        "n_pages": 1,
        "n_tables": 1,
        "multipage_table": False,
        "ocr_applied": False,
        "outside_bordered": 0,
        "table_bands": 0,
        "table_word_ratio": 0.9,
    }
    base.update(kw)
    return ModeFeatures(**base)


def test_decide_rules_in_order() -> None:
    assert decide(_features(n_pages=0))[0] == "tables"
    assert decide(_features(n_pages=0))[1] == REASON_TABLES_NONE
    # una tabla que sigue en otra página gana aunque haya cajas
    assert decide(_features(multipage_table=True, outside_bordered=50))[0] == "tables"
    assert decide(_features(ocr_applied=True, outside_bordered=50))[0] == "tables"
    assert decide(_features(outside_bordered=MIN_OUTSIDE_BORDERED))[0] == "layout"
    assert decide(_features(outside_bordered=MIN_OUTSIDE_BORDERED - 1))[0] == "tables"
    assert decide(_features(table_bands=3))[0] == "layout"
    assert decide(_features(n_tables=0))[0] == "tables"
    assert decide(_features(n_tables=0, outside_bordered=20))[0] == "layout"


def _page(n_bordered: int) -> LayoutPage:
    side = BorderSide(0.5)
    cells = [
        LayoutCell(row=i, col=0, border_left=side, bbox=(10.0, 20.0 * i, 60.0, 20.0 * i + 15.0))
        for i in range(n_bordered)
    ]
    return LayoutPage(
        page=0, width_pt=595, height_pt=842, col_widths_pt=[595], row_heights_pt=[20] * 3, cells=cells
    )


def _table(pages: tuple[int, int] = (0, 0)) -> TypedTable:
    return TypedTable(name="T", header=["a", "b"], rows=[], columns=[], pages=pages)


def test_suggest_mode_accepts_report_like_objects() -> None:
    form = SimpleNamespace(layout_pages=[_page(9)], tables=[_table()], ocr_applied=False)
    assert suggest_mode(form)[0] == "layout"
    # una tabla que cubre esas celdas las saca de "afuera"
    covering = {0: [(0.0, 0.0, 100.0, 400.0)]}
    assert suggest_mode(form, covering)[0] == "tables"
    nothing = SimpleNamespace(layout_pages=[], tables=[], ocr_applied=False)
    assert suggest_mode(nothing) == ("tables", REASON_TABLES_NONE)


def test_ground_truth_covers_every_form(
    forms_corpus: dict[str, Path], forms_ground_truth: dict[str, Any]
) -> None:
    assert set(FORMS_LAYOUT + FORMS_TABLES) == {n for n in forms_corpus}
    assert json.dumps(forms_ground_truth)  # serializable
