"""Calidad de extracción contra el corpus sintético (DESIGN.md §9), con umbrales que rompen el build.

Cada PDF se convierte con ``pipeline.convert`` y se puntúa con ``tests/scorer.py`` contra
``tests/corpus/ground_truth.json``. Los escaneos (``_scan*``) son ``slow`` (OCR) y se saltean sin tesseract.
Los resultados se acumulan para ``--corpus-report``.
"""

from __future__ import annotations

import hashlib
import time
from collections.abc import Mapping
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
import pytest
from tests.corpus import generate as gen
from tests.scorer import (
    TYPED_KINDS,
    Score,
    format_table,
    gt_for,
    gt_rows_with_header,
    match_tables,
    mismatches,
    normalize,
    score,
)

from pdf2excel.core import pipeline
from pdf2excel.core.document import classify_document, open_document
from pdf2excel.errors import UserError
from pdf2excel.models import ColumnInfo, ConversionReport, Options, TypedTable

# --------------------------------------------------------------------------- umbrales

STRICT_CELLS = 99.0
STRICT_SPURIOUS = 1.0
STRICT_TYPES = 99.0
RELAXED_CELLS = 95.0  # casos nuevos: subir a 99 cuando pasen

STRICT_DIGITAL = (
    "remito_grid",
    "remito_nogrid",
    "prueba",
    "hard_factura_ar",
    "hard_extracto_banco",
    "hard_lista_mixta",
    "con_password",
)
RELAXED_DIGITAL = ("header_dos_filas", "apaisado", "porcentajes", "numeros_us", "negativos_banco")
SCAN_MIN_CELLS = {
    "hard_factura_ar_scan": 92.0,
    "remito_grid_scan": 92.0,
    "remito_nogrid_scan": 92.0,
    "cont_sin_header_scan": 92.0,
    "hard_lista_mixta_scan": 88.0,
    "hard_factura_ar_scan90": 90.0,
    "remito_grid_scan150": 85.0,
}
# Limitaciones conocidas del core (no del corpus) con su umbral acotado: el reporte las muestra igual.
# Excepciones conocidas al umbral de filas espurias (archivo -> % tolerado). Vacío: extend_tail ya exige un
# valor numérico/fecha en el renglón de cola, así que no se traga la prosa que sigue a una tabla.
KNOWN_SPURIOUS_MAX: dict[str, float] = {}

WITH_TABLES_DIGITAL = STRICT_DIGITAL + RELAXED_DIGITAL + ("cont_sin_header", "mixto")


# --------------------------------------------------------------------------- helpers


def _opts(tmp_path: Path) -> Options:
    return Options(output_dir=tmp_path / "out", open_when_done=False)


def _convert(
    corpus_dir: Path,
    name: str,
    tmp_path: Path,
    ground_truth: dict[str, Any],
    results: dict[str, Score],
    meta: dict[str, dict[str, Any]],
    password: str | None = None,
) -> tuple[ConversionReport, Score]:
    """Convierte ``<name>.pdf``, lo puntúa contra su GT y registra el resultado para el reporte."""
    src = corpus_dir / f"{name}.pdf"
    t0 = time.perf_counter()
    report = pipeline.convert(src, _opts(tmp_path), progress=None, password=password)
    seconds = time.perf_counter() - t0
    gt = gt_for(name, ground_truth)
    result = score(report.tables, gt)
    results[src.name] = result
    meta[src.name] = {"ocr": report.ocr_applied, "seconds": round(seconds, 2)}
    return report, result


def _explain(name: str, result: Score, report: ConversionReport, gt: dict[str, Any]) -> str:
    shapes = [(len(t.rows), len(t.header)) for t in report.tables]
    lines = [f"{name}: {result} formas={shapes}", *mismatches(report.tables, gt, limit=8)]
    return "\n".join(lines)


def _xlsx_table_sheets(path: Path, n_tables: int) -> list[Any]:
    wb = openpyxl.load_workbook(path)
    return wb.worksheets[:n_tables]  # el writer crea primero una hoja por tabla, en orden


# --------------------------------------------------------------------------- scorer (acceptance)


def _table_from_gt(gt_table: dict[str, Any]) -> TypedTable:
    """TypedTable idéntica a una tabla GT (valores tipados)."""
    rows = gt_rows_with_header(gt_table)
    header = [str(c["v"]) for c in rows[0] if c is not None]

    def value(cell: Mapping[str, Any] | None) -> Any:
        if cell is None:
            return None
        if cell["t"] == "date":
            return date.fromisoformat(cell["v"])
        return cell["v"]

    body = [[value(c) for c in r] for r in rows[1:]]
    columns = [ColumnInfo(kind="text", header=h) for h in header]
    return TypedTable(name="GT", header=header, rows=body, columns=columns)


@pytest.fixture(scope="module")
def remito_gt(ground_truth: dict[str, Any]) -> dict[str, Any]:
    return gt_for("remito_grid", ground_truth)


def test_scorer_identical_table_is_perfect(remito_gt: dict[str, Any]) -> None:
    table = _table_from_gt(remito_gt["tables"][0])
    s = score([table], remito_gt)
    assert (s.cells_pct, s.rows_pct, s.spurious_rows_pct) == (100.0, 100.0, 0.0)
    assert s.spurious_cols == 0
    assert s.type_pct == 100.0
    assert (s.tables_detected, s.tables_expected) == (1, 1)


def test_scorer_spurious_row(remito_gt: dict[str, Any]) -> None:
    table = _table_from_gt(remito_gt["tables"][0])
    table.rows.append(["Página 1 de 1", None, None, None, None])
    s = score([table], remito_gt)
    assert s.cells_pct == 100.0
    assert s.spurious_rows_pct > 0
    assert s.rows_pct == 100.0


def test_scorer_untyped_number_lowers_type_pct(remito_gt: dict[str, Any]) -> None:
    table = _table_from_gt(remito_gt["tables"][0])
    table.rows[0][3] = "350.00"  # columna "Precio unitario" (float) quedó como texto
    s = score([table], remito_gt)
    assert s.cells_pct == 100.0  # el contenido sigue siendo correcto
    assert s.type_pct < 100.0
    typed_cells = sum(1 for r in remito_gt["tables"][0]["rows"] for c in r if c and c["t"] in TYPED_KINDS)
    assert s.type_pct == pytest.approx(100.0 * (typed_cells - 1) / typed_cells, abs=0.01)


def test_scorer_misc(remito_gt: dict[str, Any]) -> None:
    assert normalize(None) == ""
    assert normalize(" Mezcal   Joven ") == "mezcal joven"
    assert normalize(350) == 350.0 and normalize(1234.567) == 1234.57
    assert normalize(date(2026, 10, 15)) == "2026-10-15"
    assert normalize(datetime(2026, 10, 15, 8, 30)) == "2026-10-15"
    # sin tablas extraídas: 0 % de celdas, tablas 0/1
    s = score([], remito_gt)
    assert s.cells_pct == 0.0 and s.tables_detected == 0 and s.tables_expected == 1
    # columnas de más
    table = _table_from_gt(remito_gt["tables"][0])
    table.header.append("")
    for r in table.rows:
        r.append("x")
    assert score([table], remito_gt).spurious_cols == 1
    md = format_table({"remito_grid.pdf": s}, {"remito_grid.pdf": {"ocr": False, "seconds": 0.42}})
    assert "| remito_grid.pdf |" in md and "| 0/1 |" in md and "| no |" in md and "0,4" in md


def test_gt_for_strips_scan_suffixes(ground_truth: dict[str, Any]) -> None:
    base = gt_for("hard_factura_ar", ground_truth)
    for name in (
        "hard_factura_ar_scan.pdf",
        "hard_factura_ar_scan90",
        "hard_factura_ar_scan150.pdf",
        "hard_factura_ar_ocr.pdf",
    ):
        assert gt_for(name, ground_truth) == base
    with pytest.raises(KeyError):
        gt_for("inexistente.pdf", ground_truth)


# --------------------------------------------------------------------------- corpus / GT consistency


@pytest.fixture(scope="module")
def regenerated(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[Path], dict[str, Any]]:
    out = tmp_path_factory.mktemp("corpus-again")
    return gen.generate_with_ground_truth(out)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_generator_is_deterministic(corpus_dir: Path, regenerated: tuple[list[Path], dict[str, Any]]) -> None:
    paths, _ = regenerated
    assert len(paths) >= 22
    assert sorted(p.name for p in paths) == sorted(f"{n}.pdf" for n in gen.ALL_NAMES)
    for p in paths:
        other = corpus_dir / p.name
        assert other.is_file(), other
        if p.name in gen.NON_DETERMINISTIC:
            continue  # AES-256 usa sales aleatorias: mismo contenido, bytes distintos
        assert _sha256(p) == _sha256(other), f"{p.name} cambia entre corridas"


def test_ground_truth_json_in_sync(
    ground_truth: dict[str, Any], regenerated: tuple[list[Path], dict[str, Any]]
) -> None:
    """ground_truth.json debe ser exactamente lo que produce el generador (regenerar con --ground-truth)."""
    _, fresh = regenerated
    assert ground_truth == fresh


def test_ground_truth_shape(ground_truth: dict[str, Any]) -> None:
    for name in gen.DIGITAL_NAMES:
        entry = ground_truth[f"{name}.pdf"]
        assert entry["expected_tables"] == len(entry["tables"])
        assert entry["kinds"] and all(
            k in {"digital", "scanned", "empty", "garbage_text"} for k in entry["kinds"]
        )
        for table in entry["tables"]:
            ncols = len(table["header"])
            assert ncols >= 2 and table["rows"]
            for row in table["rows"]:
                assert len(row) == ncols
                for cell in row:
                    assert cell is None or (cell["t"] in {"float", "int", "str", "date"} and "v" in cell)
                    if cell is not None and cell["t"] == "date":
                        date.fromisoformat(cell["v"])
    for name in gen.SCAN_NAMES:
        assert f"{name}.pdf" not in ground_truth  # heredan el GT del original
        gt_for(name, ground_truth)


# --------------------------------------------------------------------------- digitales


@pytest.mark.parametrize("name", STRICT_DIGITAL)
def test_digital_strict(
    name: str,
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    gt = gt_for(name, ground_truth)
    report, s = _convert(
        corpus_dir, name, tmp_path, ground_truth, corpus_results, corpus_meta, gt.get("password")
    )
    detail = _explain(name, s, report, gt)
    assert report.ocr_applied is False and all(p.kind == "digital" for p in report.pages), detail
    assert s.tables_detected == s.tables_expected, detail
    assert s.cells_pct >= STRICT_CELLS, detail
    assert s.spurious_rows_pct <= KNOWN_SPURIOUS_MAX.get(name, STRICT_SPURIOUS), detail
    assert s.type_pct >= STRICT_TYPES, detail
    assert report.output_path is not None and report.output_path.is_file()


@pytest.mark.parametrize("name", RELAXED_DIGITAL)
def test_digital_relaxed(
    name: str,
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    gt = gt_for(name, ground_truth)
    report, s = _convert(corpus_dir, name, tmp_path, ground_truth, corpus_results, corpus_meta)
    detail = _explain(name, s, report, gt)
    assert s.tables_detected == s.tables_expected, detail
    assert s.cells_pct >= RELAXED_CELLS, detail
    assert s.type_pct >= STRICT_TYPES, detail


def test_cont_sin_header_is_one_merged_table(
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    gt = gt_for("cont_sin_header", ground_truth)
    report, s = _convert(corpus_dir, "cont_sin_header", tmp_path, ground_truth, corpus_results, corpus_meta)
    detail = _explain("cont_sin_header", s, report, gt)
    assert len(report.pages) == 3
    assert len(report.tables) == 1, detail
    table = report.tables[0]
    assert len(table.rows) == len(gt["tables"][0]["rows"]), detail
    assert table.pages == (0, 2)
    assert s.cells_pct >= STRICT_CELLS and s.spurious_rows_pct <= STRICT_SPURIOUS, detail


# --------------------------------------------------------------------------- escaneos (OCR)


@pytest.mark.slow
@pytest.mark.parametrize("name", gen.SCAN_NAMES)
def test_scanned_with_ocr(
    name: str,
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    has_tesseract: bool,
    ocr_cache_dir: Path,
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    if not has_tesseract:
        pytest.skip("tesseract no disponible")
    gt = gt_for(name, ground_truth)
    report, s = _convert(corpus_dir, name, tmp_path, ground_truth, corpus_results, corpus_meta)
    detail = _explain(name, s, report, gt)
    assert report.ocr_applied is True, detail
    assert all(p.kind == "scanned_with_text" for p in report.pages), detail
    assert s.tables_detected == s.tables_expected, detail
    assert s.cells_pct >= SCAN_MIN_CELLS[name], detail
    if name == "cont_sin_header_scan":
        assert len(report.tables[0].rows) >= 0.98 * len(gt["tables"][0]["rows"]), detail


# --------------------------------------------------------------------------- casos especiales


def test_con_password_requires_password(corpus_dir: Path, tmp_path: Path) -> None:
    with pytest.raises(UserError) as info:
        pipeline.convert(corpus_dir / "con_password.pdf", _opts(tmp_path))
    assert info.value.code == "encrypted"
    assert info.value.action == "password"
    assert not list((tmp_path / "out").glob("*.xlsx")) if (tmp_path / "out").exists() else True
    with pytest.raises(UserError) as wrong:
        pipeline.convert(corpus_dir / "con_password.pdf", _opts(tmp_path), password="incorrecta")
    assert wrong.value.code == "encrypted"


def test_sin_tablas(
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    report, s = _convert(corpus_dir, "sin_tablas", tmp_path, ground_truth, corpus_results, corpus_meta)
    assert report.tables == []
    assert s.tables_detected == 0 == s.tables_expected
    assert "no_tables" in report.warnings
    assert report.pages[0].kind == "digital"
    assert any("Estimado cliente" in f"{it.key or ''} {it.value}" for it in report.text_items)  # sale como kv
    assert report.output_path is not None
    wb = openpyxl.load_workbook(report.output_path)
    assert "Texto" in wb.sheetnames


def test_vacio(
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    report, _ = _convert(corpus_dir, "vacio", tmp_path, ground_truth, corpus_results, corpus_meta)
    assert [p.kind for p in report.pages] == ["empty"]
    assert report.tables == [] and report.ocr_applied is False
    assert report.output_path is not None and report.output_path.is_file()


def test_mixto(
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    has_tesseract: bool,
    ocr_cache_dir: Path,
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    src = corpus_dir / "mixto.pdf"
    with open_document(src) as doc:
        kinds = [p.kind for p in classify_document(doc)]
    assert kinds == ["digital", "scanned"]

    gt = gt_for("mixto", ground_truth)
    report, s = _convert(corpus_dir, "mixto", tmp_path, ground_truth, corpus_results, corpus_meta)
    detail = _explain("mixto", s, report, gt)
    assert report.pages[0].kind == "digital"
    assert report.ocr_applied is has_tesseract
    assert report.pages[1].kind == ("scanned_with_text" if has_tesseract else "scanned")
    assert s.tables_detected == 1 == s.tables_expected, detail  # la página escaneada no tiene tablas
    assert s.cells_pct >= STRICT_CELLS and s.type_pct >= STRICT_TYPES, detail
    assert report.tables[0].pages == (0, 0)


def test_texto_basura(
    corpus_dir: Path,
    ground_truth: dict[str, Any],
    ocr_cache_dir: Path,
    corpus_results: dict[str, Score],
    corpus_meta: dict[str, dict[str, Any]],
    tmp_path: Path,
) -> None:
    src = corpus_dir / "texto_basura.pdf"
    with open_document(src) as doc:
        infos = classify_document(doc)
    assert [p.kind for p in infos] == ["garbage_text"]
    assert infos[0].n_chars >= 20
    report, _ = _convert(corpus_dir, "texto_basura", tmp_path, ground_truth, corpus_results, corpus_meta)
    assert report.pages[0].kind in {"garbage_text", "scanned_with_text"}  # con tesseract pasa por OCR
    assert report.output_path is not None and report.output_path.is_file()


# --------------------------------------------------------------------------- relectura del .xlsx


def _gt_numeric_columns(gt_table: dict[str, Any]) -> list[int]:
    """Columnas cuyas celdas GT no vacías son todas float/int/date (fila de total con texto no cuenta)."""
    ncols = len(gt_table["header"])
    out = []
    for c in range(ncols):
        cells = [r[c] for r in gt_table["rows"] if r[c] is not None]
        if cells and all(cell["t"] in TYPED_KINDS for cell in cells):
            out.append(c)
    return out


@pytest.mark.parametrize("name", WITH_TABLES_DIGITAL)
def test_xlsx_numeric_columns_are_not_text(
    name: str, corpus_dir: Path, ground_truth: dict[str, Any], ocr_cache_dir: Path, tmp_path: Path
) -> None:
    gt = gt_for(name, ground_truth)
    report = pipeline.convert(corpus_dir / f"{name}.pdf", _opts(tmp_path), password=gt.get("password"))
    assert report.output_path is not None
    sheets = _xlsx_table_sheets(report.output_path, len(report.tables))
    assert len(sheets) == len(report.tables)
    chosen = match_tables(report.tables, gt)
    checked = 0
    for gt_table, idx in zip(gt["tables"], chosen, strict=True):
        assert idx is not None
        table = report.tables[idx]
        ws = sheets[idx]
        assert ws.freeze_panes == "A2"
        assert [c.value or "" for c in ws[1]][: len(table.header)] == table.header  # vacío → None en xlsx
        for c in _gt_numeric_columns(gt_table):
            for r in range(len(table.rows)):
                xlsx_value = ws.cell(row=r + 2, column=c + 1).value
                if isinstance(xlsx_value, str):
                    assert (r, c) in table.review_cells, (
                        f"{name} hoja {ws.title!r} celda ({r + 2}, {c + 1}) = {xlsx_value!r} es texto en "
                        "columna numérica y no está marcada para revisar"
                    )
                elif xlsx_value is not None:
                    assert isinstance(xlsx_value, int | float | datetime | date)
                    checked += 1
                    if not isinstance(xlsx_value, datetime | date):
                        assert normalize(xlsx_value) == normalize(table.rows[r][c])
    assert checked > 0
