"""Tests de core/pipeline.py, cli.py y __main__.py (WI-05).

Los PDFs de referencia (remito_grid.pdf, remito_nogrid.pdf, remito_grid_scan.pdf) se copian del scratchpad o
de ``tests/corpus/`` a ``tmp_path``; si no están, los tests que los usan se saltean. Los casos sintéticos
(sin tablas, cifrado, pie de página repetido) se generan con pymupdf.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import threading
import types
from collections.abc import Callable
from pathlib import Path

import openpyxl
import pymupdf
import pytest

from pdf2excel import __version__
from pdf2excel.core import pipeline
from pdf2excel.core.ocr import ocr_available
from pdf2excel.errors import UserError
from pdf2excel.models import ConversionReport, Options

CORPUS_DIRS = [
    Path("/tmp/claude-0/-home-user/932cdd7a-5305-5ce6-bcbe-bcac3f3b947f/scratchpad/t"),
    Path(__file__).parent / "corpus",
    Path(__file__).parent / "corpus" / "_generated",
]

PYTHON = sys.executable
ENV = {**os.environ, "QT_QPA_PLATFORM": "offscreen"}

REMITO_ROW0 = ["A-100", "Mezcal joven", 10, 350.0, 3500.0]


def _find_corpus_pdf(name: str) -> Path | None:
    for directory in CORPUS_DIRS:
        candidate = directory / name
        if candidate.is_file():
            return candidate
    return None


@pytest.fixture
def corpus(tmp_path: Path) -> Callable[[str], Path]:
    """Copia un PDF del corpus a ``tmp_path`` (nunca se toca el original); skip si no está."""

    def _get(name: str) -> Path:
        src = _find_corpus_pdf(name)
        if src is None:
            pytest.skip(f"PDF de corpus no disponible: {name}")
        dst = tmp_path / name
        shutil.copyfile(src, dst)
        return dst

    return _get


def _text_only_pdf(path: Path, pages: list[list[str]]) -> Path:
    """PDF de solo texto (sin tablas), una lista de líneas por página."""
    doc = pymupdf.open()
    for lines in pages:
        page = doc.new_page()
        y = 72.0
        for line in lines:
            page.insert_text((72, y), line, fontsize=11)
            y += 16
    doc.save(path)
    doc.close()
    return path


def _encrypt(src: Path, dst: Path, password: str) -> Path:
    doc = pymupdf.open(src)
    doc.save(dst, encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw=password, user_pw=password)
    doc.close()
    return dst


def _opts(tmp_path: Path, **kw: object) -> Options:
    return Options(output_dir=tmp_path / "out", open_when_done=False, **kw)  # type: ignore[arg-type]


# --------------------------------------------------------------------------- analyze


@pytest.mark.parametrize("name", ["remito_grid.pdf", "remito_nogrid.pdf"])
def test_analyze_remito(corpus: Callable[[str], Path], name: str, tmp_path: Path) -> None:
    report = pipeline.analyze(corpus(name), _opts(tmp_path))

    assert isinstance(report, ConversionReport)
    assert len(report.tables) == 1
    table = report.tables[0]
    assert len(table.rows) == 6
    assert len(table.header) == 5
    assert all(len(r) == 5 for r in table.rows)
    assert [c.kind for c in table.columns[2:5]] == ["number", "number", "number"]
    assert table.rows[0] == REMITO_ROW0
    assert table.review_cells == set()
    assert report.pages[0].kind == "digital"
    assert report.ocr_applied is False
    assert report.output_path is None
    assert report.app_version == __version__
    assert "no_tables" not in report.warnings
    assert {"classify", "tables", "typing", "total"} <= set(report.timings)
    assert report.text_items  # Cliente / Fecha / Remito N°
    assert any(it.kind == "kv" and it.key == "Cliente" for it in report.text_items)


def test_analyze_title_and_sheet_name(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    report = pipeline.analyze(corpus("remito_grid.pdf"), _opts(tmp_path))
    assert report.tables[0].name  # title_near completó el título ("Remito N° 4521")


def test_analyze_untyped(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    report = pipeline.analyze(corpus("remito_grid.pdf"), _opts(tmp_path, typed_values=False))
    table = report.tables[0]
    assert all(c.kind == "text" for c in table.columns)
    assert table.rows[0] == ["A-100", "Mezcal joven", "10", "350.00", "3500.00"]


def test_include_text_false(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    report = pipeline.analyze(corpus("remito_grid.pdf"), _opts(tmp_path, include_text=False))
    assert report.text_items == []
    assert len(report.tables) == 1


# --------------------------------------------------------------------------- convert / export


def test_convert_writes_unique_paths(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    pdf = corpus("remito_grid.pdf")
    opts = _opts(tmp_path)

    first = pipeline.convert(pdf, opts)
    assert first.output_path == tmp_path / "out" / "remito_grid.xlsx"
    assert first.output_path.is_file()
    assert "excel" in first.timings

    second = pipeline.convert(pdf, opts)
    assert second.output_path == tmp_path / "out" / "remito_grid (2).xlsx"
    assert second.output_path.is_file()
    assert first.output_path.is_file()  # nunca se sobrescribe

    wb = openpyxl.load_workbook(first.output_path)
    ws = wb.worksheets[0]
    assert ws.max_row == 7 and ws.max_column == 5
    assert [c.value for c in ws[2]] == REMITO_ROW0


def test_export_next_to_pdf_by_default(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    pdf = corpus("remito_grid.pdf")
    report = pipeline.analyze(pdf, Options(open_when_done=False))
    dst = pipeline.export(report, Options(open_when_done=False))
    assert dst == tmp_path / "remito_grid.xlsx"
    assert dst.is_file()
    assert report.output_path == dst


def test_progress_stages_in_order(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    calls: list[tuple[int, int, str]] = []
    pipeline.convert(
        corpus("remito_grid.pdf"), _opts(tmp_path), progress=lambda c, t, s: calls.append((c, t, s))
    )

    stages = [s for _, _, s in calls]
    order = []
    for s in stages:
        if not order or order[-1] != s:
            order.append(s)
    assert order == ["classify", "tables", "typing", "excel"]

    for stage in order:
        of_stage = [(c, t) for c, t, s in calls if s == stage]
        totals = {t for _, t in of_stage}
        assert len(totals) == 1, f"total inconsistente en {stage}: {totals}"
        total = totals.pop()
        assert total >= 1
        assert all(0 <= c <= total for c, _ in of_stage)
        assert of_stage[-1][0] == total  # la etapa termina completa
    assert calls[0] == (1, 1, "classify")


def test_cancel_from_first_callback(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    cancel = threading.Event()
    opts = _opts(tmp_path)

    with pytest.raises(UserError) as info:
        pipeline.convert(corpus("remito_grid.pdf"), opts, progress=lambda *_: cancel.set(), cancel=cancel)
    assert info.value.code == "cancelled"
    assert not list(tmp_path.rglob("*.xlsx"))


def test_cancel_already_set(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    cancel = threading.Event()
    cancel.set()
    with pytest.raises(UserError) as info:
        pipeline.analyze(corpus("remito_grid.pdf"), _opts(tmp_path), cancel=cancel)
    assert info.value.code == "cancelled"


def test_no_tables_pdf(tmp_path: Path) -> None:
    pdf = _text_only_pdf(
        tmp_path / "carta.pdf", [["Estimado cliente:", "Le escribimos para informarle.", "Saludos."]]
    )
    opts = _opts(tmp_path)

    report = pipeline.convert(pdf, opts)
    assert report.tables == []
    assert "no_tables" in report.warnings
    assert report.output_path is not None and report.output_path.is_file()
    assert report.output_path.name == "carta.xlsx"
    assert any("Estimado cliente" in it.value for it in report.text_items)

    wb = openpyxl.load_workbook(report.output_path)
    assert "Texto" in wb.sheetnames


def test_repeated_footer_emitted_once(tmp_path: Path) -> None:
    footer = "Empresa Ejemplo S.A. - Documento no válido como factura"
    pdf = _text_only_pdf(
        tmp_path / "multi.pdf",
        [
            ["Primera página con texto propio.", footer],
            ["Segunda página, otro texto.", footer],
            ["Tercera.", footer],
        ],
    )
    report = pipeline.analyze(pdf, _opts(tmp_path))
    values = [it.value for it in report.text_items]
    assert values.count(footer) == 1
    assert "Primera página con texto propio." in values
    assert "Tercera." in values
    assert len(report.pages) == 3


def test_encrypted_without_password(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    pdf = _encrypt(corpus("remito_grid.pdf"), tmp_path / "secreto.pdf", "clave123")
    with pytest.raises(UserError) as info:
        pipeline.analyze(pdf, _opts(tmp_path))
    assert info.value.code == "encrypted"
    assert info.value.action == "password"


def test_encrypted_with_password(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    pdf = _encrypt(corpus("remito_grid.pdf"), tmp_path / "secreto.pdf", "clave123")
    report = pipeline.convert(pdf, _opts(tmp_path), password="clave123")
    assert len(report.tables) == 1
    assert report.tables[0].rows[0] == REMITO_ROW0
    assert report.output_path is not None and report.output_path.name == "secreto.xlsx"


def test_missing_file_is_user_error(tmp_path: Path) -> None:
    with pytest.raises(UserError):
        pipeline.analyze(tmp_path / "no_existe.pdf", _opts(tmp_path))


def test_not_pdf(tmp_path: Path) -> None:
    bad = tmp_path / "notas.txt"
    bad.write_text("hola")
    with pytest.raises(UserError) as info:
        pipeline.analyze(bad, _opts(tmp_path))
    assert info.value.code == "not_pdf"


def test_unexpected_exception_becomes_user_error(
    corpus: Callable[[str], Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def boom(*_: object, **__: object) -> dict:
        raise ZeroDivisionError("explotó")

    monkeypatch.setattr(pipeline, "page_candidates", boom)
    with pytest.raises(UserError) as info:
        pipeline.analyze(corpus("remito_grid.pdf"), _opts(tmp_path))
    assert info.value.code == "unknown"
    assert isinstance(info.value.cause, ZeroDivisionError)


# --------------------------------------------------------------------------- OCR


def test_ocr_unavailable_degrades(
    corpus: Callable[[str], Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    pdf = corpus("remito_grid_scan.pdf")
    monkeypatch.setattr(pipeline, "ocr_available", lambda: False)
    calls: list[str] = []
    monkeypatch.setattr(pipeline, "ocr_document", lambda *a, **k: calls.append("ocr"))

    report = pipeline.convert(pdf, _opts(tmp_path))
    assert calls == []
    assert report.ocr_applied is False
    assert report.pages[0].kind == "scanned"
    assert "ocr_unavailable" in report.warnings
    assert "no_tables" in report.warnings
    assert report.tables == []
    assert report.output_path is not None and report.output_path.is_file()


def test_force_ocr_without_tesseract_keeps_digital_pages(
    corpus: Callable[[str], Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Forzar OCR sin tesseract no puede vaciar un PDF digital: sus páginas siguen con su texto."""
    pdf = corpus("remito_grid.pdf")
    monkeypatch.setattr(pipeline, "ocr_available", lambda: False)
    monkeypatch.setattr(pipeline, "ocr_document", lambda *a, **k: pytest.fail("no debe llamar a OCR"))

    report = pipeline.analyze(pdf, _opts(tmp_path, force_ocr=True))
    assert "ocr_unavailable" in report.warnings and "no_tables" not in report.warnings
    assert len(report.tables) == 1 and len(report.tables[0].rows) == 6
    assert report.ocr_applied is False


def test_export_explicit_destination(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    """'Guardar con otro nombre': el destino se respeta tal cual (no <pdf>.xlsx en esa carpeta)."""
    report = pipeline.analyze(corpus("remito_grid.pdf"), _opts(tmp_path))
    chosen = tmp_path / "otra" / "remito_para_contabilidad.xlsx"
    chosen.parent.mkdir()
    out = pipeline.export(report, _opts(tmp_path), dst=chosen)
    assert out == chosen and chosen.is_file() and report.output_path == chosen
    assert sorted(p.name for p in chosen.parent.iterdir()) == ["remito_para_contabilidad.xlsx"]


@pytest.mark.slow
@pytest.mark.skipif(not ocr_available(), reason="tesseract no disponible")
def test_analyze_scanned_with_ocr(
    corpus: Callable[[str], Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import pdf2excel.core.ocr as ocr_mod

    monkeypatch.setattr(ocr_mod, "default_cache_dir", lambda: tmp_path / "cache")
    calls: list[tuple[int, int, str]] = []

    report = pipeline.analyze(
        corpus("remito_grid_scan.pdf"), _opts(tmp_path), progress=lambda c, t, s: calls.append((c, t, s))
    )

    assert report.ocr_applied is True
    assert report.pages[0].kind == "scanned_with_text"
    assert "scanned" in report.warnings
    assert "ocr" in report.timings
    assert "ocr" in {s for _, _, s in calls}
    assert len(report.tables) == 1
    table = report.tables[0]
    assert len(table.header) == 5
    assert len(table.rows) == 6  # 7×5 contando el encabezado
    assert any(r[0] == "A-100" for r in table.rows)


# --------------------------------------------------------------------------- export_batch


def test_export_batch_single_workbook(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    opts = _opts(tmp_path, single_workbook=True)
    reports = [
        pipeline.analyze(corpus("remito_grid.pdf"), opts),
        pipeline.analyze(corpus("remito_nogrid.pdf"), opts),
    ]

    dst = pipeline.export_batch(reports, opts)
    assert dst.parent == tmp_path / "out"
    assert dst.name.startswith("Conversion ") and dst.suffix == ".xlsx"
    assert dst.is_file()
    assert all(r.output_path == dst for r in reports)

    wb = openpyxl.load_workbook(dst)
    table_sheets = [n for n in wb.sheetnames if n.startswith("remito_")]
    assert len(table_sheets) == 2
    assert any(n.startswith("remito_grid ·") for n in table_sheets)
    assert any(n.startswith("remito_nogrid ·") for n in table_sheets)
    assert [c.value for c in wb[table_sheets[0]][2]] == REMITO_ROW0

    again = pipeline.export_batch(reports, opts)
    assert again != dst and again.is_file()  # tampoco sobrescribe


def test_export_batch_empty() -> None:
    with pytest.raises(UserError):
        pipeline.export_batch([], Options())


# --------------------------------------------------------------------------- CLI (subprocesos)


def _run(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [PYTHON, "-m", "pdf2excel", *args],
        capture_output=True,
        text=True,
        env=ENV,
        timeout=300,
        check=False,
    )


def test_cli_converts(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    pdf = corpus("remito_grid.pdf")
    out = tmp_path / "salida"
    proc = _run("--cli", str(pdf), "-o", str(out))
    assert proc.returncode == 0, proc.stderr
    assert (out / "remito_grid.xlsx").is_file()
    assert "remito_grid.pdf" in proc.stdout and "remito_grid.xlsx" in proc.stdout
    assert "1 tabla" in proc.stdout and "6 filas" in proc.stdout


def test_cli_json_and_failure_exit_code(corpus: Callable[[str], Path], tmp_path: Path) -> None:
    pdf = corpus("remito_grid.pdf")
    missing = tmp_path / "no_esta.pdf"
    proc = _run("--cli", "--json", str(pdf), str(missing), "-o", str(tmp_path / "o"))
    assert proc.returncode == 1
    data = json.loads(proc.stdout)
    assert [d["ok"] for d in data] == [True, False]
    assert data[0]["tables"] == 1 and data[0]["rows"] == 6
    assert data[0]["output"].endswith("remito_grid.xlsx")
    assert "error" in data[1] and data[1]["error"]["code"]


def test_cli_selftest_ok(corpus: Callable[[str], Path]) -> None:
    proc = _run("--selftest", str(corpus("remito_grid.pdf")))
    assert proc.returncode == 0, proc.stderr
    assert "selftest ok" in proc.stdout


def test_selftest_uses_isolated_ocr_cache(
    corpus: Callable[[str], Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """El selftest no puede aprovechar la caché del usuario: con caché poblada pasaría con tesseract roto."""
    from pdf2excel import cli
    from pdf2excel.core import ocr

    seen: list[Path] = []
    real_convert = pipeline.convert

    def spying_convert(src: Path, opts: Options, *a: object, **k: object) -> ConversionReport:
        seen.append(ocr.default_cache_dir())
        return real_convert(src, opts, *a, **k)

    monkeypatch.setattr(pipeline, "convert", spying_convert)
    monkeypatch.delenv(ocr.CACHE_DIR_ENV, raising=False)
    user_cache = ocr.default_cache_dir()
    assert cli.selftest(corpus("remito_grid.pdf")) == 0
    assert len(seen) == 1 and seen[0] != user_cache
    assert "pdf2excel-selftest-" in str(seen[0]) and not seen[0].exists()  # tmpdir ya borrado
    assert os.environ.get(ocr.CACHE_DIR_ENV) is None  # no deja la variable puesta


def test_selftest_detects_vendored_tesseract_not_in_use(
    corpus: Callable[[str], Path],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    import pdf2excel.resources_path as rp
    from pdf2excel import cli
    from pdf2excel.core import ocr

    vendored = tmp_path / "tesseract"
    vendored.mkdir()
    monkeypatch.setattr(rp, "tesseract_dir", lambda: vendored)
    monkeypatch.delenv(ocr.TESSERACT_ENV, raising=False)
    monkeypatch.setattr(ocr, "find_tesseract", lambda: tmp_path / "otro" / "tesseract")
    assert cli.selftest(corpus("remito_grid.pdf")) == 1
    assert "vendoreado" in capsys.readouterr().err
    # Con el vendoreado en uso, pasa.
    monkeypatch.setattr(ocr, "find_tesseract", lambda: vendored / "tesseract")
    assert cli.selftest(corpus("remito_grid.pdf")) == 0


def test_cli_selftest_missing_file(tmp_path: Path) -> None:
    proc = _run("--selftest", str(tmp_path / "archivo_inexistente.pdf"))
    assert proc.returncode == 1
    assert "selftest ok" not in proc.stdout
    assert "Traceback" in proc.stderr


def test_cli_no_args_is_usage_error() -> None:
    from pdf2excel import cli

    assert cli.main(["--cli"]) == 2


def test_cli_single_workbook_in_process(
    corpus: Callable[[str], Path], tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from pdf2excel import cli

    code = cli.main(["--cli", "--single-workbook", str(corpus("remito_grid.pdf")), "-o", str(tmp_path / "o")])
    assert code == 0
    files = list((tmp_path / "o").glob("Conversion *.xlsx"))
    assert len(files) == 1
    assert "remito_grid.pdf" in capsys.readouterr().out


# --------------------------------------------------------------------------- __main__


def test_main_dispatches_to_cli(monkeypatch: pytest.MonkeyPatch) -> None:
    import pdf2excel.__main__ as entry
    import pdf2excel.cli as cli

    seen: list[list[str]] = []
    monkeypatch.setattr(cli, "main", lambda argv=None: seen.append(list(argv or [])) or 7)
    assert entry.main(["--cli", "a.pdf"]) == 7
    assert seen == [["--cli", "a.pdf"]]
    assert entry.wants_cli(["--selftest", "x.pdf"])
    assert entry.wants_cli(["x.pdf", "-o", "dir"])
    assert not entry.wants_cli([])


def test_main_dispatches_to_gui(monkeypatch: pytest.MonkeyPatch) -> None:
    import pdf2excel.__main__ as entry

    received: list[list[str]] = []
    fake_app = types.ModuleType("pdf2excel.app")
    fake_app.main = lambda argv: received.append(list(argv)) or 0  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "pdf2excel.app", fake_app)
    monkeypatch.setenv("QT_QPA_PLATFORM", "offscreen")

    assert entry.main(["remito.pdf"]) == 0  # "Enviar a": PDFs precargados en la GUI
    assert received == [["remito.pdf"]]
    assert entry.main([]) == 0
    assert received[-1] == []


def test_main_no_display_with_pdfs_uses_cli(monkeypatch: pytest.MonkeyPatch) -> None:
    import pdf2excel.__main__ as entry

    if sys.platform.startswith("win") or sys.platform == "darwin":
        pytest.skip("en Windows/macOS siempre hay pantalla")
    for var in ("DISPLAY", "WAYLAND_DISPLAY", "QT_QPA_PLATFORM"):
        monkeypatch.delenv(var, raising=False)
    assert not entry.has_display()
    assert entry.wants_cli(["remito.pdf"])
    assert not entry.wants_cli([])
