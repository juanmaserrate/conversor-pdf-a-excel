"""Tests del ítem base: modelos, errores, rutas de recursos y logging."""

from __future__ import annotations

import logging
import subprocess
from collections.abc import Iterator
from datetime import date
from pathlib import Path

import pytest

import pdf2excel
from pdf2excel import errors, logging_setup, models, resources_path
from pdf2excel.errors import DEFAULT_TITLES, ErrorCode, UserError, from_exception
from pdf2excel.models import (
    STAGES,
    ColumnInfo,
    ConversionReport,
    Options,
    PageInfo,
    RawTable,
    TextItem,
    TypedTable,
)

# --------------------------------------------------------------------------- paquete


def test_package_metadata() -> None:
    assert pdf2excel.__version__ == "0.1.0"
    assert pdf2excel.APP_NAME == "Conversor PDF a Excel"
    assert pdf2excel.APP_ID == "ConversorPDFaExcel"
    assert STAGES == ("classify", "ocr", "tables", "layout", "typing", "excel")


def test_models_module_has_no_logic() -> None:
    # Solo dataclasses y alias: ninguna función de módulo.
    import inspect

    funcs = [n for n, o in vars(models).items() if inspect.isfunction(o) and o.__module__ == models.__name__]
    assert funcs == []


# --------------------------------------------------------------------------- dataclasses


def test_dataclasses_instantiate_with_defaults() -> None:
    page = PageInfo(index=0, kind="digital", n_chars=120, image_coverage=0.0)
    assert page.rotation == 0

    raw = RawTable(
        page=0, page_end=0, bbox=(0.0, 0.0, 10.0, 10.0), rows=[["a"]], source="hybrid", quality=0.9
    )
    assert raw.title is None

    col = ColumnInfo(kind="number")
    assert col.locale is None and col.number_format is None and col.header == ""

    table = TypedTable(name="T 1", header=["a", "b"], rows=[[1, "x"]], columns=[col, ColumnInfo(kind="text")])
    assert table.review_cells == set()
    assert table.total_rows == set()
    assert table.pages == (0, 0)
    assert table.source == "" and table.quality == 0.0

    item = TextItem(page=0, kind="kv", key="Cliente", value="Bar El Catorce")
    assert item.key == "Cliente"

    opts = Options()
    assert opts.typed_values is True
    assert opts.merge_continuations is True
    assert opts.include_text is True
    assert opts.one_sheet_per_page is False
    assert opts.force_ocr is False
    assert opts.ocr_lang == "spa+eng"
    assert opts.output_dir is None
    assert opts.open_when_done is True
    assert opts.show_origin_sheet is False
    assert opts.single_workbook is False

    report = ConversionReport(source=Path("x.pdf"), pages=[page], tables=[table], text_items=[item])
    assert report.warnings == [] and report.timings == {}
    assert report.ocr_applied is False and report.output_path is None and report.app_version == ""


def test_default_factories_are_independent() -> None:
    a = TypedTable(name="a", header=[], rows=[], columns=[])
    b = TypedTable(name="b", header=[], rows=[], columns=[])
    a.review_cells.add((0, 0))
    a.total_rows.add(1)
    assert b.review_cells == set() and b.total_rows == set()

    r1 = ConversionReport(source=Path("a.pdf"), pages=[], tables=[], text_items=[])
    r2 = ConversionReport(source=Path("b.pdf"), pages=[], tables=[], text_items=[])
    r1.warnings.append("w")
    r1.timings["ocr"] = 1.0
    assert r2.warnings == [] and r2.timings == {}


def test_report_n_rows_and_n_review() -> None:
    t1 = TypedTable(
        name="t1",
        header=["Fecha", "Importe"],
        rows=[[date(2026, 10, 15), 1234.56], [date(2026, 10, 16), "s/d"]],
        columns=[ColumnInfo(kind="date"), ColumnInfo(kind="number", locale="ar")],
        review_cells={(1, 1)},
    )
    t2 = TypedTable(name="t2", header=["a"], rows=[[1], [2], [3]], columns=[ColumnInfo(kind="number")])
    report = ConversionReport(source=Path("x.pdf"), pages=[], tables=[t1, t2], text_items=[])
    assert report.n_rows == 5
    assert report.n_review == 1
    empty = ConversionReport(source=Path("x.pdf"), pages=[], tables=[], text_items=[])
    assert empty.n_rows == 0 and empty.n_review == 0


# --------------------------------------------------------------------------- UserError


def test_user_error_default_title_is_spanish() -> None:
    err = UserError("encrypted")
    assert err.title
    assert err.title == DEFAULT_TITLES["encrypted"]
    assert "contraseña" in err.title.lower()
    assert err.code == "encrypted"
    assert err.action is None
    assert err.cause is None
    assert err.detail  # también hay detalle por defecto
    assert isinstance(err, Exception)


def test_user_error_custom_fields_and_cause() -> None:
    cause = RuntimeError("boom")
    err = UserError("file_in_use", title="T", detail="D", action="retry", cause=cause)
    assert (err.title, err.detail, err.action) == ("T", "D", "retry")
    assert err.cause is cause
    assert err.__cause__ is cause
    assert "T" in str(err) and "D" in str(err)
    assert "file_in_use" in repr(err)


def test_all_codes_have_titles_in_spanish() -> None:
    codes: tuple[ErrorCode, ...] = (
        "encrypted",
        "damaged",
        "not_pdf",
        "no_permission",
        "file_in_use",
        "ocr_unavailable",
        "ocr_failed",
        "cancelled",
        "unknown",
    )
    for code in codes:
        assert DEFAULT_TITLES[code].strip()
        assert code in errors.DEFAULT_DETAILS
        assert code in errors.DEFAULT_ACTIONS
    assert errors.DEFAULT_ACTIONS["file_in_use"] == "retry"
    assert errors.DEFAULT_ACTIONS["no_permission"] == "choose_folder"
    assert errors.DEFAULT_ACTIONS["encrypted"] == "password"


# --------------------------------------------------------------------------- from_exception


def test_from_exception_passthrough_user_error() -> None:
    err = UserError("cancelled")
    assert from_exception(err) is err


def test_from_exception_permission_xlsx_is_file_in_use() -> None:
    err = from_exception(PermissionError(13, "x", "/a/b.xlsx"))
    assert err.code == "file_in_use"
    assert err.action == "retry"
    assert "b.xlsx" in err.detail
    assert isinstance(err.cause, PermissionError)


def test_from_exception_permission_xlsx_via_path_argument() -> None:
    err = from_exception(PermissionError("denied"), path=Path("/salida/remito (2).XLSX"))
    assert err.code == "file_in_use" and err.action == "retry"


def test_from_exception_permission_dir_is_no_permission() -> None:
    err = from_exception(PermissionError(13, "x", "/dir"))
    assert err.code == "no_permission"
    assert err.action == "choose_folder"


def test_from_exception_tesseract_missing_is_ocr_unavailable() -> None:
    err = from_exception(FileNotFoundError("tesseract not found"))
    assert err.code == "ocr_unavailable"
    # Como la lanza subprocess: filename = ejecutable
    err2 = from_exception(FileNotFoundError(2, "No such file or directory", "tesseract"))
    assert err2.code == "ocr_unavailable"


def test_from_exception_other_file_not_found_is_unknown() -> None:
    err = from_exception(FileNotFoundError(2, "No such file", "/x/y.pdf"))
    assert err.code == "unknown"


def test_from_exception_unknown_keeps_cause() -> None:
    exc = ValueError("boom")
    err = from_exception(exc)
    assert err.code == "unknown"
    assert err.cause is exc
    assert "boom" in err.detail
    assert err.title == DEFAULT_TITLES["unknown"]


def test_from_exception_not_pdf() -> None:
    assert from_exception(ValueError("Not a PDF file")).code == "not_pdf"
    err = from_exception(ValueError("cannot parse"), path=Path("/docs/foto.jpg"))
    assert err.code == "not_pdf" and err.action == "skip"
    assert "foto.jpg" in err.detail


def test_from_exception_pymupdf_errors_are_damaged() -> None:
    pymupdf = pytest.importorskip("pymupdf")
    err = from_exception(pymupdf.FileDataError("cannot open broken document"), path=Path("/a/remito.pdf"))
    assert err.code == "damaged" and err.action == "skip"
    assert "remito.pdf" in err.detail
    assert from_exception(pymupdf.EmptyFileError("cannot open empty document")).code == "damaged"
    # Documento dañado pero con otra extensión → not_pdf
    assert from_exception(pymupdf.FileDataError("x"), path=Path("/a/lista.docx")).code == "not_pdf"


def test_from_exception_runtime_cannot_open_is_damaged() -> None:
    assert from_exception(RuntimeError("cannot open file.pdf: broken xref")).code == "damaged"
    assert from_exception(RuntimeError("something else")).code == "unknown"


def test_from_exception_subprocess_errors_are_ocr_failed() -> None:
    assert from_exception(subprocess.CalledProcessError(1, ["tesseract"])).code == "ocr_failed"
    assert from_exception(subprocess.TimeoutExpired(["tesseract"], 60)).code == "ocr_failed"


# --------------------------------------------------------------------------- recursos


def test_resource_dir_and_icons_exist() -> None:
    rd = resources_path.resource_dir()
    assert rd.is_dir(), rd
    assert rd.name == "resources"
    assert resources_path.icon_path("app.png").is_file()
    assert resources_path.icon_path("app.svg").is_file()
    assert resources_path.icon_path("app.ico").is_file()
    assert resources_path.tessdata_dir() == rd / "tessdata"
    assert resources_path.tesseract_dir() == rd / "tesseract"


def test_resource_dir_frozen_uses_meipass(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    import sys

    monkeypatch.delenv("PDF2EXCEL_RESOURCES", raising=False)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
    assert resources_path.resource_dir() == tmp_path
    assert resources_path.icon_path("app.ico") == tmp_path / "icons" / "app.ico"


def test_resource_dir_env_override(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("PDF2EXCEL_RESOURCES", str(tmp_path))
    assert resources_path.resource_dir() == tmp_path


def test_icon_png_is_256px() -> None:
    PIL_Image = pytest.importorskip("PIL.Image")
    with PIL_Image.open(resources_path.icon_path("app.png")) as im:
        assert im.size == (256, 256)
    with PIL_Image.open(resources_path.icon_path("app.ico")) as ico:
        assert ico.format == "ICO"
        sizes = {s[0] for s in getattr(ico, "info", {}).get("sizes", set())}
        assert {16, 32, 48, 256} <= sizes


# --------------------------------------------------------------------------- logging


@pytest.fixture
def clean_root_logger() -> Iterator[None]:
    root = logging.getLogger()
    before = list(root.handlers)
    yield
    for h in root.handlers:
        if h not in before:
            root.removeHandler(h)
            h.close()


@pytest.mark.usefixtures("clean_root_logger")
def test_setup_logging_creates_file(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(logging_setup, "log_dir", lambda: tmp_path / "logs")
    path = logging_setup.setup_logging()
    assert path == tmp_path / "logs" / "app.log"
    assert path.is_file()
    logging.getLogger("pdf2excel.test").info("hola")
    for h in logging.getLogger().handlers:
        h.flush()
    assert "hola" in path.read_text(encoding="utf-8")


@pytest.mark.usefixtures("clean_root_logger")
def test_setup_logging_is_idempotent(tmp_path: Path) -> None:
    from logging.handlers import RotatingFileHandler

    logging_setup.setup_logging(directory=tmp_path)
    logging_setup.setup_logging(directory=tmp_path)
    handlers = [
        h for h in logging.getLogger().handlers
        if isinstance(h, RotatingFileHandler) and Path(h.baseFilename) == (tmp_path / "app.log").resolve()
    ]  # fmt: skip
    assert len(handlers) == 1
    assert handlers[0].maxBytes == 1_000_000
    assert handlers[0].backupCount == 3


def test_log_dir_uses_app_id() -> None:
    assert "ConversorPDFaExcel" in str(logging_setup.log_dir())
    assert logging_setup.log_path().name == "app.log"
