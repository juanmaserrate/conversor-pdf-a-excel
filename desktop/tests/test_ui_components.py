"""Tests de los componentes de UI (WI-07) con pytest-qt y QT_QPA_PLATFORM=offscreen."""

from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import sys
import threading
import time
import types
from datetime import date
from pathlib import Path
from typing import Any, get_args

import pytest
from PySide6.QtCore import QMimeData, QPointF, QSettings, QStandardPaths, Qt, QUrl
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont, QImage
from PySide6.QtTest import QSignalSpy
from PySide6.QtWidgets import QApplication, QFileDialog

from pdf2excel.errors import ErrorCode, UserError
from pdf2excel.models import ColumnInfo, ConversionReport, Options, PageInfo, TypedTable
from pdf2excel.ui import strings_es as S
from pdf2excel.ui.drop_zone import DropZone
from pdf2excel.ui.file_list import FileListWidget
from pdf2excel.ui.preview_model import TablePreviewModel, _fallback_format
from pdf2excel.ui.settings_dialog import SettingsDialog, load_options, save_options

pytestmark = pytest.mark.ui

ALL_CODES: tuple[str, ...] = get_args(ErrorCode)


# --------------------------------------------------------------------------- fixtures


@pytest.fixture(autouse=True)
def isolated_settings(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """QSettings en un INI dentro de tmp_path: no toca la configuración del usuario ni se filtra entre tests.

    ``QSettings.setPath`` solo lo respeta el constructor con formato explícito, así que los dos puntos que
    crean ``QSettings(org, app)`` (opciones y última carpeta del DropZone) se redirigen a ese constructor.
    """
    from pdf2excel.ui import drop_zone, settings_dialog

    QStandardPaths.setTestModeEnabled(True)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path / "qsettings"))
    org_default, app_default = settings_dialog.SETTINGS_ORG, settings_dialog.SETTINGS_APP

    def make_settings(org: str = org_default, app: str = app_default) -> QSettings:
        return QSettings(QSettings.Format.IniFormat, QSettings.Scope.UserScope, org, app)

    monkeypatch.setattr(settings_dialog, "_settings", make_settings)
    monkeypatch.setattr(drop_zone, "QSettings", make_settings)
    assert str(tmp_path) in make_settings().fileName()


@pytest.fixture
def pipeline_stub(monkeypatch: pytest.MonkeyPatch) -> types.ModuleType:
    """Módulo ``pdf2excel.core.pipeline`` sintético: el worker lo importa de forma diferida en run()."""
    import pdf2excel.core as core_pkg

    stub = types.ModuleType("pdf2excel.core.pipeline")
    monkeypatch.setitem(sys.modules, "pdf2excel.core.pipeline", stub)
    monkeypatch.setattr(core_pkg, "pipeline", stub, raising=False)
    return stub


def make_table(n_rows: int = 250) -> TypedTable:
    header = ["Producto", "Código", "Fecha", "Importe", "IVA"]
    columns = [
        ColumnInfo(kind="text", header="Producto"),
        ColumnInfo(kind="id", header="Código"),
        ColumnInfo(kind="date", number_format="dd/mm/yyyy", header="Fecha"),
        ColumnInfo(kind="number", locale="ar", number_format="#,##0.00", header="Importe"),
        ColumnInfo(kind="percent", number_format="0.00%", header="IVA"),
    ]
    rows: list[list[Any]] = []
    for i in range(n_rows):
        rows.append([f"Item {i}", f"A-{100 + i}", date(2026, 10, 15), 1234.5 + i, 0.21])
    rows[1][3] = "s/d"  # celda que no parseó
    return TypedTable(
        name="Tabla 1",
        header=header,
        rows=rows,
        columns=columns,
        review_cells={(1, 3)},
        total_rows={n_rows - 1},
        pages=(0, 1),
    )


def make_report(path: Path, n_tables: int = 1, ocr: bool = False, n_pages: int = 3) -> ConversionReport:
    tables = [make_table(10) for _ in range(n_tables)]
    pages = [
        PageInfo(index=i, kind="scanned" if ocr else "digital", n_chars=500, image_coverage=0.0)
        for i in range(n_pages)
    ]
    return ConversionReport(source=path, pages=pages, tables=tables, text_items=[], ocr_applied=ocr)


# --------------------------------------------------------------------------- strings_es


def test_all_strings_non_empty() -> None:
    def check(value: Any, name: str) -> None:
        if isinstance(value, str):
            assert value.strip(), f"{name} está vacío"
        elif isinstance(value, dict):
            assert value, f"{name} está vacío"
            for k, v in value.items():
                check(v, f"{name}[{k!r}]")
        elif isinstance(value, list | tuple):
            assert value, f"{name} está vacío"
            for i, v in enumerate(value):
                if v is not None:  # la etiqueta de acción puede ser None
                    check(v, f"{name}[{i}]")

    names = [n for n in dir(S) if n.isupper()]
    assert names
    for name in names:
        check(getattr(S, name), name)


def test_errors_cover_all_codes() -> None:
    assert set(S.ERRORS) == set(ALL_CODES)
    assert set(S.ERROR_SHORT) == set(ALL_CODES)
    for code in ALL_CODES:
        title, detail, action = S.ERRORS[code]  # type: ignore[index]
        assert title and detail
        assert action is None or action
    assert S.ERRORS["file_in_use"] == (
        "No se pudo guardar el Excel",
        "{name} está abierto en Excel. Cerralo y tocá Reintentar.",
        "Reintentar",
    )
    assert set(S.STAGE_TEXT) == {"classify", "ocr", "tables", "typing", "excel"}


def test_string_helpers(tmp_path: Path) -> None:
    assert S.plural(1, "tabla", "tablas") == "1 tabla"
    assert S.plural(3, "tabla", "tablas") == "3 tablas"
    rep = make_report(tmp_path / "remito.pdf", n_tables=3, n_pages=3)
    assert S.file_status(rep) == "3 tablas · 3 pág."
    rep_ocr = make_report(tmp_path / "remito.pdf", n_tables=1, ocr=True, n_pages=2)
    assert S.file_status(rep_ocr) == "1 tabla · 2 pág. · escaneado, se leyó con OCR"
    empty = make_report(tmp_path / "remito.pdf", n_tables=0)
    assert S.file_status(empty) == S.NO_TABLES

    rep.output_path = tmp_path / "remito.xlsx"
    assert S.done_line(rep) == "✔ remito.pdf → remito.xlsx (3 tablas, 30 filas)"
    err = UserError("encrypted")
    assert S.error_line("factura.pdf", err) == "⚠ factura.pdf: no se pudo leer, está protegido con contraseña"
    _title, detail, action = S.error_texts(UserError("file_in_use"), "remito.xlsx")
    assert "remito.xlsx está abierto en Excel" in detail and action == "Reintentar"
    assert S.done_title(5) == "¡Listo! Guardamos 5 archivos"
    assert S.review_cells_text(4) == "4 celdas marcadas para revisar en amarillo"
    assert S.progress_text(2, 5, 12, 50) == "Archivo 2 de 5 · página 12 de 50"


# --------------------------------------------------------------------------- DropZone


def test_expand_paths(tmp_path: Path) -> None:
    (tmp_path / "remito.pdf").write_bytes(b"%PDF-1.4")
    (tmp_path / "foto.jpg").write_bytes(b"\xff\xd8")
    folder = tmp_path / "carpeta"
    (folder / "sub").mkdir(parents=True)
    (folder / "b.pdf").write_bytes(b"%PDF")
    (folder / "A.PDF").write_bytes(b"%PDF")
    (folder / "notas.txt").write_text("x")
    (folder / "sub" / "c.pdf").write_bytes(b"%PDF")

    result = DropZone.expand_paths([str(tmp_path / "remito.pdf"), str(tmp_path / "foto.jpg"), str(folder)])
    assert result == [
        str(tmp_path / "remito.pdf"),
        str(folder / "A.PDF"),
        str(folder / "b.pdf"),
        str(folder / "sub" / "c.pdf"),
    ]
    # sin duplicados, ignora vacíos y solo .pdf
    assert DropZone.expand_paths(["/tmp/x/remito.pdf", "/tmp/x/remito.pdf", "", "/tmp/x/foto.jpg"]) == [
        "/tmp/x/remito.pdf"
    ]


def _mime(paths: list[str]) -> QMimeData:
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(p) for p in paths])
    return mime


def test_drop_zone_filters_non_pdf(qtbot: Any) -> None:
    zone = DropZone()
    qtbot.addWidget(zone)
    zone.show()
    assert zone.accessibleName() == S.DROP_TITLE
    assert zone.focusPolicy() == Qt.FocusPolicy.StrongFocus
    assert zone.acceptDrops()

    got: list[list[str]] = []
    zone.files_dropped.connect(got.append)
    mime = _mime(["/tmp/x/remito.pdf", "/tmp/x/foto.jpg"])
    app = QApplication.instance()
    assert app is not None

    enter = QDragEnterEvent(
        QPointF(10, 10),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    app.sendEvent(zone, enter)
    assert enter.isAccepted()
    assert zone.property("active") is True

    drop = QDropEvent(
        QPointF(10, 10),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    app.sendEvent(zone, drop)
    assert got == [[os.path.abspath("/tmp/x/remito.pdf")]]
    assert zone.property("active") is False

    # Solo archivos que no son PDF: no se acepta ni se emite.
    only_jpg = _mime(["/tmp/x/foto.jpg"])
    enter2 = QDragEnterEvent(
        QPointF(10, 10),
        Qt.DropAction.CopyAction,
        only_jpg,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    app.sendEvent(zone, enter2)
    assert not enter2.isAccepted()
    drop2 = QDropEvent(
        QPointF(10, 10),
        Qt.DropAction.CopyAction,
        only_jpg,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    app.sendEvent(zone, drop2)
    assert len(got) == 1


def test_drop_zone_keyboard_opens_dialog(qtbot: Any, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    zone = DropZone()
    qtbot.addWidget(zone)
    zone.show()
    calls: list[tuple[Any, ...]] = []

    def fake_dialog(parent: Any, title: str, directory: str, filt: str) -> tuple[list[str], str]:
        calls.append((title, directory, filt))
        return [str(tmp_path / "uno.pdf"), str(tmp_path / "foto.jpg")], filt

    monkeypatch.setattr(QFileDialog, "getOpenFileNames", staticmethod(fake_dialog))
    got: list[list[str]] = []
    zone.files_dropped.connect(got.append)
    qtbot.keyClick(zone, Qt.Key.Key_Return)
    assert calls and calls[0][0] == S.FILE_DIALOG_TITLE and calls[0][2] == S.FILE_DIALOG_FILTER
    assert got == [[str(tmp_path / "uno.pdf")]]
    # Recuerda la última carpeta.
    qtbot.mouseClick(zone, Qt.MouseButton.LeftButton)
    assert calls[1][1] == str(tmp_path)


# --------------------------------------------------------------------------- TablePreviewModel


def test_preview_model_roles(qtbot: Any) -> None:
    model = TablePreviewModel()
    table = make_table(250)
    model.set_table(table)
    assert model.rowCount() == 200
    assert model.is_truncated() and model.total_rows() == 250
    assert model.columnCount() == 5

    idx = model.index(0, 3)
    assert model.data(idx, Qt.ItemDataRole.DisplayRole) == "1.234,50"
    assert model.data(model.index(0, 2), Qt.ItemDataRole.DisplayRole) == "15/10/2026"
    assert model.data(model.index(0, 1), Qt.ItemDataRole.DisplayRole) == "A-100"
    assert model.data(model.index(0, 4), Qt.ItemDataRole.DisplayRole) == "21,00 %"

    align_num = model.data(idx, Qt.ItemDataRole.TextAlignmentRole)
    assert align_num & Qt.AlignmentFlag.AlignRight
    align_txt = model.data(model.index(0, 0), Qt.ItemDataRole.TextAlignmentRole)
    assert align_txt & Qt.AlignmentFlag.AlignLeft
    assert not (align_txt & Qt.AlignmentFlag.AlignRight)

    review = model.index(1, 3)
    assert model.data(review, Qt.ItemDataRole.BackgroundRole) == QColor("#FFF2CC")
    assert model.data(review, Qt.ItemDataRole.DisplayRole) == "s/d"
    assert model.data(review, Qt.ItemDataRole.ToolTipRole) == "Revisar: no se pudo interpretar"
    assert model.data(idx, Qt.ItemDataRole.BackgroundRole) is None
    assert model.data(idx, Qt.ItemDataRole.ToolTipRole) is None

    for section, name in enumerate(table.header):
        assert model.headerData(section, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole) == name
    assert model.headerData(0, Qt.Orientation.Vertical, Qt.ItemDataRole.DisplayRole) == "1"

    assert not (model.flags(idx) & Qt.ItemFlag.ItemIsEditable)
    assert model.flags(idx) & Qt.ItemFlag.ItemIsSelectable

    # Negrita en filas de total (si quedan dentro de las 200 visibles).
    model.set_table(table, max_rows=250)
    font = model.data(model.index(249, 0), Qt.ItemDataRole.FontRole)
    assert isinstance(font, QFont) and font.bold()
    assert model.data(model.index(0, 0), Qt.ItemDataRole.FontRole) is None

    # None → "" y tabla vacía → 0 filas.
    table.rows[0][0] = None
    assert model.data(model.index(0, 0), Qt.ItemDataRole.DisplayRole) == ""
    model.set_table(None)
    assert model.rowCount() == 0 and model.columnCount() == 0


def test_preview_fallback_formatter_matches_ar() -> None:
    num = ColumnInfo(kind="number", number_format="#,##0.00")
    assert _fallback_format(1234.5, num) == "1.234,50"
    assert _fallback_format(1234, ColumnInfo(kind="number")) == "1.234"
    assert _fallback_format(0.21, ColumnInfo(kind="percent")) == "21,00 %"
    assert _fallback_format(date(2026, 10, 15), ColumnInfo(kind="date")) == "15/10/2026"
    assert _fallback_format(None, num) == ""


# --------------------------------------------------------------------------- FileListWidget


def test_file_list_status_and_accessibility(qtbot: Any) -> None:
    widget = FileListWidget()
    qtbot.addWidget(widget)
    widget.show()
    changed: list[str] = []
    widget.current_file_changed.connect(changed.append)

    widget.add_file("/tmp/x/remito.pdf")
    widget.add_file("/tmp/x/factura.pdf")
    widget.add_file("/tmp/x/remito.pdf")  # duplicado: se ignora
    assert widget.paths() == ["/tmp/x/remito.pdf", "/tmp/x/factura.pdf"]
    assert widget.status_text("/tmp/x/remito.pdf") == S.FILE_STATUS_QUEUED
    assert widget.accessible_description("/tmp/x/remito.pdf") == S.FILE_STATUS_QUEUED
    assert changed == ["/tmp/x/remito.pdf"]

    widget.set_status("/tmp/x/remito.pdf", "3 tablas · 3 pág.", "ok")
    assert widget.accessible_description("/tmp/x/remito.pdf") == "3 tablas · 3 pág."
    assert widget.status_text("/tmp/x/remito.pdf") == "3 tablas · 3 pág."
    assert widget.state_of("/tmp/x/remito.pdf") == "ok"
    item = widget.item(0)
    assert item.data(Qt.ItemDataRole.AccessibleDescriptionRole) == "3 tablas · 3 pág."
    row = widget.itemWidget(item)
    assert row is not None and row.glyph_label.text() == "✔"

    widget.set_status("/tmp/x/factura.pdf", "No se pudo leer", "error")
    assert widget.itemWidget(widget.item(1)).glyph_label.text() == "✖"

    img = QImage(400, 560, QImage.Format.Format_RGB32)
    img.fill(QColor("white"))
    widget.set_thumbnail("/tmp/x/remito.pdf", img)
    pix = row.thumb.pixmap()
    assert pix.width() <= 48 and pix.height() <= 68 and pix.height() > 0

    widget.select_file("/tmp/x/factura.pdf")
    assert changed[-1] == "/tmp/x/factura.pdf"
    widget.remove_file("/tmp/x/remito.pdf")
    assert widget.paths() == ["/tmp/x/factura.pdf"]
    widget.remove_file("/no/existe.pdf")
    assert widget.count() == 1


# --------------------------------------------------------------------------- settings


def test_options_roundtrip(qtbot: Any, tmp_path: Path) -> None:
    assert load_options() == Options()
    modified = Options(
        typed_values=False,
        merge_continuations=False,
        include_text=False,
        one_sheet_per_page=True,
        force_ocr=True,
        ocr_lang="spa",
        output_dir=tmp_path / "salida",
        open_when_done=False,
        show_origin_sheet=True,
        single_workbook=True,
    )
    save_options(modified)
    loaded = load_options()
    assert loaded == modified
    assert isinstance(loaded.output_dir, Path)

    save_options(Options())
    assert load_options() == Options()
    assert load_options().output_dir is None


def test_settings_dialog_options(qtbot: Any) -> None:
    opts = Options(typed_values=False, ocr_lang="eng", force_ocr=True)
    dialog = SettingsDialog(opts)
    qtbot.addWidget(dialog)
    assert dialog.lang_combo.currentData() == "eng"
    assert dialog.force_ocr_check.isChecked()
    assert dialog.force_ocr_check.toolTip() == "Usalo si el PDF tiene texto pero sale con caracteres raros"
    assert dialog.ok_button.text() == "Aceptar" and dialog.cancel_button.text() == "Cancelar"
    assert "0.1.0" in dialog.version_label.text()

    dialog.lang_combo.setCurrentIndex(dialog.lang_combo.findData("spa"))
    dialog.single_workbook_check.setChecked(True)
    dialog.open_check.setChecked(False)
    result = dialog.options()
    assert result.ocr_lang == "spa"
    assert result.single_workbook is True
    assert result.open_when_done is False
    assert result.force_ocr is True
    assert result.typed_values is False  # no está en el diálogo: se conserva


# --------------------------------------------------------------------------- JobQueue


def test_job_queue_analyze_and_export(qtbot: Any, pipeline_stub: types.ModuleType, tmp_path: Path) -> None:
    from pdf2excel.jobs.worker import JobQueue

    src = tmp_path / "remito.pdf"

    def fake_analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        assert threading.current_thread() is not threading.main_thread()
        progress(1, 3, "classify")
        progress(3, 3, "tables")
        return make_report(path, n_tables=2)

    def fake_export(report: ConversionReport, opts: Options) -> Path:
        out = report.source.with_suffix(".xlsx")
        report.output_path = out
        return out

    pipeline_stub.analyze = fake_analyze  # type: ignore[attr-defined]
    pipeline_stub.export = fake_export  # type: ignore[attr-defined]

    queue = JobQueue(max_threads=2)
    progress_spy = QSignalSpy(queue.signals.progress)
    analyzed: list[tuple[str, Any]] = []
    exported: list[tuple[str, Any]] = []
    failed: list[tuple[str, Any]] = []
    queue.signals.analyzed.connect(lambda p, r: analyzed.append((p, r)))
    queue.signals.exported.connect(lambda p, r: exported.append((p, r)))
    queue.signals.failed.connect(lambda p, e: failed.append((p, e)))

    with qtbot.waitSignal(queue.signals.all_done, timeout=5000):
        queue.submit_analyze([str(src)], Options())
        assert queue.is_busy()
    assert not failed
    assert len(analyzed) == 1 and analyzed[0][0] == str(src)
    report = analyzed[0][1]
    assert isinstance(report, ConversionReport) and len(report.tables) == 2
    assert progress_spy.count() >= 2
    first = progress_spy.at(0)
    assert first[0] == str(src) and first[1] == 1 and first[2] == 3 and first[3] == "classify"
    assert not queue.is_busy()

    with qtbot.waitSignal(queue.signals.all_done, timeout=5000):
        queue.submit_export([report], Options())
    assert len(exported) == 1
    assert exported[0][1].output_path == src.with_suffix(".xlsx")
    assert not failed
    queue.wait_for_done(2000)


def test_job_queue_failed_damaged(qtbot: Any, pipeline_stub: types.ModuleType, tmp_path: Path) -> None:
    pymupdf = pytest.importorskip("pymupdf")
    from pdf2excel.jobs.worker import JobQueue

    def broken_analyze(path: Path, *args: Any, **kwargs: Any) -> Any:
        raise pymupdf.FileDataError(f"cannot open {path}: No such file or directory")

    pipeline_stub.analyze = broken_analyze  # type: ignore[attr-defined]
    queue = JobQueue(max_threads=1)
    failed: list[tuple[str, Any]] = []
    queue.signals.failed.connect(lambda p, e: failed.append((p, e)))

    with qtbot.waitSignal(queue.signals.all_done, timeout=5000):
        queue.submit_analyze([str(tmp_path / "roto.pdf")], Options())
    assert len(failed) == 1
    path, err = failed[0]
    assert path == str(tmp_path / "roto.pdf")
    assert isinstance(err, UserError) and err.code == "damaged"
    queue.wait_for_done(2000)


def test_job_queue_user_error_passthrough(
    qtbot: Any, pipeline_stub: types.ModuleType, tmp_path: Path
) -> None:
    from pdf2excel.jobs.worker import JobQueue

    def locked_analyze(path: Path, *args: Any, **kwargs: Any) -> Any:
        raise UserError("encrypted", action="password")

    pipeline_stub.analyze = locked_analyze  # type: ignore[attr-defined]
    queue = JobQueue()
    failed: list[tuple[str, Any]] = []
    queue.signals.failed.connect(lambda p, e: failed.append((p, e)))
    with qtbot.waitSignal(queue.signals.all_done, timeout=5000):
        queue.submit_analyze([str(tmp_path / "a.pdf"), str(tmp_path / "b.pdf")], Options(), {"a.pdf": "x"})
    assert sorted(p for p, _ in failed) == sorted([str(tmp_path / "a.pdf"), str(tmp_path / "b.pdf")])
    assert all(e.code == "encrypted" and e.action == "password" for _, e in failed)
    queue.wait_for_done(2000)


def test_job_queue_cancel_all_stops_queue(
    qtbot: Any, pipeline_stub: types.ModuleType, tmp_path: Path
) -> None:
    from pdf2excel.jobs.worker import JobQueue

    started = threading.Event()

    def slow_analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        started.set()
        deadline = time.monotonic() + 0.3
        while time.monotonic() < deadline:
            if cancel.is_set():
                raise UserError("cancelled")
            time.sleep(0.01)
        return make_report(path)

    pipeline_stub.analyze = slow_analyze  # type: ignore[attr-defined]
    queue = JobQueue(max_threads=1)
    analyzed: list[str] = []
    failed: list[tuple[str, Any]] = []
    queue.signals.analyzed.connect(lambda p, r: analyzed.append(p))
    queue.signals.failed.connect(lambda p, e: failed.append((p, e)))
    paths = [str(tmp_path / f"f{i}.pdf") for i in range(6)]

    with qtbot.waitSignal(queue.signals.all_done, timeout=5000):
        queue.submit_analyze(paths, Options())
        assert started.wait(2.0)
        queue.cancel_all()
    queue.wait_for_done(2000)
    qtbot.wait(50)

    # El que estaba corriendo se cortó (cancelled); los encolados nunca arrancaron.
    assert analyzed == []
    assert len(failed) <= 1
    assert all(e.code == "cancelled" for _, e in failed)
    assert not queue.is_busy() and queue.pending_count() == 0

    # Un lote nuevo después de cancelar funciona con un Event fresco.
    pipeline_stub.analyze = lambda path, *a, **k: make_report(path)  # type: ignore[attr-defined]
    with qtbot.waitSignal(queue.signals.all_done, timeout=5000):
        queue.submit_analyze([paths[0]], Options())
    assert analyzed == [paths[0]]
    queue.wait_for_done(2000)


def test_job_queue_empty_submit_emits_all_done(qtbot: Any) -> None:
    from pdf2excel.jobs.worker import JobQueue

    queue = JobQueue()
    spy = QSignalSpy(queue.signals.all_done)
    queue.submit_analyze([], Options())
    assert spy.count() == 1
    assert not queue.is_busy()
