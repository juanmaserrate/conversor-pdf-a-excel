"""Tests del flujo completo de la UI (WI-08): MainWindow + 4 pantallas, con pytest-qt y offscreen.

``pdf2excel.core.pipeline.analyze/export`` se reemplazan por stubs rápidos (el worker los importa de
forma diferida), salvo el test de integración real con ``remito_grid.pdf``. Marcados ``ui``.
"""

from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import itertools
import shutil
import subprocess
import sys
import threading
import time
import types
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

import pytest
from PySide6.QtCore import QSettings, QStandardPaths, Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtTest import QSignalSpy
from PySide6.QtWidgets import QApplication, QDialog, QFileDialog, QLineEdit, QMessageBox, QWidget

from pdf2excel.errors import UserError
from pdf2excel.models import ColumnInfo, ConversionReport, Options, PageInfo, TypedTable
from pdf2excel.ui import main_window as mw
from pdf2excel.ui import strings_es as S
from pdf2excel.ui.file_list import FileRow
from pdf2excel.ui.main_window import MainWindow, PasswordDialog, analysis_key, is_large_file
from pdf2excel.ui.pages.done_page import DonePage
from pdf2excel.ui.pages.drop_page import DropPage
from pdf2excel.ui.pages.progress_page import ProgressPage
from pdf2excel.ui.pages.review_page import ReviewPage, as_text_table
from pdf2excel.ui.settings_dialog import SettingsDialog, load_options

pytestmark = pytest.mark.ui

CORPUS_DIRS = (
    Path("/tmp/claude-0/-home-user/932cdd7a-5305-5ce6-bcbe-bcac3f3b947f/scratchpad/t"),
    Path(__file__).parent / "corpus",
    Path(__file__).parent / "corpus" / "_generated",
)
REMITO = "remito_grid.pdf"
AnalyzeFn = Callable[[Path, Options, Any, threading.Event, Any], ConversionReport]


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


@pytest.fixture(autouse=True)
def no_open_url(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    """``QDesktopServices.openUrl`` no abre nada: registra las URLs."""
    opened: list[str] = []
    monkeypatch.setattr(QDesktopServices, "openUrl", staticmethod(lambda url: opened.append(url.toString())))
    return opened


@pytest.fixture
def pipeline_stub(monkeypatch: pytest.MonkeyPatch) -> types.ModuleType:
    """Módulo ``pdf2excel.core.pipeline`` sintético (el worker lo importa dentro de run())."""
    import pdf2excel.core as core_pkg

    stub = types.ModuleType("pdf2excel.core.pipeline")
    monkeypatch.setitem(sys.modules, "pdf2excel.core.pipeline", stub)
    monkeypatch.setattr(core_pkg, "pipeline", stub, raising=False)
    stub.analyze = fake_analyze  # type: ignore[attr-defined]
    stub.export = fake_export  # type: ignore[attr-defined]
    return stub


@pytest.fixture
def window_factory(qtbot: Any) -> Iterator[Callable[..., MainWindow]]:
    """Crea ventanas y, al terminar, cancela y espera la cola para que el cierre nunca pregunte nada."""
    windows: list[MainWindow] = []

    def make(initial_files: list[str] | None = None) -> MainWindow:
        window = MainWindow(initial_files=initial_files)
        qtbot.addWidget(window)
        window.show()
        windows.append(window)
        return window

    yield make
    for window in windows:
        window.queue.cancel_all()
        window.queue.wait_for_done(5000)
        if window.password_dialog is not None:
            window.password_dialog.reject()


def find_corpus_pdf(name: str) -> Path | None:
    for directory in CORPUS_DIRS:
        candidate = directory / name
        if candidate.is_file():
            return candidate
    return None


@pytest.fixture
def remito(tmp_path: Path) -> Path:
    src = find_corpus_pdf(REMITO)
    if src is None:
        pytest.skip(f"PDF de corpus no disponible: {REMITO}")
    dst = tmp_path / REMITO
    shutil.copy(src, dst)
    return dst


# --------------------------------------------------------------------------- stubs


def make_remito_report(path: Path, typed: bool = True) -> ConversionReport:
    """Reporte equivalente al análisis real de remito_grid.pdf (6 filas × 5 columnas, 350,00 en (0,3))."""
    header = ["Código", "Producto", "Cantidad", "Precio unitario", "Total"]
    rows: list[list[Any]] = [
        ["A-100", "Mezcal joven", 10, 350.0, 3500.0],
        ["A-101", "Tequila reposado", 5, 420.5, 2102.5],
        ["A-102", "Agua mineral", 48, 12.9, 619.2],
        ["A-103", "Hielo bolsa 5kg", 20, 35.0, 700.0],
        ["A-104", "Limón kg", 8, 28.75, 230.0],
        ["A-105", "Sal de gusano", 3, 90.0, 270.0],
    ]
    columns = [
        ColumnInfo(kind="id", header=header[0]),
        ColumnInfo(kind="text", header=header[1]),
        ColumnInfo(kind="number", locale="ar", number_format="#,##0", header=header[2]),
        ColumnInfo(kind="number", locale="ar", number_format="#,##0.00", header=header[3]),
        ColumnInfo(kind="number", locale="ar", number_format="#,##0.00", header=header[4]),
    ]
    table = TypedTable(name="Remito N° 4521", header=header, rows=rows, columns=columns, pages=(0, 0))
    if not typed:
        table = as_text_table(table)
    return ConversionReport(
        source=path,
        pages=[PageInfo(index=0, kind="digital", n_chars=320, image_coverage=0.0)],
        tables=[table],
        text_items=[],
    )


def fake_analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
    progress(1, 1, "classify")
    progress(1, 1, "tables")
    progress(1, 1, "typing")
    return make_remito_report(Path(path), typed=opts.typed_values)


def fake_export(report: ConversionReport, opts: Options) -> Path:
    folder = Path(opts.output_dir) if opts.output_dir else Path(report.source).parent
    out = folder / f"{Path(report.source).stem}.xlsx"
    report.output_path = out
    return out


def slow_analyze_factory(seconds: float = 2.0, started: threading.Event | None = None) -> AnalyzeFn:
    def slow_analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        if started is not None:
            started.set()
        progress(0, 1, "classify")
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            if cancel.is_set():
                raise UserError("cancelled")
            time.sleep(0.01)
        return make_remito_report(Path(path))

    return slow_analyze


def wait_page(qtbot: Any, window: MainWindow, page: QWidget, timeout: int = 5000) -> None:
    qtbot.waitUntil(lambda: window.stack.currentWidget() is page, timeout=timeout)


def wait_idle(qtbot: Any, window: MainWindow, timeout: int = 5000) -> None:
    qtbot.waitUntil(lambda: not window.queue.is_busy() and window.phase == "idle", timeout=timeout)


# --------------------------------------------------------------------------- (1) + (2) flujo con stubs


def test_flow_review_then_save(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    no_open_url: list[str],
) -> None:
    pdf = tmp_path / "remito_grid.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    visited: list[int] = []

    window = window_factory([str(pdf)])
    window.stack.currentChanged.connect(visited.append)
    assert window.windowTitle() == S.APP_TITLE
    assert window.stack.currentWidget() is window.drop_page  # arranca en la pantalla 1
    assert window.size().width() == 860 and window.size().height() == 560
    assert window.minimumSize().width() == 640 and window.minimumSize().height() == 440
    assert window.menuBar().actions() == []

    qtbot.waitUntil(lambda: str(pdf) in window.reports, timeout=5000)
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    review = window.review_page
    assert review.file_list.status_text(str(pdf)) == "1 tabla · 1 pág."
    assert review.file_list.state_of(str(pdf)) == "ok"
    model = review.table_view.model()
    assert model.rowCount() == 6 and model.columnCount() == 5
    assert model.index(0, 3).data(Qt.ItemDataRole.DisplayRole) == "350,00"
    assert review.table_combo.currentText() == "Tabla 1 de 1"
    assert "Remito N° 4521" in review.sheet_label.text()
    assert review.save_button.isEnabled() and review.save_button.property("primary") is True
    assert review.typed_check.isChecked()

    # (2) Guardar Excel → Convirtiendo → Listo
    review.set_output_dir(tmp_path / "salida")
    assert review.output_dir() == tmp_path / "salida"
    with qtbot.waitSignal(window.queue.signals.exported, timeout=5000):
        qtbot.mouseClick(review.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    assert window.stack.indexOf(window.progress_page) in visited  # pasó por la pantalla 3
    done = window.done_page
    lines = done.lines()
    assert len(lines) == 1 and lines[0].startswith("✔")
    assert "remito_grid.pdf → remito_grid.xlsx (1 tabla, 6 filas)" in lines[0]
    assert done.open_excel_button.isVisible() and done.open_folder_button.isVisible()
    assert done.title_label.text() == S.DONE_TITLE_ONE
    assert not done.review_note.isVisible()
    # open_when_done (por defecto) con un solo archivo abre el Excel.
    assert no_open_url and no_open_url[0].endswith("remito_grid.xlsx")
    # Las opciones elegidas quedan guardadas.
    assert load_options().output_dir == tmp_path / "salida"

    # Convertir otros → pantalla 1 limpia.
    qtbot.mouseClick(done.convert_more_button, Qt.MouseButton.LeftButton)
    assert window.stack.currentWidget() is window.drop_page
    assert window.reports == {} and window.order == [] and review.file_list.count() == 0


# --------------------------------------------------------------------------- integración real + (6) ticks


def test_real_conversion_keeps_event_loop_alive(
    qtbot: Any, window_factory: Callable[..., MainWindow], remito: Path, tmp_path: Path
) -> None:
    pytest.importorskip("pymupdf")
    ticks: list[float] = []
    timer = QTimer()
    timer.setInterval(50)
    timer.timeout.connect(lambda: ticks.append(time.monotonic()))
    timer.start()

    window = window_factory([str(remito)])
    window.opts.open_when_done = False
    qtbot.waitUntil(lambda: str(remito) in window.reports, timeout=30000)
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window, timeout=30000)
    review = window.review_page
    assert review.file_list.status_text(str(remito)) == "1 tabla · 1 pág."
    model = review.table_view.model()
    assert model.rowCount() == 6 and model.columnCount() == 5
    assert model.index(0, 3).data(Qt.ItemDataRole.DisplayRole) == "350,00"
    # Miniatura real de la primera página.
    row = review.file_list.itemWidget(review.file_list.item(0))
    assert isinstance(row, FileRow) and row.thumb.pixmap().height() > 0

    review.set_output_dir(tmp_path)
    with qtbot.waitSignal(window.queue.signals.exported, timeout=30000):
        qtbot.mouseClick(review.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page, timeout=30000)
    # En un proceso ya "caliente" la conversión entera tarda < 0,5 s: seguir muestreando hasta 10 ticks.
    qtbot.waitUntil(lambda: len(ticks) >= 10, timeout=5000)
    timer.stop()

    assert (tmp_path / "remito_grid.xlsx").exists()
    lines = window.done_page.lines()
    assert lines and lines[0].startswith("✔") and "remito_grid.xlsx" in lines[0]
    assert window.done_page.open_excel_button.isVisible()

    assert len(ticks) >= 10
    gaps = [b - a for a, b in itertools.pairwise(ticks)]
    assert max(gaps) < 0.300, f"gap máximo del event loop: {max(gaps) * 1000:.0f} ms"


# --------------------------------------------------------------------------- (3) contraseña


def test_encrypted_password_dialog(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    pdf = tmp_path / "protegido.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    seen_passwords: list[Any] = []

    def locked_analyze(path: Path, opts: Options, progress: Any, cancel: Any, password: Any) -> Any:
        seen_passwords.append(password)
        if password == "1234":
            return make_remito_report(Path(path))
        raise UserError("encrypted", action="password")

    pipeline_stub.analyze = locked_analyze  # type: ignore[attr-defined]
    window = window_factory([str(pdf)])

    qtbot.waitUntil(lambda: window.password_dialog is not None and window.password_dialog.isVisible())
    dialog = window.password_dialog
    assert isinstance(dialog, PasswordDialog)
    assert isinstance(dialog.password_edit, QLineEdit)
    assert dialog.password_edit.echoMode() == QLineEdit.EchoMode.Password
    assert dialog.continue_button.text() == S.BTN_CONTINUE and dialog.skip_button.text() == S.BTN_SKIP

    # Saltar → fila en error con el título de ERRORS["encrypted"].
    qtbot.mouseClick(dialog.skip_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: window.password_dialog is None)
    wait_idle(qtbot, window)
    assert window.stack.currentWidget() is window.review_page
    assert window.review_page.file_list.status_text(str(pdf)) == S.ERRORS["encrypted"][0]
    assert window.review_page.file_list.state_of(str(pdf)) == "error"
    assert not window.review_page.save_button.isEnabled()
    assert window.errors[str(pdf)].code == "encrypted"
    assert "⚠" in window.review_page.review_label.text()

    # Segundo intento con contraseña correcta: se reenvía con {path: pwd} y llega el reporte.
    other = tmp_path / "protegido2.pdf"
    other.write_bytes(b"%PDF-1.4 stub")
    window.start_analysis([str(other)])
    qtbot.waitUntil(lambda: window.password_dialog is not None and window.password_dialog.isVisible())
    dialog = window.password_dialog
    assert dialog is not None and dialog.path == str(other)
    dialog.password_edit.setText("1234")
    with qtbot.waitSignal(window.queue.signals.analyzed, timeout=5000):
        qtbot.mouseClick(dialog.continue_button, Qt.MouseButton.LeftButton)
    wait_idle(qtbot, window)
    assert seen_passwords[-1] == "1234"
    assert window.review_page.file_list.state_of(str(other)) == "ok"
    assert window.review_page.save_button.isEnabled()


# --------------------------------------------------------------------------- (4) Excel abierto


def test_file_in_use_retry(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    attempts: list[int] = []

    def locked_export(report: ConversionReport, opts: Options) -> Path:
        attempts.append(1)
        if len(attempts) == 1:
            raise UserError("file_in_use", action="retry")
        return fake_export(report, opts)

    pipeline_stub.export = locked_export  # type: ignore[attr-defined]
    boxes: list[tuple[str, list[str]]] = []

    def fake_exec(self: QMessageBox) -> int:
        labels = [b.text() for b in self.buttons()]
        boxes.append((self.windowTitle(), labels))
        retry = next(b for b in self.buttons() if b.text() == S.BTN_RETRY)
        retry.click()  # como el usuario: QMessageBox registra clickedButton() y cierra
        return 0

    monkeypatch.setattr(QMessageBox, "exec", fake_exec)

    window = window_factory([str(pdf)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page, timeout=10000)

    assert len(boxes) == 1
    title, labels = boxes[0]
    assert title == S.ERRORS["file_in_use"][0]
    assert sorted(labels) == sorted([S.BTN_RETRY, S.BTN_SAVE_AS, S.BTN_SKIP])
    assert len(attempts) == 2
    lines = window.done_page.lines()
    assert len(lines) == 1 and lines[0].startswith("✔")
    assert not window.errors


def test_file_in_use_skip_shows_warning_line(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")

    def locked_export(report: ConversionReport, opts: Options) -> Path:
        raise UserError("file_in_use", action="retry")

    pipeline_stub.export = locked_export  # type: ignore[attr-defined]
    # exec() reemplazado devolviendo el índice (orden de alta: Reintentar, Guardar con otro nombre, Saltar).
    monkeypatch.setattr(QMessageBox, "exec", lambda self: 2)

    window = window_factory([str(pdf)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    lines = window.done_page.lines()
    assert lines == [S.error_line("remito.pdf", UserError("file_in_use"))]
    assert lines[0].startswith("⚠")
    assert window.done_page.title_label.text() == S.DONE_NOTHING
    assert not window.done_page.open_excel_button.isVisible()


# --------------------------------------------------------------------------- (5) cancelación


def test_cancel_during_analysis(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    fast = tmp_path / "chico.pdf"
    fast.write_bytes(b"%PDF-1.4 stub")
    big = tmp_path / "grande.pdf"
    big.write_bytes(b"%PDF-1.4" + b"\0" * (2 * 1024 * 1024 + 10))  # > 2 MB → pantalla de progreso
    assert is_large_file(str(big)) and not is_large_file(str(fast))
    started = threading.Event()
    slow = slow_analyze_factory(4.0, started)  # no puede terminar solo antes de que cancelemos

    def analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        if Path(path).name == "grande.pdf":
            return slow(path, opts, progress, cancel, password)
        return fake_analyze(path, opts, progress, cancel, password)

    pipeline_stub.analyze = analyze  # type: ignore[attr-defined]
    window = window_factory()
    failed: list[tuple[str, UserError]] = []
    window.queue.signals.failed.connect(lambda p, e: failed.append((p, e)))

    window.start_analysis([str(big), str(fast)])
    assert window.stack.currentWidget() is window.progress_page
    assert window.progress_page.mode() == "analyze"
    assert started.wait(2.0)
    # Esperar por estado y no por señal: `analyzed` del chico puede haberse emitido antes de conectar
    # un waitSignal (las conexiones encoladas solo entregan a quien ya estaba conectado al emitir).
    qtbot.waitUntil(lambda: str(fast) in window.reports, timeout=5000)
    wait_page(qtbot, window, window.review_page)  # el primero que termina ya se muestra
    assert str(big) not in window.reports and window.queue.is_busy()

    t0 = time.monotonic()
    window.progress_page.cancel_requested.emit()
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=3000)
    assert time.monotonic() - t0 < 3.0
    wait_idle(qtbot, window)
    assert window.stack.currentWidget() is window.review_page
    assert window.review_page.file_list.state_of(str(fast)) == "ok"
    assert window.review_page.file_list.state_of(str(big)) == "error"
    assert window.review_page.file_list.status_text(str(big)) == S.ERRORS["cancelled"][0]
    assert window.errors[str(big)].code == "cancelled"
    assert all(e.code == "cancelled" for _p, e in failed)
    assert window.review_page.save_button.isEnabled()  # lo analizado se conserva


def test_cancel_everything_returns_to_drop_page(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    big = tmp_path / "grande.pdf"
    big.write_bytes(b"%PDF-1.4" + b"\0" * (2 * 1024 * 1024 + 10))
    started = threading.Event()
    pipeline_stub.analyze = slow_analyze_factory(2.0, started)  # type: ignore[attr-defined]
    window = window_factory([str(big)])
    wait_page(qtbot, window, window.progress_page)
    assert started.wait(2.0)
    qtbot.mouseClick(window.progress_page.cancel_button, Qt.MouseButton.LeftButton)
    assert window.progress_page.is_cancelling()
    assert window.progress_page.stage_label.text() == S.PROGRESS_CANCELLING
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=3000)
    wait_idle(qtbot, window)
    assert window.stack.currentWidget() is window.drop_page
    assert not window.drop_page.is_busy()
    assert window.order == []


def test_cancel_during_export_keeps_exported(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    pdfs = [tmp_path / f"r{i}.pdf" for i in range(3)]
    for pdf in pdfs:
        pdf.write_bytes(b"%PDF-1.4 stub")
    first_exported = threading.Event()
    release = threading.Event()

    def export(report: ConversionReport, opts: Options) -> Path:
        out = fake_export(report, opts)
        if Path(report.source).name == "r0.pdf":
            first_exported.set()
        else:
            release.wait(3.0)  # los demás esperan hasta que el test cancele
        return out

    pipeline_stub.export = export  # type: ignore[attr-defined]
    window = window_factory([str(p) for p in pdfs])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    assert len(window.reports) == 3

    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    assert window.stack.currentWidget() is window.progress_page
    assert window.progress_page.mode() == "export"
    assert first_exported.wait(3.0)
    qtbot.waitUntil(lambda: str(pdfs[0]) in window.exported, timeout=3000)
    window.progress_page.cancel_requested.emit()
    release.set()
    wait_page(qtbot, window, window.done_page)
    wait_idle(qtbot, window)
    lines = window.done_page.lines()
    assert lines[0] == S.DONE_CANCELLED
    assert any(line.startswith("✔ r0.pdf") for line in lines)
    assert window.done_page.open_excel_button.isVisible()


# --------------------------------------------------------------------------- (7) teclado


def test_escape_on_review_emits_back(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    window = window_factory([str(pdf)])
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    spy = QSignalSpy(window.review_page.back_requested)
    qtbot.keyClick(window.review_page, Qt.Key.Key_Escape)
    assert spy.count() == 1
    assert window.stack.currentWidget() is window.drop_page
    assert window.reports == {}

    # ReviewPage suelta: Esc también emite back_requested, Enter = Guardar (si está habilitado).
    page = ReviewPage()
    qtbot.addWidget(page)
    page.show()
    back = QSignalSpy(page.back_requested)
    save = QSignalSpy(page.save_requested)
    qtbot.keyClick(page, Qt.Key.Key_Escape)
    assert back.count() == 1
    page.primary_action()
    assert save.count() == 0  # sin reportes, Guardar está deshabilitado
    page.add_report(str(pdf), make_remito_report(pdf))
    page.primary_action()
    assert save.count() == 1


def test_ctrl_o_opens_file_dialog(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    calls: list[str] = []

    def fake_dialog(parent: Any, title: str, directory: str, filt: str) -> tuple[list[str], str]:
        calls.append(title)
        return [str(pdf)], filt

    monkeypatch.setattr(QFileDialog, "getOpenFileNames", staticmethod(fake_dialog))
    window = window_factory()
    assert window.focusWidget() is window.drop_page.drop_zone or window.drop_page.drop_zone.hasFocus()
    with qtbot.waitSignal(window.queue.signals.analyzed, timeout=5000):
        qtbot.keyClick(window, Qt.Key.Key_O, Qt.KeyboardModifier.ControlModifier)
    assert calls == [S.FILE_DIALOG_TITLE]
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    assert window.order == [str(pdf)]

    # Enter sobre un botón con foco activa ese botón (Volver), no la acción primaria.
    window.review_page.back_button.setFocus()
    qtbot.keyClick(window, Qt.Key.Key_Return)
    assert window.stack.currentWidget() is window.drop_page


def test_enter_on_done_page_converts_more(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    window = window_factory([str(pdf)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.keyClick(window.review_page.table_view, Qt.Key.Key_Return)  # Enter = Guardar Excel
    wait_page(qtbot, window, window.done_page)
    assert window.done_page.convert_more_button.hasFocus()
    qtbot.keyClick(window, Qt.Key.Key_Return)
    assert window.stack.currentWidget() is window.drop_page


# --------------------------------------------------------------------------- casilla de tipado


def test_typed_toggle_rerenders_and_reanalyzes(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    analyzed_with: list[bool] = []

    def analyze(path: Path, opts: Options, progress: Any, cancel: Any, password: Any) -> Any:
        analyzed_with.append(opts.typed_values)
        return fake_analyze(path, opts, progress, cancel, password)

    pipeline_stub.analyze = analyze  # type: ignore[attr-defined]
    window = window_factory([str(pdf)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    review = window.review_page
    model = review.table_view.model()
    assert model.index(0, 3).data(Qt.ItemDataRole.DisplayRole) == "350,00"
    assert model.index(0, 2).data(Qt.ItemDataRole.TextAlignmentRole) & Qt.AlignmentFlag.AlignRight

    review.typed_check.setChecked(False)  # sin reanalizar: str(valor) sin formato
    assert analyzed_with == [True]
    model = review.table_view.model()
    assert model.index(0, 3).data(Qt.ItemDataRole.DisplayRole) == "350.0"
    assert not (model.index(0, 2).data(Qt.ItemDataRole.TextAlignmentRole) & Qt.AlignmentFlag.AlignRight)
    review.typed_check.setChecked(True)
    assert review.table_view.model().index(0, 3).data(Qt.ItemDataRole.DisplayRole) == "350,00"

    # Guardar con la casilla desmarcada reanaliza con typed_values=False y después exporta.
    review.typed_check.setChecked(False)
    qtbot.mouseClick(review.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    wait_idle(qtbot, window)
    assert analyzed_with == [True, False]
    assert load_options().typed_values is False
    assert analysis_key(window.opts) == analysis_key(Options(typed_values=False))
    assert window.done_page.lines()[0].startswith("✔")


# --------------------------------------------------------------------------- errores de análisis y lote


def test_damaged_file_does_not_stop_batch(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    good = tmp_path / "bueno.pdf"
    bad = tmp_path / "roto.pdf"
    for pdf in (good, bad):
        pdf.write_bytes(b"%PDF-1.4 stub")

    def analyze(path: Path, opts: Options, progress: Any, cancel: Any, password: Any) -> Any:
        if Path(path).name == "roto.pdf":
            raise UserError("damaged", action="skip")
        return fake_analyze(path, opts, progress, cancel, password)

    pipeline_stub.analyze = analyze  # type: ignore[attr-defined]
    window = window_factory([str(good), str(bad)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    assert window.review_page.file_list.state_of(str(bad)) == "error"
    assert window.review_page.file_list.status_text(str(bad)) == "No pudimos leer roto.pdf"
    assert window.review_page.save_button.isEnabled()

    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    lines = window.done_page.lines()
    assert len(lines) == 2
    assert lines[0].startswith("✔ bueno.pdf")
    assert lines[1] == S.error_line("roto.pdf", UserError("damaged"))
    assert window.done_page.title_label.text() == S.DONE_TITLE_ONE


def test_single_workbook_uses_export_batch(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    pdfs = [tmp_path / "a.pdf", tmp_path / "b.pdf"]
    for pdf in pdfs:
        pdf.write_bytes(b"%PDF-1.4 stub")
    batches: list[int] = []

    def export_batch(reports: list[ConversionReport], opts: Options) -> Path:
        batches.append(len(reports))
        out = tmp_path / "Conversion 2026-11-01.xlsx"
        for r in reports:
            r.output_path = out
        return out

    def export_single(report: ConversionReport, opts: Options) -> Path:
        raise AssertionError("con single_workbook no debe llamarse export()")

    pipeline_stub.export_batch = export_batch  # type: ignore[attr-defined]
    pipeline_stub.export = export_single  # type: ignore[attr-defined]
    window = window_factory([str(p) for p in pdfs])
    window.opts.open_when_done = False
    window.opts.single_workbook = True
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    wait_idle(qtbot, window)
    assert batches == [2]
    lines = window.done_page.lines()
    assert len(lines) == 2
    assert all(line.startswith("✔") and "Conversion 2026-11-01.xlsx" in line for line in lines)
    assert window.done_page.title_label.text() == S.done_title(2, 1)
    assert window.done_page.title_label.text() == "¡Listo! Guardamos 1 archivo con 2 PDF"


def test_settings_gear_saves_options(
    qtbot: Any, window_factory: Callable[..., MainWindow], monkeypatch: pytest.MonkeyPatch
) -> None:
    def fake_exec(self: SettingsDialog) -> int:
        self.single_workbook_check.setChecked(True)
        self.lang_combo.setCurrentIndex(self.lang_combo.findData("spa"))
        return int(QDialog.DialogCode.Accepted)

    monkeypatch.setattr(SettingsDialog, "exec", fake_exec)
    window = window_factory()
    assert window.settings_button.accessibleName() == S.BTN_SETTINGS
    qtbot.mouseClick(window.settings_button, Qt.MouseButton.LeftButton)
    assert window.opts.single_workbook is True and window.opts.ocr_lang == "spa"
    assert load_options().ocr_lang == "spa"


def test_close_while_busy_asks_and_cancels(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf = tmp_path / "lento.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    started = threading.Event()
    pipeline_stub.analyze = slow_analyze_factory(2.0, started)  # type: ignore[attr-defined]
    questions: list[str] = []
    answers = [QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes]

    def fake_question(parent: Any, title: str, text: str, *args: Any, **kwargs: Any) -> Any:
        questions.append(text)
        return answers.pop(0)

    monkeypatch.setattr(QMessageBox, "question", staticmethod(fake_question))
    window = window_factory([str(pdf)])
    qtbot.waitUntil(started.is_set, timeout=3000)
    assert window.queue.is_busy()
    assert window.close() is False  # "No": sigue abierta
    assert window.isVisible()
    t0 = time.monotonic()
    assert window.close() is True  # "Sí": cancela y cierra
    assert time.monotonic() - t0 < 3.5
    assert questions == ["¿Cancelar la conversión en curso?"] * 2
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=3000)


def _single_workbook_stub(stub: types.ModuleType, out: Path) -> None:
    def export_batch(reports: list[ConversionReport], opts: Options) -> Path:
        for r in reports:
            r.output_path = out
        return out

    stub.export_batch = export_batch  # type: ignore[attr-defined]


@pytest.mark.parametrize("n_pdfs", [2, 4])
def test_single_workbook_done_page_opens_the_book_once(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    no_open_url: list[str],
    n_pdfs: int,
) -> None:
    """Libro único: 'Abrir Excel al terminar' abre ese libro una vez, el botón se ve y no repite aperturas."""
    pdfs = [tmp_path / f"r{i}.pdf" for i in range(n_pdfs)]
    for pdf in pdfs:
        pdf.write_bytes(b"%PDF-1.4 stub")
    out = tmp_path / "Conversion 2026-11-01.xlsx"
    _single_workbook_stub(pipeline_stub, out)
    window = window_factory([str(p) for p in pdfs])
    window.opts.single_workbook = True
    window.opts.open_when_done = True
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    wait_idle(qtbot, window)
    assert no_open_url == [QUrl.fromLocalFile(str(out)).toString()]  # auto-apertura: una sola vez
    assert window.done_page.open_excel_button.isVisible()  # un libro, aunque sean 4 PDF
    assert window.done_page.output_paths() == [out]
    assert window.done_page.title_label.text() == f"¡Listo! Guardamos 1 archivo con {n_pdfs} PDF"
    no_open_url.clear()
    qtbot.mouseClick(window.done_page.open_excel_button, Qt.MouseButton.LeftButton)
    assert no_open_url == [QUrl.fromLocalFile(str(out)).toString()]
    no_open_url.clear()
    qtbot.mouseClick(window.done_page.open_folder_button, Qt.MouseButton.LeftButton)
    assert no_open_url == [QUrl.fromLocalFile(str(out.parent)).toString()]


# --------------------------------------------------------------------------- hallazgos de la revisión


def test_save_as_uses_the_chosen_filename(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """'Guardar con otro nombre' escribe el archivo que tipeó el usuario, no <pdf>.xlsx en esa carpeta."""
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    attempts: list[Path | None] = []

    def locked_export(report: ConversionReport, opts: Options, dst: Path | None = None) -> Path:
        attempts.append(dst)
        if dst is None:
            raise UserError("file_in_use", action="retry")
        dst.write_bytes(b"xlsx")
        report.output_path = dst
        return dst

    pipeline_stub.export = locked_export  # type: ignore[attr-defined]
    chosen = tmp_path / "otra" / "remito_para_contabilidad"  # sin extensión: se agrega .xlsx
    chosen.parent.mkdir()
    monkeypatch.setattr(QMessageBox, "exec", lambda self: 1)  # índice 1 = "Guardar con otro nombre"
    monkeypatch.setattr(
        QFileDialog, "getSaveFileName", staticmethod(lambda *a, **k: (str(chosen), "Excel (*.xlsx)"))
    )
    window = window_factory([str(pdf)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page, timeout=10000)
    wait_idle(qtbot, window)
    expected = chosen.with_suffix(".xlsx")
    assert attempts == [None, expected]
    assert window.exported[str(pdf)].output_path == expected and expected.is_file()
    assert window.done_page.lines() == [S.done_line(window.exported[str(pdf)])]
    assert "remito_para_contabilidad.xlsx" in window.done_page.lines()[0]
    assert not window.errors


def _late_analyze(delay: float, outcome: str) -> AnalyzeFn:
    """Analiza ignorando `cancel` durante `delay` s (como un tramo real entre dos chequeos)."""

    def analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        if Path(path).name.startswith("lento"):
            progress(0, 1, "classify")
            time.sleep(delay)
            if outcome == "error":
                raise UserError("damaged")
            return make_remito_report(Path(path))
        return fake_analyze(path, opts, progress, cancel, password)

    return analyze


def test_stale_report_after_back_is_ignored(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    """Un análisis que termina después de 'Volver' no reaparece en 'Esto encontré' ni habilita Guardar."""
    fast, slow = tmp_path / "rapido.pdf", tmp_path / "lento.pdf"
    for pdf in (fast, slow):
        pdf.write_bytes(b"%PDF-1.4 stub")
    pipeline_stub.analyze = _late_analyze(0.8, "report")  # type: ignore[attr-defined]
    window = window_factory()
    window.start_analysis([str(fast), str(slow)])
    qtbot.waitUntil(lambda: str(fast) in window.reports, timeout=5000)
    wait_page(qtbot, window, window.review_page)
    assert window.queue.is_busy()
    qtbot.mouseClick(window.review_page.back_button, Qt.MouseButton.LeftButton)
    assert window.stack.currentWidget() is window.drop_page
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=5000)
    qtbot.wait(300)  # que se procesen las señales encoladas del job tardío
    assert window.reports == {} and window.errors == {} and window.review_page.paths() == []
    assert not window.review_page.save_enabled() and window.phase == "idle"
    assert window.stack.currentWidget() is window.drop_page

    nuevo = tmp_path / "nuevo.pdf"
    nuevo.write_bytes(b"%PDF-1.4 stub")
    window.start_analysis([str(nuevo)])
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    assert window.review_page.paths() == [str(nuevo)] and list(window.reports) == [str(nuevo)]


def test_stale_error_after_back_does_not_leak_into_next_done_page(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    fast, slow = tmp_path / "rapido.pdf", tmp_path / "lento_roto.pdf"
    for pdf in (fast, slow):
        pdf.write_bytes(b"%PDF-1.4 stub")
    pipeline_stub.analyze = _late_analyze(0.8, "error")  # type: ignore[attr-defined]
    window = window_factory()
    window.opts.open_when_done = False
    window.start_analysis([str(fast), str(slow)])
    qtbot.waitUntil(lambda: str(fast) in window.reports, timeout=5000)
    wait_page(qtbot, window, window.review_page)
    qtbot.mouseClick(window.review_page.back_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=5000)
    qtbot.wait(300)
    assert window.errors == {}

    nuevo = tmp_path / "nuevo.pdf"
    nuevo.write_bytes(b"%PDF-1.4 stub")
    window.start_analysis([str(nuevo)])
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    qtbot.mouseClick(window.review_page.save_button, Qt.MouseButton.LeftButton)
    wait_page(qtbot, window, window.done_page)
    wait_idle(qtbot, window)
    lines = window.done_page.lines()
    assert len(lines) == 1 and lines[0].startswith("✔ nuevo.pdf")
    assert window.done_page.errors() == {}


def test_enter_on_typed_checkbox_saves_instead_of_toggling(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    """Enter con el foco en 'Convertir números y fechas' = acción primaria (Guardar), no alternar."""
    pdf = tmp_path / "remito.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    window = window_factory([str(pdf)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    check = window.review_page.typed_check
    assert check.isChecked()
    check.setFocus(Qt.FocusReason.TabFocusReason)
    qtbot.waitUntil(lambda: QApplication.focusWidget() is check, timeout=2000)
    qtbot.keyClick(check, Qt.Key.Key_Return)
    wait_page(qtbot, window, window.done_page)
    assert check.isChecked()
    assert window.exported and window.opts.typed_values is True


def test_ocr_stage_switches_to_progress_page_with_cancel(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    """Un escaneo chico (< 2 MB) que pasa por OCR muestra la pantalla de progreso (barra + Cancelar)."""
    pdf = tmp_path / "escaneo_chico.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub" + b"\0" * (300 * 1024))
    started = threading.Event()

    def analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        progress(0, 12, "classify")
        for i in range(1, 13):
            progress(i, 12, "ocr")
            started.set()
            deadline = time.monotonic() + 0.2
            while time.monotonic() < deadline:
                if cancel.is_set():
                    raise UserError("cancelled")
                time.sleep(0.01)
        return make_remito_report(Path(path))

    pipeline_stub.analyze = analyze  # type: ignore[attr-defined]
    window = window_factory()
    assert not is_large_file(str(pdf))
    window.start_analysis([str(pdf)])
    assert started.wait(2.0)
    wait_page(qtbot, window, window.progress_page, timeout=2000)
    assert not window.drop_page.is_busy()
    assert window.progress_page.cancel_button.isVisible() and window.progress_page.cancel_button.isEnabled()
    qtbot.waitUntil(lambda: "OCR" in window.progress_page.stage_label.text(), timeout=2000)
    qtbot.mouseClick(window.progress_page.cancel_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=3000)
    wait_page(qtbot, window, window.drop_page)


def test_slow_analysis_on_drop_page_switches_to_progress(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    """Sin OCR pero con un análisis que se alarga, a los SLOW_ANALYSIS_MS se pasa a la pantalla 3."""
    pdf = tmp_path / "lento.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    started = threading.Event()
    pipeline_stub.analyze = slow_analyze_factory(3.0, started)  # type: ignore[attr-defined]
    window = window_factory()
    window._slow_timer.setInterval(300)
    window.start_analysis([str(pdf)])
    assert started.wait(2.0)
    assert window.stack.currentWidget() is window.drop_page and window.drop_page.is_busy()
    wait_page(qtbot, window, window.progress_page, timeout=2000)
    assert not window.drop_page.is_busy()
    qtbot.mouseClick(window.progress_page.cancel_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.queue.is_busy(), timeout=3000)


def test_save_disabled_while_added_files_are_analyzing(
    qtbot: Any, pipeline_stub: types.ModuleType, window_factory: Callable[..., MainWindow], tmp_path: Path
) -> None:
    """Agregar PDF desde 'Esto encontré' deshabilita Guardar hasta que terminen de analizarse."""
    first, slow = tmp_path / "primero.pdf", tmp_path / "lento.pdf"
    for pdf in (first, slow):
        pdf.write_bytes(b"%PDF-1.4 stub")
    started = threading.Event()

    def analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        if Path(path).name == "lento.pdf":
            return slow_analyze_factory(1.0, started)(path, opts, progress, cancel, password)
        return fake_analyze(path, opts, progress, cancel, password)

    pipeline_stub.analyze = analyze  # type: ignore[attr-defined]
    window = window_factory([str(first)])
    window.opts.open_when_done = False
    wait_page(qtbot, window, window.review_page)
    wait_idle(qtbot, window)
    assert window.review_page.save_enabled()
    window.start_analysis([str(slow)])  # lo que hace Ctrl+O desde la pantalla 2
    assert started.wait(2.0)
    assert not window.review_page.save_enabled()
    assert window.stack.currentWidget() is window.review_page
    wait_idle(qtbot, window, timeout=5000)
    assert window.review_page.save_enabled()
    assert window.review_page.paths() == [str(first), str(slow)]


def test_close_with_stubborn_job_defers_until_it_ends(
    qtbot: Any,
    pipeline_stub: types.ModuleType,
    window_factory: Callable[..., MainWindow],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Un job que no honra cancel en CLOSE_WAIT_MS: la ventana se oculta y se cierra cuando termina, sin
    'Signal source has been deleted' ni 'Fallo inesperado' en el log."""
    pdf = tmp_path / "terco.pdf"
    pdf.write_bytes(b"%PDF-1.4 stub")
    started = threading.Event()
    finished = threading.Event()

    def analyze(path: Path, opts: Options, progress: Any, cancel: threading.Event, password: Any) -> Any:
        started.set()
        progress(0, 1, "classify")
        time.sleep(1.2)  # tramo sin chequeo de cancelación
        progress(1, 1, "tables")
        finished.set()
        return make_remito_report(Path(path))

    pipeline_stub.analyze = analyze  # type: ignore[attr-defined]
    monkeypatch.setattr(mw, "CLOSE_WAIT_MS", 150)
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes))
    window = window_factory([str(pdf)])
    qtbot.waitUntil(started.is_set, timeout=3000)
    with caplog.at_level("INFO"):
        assert window.close() is False  # diferido: el job sigue vivo
        assert window.is_closing() and not window.isVisible()
        assert window.queue.is_busy()
        qtbot.waitUntil(
            lambda: not window.queue.is_busy() and not window._close_timer.isActive(), timeout=5000
        )
    assert finished.is_set()
    assert not window.isVisible()
    assert "Fallo inesperado" not in caplog.text and "Signal source has been deleted" not in caplog.text


# --------------------------------------------------------------------------- pantallas sueltas


def test_drop_page_busy_and_update(qtbot: Any) -> None:
    page = DropPage()
    qtbot.addWidget(page)
    page.show()
    assert page.accessibleName() == S.DROP_TITLE
    assert page.privacy_label.text() == S.PRIVACY_NOTE
    assert not page.busy_box.isVisible() and not page.update_visible()

    page.set_busy(True)
    assert page.busy_box.isVisible() and page.busy_label.text() == S.ANALYZING
    assert not page.drop_zone.isEnabled() and not page.drop_zone.acceptDrops()
    got: list[list[str]] = []
    page.files_chosen.connect(got.append)
    page.drop_zone.files_dropped.emit(["/x/a.pdf"])
    assert got == []  # ocupado: se ignora
    page.set_busy(False)
    assert page.drop_zone.isEnabled() and page.drop_zone.acceptDrops()
    page.drop_zone.files_dropped.emit(["/x/a.pdf"])
    assert got == [["/x/a.pdf"]]

    page.show_update("0.2.0", "https://example.com/setup.exe")
    assert page.update_visible()
    assert "Hay una versión nueva (0.2.0)" in page.update_label.text()
    assert 'href="https://example.com/setup.exe"' in page.update_label.text()
    assert page.update_label.objectName() == "UpdateNote"
    dismissed = QSignalSpy(page.update_dismissed)
    qtbot.mouseClick(page.update_close, Qt.MouseButton.LeftButton)
    assert dismissed.count() == 1 and not page.update_visible()


def test_progress_page_texts(qtbot: Any) -> None:
    page = ProgressPage()
    qtbot.addWidget(page)
    page.show()
    page.set_mode("export")
    assert page.title_label.text() == S.PROGRESS_TITLE
    page.set_progress("/x/remito.pdf", 12, 50, "tables", 2, 5)
    assert page.file_label.text() == "remito.pdf"
    assert page.counter_label.text() == "Archivo 2 de 5 · página 12 de 50"
    assert page.bar.maximum() == 50 and page.bar.value() == 12
    assert page.stage_label.text() == S.STAGE_TEXT["tables"]
    page.set_progress("/x/remito.pdf", 1, 1, "excel", 2, 5)
    assert page.counter_label.text() == "Archivo 2 de 5"
    assert page.stage_label.text() == S.STAGE_TEXT["excel"]
    spy = QSignalSpy(page.cancel_requested)
    qtbot.keyClick(page, Qt.Key.Key_Escape)
    assert spy.count() == 1 and page.is_cancelling() and not page.cancel_button.isEnabled()
    qtbot.keyClick(page, Qt.Key.Key_Escape)
    assert spy.count() == 1  # cancelar dos veces no emite dos veces
    page.reset()
    assert page.cancel_button.isEnabled() and page.bar.value() == 0
    page.set_mode("analyze")
    assert page.title_label.text() == S.ANALYZING


def test_done_page_results(qtbot: Any, tmp_path: Path) -> None:
    page = DonePage()
    qtbot.addWidget(page)
    page.show()
    reports = []
    for i in range(4):
        rep = make_remito_report(tmp_path / f"r{i}.pdf")
        rep.output_path = tmp_path / f"r{i}.xlsx"
        reports.append(rep)
    reports[0].tables[0].review_cells = {(0, 3), (1, 3)}
    errors = {str(tmp_path / "roto.pdf"): UserError("damaged")}

    page.set_results(reports, errors)
    assert page.title_label.text() == "¡Listo! Guardamos 4 archivos"
    lines = page.lines()
    assert len(lines) == 5
    assert lines[0] == "✔ r0.pdf → r0.xlsx (1 tabla, 6 filas)"
    assert lines[-1] == "⚠ roto.pdf: no se pudo leer, parece dañado o no es un PDF"
    assert page.review_note.isVisible()
    assert page.review_note.text() == "⚠ 2 celdas marcadas para revisar en amarillo"
    assert not page.open_excel_button.isVisible()  # más de 3 archivos
    assert page.open_folder_button.isVisible()
    assert page.convert_more_button.property("primary") is True

    reports[1].tables[0].review_cells = {(0, 3)}
    page.set_results(reports[1:2], {})
    assert page.review_note.text() == "⚠ 1 celda marcada para revisar en amarillo"
    reports[1].tables[0].review_cells = set()

    page.set_results(reports[:3], {})
    assert page.open_excel_button.isVisible()
    assert not page.review_note.isVisible() or reports[0].n_review > 0

    # Libro único del lote: 4 PDF → 1 .xlsx: el botón se ve, la lista de libros no repite y el título lo dice.
    for rep in reports:
        rep.output_path = tmp_path / "Conversion 2026-11-01.xlsx"
    page.set_results(reports, {})
    assert page.open_excel_button.isVisible()
    assert page.output_paths() == [tmp_path / "Conversion 2026-11-01.xlsx"]
    assert page.title_label.text() == "¡Listo! Guardamos 1 archivo con 4 PDF"
    for i, rep in enumerate(reports):
        rep.output_path = tmp_path / f"r{i}.xlsx"
    page.set_results([], {}, cancelled=True)
    assert page.lines() == [S.DONE_CANCELLED]
    assert page.title_label.text() == S.DONE_NOTHING
    assert not page.open_excel_button.isVisible() and not page.open_folder_button.isVisible()
    more = QSignalSpy(page.convert_more)
    qtbot.keyClick(page, Qt.Key.Key_Escape)
    assert more.count() == 1


def test_review_page_folder_selector(qtbot: Any, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    page = ReviewPage()
    qtbot.addWidget(page)
    page.show()
    assert page.output_dir() is None
    assert page.folder_combo.currentText() == S.SAVE_NEXT_TO_PDF
    assert page.folder_combo.itemText(page.folder_combo.count() - 1) == S.SAVE_CHOOSE

    chosen = tmp_path / "salida"
    chosen.mkdir()
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: str(chosen)))
    page.folder_combo.activated.emit(page.folder_combo.count() - 1)
    assert page.output_dir() == chosen
    assert page.folder_combo.currentText() == S.SAVE_IN_FOLDER_FMT.format(folder="salida")
    assert page.folder_combo.count() == 3

    # Cancelar el diálogo vuelve a la opción anterior.
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: ""))
    page.folder_combo.activated.emit(page.folder_combo.count() - 1)
    assert page.output_dir() == chosen
    page.folder_combo.setCurrentIndex(0)
    assert page.output_dir() is None
    page.set_output_dir(None)
    assert page.folder_combo.count() == 2

    # Varias tablas: combo "Tabla i de n" y nombre de hoja.
    report = make_remito_report(tmp_path / "r.pdf")
    second = as_text_table(report.tables[0])
    second.name = "Totales"
    second.pages = (1, 2)
    report.tables.append(second)
    page.add_report(str(tmp_path / "r.pdf"), report)
    assert [page.table_combo.itemText(i) for i in range(2)] == ["Tabla 1 de 2", "Tabla 2 de 2"]
    assert page.table_combo.isEnabled()
    page.table_combo.setCurrentIndex(1)
    assert page.sheet_label.text() == "Totales · páginas 2 a 3"
    assert page.review_label.text() == S.REVIEW_NONE

    empty = ConversionReport(source=tmp_path / "t.pdf", pages=[], tables=[], text_items=[])
    page.add_report(str(tmp_path / "t.pdf"), empty)
    page.file_list.select_file(str(tmp_path / "t.pdf"))
    assert page.file_list.status_text(str(tmp_path / "t.pdf")) == S.NO_TABLES
    assert page.review_label.text() == S.NO_TABLES_LONG
    assert page.table_view.model().rowCount() == 0


# --------------------------------------------------------------------------- python -m pdf2excel


def test_main_module_autoquit_exits_zero(tmp_path: Path) -> None:
    env = dict(os.environ)
    env.update(
        {
            "QT_QPA_PLATFORM": "offscreen",
            "PDF2EXCEL_AUTOQUIT_MS": "500",
            "PDF2EXCEL_UPDATE_URL": "file:///dev/null",  # esquema no permitido: el chequeo termina ya
        }
    )
    result = subprocess.run(
        [sys.executable, "-m", "pdf2excel"],
        env=env,
        cwd=str(tmp_path),
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stderr
