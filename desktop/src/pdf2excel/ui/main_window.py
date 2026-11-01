"""Ventana principal: ``QStackedWidget`` con las 4 pantallas; solo coordina señales (DESIGN.md §7).

Nunca llama al core directamente: todo pasa por ``JobQueue`` (``jobs/worker.py``). La única excepción
tolerada es la miniatura de la primera página (``pymupdf`` → ``get_pixmap(dpi=40)``), que es rápida y se
hace en el hilo de la UI cuando llega el análisis.
"""

from __future__ import annotations

import logging
import os
import sys
from collections import deque
from dataclasses import replace
from pathlib import Path
from typing import Literal

from PySide6.QtCore import Qt, QTimer, QUrl
from PySide6.QtGui import QCloseEvent, QDesktopServices, QImage, QKeyEvent, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractButton,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from pdf2excel.errors import UserError
from pdf2excel.jobs.worker import JobQueue
from pdf2excel.models import ConversionReport, EffectiveMode, Options
from pdf2excel.ui import strings_es as S
from pdf2excel.ui.pages.done_page import MAX_OPEN_EXCEL, DonePage
from pdf2excel.ui.pages.drop_page import DropPage
from pdf2excel.ui.pages.progress_page import ProgressPage
from pdf2excel.ui.pages.review_page import ReviewPage
from pdf2excel.ui.settings_dialog import SettingsDialog, load_options, save_options

log = logging.getLogger(__name__)

WINDOW_SIZE = (860, 560)
MIN_WINDOW_SIZE = (640, 440)
LARGE_FILE_BYTES = 2 * 1024 * 1024  # > 2 MB: probablemente escaneado → pantalla de progreso
THUMBNAIL_DPI = 40
CLOSE_WAIT_MS = 3000
CLOSE_POLL_MS = 200  # con un job que no honró la cancelación: la ventana se oculta y se espera de a 200 ms
SLOW_ANALYSIS_MS = 3000  # análisis en la pantalla 1 que supera esto → pantalla de progreso (con Cancelar)

# Texto que todavía no está en strings_es (pedido en deviations): confirmación al cerrar con trabajo en curso.
CLOSE_CONFIRM = S.CLOSE_CONFIRM  # alias (el texto vive en strings_es)
EXCEL_FILTER = "Excel (*.xlsx)"

Phase = Literal["idle", "analyze", "reanalyze", "export"]
Page = DropPage | ReviewPage | ProgressPage | DonePage

# Opciones que cambian el resultado del análisis: si cambian entre analizar y guardar, se reanaliza.
_ANALYSIS_FIELDS = (
    "typed_values",
    "merge_continuations",
    "include_text",
    "one_sheet_per_page",
    "force_ocr",
    "ocr_lang",
    "layout",
    "layout_images",
)


def analysis_key(opts: Options) -> tuple[object, ...]:
    return tuple(getattr(opts, name) for name in _ANALYSIS_FIELDS)


def is_large_file(path: str, threshold: int = LARGE_FILE_BYTES) -> bool:
    """Heurística "grande o escaneado": tamaño > 2 MB (los escaneos pesan; los vectoriales no)."""
    try:
        return os.path.getsize(path) > threshold
    except OSError:
        return False


def thumbnail_for(path: str, password: str | None = None, dpi: int = THUMBNAIL_DPI) -> QImage | None:
    """Miniatura de la primera página o ``None`` si no se pudo (nunca lanza).

    Lee ``needs_pass`` ANTES de autenticar (en pymupdf 2.0 leerlo después vacía la extracción de texto
    de ese documento; acá no extraemos texto, pero no cuesta nada hacerlo bien).
    """
    if not os.path.isfile(path):
        return None
    # El worker ya importó pymupdf antes de emitir `analyzed`; si no está cargado (stubs en tests) no se
    # importa acá: hacerlo en el hilo de la UI congelaría la ventana ~1 s.
    pymupdf = sys.modules.get("pymupdf")
    if pymupdf is None:
        return None
    try:
        doc = pymupdf.open(path, filetype="pdf")
        try:
            if doc.needs_pass and not (password and doc.authenticate(password) > 0):
                return None
            if doc.page_count == 0:
                return None
            data = doc[0].get_pixmap(dpi=dpi).tobytes("png")
        finally:
            doc.close()
    except Exception as exc:
        log.debug("Sin miniatura para %s: %r", path, exc)
        return None
    image = QImage.fromData(data)  # PNG autodetectado por la cabecera
    return None if image.isNull() else image


class PasswordDialog(QDialog):
    """ "Este PDF está protegido con contraseña": campo de contraseña + Continuar / Saltar."""

    def __init__(self, path: str, parent: QWidget | None = None, wrong: bool = False) -> None:
        super().__init__(parent)
        self.path = path
        name = os.path.basename(path)
        title, detail, _action = S.error_texts(UserError("encrypted"), name)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(380)
        self.setAccessibleName(title)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        self.title_label = QLabel(f"<b>{title}</b>", self)
        self.title_label.setTextFormat(Qt.TextFormat.RichText)
        self.title_label.setWordWrap(True)
        layout.addWidget(self.title_label)
        self.prompt_label = QLabel(S.PASSWORD_PROMPT_FMT.format(name=name), self)
        self.prompt_label.setWordWrap(True)
        layout.addWidget(self.prompt_label)

        self.password_edit = QLineEdit(self)
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_edit.setPlaceholderText(S.PASSWORD_PLACEHOLDER)
        self.password_edit.setAccessibleName(S.PASSWORD_PROMPT_FMT.format(name=name))
        self.password_edit.setClearButtonEnabled(True)
        self.prompt_label.setBuddy(self.password_edit)
        layout.addWidget(self.password_edit)

        self.hint_label = QLabel(S.PASSWORD_WRONG if wrong else detail, self)
        self.hint_label.setObjectName("WarnNote" if wrong else "Muted")
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        self.buttons = QDialogButtonBox(self)
        self.continue_button = self.buttons.addButton(S.BTN_CONTINUE, QDialogButtonBox.ButtonRole.AcceptRole)
        self.skip_button = self.buttons.addButton(S.BTN_SKIP, QDialogButtonBox.ButtonRole.RejectRole)
        self.continue_button.setProperty("primary", True)
        style = self.continue_button.style()
        style.unpolish(self.continue_button)  # la propiedad se seteó después de crear el botón
        style.polish(self.continue_button)
        self.continue_button.setDefault(True)
        self.continue_button.setAccessibleName(S.BTN_CONTINUE)
        self.skip_button.setAccessibleName(S.BTN_SKIP)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.password_edit.setFocus()

    def password(self) -> str:
        return self.password_edit.text()


class MainWindow(QMainWindow):
    """Arrastrá → Esto encontré → Convirtiendo → Listo. Estado del lote: ``reports``/``errors``/``order``."""

    def __init__(self, initial_files: list[str] | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(S.APP_TITLE)
        self.resize(*WINDOW_SIZE)
        self.setMinimumSize(*MIN_WINDOW_SIZE)
        self.setAccessibleName(S.APP_TITLE)

        self.opts: Options = load_options()
        self.queue = JobQueue(2, self)
        self.reports: dict[str, ConversionReport] = {}
        self.errors: dict[str, UserError] = {}
        self.order: list[str] = []
        self.exported: dict[str, ConversionReport] = {}
        self.password_dialog: PasswordDialog | None = None
        self._passwords: dict[str, str] = {}
        self._password_queue: deque[str] = deque()
        self._phase: Phase = "idle"
        self._cancelled = False
        self._analysis_key: tuple[object, ...] | None = None
        # Formato del Excel que el usuario eligió a mano por archivo (sobrevive a un reanálisis).
        self._mode_overrides: dict[str, EffectiveMode] = {}
        self._batch: list[str] = []  # archivos del lote en curso (para "Archivo i de n")
        self._closing = False  # closeEvent aceptado o diferido (esperando un job que no honró cancel)

        # Análisis "corto" en la pantalla 1 que se alarga (escaneo chico, OCR) → pasar a progreso.
        self._slow_timer = QTimer(self)
        self._slow_timer.setSingleShot(True)
        self._slow_timer.setInterval(SLOW_ANALYSIS_MS)
        self._slow_timer.timeout.connect(self._on_slow_analysis)
        # Cierre diferido: la ventana ya se ocultó y falta que termine el último job.
        self._close_timer = QTimer(self)
        self._close_timer.setInterval(CLOSE_POLL_MS)
        self._close_timer.timeout.connect(self._poll_deferred_close)

        # ---- widgets
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        top = QHBoxLayout()
        top.setContentsMargins(8, 6, 8, 0)
        top.addStretch(1)
        self.settings_button = QToolButton(central)
        self.settings_button.setObjectName("SettingsButton")
        self.settings_button.setText("⚙")
        self.settings_button.setAutoRaise(True)
        self.settings_button.setToolTip(S.BTN_SETTINGS)
        self.settings_button.setAccessibleName(S.BTN_SETTINGS)
        self.settings_button.setCursor(Qt.CursorShape.PointingHandCursor)
        font = self.settings_button.font()
        font.setPointSize(max(font.pointSize(), 10) + 4)
        self.settings_button.setFont(font)
        self.settings_button.clicked.connect(self.open_settings)
        top.addWidget(self.settings_button)
        root.addLayout(top)

        self.stack = QStackedWidget(central)
        self.drop_page = DropPage(self.stack)
        self.review_page = ReviewPage(self.stack)
        self.progress_page = ProgressPage(self.stack)
        self.done_page = DonePage(self.stack)
        for page in (self.drop_page, self.review_page, self.progress_page, self.done_page):
            self.stack.addWidget(page)
        root.addWidget(self.stack, 1)
        self.setCentralWidget(central)

        self.review_page.set_typed_values(self.opts.typed_values)
        self.review_page.set_output_dir(self.opts.output_dir)

        # ---- señales de las pantallas
        self.drop_page.files_chosen.connect(self.start_analysis)
        self.review_page.save_requested.connect(self._on_save_requested)
        self.review_page.mode_changed.connect(self._on_mode_changed)
        self.review_page.back_requested.connect(self._on_back_requested)
        self.progress_page.cancel_requested.connect(self._on_cancel_requested)
        self.done_page.convert_more.connect(self._on_convert_more)
        self.done_page.open_excel.connect(self._on_open_excel)
        self.done_page.open_folder.connect(self._on_open_folder)

        # ---- señales de la cola (llegan en el hilo principal)
        self.queue.signals.progress.connect(self._on_progress)
        self.queue.signals.analyzed.connect(self._on_analyzed)
        self.queue.signals.exported.connect(self._on_exported)
        self.queue.signals.failed.connect(self._on_failed)
        self.queue.signals.all_done.connect(self._on_all_done)

        # ---- atajos
        self.open_shortcut = QShortcut(QKeySequence.StandardKey.Open, self)
        self.open_shortcut.activated.connect(self.choose_files)

        self.stack.setCurrentWidget(self.drop_page)
        self.drop_page.focus_drop_zone()

        if initial_files:
            files = [os.path.abspath(f) for f in initial_files if f.lower().endswith(".pdf")]
            if files:
                QTimer.singleShot(0, lambda: self.start_analysis(files))

    # ------------------------------------------------------------------ helpers de estado

    @property
    def phase(self) -> Phase:
        return self._phase

    def current_page(self) -> Page:
        widget = self.stack.currentWidget()
        if isinstance(widget, DropPage | ReviewPage | ProgressPage | DonePage):
            return widget
        return self.drop_page

    def _show(self, page: QWidget) -> None:
        if self.stack.currentWidget() is not page:
            self.stack.setCurrentWidget(page)
        page.setFocus(Qt.FocusReason.OtherFocusReason)

    def _show_progress(self, mode: Literal["analyze", "export"]) -> None:
        self._slow_timer.stop()
        self.progress_page.reset()
        self.progress_page.set_mode(mode)
        self.drop_page.set_busy(False)
        self._show(self.progress_page)

    def _on_slow_analysis(self) -> None:
        """El análisis en la pantalla 1 ya lleva ``SLOW_ANALYSIS_MS``: mostrar progreso y Cancelar."""
        if self._phase == "analyze" and self.current_page() is self.drop_page and self.drop_page.is_busy():
            self._show_progress("analyze")

    def _belongs_to_batch(self, path: str) -> bool:
        """¿La señal es del lote vigente? Tras "Volver"/cancelar un job puede terminar tarde e informar."""
        return self._phase != "idle" and path in self.order

    def _show_review(self) -> None:
        self.drop_page.set_busy(False)
        # Guardar recién cuando terminó todo el lote (lo habilita _finish_analysis); mientras tanto la lista
        # muestra qué falta.
        self.review_page.set_save_enabled(bool(self.reports) and not self.queue.is_busy())
        if self.stack.currentWidget() is not self.review_page:
            self._show(self.review_page)
            self.review_page.save_button.setFocus(Qt.FocusReason.OtherFocusReason)

    def _reset_to_drop(self) -> None:
        self.reports.clear()
        self.errors.clear()
        self.order.clear()
        self.exported.clear()
        self._passwords.clear()
        self._password_queue.clear()
        self._mode_overrides.clear()
        self._batch = []
        self._phase = "idle"
        self._cancelled = False
        self._analysis_key = None
        self._slow_timer.stop()
        self.review_page.clear()
        self.progress_page.reset()
        self.drop_page.set_busy(False)
        self._show(self.drop_page)
        self.drop_page.focus_drop_zone()

    def _file_position(self, path: str) -> tuple[int, int]:
        batch = self._batch or self.order
        index = batch.index(path) + 1 if path in batch else 1
        return index, max(1, len(batch))

    def _passwords_for(self, paths: list[str]) -> dict[str, str]:
        return {p: self._passwords[p] for p in paths if p in self._passwords}

    # ------------------------------------------------------------------ flujo: análisis

    def choose_files(self) -> None:
        """Ctrl+O: elegir archivos (en la pantalla 1, o para agregar más en "Esto encontré")."""
        if self._phase in ("idle", "analyze") and self.current_page() in (self.drop_page, self.review_page):
            self.drop_page.drop_zone.choose_files()

    def start_analysis(self, paths: list[str]) -> None:
        """Agrega los archivos a la lista y los manda a analizar en segundo plano."""
        if self._phase in ("export", "reanalyze"):
            return
        new = [p for p in paths if p not in self.order]
        if not new:
            return
        if self._phase == "idle":
            self._analysis_key = analysis_key(self.opts)
            self._cancelled = False
        self._phase = "analyze"
        for path in new:
            self.order.append(path)
            self.review_page.add_file(path)
        self._batch = list(self.order)

        if self.current_page() is self.drop_page:
            if any(is_large_file(p) for p in new):
                self._show_progress("analyze")
            else:
                self.drop_page.set_busy(True)
                self._slow_timer.start()
        elif self.current_page() is self.review_page:
            # Se agregaron archivos desde "Esto encontré": Guardar recién cuando terminen de analizarse.
            self.review_page.set_save_enabled(False)
        self.queue.submit_analyze(new, self.opts, self._passwords_for(new))

    def _on_progress(self, path: str, current: int, total: int, stage: str) -> None:
        if not self._belongs_to_batch(path):
            return
        if stage == "excel":
            if path in self.reports and current < total:
                self.review_page.set_file_status(path, S.FILE_STATUS_EXPORTING, "working")
        elif path not in self.reports or self._phase == "reanalyze":
            text = S.FILE_STATUS_OCR if stage == "ocr" else S.FILE_STATUS_WORKING
            self.review_page.set_file_status(path, text, "working")
        if stage == "ocr" and self._phase == "analyze" and self.current_page() is self.drop_page:
            # OCR = segundos por página: barra por página y botón Cancelar (DESIGN §7, pantalla 3).
            self._show_progress("analyze")
        index, count = self._file_position(path)
        self.progress_page.set_progress(path, current, total, stage, index, count)

    def _on_analyzed(self, path: str, report: ConversionReport) -> None:
        if not self._belongs_to_batch(path):
            log.debug("Resultado tardío descartado para %s (lote ya cerrado)", path)
            return
        override = self._mode_overrides.get(path)
        if override is not None and self._phase == "reanalyze":
            report.mode = override if (override == "tables" or report.layout_pages) else "tables"
        self.reports[path] = report
        self.errors.pop(path, None)
        self.review_page.add_report(path, report)
        image = thumbnail_for(path, self._passwords.get(path))
        if image is not None:
            self.review_page.set_thumbnail(path, image)
        if self._phase == "analyze":
            self._show_review()
        self._maybe_finish()

    def _on_failed(self, path: str, err: UserError) -> None:
        if err.code == "cancelled" or not self._belongs_to_batch(path):
            self._maybe_finish()
            return
        if self._phase in ("analyze", "reanalyze"):
            if err.code == "encrypted" and not self._cancelled:
                self._password_queue.append(path)
                self._next_password_dialog()
                return
            self.errors[path] = err
            self.review_page.add_error(path, err)
            if self._phase == "analyze":
                self._show_review()
        elif self._phase == "export":
            report = self.reports.get(path)
            if report is not None and err.code == "file_in_use" and not self._cancelled:
                self._handle_file_in_use(path, report, err)
            elif report is not None and err.code == "no_permission" and not self._cancelled:
                self._handle_no_permission(path, report, err)
            else:
                self.errors[path] = err
        else:
            self.errors[path] = err
        self._maybe_finish()

    # ------------------------------------------------------------------ contraseñas

    def _next_password_dialog(self) -> None:
        if self.password_dialog is not None or not self._password_queue:
            return
        path = self._password_queue.popleft()
        wrong = path in self._passwords
        dialog = PasswordDialog(path, self, wrong=wrong)
        dialog.finished.connect(lambda result, d=dialog: self._on_password_finished(d, int(result)))
        self.password_dialog = dialog
        dialog.open()  # no bloquea el event loop: las señales de la cola siguen llegando

    def _on_password_finished(self, dialog: PasswordDialog, result: int) -> None:
        path = dialog.path
        if self.password_dialog is dialog:
            self.password_dialog = None
        dialog.deleteLater()
        if result == int(QDialog.DialogCode.Accepted) and dialog.password():
            self._passwords[path] = dialog.password()
            if self._phase == "idle":
                self._phase = "analyze"
            self.review_page.set_file_status(path, S.FILE_STATUS_WORKING, "working")
            self.queue.submit_analyze([path], self.opts, {path: dialog.password()})
        else:
            self._passwords.pop(path, None)
            err = UserError("encrypted", action="password")
            self.errors[path] = err
            self.review_page.add_error(path, err)
            if self._phase == "analyze":
                self._show_review()
        self._next_password_dialog()
        self._maybe_finish()

    # ------------------------------------------------------------------ flujo: guardar

    def _on_save_requested(self) -> None:
        if not self.reports or self._phase not in ("idle", "analyze"):
            return
        if self.queue.is_busy():
            return  # todavía hay archivos analizándose: la lista lo muestra
        self.opts.typed_values = self.review_page.typed_values()
        self.opts.output_dir = self.review_page.output_dir()
        save_options(self.opts)
        self._cancelled = False
        if analysis_key(self.opts) != self._analysis_key:
            self._reanalyze_then_export()
        else:
            self._start_export()

    def _on_mode_changed(self, path: str, mode: str) -> None:
        if mode in ("layout", "tables"):
            self._mode_overrides[path] = mode  # type: ignore[assignment]

    @staticmethod
    def _sanitize_modes(reports: list[ConversionReport]) -> None:
        """ "Igual al PDF" solo si el reporte trae el diseño de las páginas; si no, cae a tablas."""
        for report in reports:
            if report.mode == "layout" and not report.layout_pages:
                report.mode = "tables"

    def _apply_layout_preference(self) -> None:
        """Cambió "Formato del Excel" en Opciones avanzadas: manda sobre las elecciones previas."""
        self._mode_overrides.clear()
        for report in self.reports.values():
            if self.opts.layout == "auto":
                report.mode = report.suggested_mode
            else:
                report.mode = self.opts.layout
        self._sanitize_modes(list(self.reports.values()))
        current = self.review_page.current_path()
        if current in self.reports:
            self.review_page.show_report(current)

    def _reanalyze_then_export(self) -> None:
        """Cambió una opción que afecta el análisis (p. ej. el tipado): se reanaliza y después se guarda."""
        paths = [p for p in self.order if p in self.reports]
        self._phase = "reanalyze"
        self._batch = paths
        self._show_progress("analyze")
        self.queue.submit_analyze(paths, self.opts, self._passwords_for(paths))

    def _start_export(self) -> None:
        reports = [self.reports[p] for p in self.order if p in self.reports]
        self._sanitize_modes(reports)
        self._phase = "export"
        self.exported = {}
        self._batch = [str(r.source) for r in reports]
        if not reports:
            self._finish_export()
            return
        self._show_progress("export")
        if self.opts.single_workbook and len(reports) > 1:
            self.queue.submit_batch_export(reports, self.opts)
        else:
            self.queue.submit_export(reports, self.opts)

    def _on_exported(self, path: str, report: ConversionReport) -> None:
        self.exported[path] = report
        self.errors.pop(path, None)
        if path in self.reports:
            self.review_page.set_file_status(path, S.file_status(report), "ok")
        self._maybe_finish()

    @staticmethod
    def _resolve_clicked(
        box: QMessageBox, result: object, added: list[QAbstractButton]
    ) -> QAbstractButton | None:
        """Botón elegido: ``clickedButton()``; si no hay (``exec`` reemplazado en tests), el índice devuelto.

        Con botones personalizados ``QMessageBox.exec()`` devuelve el índice en el orden en que se agregaron.
        """
        clicked = box.clickedButton()
        if clicked is None and isinstance(result, int) and 0 <= result < len(added):
            clicked = added[result]
        return clicked

    def _handle_file_in_use(self, path: str, report: ConversionReport, err: UserError) -> None:
        xlsx_name = report.output_path.name if report.output_path else f"{Path(path).stem}.xlsx"
        title, detail, _action = S.error_texts(err, xlsx_name)
        box = QMessageBox(QMessageBox.Icon.Warning, title, detail, QMessageBox.StandardButton.NoButton, self)
        box.setInformativeText(err.detail if err.detail and err.detail not in detail else "")
        retry = box.addButton(S.BTN_RETRY, QMessageBox.ButtonRole.AcceptRole)
        save_as = box.addButton(S.BTN_SAVE_AS, QMessageBox.ButtonRole.ActionRole)
        skip = box.addButton(S.BTN_SKIP, QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(retry)
        box.setEscapeButton(skip)
        clicked = self._resolve_clicked(box, box.exec(), [retry, save_as, skip])
        if clicked is retry:
            self.queue.submit_export([report], self.opts)
        elif clicked is save_as:
            folder = self.opts.output_dir or Path(path).parent
            suggested = str(folder / xlsx_name)
            chosen, _filter = QFileDialog.getSaveFileName(self, S.BTN_SAVE_AS, suggested, EXCEL_FILTER)
            if chosen:
                target = Path(chosen)
                if target.suffix.lower() != ".xlsx":
                    target = target.with_suffix(".xlsx")
                report.output_path = target
                # Destino explícito: el nombre que tipeó el usuario, no <pdf>.xlsx en esa carpeta.
                self.queue.submit_export([report], replace(self.opts, output_dir=target.parent), dst=target)
            else:
                self.errors[path] = err
        else:
            self.errors[path] = err

    def _handle_no_permission(self, path: str, report: ConversionReport, err: UserError) -> None:
        title, detail, _action = S.error_texts(err, os.path.basename(path))
        box = QMessageBox(QMessageBox.Icon.Warning, title, detail, QMessageBox.StandardButton.NoButton, self)
        choose = box.addButton(S.BTN_CHOOSE_FOLDER, QMessageBox.ButtonRole.AcceptRole)
        skip = box.addButton(S.BTN_SKIP, QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(choose)
        box.setEscapeButton(skip)
        clicked = self._resolve_clicked(box, box.exec(), [choose, skip])
        if clicked is choose:
            start = str(self.opts.output_dir or Path(path).parent)
            chosen = QFileDialog.getExistingDirectory(self, S.FOLDER_DIALOG_TITLE, start)
            if chosen:
                self.opts.output_dir = Path(chosen)
                self.review_page.set_output_dir(self.opts.output_dir)
                save_options(self.opts)
                self.queue.submit_export([report], self.opts)
                return
        self.errors[path] = err

    # ------------------------------------------------------------------ fin de lote

    def _on_all_done(self) -> None:
        self._maybe_finish()

    def _password_pending(self) -> bool:
        return bool(self._password_queue) or self.password_dialog is not None

    def _results_complete(self) -> bool:
        """True si cada archivo del lote ya tiene resultado (o se canceló).

        ``all_done`` puede procesarse cuando un reintento ya terminó pero su ``analyzed``/``exported`` sigue
        en la cola de eventos: por eso no alcanza con ``is_busy()``.
        """
        if self._cancelled:
            return True
        if self._phase == "export":
            return all(p in self.exported or p in self.errors for p in self._batch)
        paths = self._batch if self._phase == "reanalyze" else self.order
        return all(p in self.reports or p in self.errors for p in paths)

    def _maybe_finish(self) -> None:
        """Cierra la fase en curso cuando no queda nada corriendo ni resultados por procesar."""
        if self._phase == "idle" or self.queue.is_busy() or self._password_pending():
            return
        if not self._results_complete():
            return
        if self._phase == "analyze":
            self._finish_analysis()
        elif self._phase == "reanalyze":
            self._analysis_key = analysis_key(self.opts)
            if self._cancelled:
                self._finish_analysis()
            else:
                self._start_export()
        elif self._phase == "export":
            self._finish_export()

    def _finish_analysis(self) -> None:
        self._phase = "idle"
        self.drop_page.set_busy(False)
        self.progress_page.reset()
        for path in self.order:
            if path not in self.reports and path not in self.errors:
                err = UserError("cancelled")
                self.errors[path] = err
                self.review_page.set_file_status(path, S.ERRORS["cancelled"][0], "error")
        if self.reports or (self.order and not self._cancelled):
            self._show_review()
            self.review_page.set_save_enabled(bool(self.reports))
        else:
            self._reset_to_drop()

    def _finish_export(self) -> None:
        self._phase = "idle"
        self.progress_page.reset()
        exported = [self.exported[p] for p in self.order if p in self.exported]
        for path in self.order:
            if path in self.reports and path not in self.exported and path not in self.errors:
                self.errors[path] = UserError("cancelled")
        errors = {p: e for p, e in self.errors.items() if p in self.order}  # solo el lote vigente
        self.done_page.set_results(exported, errors, cancelled=self._cancelled)
        self._show(self.done_page)
        outputs = self.done_page.output_paths()
        if self.opts.open_when_done and len(outputs) == 1 and not self._cancelled and not self._closing:
            self._open_paths(outputs)  # un solo libro (un PDF, o el lote en un único libro)

    # ------------------------------------------------------------------ acciones de las pantallas

    def _on_cancel_requested(self) -> None:
        self._cancelled = True
        self.progress_page.set_cancelling()
        self.queue.cancel_all()
        self._maybe_finish()  # si no corría nada, cancel_all no emite all_done

    def _on_back_requested(self) -> None:
        self._cancelled = True
        self.queue.cancel_all()
        self._reset_to_drop()

    def _on_convert_more(self) -> None:
        self._reset_to_drop()

    def _on_open_excel(self) -> None:
        self._open_paths(self.done_page.output_paths()[:MAX_OPEN_EXCEL])  # libros distintos, sin repetir

    def _on_open_folder(self) -> None:
        outputs = self.done_page.output_paths()
        if outputs:
            self._open_paths([outputs[0].parent])

    @staticmethod
    def _open_paths(paths: list[Path]) -> None:
        for path in paths:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def open_settings(self) -> None:
        dialog = SettingsDialog(self.opts, self)
        if dialog.exec() == int(QDialog.DialogCode.Accepted):
            previous_layout = self.opts.layout
            self.opts = dialog.options()
            save_options(self.opts)
            if self.opts.layout != previous_layout:
                self._apply_layout_preference()

    # ------------------------------------------------------------------ teclado y cierre

    def keyPressEvent(self, event: QKeyEvent) -> None:
        page = self.current_page()
        if event.matches(QKeySequence.StandardKey.Open):
            # Respaldo del QShortcut: los atajos de ventana exigen ventana activa (no la hay en offscreen).
            event.accept()
            self.choose_files()
            return
        if event.key() == Qt.Key.Key_Escape:
            event.accept()
            if page is self.drop_page:
                if self.drop_page.is_busy():
                    self._on_cancel_requested()
                return
            page.escape_action()
            return
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            event.accept()
            focus = QApplication.focusWidget()
            if isinstance(focus, QPushButton | QToolButton) and focus.isEnabled() and focus.isVisible():
                focus.click()  # Enter sobre un botón de acción con foco activa ESE botón
                return
            # Casillas y radios (también QAbstractButton) no se alternan con Enter: va la acción primaria.
            page.primary_action()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.queue.is_busy() and not self._closing:
            answer = QMessageBox.question(
                self,
                S.APP_TITLE,
                CLOSE_CONFIRM,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self._cancelled = True
        self._closing = True
        self._slow_timer.stop()
        self.queue.cancel_all()
        if self.queue.wait_for_done(CLOSE_WAIT_MS):
            self._close_timer.stop()
            event.accept()
            return
        # Un job sigue en un tramo sin chequeo de cancelación (find_tables, typing, un libro grande).
        # Cerrar igual dejaría el proceso vivo e invisible hasta que termine y el worker emitiría sobre
        # señales destruidas: se oculta la ventana y se cierra de verdad cuando la cola queda libre.
        log.info("Cerrando: esperando a que termine el último trabajo…")
        self.hide()
        self._close_timer.start()
        event.ignore()

    def _poll_deferred_close(self) -> None:
        if self.queue.is_busy():
            return
        self._close_timer.stop()
        self.close()  # ya no hay trabajo: closeEvent acepta sin preguntar

    def is_closing(self) -> bool:
        """True desde que el usuario confirmó el cierre (aunque falte que termine un job)."""
        return self._closing


__all__ = [
    "CLOSE_CONFIRM",
    "CLOSE_POLL_MS",
    "CLOSE_WAIT_MS",
    "LARGE_FILE_BYTES",
    "SLOW_ANALYSIS_MS",
    "MainWindow",
    "PasswordDialog",
    "analysis_key",
    "is_large_file",
    "thumbnail_for",
]
