"""Pantalla 2 "Esto encontré": lista de archivos + vista previa tipada + la única opción visible."""

from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QImage, QKeyEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from pdf2excel.errors import UserError
from pdf2excel.models import CellValue, ColumnInfo, ConversionReport, TypedTable
from pdf2excel.ui import strings_es as S
from pdf2excel.ui.file_list import FileListWidget, FileState
from pdf2excel.ui.preview_model import TablePreviewModel

# Índices fijos del selector de carpeta; la carpeta elegida se inserta entre ambos.
_IDX_NEXT_TO_PDF = 0
_FOLDER_ROLE = int(Qt.ItemDataRole.UserRole) + 1
_KIND_NEXT = "next"
_KIND_FOLDER = "folder"
_KIND_CHOOSE = "choose"


def as_text_table(table: TypedTable) -> TypedTable:
    """Copia de ``table`` como texto plano: ``str(valor)`` sin formato, columnas ``text``, sin revisión.

    Es lo que se muestra cuando la casilla "Convertir números y fechas" está desmarcada, sin reanalizar.
    """
    rows: list[list[CellValue]] = [
        [None if value is None else str(value) for value in row] for row in table.rows
    ]
    columns = [ColumnInfo(kind="text", header=col.header) for col in table.columns]
    return TypedTable(
        name=table.name,
        header=list(table.header),
        rows=rows,
        columns=columns,
        review_cells=set(),
        total_rows=set(table.total_rows),
        pages=table.pages,
        source=table.source,
        quality=table.quality,
    )


def _pages_text(table: TypedTable) -> str:
    first, last = table.pages
    if last > first:
        return S.TABLE_PAGES_FMT.format(a=first + 1, b=last + 1)
    return S.TABLE_PAGE_FMT.format(p=first + 1)


def report_state(report: ConversionReport) -> FileState:
    """Estado del archivo en la lista: ``warn`` si hubo OCR o avisos, si no ``ok``."""
    return "warn" if report.ocr_applied or report.warnings else "ok"


class ReviewPage(QWidget):
    """Lista a la izquierda, vista previa a la derecha, casilla + selector de carpeta + botones abajo."""

    save_requested = Signal()
    back_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("ReviewPage")
        self.setAccessibleName(S.REVIEW_TITLE)

        self._reports: dict[str, ConversionReport] = {}
        self._errors: dict[str, UserError] = {}
        self._current: str | None = None
        self._output_dir: Path | None = None
        self._last_folder_index = _IDX_NEXT_TO_PDF

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(10)

        self.title_label = QLabel(S.REVIEW_TITLE, self)
        self.title_label.setObjectName("PageTitle")
        layout.addWidget(self.title_label)
        self.subtitle_label = QLabel(S.REVIEW_SUBTITLE, self)
        self.subtitle_label.setObjectName("Muted")
        self.subtitle_label.setWordWrap(True)
        layout.addWidget(self.subtitle_label)

        # ---- cuerpo: splitter lista | vista previa
        self.splitter = QSplitter(Qt.Orientation.Horizontal, self)
        self.splitter.setChildrenCollapsible(False)

        self.file_list = FileListWidget(self.splitter)
        self.file_list.current_file_changed.connect(self._on_current_file_changed)
        self.splitter.addWidget(self.file_list)

        right = QWidget(self.splitter)
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        table_row = QHBoxLayout()
        table_row.setSpacing(8)
        self.table_combo = QComboBox(right)
        self.table_combo.setAccessibleName(S.TABLE_COMBO_FMT.format(i=1, n=1))
        self.table_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.table_combo.currentIndexChanged.connect(self._on_table_changed)
        table_row.addWidget(self.table_combo)
        self.sheet_label = QLabel(right)
        self.sheet_label.setObjectName("Muted")
        self.sheet_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        table_row.addWidget(self.sheet_label, 1)
        right_layout.addLayout(table_row)

        self.preview_model = TablePreviewModel(self)
        self.table_view = QTableView(right)
        self.table_view.setObjectName("PreviewTable")
        self.table_view.setAccessibleName(S.REVIEW_TITLE)
        self.table_view.setModel(self.preview_model)
        self.table_view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_view.setWordWrap(False)
        self.table_view.setCornerButtonEnabled(False)
        h_header = self.table_view.horizontalHeader()
        h_header.setHighlightSections(False)
        h_header.setStretchLastSection(True)
        h_header.setSectionsMovable(False)
        h_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table_view.verticalHeader().setDefaultSectionSize(22)
        self.table_view.verticalHeader().setVisible(False)
        right_layout.addWidget(self.table_view, 1)

        self.review_label = QLabel(S.PREVIEW_EMPTY, right)
        self.review_label.setObjectName("Muted")
        self.review_label.setWordWrap(True)
        right_layout.addWidget(self.review_label)

        self.scan_note = QLabel(S.SCANNED_NOTE, right)
        self.scan_note.setObjectName("WarnNote")
        self.scan_note.setWordWrap(True)
        self.scan_note.hide()
        right_layout.addWidget(self.scan_note)

        self.splitter.addWidget(right)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 2)
        layout.addWidget(self.splitter, 1)

        # ---- pie: opción + selector de carpeta + botones
        bottom = QHBoxLayout()
        bottom.setSpacing(12)
        self.typed_check = QCheckBox(S.OPT_TYPED, self)
        self.typed_check.setChecked(True)
        self.typed_check.setAccessibleName(S.OPT_TYPED)
        self.typed_check.toggled.connect(self._on_typed_toggled)
        bottom.addWidget(self.typed_check)
        bottom.addStretch(1)

        save_label = QLabel(S.SAVE_LABEL, self)
        bottom.addWidget(save_label)
        self.folder_combo = QComboBox(self)
        self.folder_combo.setAccessibleName(S.SAVE_LABEL)
        self.folder_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.folder_combo.addItem(S.SAVE_NEXT_TO_PDF, _KIND_NEXT)
        self.folder_combo.addItem(S.SAVE_CHOOSE, _KIND_CHOOSE)
        self.folder_combo.activated.connect(self._on_folder_activated)
        save_label.setBuddy(self.folder_combo)
        bottom.addWidget(self.folder_combo)
        layout.addLayout(bottom)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        buttons.addStretch(1)
        self.back_button = QPushButton(S.BTN_BACK, self)
        self.back_button.setAccessibleName(S.BTN_BACK)
        self.back_button.setAutoDefault(False)
        self.back_button.clicked.connect(self.back_requested)
        buttons.addWidget(self.back_button)
        self.save_button = QPushButton(S.BTN_SAVE, self)
        self.save_button.setObjectName("SaveButton")
        self.save_button.setProperty("primary", True)
        self.save_button.setDefault(True)
        self.save_button.setAccessibleName(S.BTN_SAVE)
        self.save_button.setAccessibleDescription(S.REVIEW_SUBTITLE)
        self.save_button.clicked.connect(self._on_save_clicked)
        buttons.addWidget(self.save_button)
        layout.addLayout(buttons)

        # Orden de Tab: lista → tabla → casilla → carpeta → Volver → Guardar.
        QWidget.setTabOrder(self.file_list, self.table_combo)
        QWidget.setTabOrder(self.table_combo, self.table_view)
        QWidget.setTabOrder(self.table_view, self.typed_check)
        QWidget.setTabOrder(self.typed_check, self.folder_combo)
        QWidget.setTabOrder(self.folder_combo, self.back_button)
        QWidget.setTabOrder(self.back_button, self.save_button)

        self.set_save_enabled(False)

    # ------------------------------------------------------------------ archivos

    def add_file(self, path: str) -> None:
        """Agrega el archivo a la lista en estado "En cola" (si no estaba)."""
        self.file_list.add_file(path)

    def set_file_status(self, path: str, text: str, state: FileState) -> None:
        self.file_list.set_status(path, text, state)

    def set_thumbnail(self, path: str, image: QImage) -> None:
        self.file_list.set_thumbnail(path, image)

    def add_report(self, path: str, report: ConversionReport) -> None:
        """Guarda el reporte, actualiza el estado de la fila y, si es el actual (o el primero), lo muestra."""
        self.file_list.add_file(path)
        self._reports[path] = report
        self._errors.pop(path, None)
        self.file_list.set_status(path, S.file_status(report), report_state(report))
        if self._current is None or self._current == path:
            self.file_list.select_file(path)
            self.show_report(path)
        self.set_save_enabled(bool(self._reports))

    def add_error(self, path: str, err: UserError) -> None:
        """Fila en estado error con el título de ``ERRORS[code]``; la vista previa queda vacía."""
        self.file_list.add_file(path)
        self._errors[path] = err
        self._reports.pop(path, None)
        title, _detail, _action = S.error_texts(err, os.path.basename(path))
        self.file_list.set_status(path, title, "error")
        if self._current is None or self._current == path:
            self.file_list.select_file(path)
            self.show_report(path)
        self.set_save_enabled(bool(self._reports))

    def remove_file(self, path: str) -> None:
        self.file_list.remove_file(path)
        self._reports.pop(path, None)
        self._errors.pop(path, None)
        if self._current == path:
            self._current = None
            self._show_empty(S.PREVIEW_EMPTY)
        self.set_save_enabled(bool(self._reports))

    def clear(self) -> None:
        """Vuelve al estado inicial (sin archivos); conserva la casilla y la carpeta elegida."""
        self._reports.clear()
        self._errors.clear()
        self._current = None
        self.file_list.clear()
        self._show_empty(S.PREVIEW_EMPTY)
        self.set_save_enabled(False)

    def reports(self) -> dict[str, ConversionReport]:
        return dict(self._reports)

    def paths(self) -> list[str]:
        return self.file_list.paths()

    def current_path(self) -> str | None:
        return self._current

    # ------------------------------------------------------------------ vista previa

    def show_report(self, path: str) -> None:
        """Muestra la vista previa del archivo ``path`` (o su error, si falló)."""
        self._current = path
        report = self._reports.get(path)
        if report is None:
            err = self._errors.get(path)
            if err is None:
                self._show_empty(S.PREVIEW_EMPTY)
            else:
                title, detail, _action = S.error_texts(err, os.path.basename(path))
                self._show_empty(f"⚠ {title}. {detail}".strip())
            return

        self.table_combo.blockSignals(True)
        self.table_combo.clear()
        n = len(report.tables)
        for i in range(n):
            self.table_combo.addItem(S.TABLE_COMBO_FMT.format(i=i + 1, n=n), i)
        self.table_combo.setEnabled(n > 1)
        self.table_combo.blockSignals(False)
        self.scan_note.setVisible(report.ocr_applied and n > 0)
        if n == 0:
            self.table_combo.setVisible(False)
            self.sheet_label.setText(S.NO_TABLES)
            self.preview_model.set_table(None)
            self._set_review_text(S.NO_TABLES_LONG, warn=False)
            return
        self.table_combo.setVisible(True)
        self.table_combo.setCurrentIndex(0)
        self._render_table(0)

    def _show_empty(self, message: str) -> None:
        self.table_combo.blockSignals(True)
        self.table_combo.clear()
        self.table_combo.blockSignals(False)
        self.table_combo.setVisible(False)
        self.sheet_label.setText("")
        self.scan_note.hide()
        self.preview_model.set_table(None)
        self._set_review_text(message, warn=False)

    def _render_table(self, index: int) -> None:
        report = self._reports.get(self._current or "")
        if report is None or not (0 <= index < len(report.tables)):
            return
        table = report.tables[index]
        shown = table if self.typed_values() else as_text_table(table)
        self.preview_model.set_table(shown)
        self.table_view.resizeColumnsToContents()
        name = table.name or f"Tabla {index + 1}"
        self.sheet_label.setText(f"{name} · {_pages_text(table)}")

        n_review = len(table.review_cells) if self.typed_values() else 0
        parts = [S.review_cells_text(n_review)]
        if self.preview_model.is_truncated():
            parts.append(
                S.PREVIEW_TRUNCATED_FMT.format(n=self.preview_model.rowCount(), total=len(table.rows))
            )
        self._set_review_text(" · ".join(parts), warn=n_review > 0)

    def _set_review_text(self, text: str, warn: bool) -> None:
        self.review_label.setText(text)
        self.review_label.setAccessibleName(text)
        name = "WarnNote" if warn else "Muted"
        if self.review_label.objectName() != name:
            self.review_label.setObjectName(name)
            style = self.review_label.style()
            style.unpolish(self.review_label)
            style.polish(self.review_label)

    def current_table_index(self) -> int:
        return max(0, self.table_combo.currentIndex())

    # ------------------------------------------------------------------ opción y carpeta

    def typed_values(self) -> bool:
        return self.typed_check.isChecked()

    def set_typed_values(self, typed: bool) -> None:
        self.typed_check.setChecked(typed)

    def output_dir(self) -> Path | None:
        """Carpeta elegida o ``None`` = junto a cada PDF."""
        data = self.folder_combo.currentData()
        return self._output_dir if data == _KIND_FOLDER else None

    def set_output_dir(self, folder: Path | None) -> None:
        """Preselecciona la carpeta (p. ej. la guardada en las opciones)."""
        if folder is None:
            self._output_dir = None
            self._remove_folder_item()
            self.folder_combo.setCurrentIndex(_IDX_NEXT_TO_PDF)
        else:
            self._apply_folder(Path(folder))
        self._last_folder_index = self.folder_combo.currentIndex()

    def set_save_enabled(self, enabled: bool) -> None:
        self.save_button.setEnabled(enabled)

    def save_enabled(self) -> bool:
        return self.save_button.isEnabled()

    def _folder_item_index(self) -> int:
        return self.folder_combo.findData(_KIND_FOLDER)

    def _remove_folder_item(self) -> None:
        idx = self._folder_item_index()
        if idx >= 0:
            self.folder_combo.removeItem(idx)

    def _apply_folder(self, folder: Path) -> None:
        self._output_dir = folder
        text = S.SAVE_IN_FOLDER_FMT.format(folder=folder.name or str(folder))
        idx = self._folder_item_index()
        if idx < 0:
            idx = _IDX_NEXT_TO_PDF + 1
            self.folder_combo.insertItem(idx, text, _KIND_FOLDER)
        else:
            self.folder_combo.setItemText(idx, text)
        self.folder_combo.setItemData(idx, str(folder), Qt.ItemDataRole.ToolTipRole)
        self.folder_combo.setCurrentIndex(idx)

    def _on_folder_activated(self, index: int) -> None:
        kind = self.folder_combo.itemData(index)
        if kind == _KIND_CHOOSE:
            start = str(self._output_dir) if self._output_dir else os.path.expanduser("~")
            chosen = QFileDialog.getExistingDirectory(self, S.FOLDER_DIALOG_TITLE, start)
            if chosen:
                self._apply_folder(Path(chosen))
            else:
                self.folder_combo.setCurrentIndex(self._last_folder_index)
        self._last_folder_index = self.folder_combo.currentIndex()

    def _on_typed_toggled(self, _checked: bool) -> None:
        if self._current is not None and self._current in self._reports:
            self._render_table(self.current_table_index())

    def _on_table_changed(self, index: int) -> None:
        if index >= 0:
            self._render_table(index)

    def _on_current_file_changed(self, path: str) -> None:
        if path != self._current:
            self.show_report(path)

    def _on_save_clicked(self) -> None:
        if self.save_button.isEnabled():
            self.save_requested.emit()

    # ------------------------------------------------------------------ teclado

    def primary_action(self) -> None:
        """Enter = Guardar Excel (si está habilitado)."""
        self._on_save_clicked()

    def escape_action(self) -> None:
        """Esc = Volver."""
        self.back_requested.emit()

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape:
            event.accept()
            self.escape_action()
            return
        super().keyPressEvent(event)


__all__ = ["ReviewPage", "as_text_table", "report_state"]
