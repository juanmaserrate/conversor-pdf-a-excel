"""Pantalla 4 "Listo": resumen ✔/⚠ por archivo, aviso de celdas para revisar y acciones finales."""

from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from pdf2excel.errors import UserError
from pdf2excel.models import ConversionReport
from pdf2excel.ui import strings_es as S

MAX_OPEN_EXCEL = 3  # "Abrir Excel" solo si se guardaron hasta 3 libros .xlsx distintos


def unique_outputs(reports: list[ConversionReport]) -> list[Path]:
    """Libros .xlsx distintos, en orden estable (con libro único del lote, N reportes → 1 ruta)."""
    return list(dict.fromkeys(r.output_path for r in reports if r.output_path is not None))


class DonePage(QWidget):
    """``set_results(reports, errors)`` arma el resumen; los botones emiten señales para la ventana."""

    open_excel = Signal()
    open_folder = Signal()
    convert_more = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DonePage")
        self.setAccessibleName(S.DONE_TITLE_ONE)
        self._reports: list[ConversionReport] = []
        self._errors: dict[str, UserError] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 24, 32, 20)
        layout.setSpacing(12)

        self.title_label = QLabel(S.DONE_TITLE_ONE, self)
        self.title_label.setObjectName("PageTitle")
        self.title_label.setWordWrap(True)
        layout.addWidget(self.title_label)

        self.results_list = QListWidget(self)
        self.results_list.setObjectName("DoneList")
        self.results_list.setAccessibleName(S.DONE_TITLE_ONE)
        self.results_list.setWordWrap(True)
        self.results_list.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.results_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.results_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.results_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.results_list.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        layout.addWidget(self.results_list, 1)

        self.review_note = QLabel("", self)
        self.review_note.setObjectName("WarnNote")
        self.review_note.setWordWrap(True)
        self.review_note.hide()
        layout.addWidget(self.review_note)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        self.open_excel_button = QPushButton(S.BTN_OPEN_EXCEL, self)
        self.open_excel_button.setAccessibleName(S.BTN_OPEN_EXCEL)
        self.open_excel_button.setAutoDefault(False)
        self.open_excel_button.clicked.connect(self.open_excel)
        buttons.addWidget(self.open_excel_button)
        self.open_folder_button = QPushButton(S.BTN_OPEN_FOLDER, self)
        self.open_folder_button.setAccessibleName(S.BTN_OPEN_FOLDER)
        self.open_folder_button.setAutoDefault(False)
        self.open_folder_button.clicked.connect(self.open_folder)
        buttons.addWidget(self.open_folder_button)
        buttons.addStretch(1)
        self.convert_more_button = QPushButton(S.BTN_CONVERT_MORE, self)
        self.convert_more_button.setProperty("primary", True)
        self.convert_more_button.setDefault(True)
        self.convert_more_button.setAccessibleName(S.BTN_CONVERT_MORE)
        self.convert_more_button.clicked.connect(self.convert_more)
        buttons.addWidget(self.convert_more_button)
        layout.addLayout(buttons)

        QWidget.setTabOrder(self.results_list, self.open_excel_button)
        QWidget.setTabOrder(self.open_excel_button, self.open_folder_button)
        QWidget.setTabOrder(self.open_folder_button, self.convert_more_button)
        self.setFocusProxy(self.convert_more_button)

    # ------------------------------------------------------------------ API pública

    def set_results(
        self,
        reports: list[ConversionReport],
        errors: dict[str, UserError],
        cancelled: bool = False,
    ) -> None:
        """``reports`` = exportados (en orden); ``errors`` = ruta → error (análisis o guardado)."""
        self._reports = list(reports)
        self._errors = dict(errors)
        outputs = unique_outputs(self._reports)
        title = S.done_title(len(reports), len(outputs))
        self.title_label.setText(title)
        self.setAccessibleName(title)
        self.results_list.setAccessibleName(title)

        self.results_list.clear()
        if cancelled:
            self._add_line(S.DONE_CANCELLED)
        for report in reports:
            self._add_line(S.done_line(report))
        for path, err in errors.items():
            self._add_line(S.error_line(os.path.basename(path), err))

        n_review = sum(r.n_review for r in reports)
        if n_review > 0:
            # DONE_REVIEW_FMT no contempla el singular; review_cells_text sí ("1 celda marcada…").
            self.review_note.setText(f"⚠ {S.review_cells_text(n_review)}")
            self.review_note.setAccessibleName(self.review_note.text())
            self.review_note.show()
        else:
            self.review_note.hide()

        n_books = len(outputs)
        self.open_excel_button.setVisible(1 <= n_books <= MAX_OPEN_EXCEL)
        self.open_folder_button.setVisible(len(reports) >= 1)

    def lines(self) -> list[str]:
        """Textos de la lista, en orden (para tests y accesibilidad)."""
        return [self.results_list.item(i).text() for i in range(self.results_list.count())]

    def reports(self) -> list[ConversionReport]:
        return list(self._reports)

    def output_paths(self) -> list[Path]:
        """Libros .xlsx guardados, sin duplicados (lo que abre "Abrir Excel")."""
        return unique_outputs(self._reports)

    def errors(self) -> dict[str, UserError]:
        return dict(self._errors)

    def clear(self) -> None:
        self.set_results([], {})

    def primary_action(self) -> None:
        """Enter = Convertir otros."""
        self.convert_more.emit()

    def escape_action(self) -> None:
        """Esc = volver al inicio (Convertir otros)."""
        self.convert_more.emit()

    # ------------------------------------------------------------------ interno

    def _add_line(self, text: str) -> None:
        item = QListWidgetItem(text)
        item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        item.setToolTip(text)
        item.setData(Qt.ItemDataRole.AccessibleTextRole, text)
        self.results_list.addItem(item)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape:
            event.accept()
            self.escape_action()
            return
        super().keyPressEvent(event)


__all__ = ["MAX_OPEN_EXCEL", "DonePage", "unique_outputs"]
