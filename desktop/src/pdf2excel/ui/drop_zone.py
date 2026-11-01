"""Zona de arrastre de la pantalla 1: acepta PDF y carpetas, o abre el diálogo de archivos."""

from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import QSettings, Qt, Signal
from PySide6.QtGui import (
    QDragEnterEvent,
    QDragLeaveEvent,
    QDragMoveEvent,
    QDropEvent,
    QKeyEvent,
    QMouseEvent,
    QPixmap,
)
from PySide6.QtWidgets import QFileDialog, QFrame, QLabel, QSizePolicy, QVBoxLayout, QWidget

from pdf2excel.resources_path import icon_path
from pdf2excel.ui import strings_es as S

SETTINGS_ORG = "ConversorPDF"
SETTINGS_APP = "ConversorPDFaExcel"
LAST_DIR_KEY = "last_dir"
ICON_SIZE = 96


def _is_pdf(path: str) -> bool:
    return path.lower().endswith(".pdf")


class DropZone(QFrame):
    """Marco punteado grande: ícono + "Arrastrá tus PDF acá" + "o hacé clic para elegirlos".

    Emite ``files_dropped(list[str])`` con rutas absolutas de archivos ``.pdf`` (las carpetas se expanden
    recursivamente). Clic, Enter o Espacio abren ``QFileDialog``.
    """

    files_dropped = Signal(list)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAccessibleName(S.DROP_TITLE)
        self.setAccessibleDescription(S.DROP_SUBTITLE)
        self.setProperty("active", False)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMinimumHeight(220)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(8)
        layout.addStretch(1)

        self._icon = QLabel(self)
        self._icon.setObjectName("DropIcon")
        self._icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pix = QPixmap(str(icon_path("app.png")))
        if not pix.isNull():
            self._icon.setPixmap(
                pix.scaled(
                    ICON_SIZE,
                    ICON_SIZE,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        layout.addWidget(self._icon)

        self._title = QLabel(S.DROP_TITLE, self)
        self._title.setObjectName("DropTitle")
        self._title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self._title.font()
        font.setPointSize(max(font.pointSize(), 10) + 8)
        font.setBold(True)
        self._title.setFont(font)
        layout.addWidget(self._title)

        self._subtitle = QLabel(S.DROP_SUBTITLE, self)
        self._subtitle.setObjectName("DropSubtitle")
        self._subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._subtitle)
        layout.addStretch(1)

    # ------------------------------------------------------------------ rutas (puro, testeable)

    @staticmethod
    def expand_paths(paths: list[str]) -> list[str]:
        """Filtra ``.pdf`` (sin importar mayúsculas) y expande carpetas recursivamente, ordenadas.

        Devuelve rutas absolutas sin duplicados, en el orden recibido. Un ``.pdf`` que no existe se deja
        pasar igual (lo reporta después el motor con un error entendible).
        """
        result: list[str] = []
        seen: set[str] = set()

        def add(p: str) -> None:
            absolute = os.path.abspath(p)
            key = os.path.normcase(absolute)
            if key not in seen:
                seen.add(key)
                result.append(absolute)

        for raw in paths:
            if not raw:
                continue
            if os.path.isdir(raw):
                found = [str(f) for f in Path(raw).rglob("*") if f.is_file() and _is_pdf(f.name)]
                for f in sorted(found, key=str.lower):
                    add(f)
            elif _is_pdf(raw):
                add(raw)
        return result

    @staticmethod
    def _urls_to_paths(event: QDragEnterEvent | QDragMoveEvent | QDropEvent) -> list[str]:
        mime = event.mimeData()
        if mime is None or not mime.hasUrls():
            return []
        return [u.toLocalFile() for u in mime.urls() if u.isLocalFile()]

    @staticmethod
    def _accepts(paths: list[str]) -> bool:
        return any(_is_pdf(p) or os.path.isdir(p) for p in paths)

    # ------------------------------------------------------------------ estado visual

    def _set_active(self, active: bool) -> None:
        if self.property("active") == active:
            return
        self.setProperty("active", active)
        style = self.style()
        style.unpolish(self)
        style.polish(self)
        self.update()

    # ------------------------------------------------------------------ drag & drop

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if self._accepts(self._urls_to_paths(event)):
            event.acceptProposedAction()
            self._set_active(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        if self.property("active"):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:
        self._set_active(False)
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:
        self._set_active(False)
        files = self.expand_paths(self._urls_to_paths(event))
        if files:
            event.acceptProposedAction()
            self.files_dropped.emit(files)
        else:
            event.ignore()

    # ------------------------------------------------------------------ clic / teclado

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            event.accept()
            self.choose_files()
        else:
            super().mousePressEvent(event)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            event.accept()
            self.choose_files()
        else:
            super().keyPressEvent(event)

    def choose_files(self) -> None:
        """Abre el diálogo nativo (filtro *.pdf) y emite ``files_dropped`` con lo elegido."""
        settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        last_dir = str(settings.value(LAST_DIR_KEY, "", type=str) or "")
        if not last_dir or not os.path.isdir(last_dir):
            last_dir = os.path.expanduser("~")
        files, _selected_filter = QFileDialog.getOpenFileNames(
            self, S.FILE_DIALOG_TITLE, last_dir, S.FILE_DIALOG_FILTER
        )
        if not files:
            return
        settings.setValue(LAST_DIR_KEY, os.path.dirname(files[0]))
        expanded = self.expand_paths(list(files))
        if expanded:
            self.files_dropped.emit(expanded)


__all__ = ["LAST_DIR_KEY", "SETTINGS_APP", "SETTINGS_ORG", "DropZone"]
