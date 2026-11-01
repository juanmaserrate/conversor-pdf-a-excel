"""Lista de archivos de la pantalla 2: miniatura + nombre + estado (ícono y texto, nunca solo color)."""

from __future__ import annotations

import os
from typing import Literal

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPixmap, QResizeEvent
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from pdf2excel.ui import strings_es as S

FileState = Literal["pending", "working", "ok", "warn", "error"]

THUMB_W, THUMB_H = 48, 68
PATH_ROLE = int(Qt.ItemDataRole.UserRole) + 1

# Glifo + color por estado. El glifo va siempre junto al texto: el color es solo refuerzo.
STATE_GLYPHS: dict[str, str] = {
    "pending": "○",
    "working": "⏳",
    "ok": "✔",
    "warn": "⚠",
    "error": "✖",
}
STATE_COLORS: dict[str, str] = {
    "pending": "#6E7781",
    "working": "#1F6FEB",
    "ok": "#1A7F37",
    "warn": "#9A6700",
    "error": "#CF222E",
}


def _placeholder_thumbnail() -> QPixmap:
    """Hoja gris con borde, para antes de que llegue la miniatura real."""
    pix = QPixmap(THUMB_W, THUMB_H)
    pix.fill(QColor("#F6F8FA"))
    painter = QPainter(pix)
    painter.setPen(QColor("#D0D7DE"))
    painter.drawRect(0, 0, THUMB_W - 1, THUMB_H - 1)
    painter.setPen(QColor("#9AA4B2"))
    painter.drawText(pix.rect(), int(Qt.AlignmentFlag.AlignCenter), "PDF")
    painter.end()
    return pix


class FileRow(QWidget):
    """Widget de cada ítem: miniatura a la izquierda, nombre arriba y estado (glifo + texto) abajo."""

    def __init__(self, path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.path = path
        self.state: FileState = "pending"
        self.status_text = S.FILE_STATUS_QUEUED

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(10)

        self.thumb = QLabel(self)
        self.thumb.setFixedSize(THUMB_W, THUMB_H)
        self.thumb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumb.setPixmap(_placeholder_thumbnail())
        layout.addWidget(self.thumb)

        texts = QVBoxLayout()
        texts.setSpacing(2)
        self.name_label = QLabel(os.path.basename(path), self)
        self.name_label.setObjectName("FileName")
        font = self.name_label.font()
        font.setBold(True)
        self.name_label.setFont(font)
        self.name_label.setToolTip(path)
        self.name_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        texts.addWidget(self.name_label)

        status_row = QHBoxLayout()
        status_row.setSpacing(4)
        self.glyph_label = QLabel(self)
        self.glyph_label.setObjectName("FileStateGlyph")
        self.status_label = QLabel(self)
        self.status_label.setObjectName("FileStatus")
        self.status_label.setWordWrap(True)
        status_row.addWidget(self.glyph_label)
        status_row.addWidget(self.status_label, 1)
        texts.addLayout(status_row)
        texts.addStretch(1)
        layout.addLayout(texts, 1)

        self.setAccessibleName(os.path.basename(path))
        self.set_status(self.status_text, "pending")

    def set_status(self, text: str, state: FileState) -> None:
        self.state = state
        self.status_text = text
        glyph = STATE_GLYPHS.get(state, STATE_GLYPHS["pending"])
        color = STATE_COLORS.get(state, STATE_COLORS["pending"])
        self.glyph_label.setText(glyph)
        self.glyph_label.setStyleSheet(f"color: {color}; font-weight: 600;")
        self.status_label.setText(text)
        self.status_label.setToolTip(text)
        self.setAccessibleDescription(text)

    def set_thumbnail(self, image: QImage) -> None:
        if image.isNull():
            return
        pix = QPixmap.fromImage(image).scaled(
            THUMB_W, THUMB_H, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
        )
        self.thumb.setPixmap(pix)


class FileListWidget(QListWidget):
    """``QListWidget`` con un ``FileRow`` por archivo, indexado por ruta."""

    current_file_changed = Signal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("FileList")
        self.setAccessibleName(S.REVIEW_TITLE)
        self.setIconSize(QSize(THUMB_W, THUMB_H))
        self.setUniformItemSizes(False)
        self.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setMinimumWidth(240)
        self.currentItemChanged.connect(self._on_current_changed)

    # ------------------------------------------------------------------ helpers internos

    def _item_for(self, path: str) -> QListWidgetItem | None:
        for i in range(self.count()):
            item = self.item(i)
            if item is not None and item.data(PATH_ROLE) == path:
                return item
        return None

    def _row_for(self, path: str) -> FileRow | None:
        item = self._item_for(path)
        if item is None:
            return None
        widget = self.itemWidget(item)
        return widget if isinstance(widget, FileRow) else None

    def _on_current_changed(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is not None:
            self.current_file_changed.emit(str(current.data(PATH_ROLE)))

    def _update_size_hint(self, item: QListWidgetItem, row: FileRow) -> None:
        """Alto según el ancho visible (el estado puede ocupar 2 renglones), nunca más ancho que la lista."""
        width = max(self.viewport().width(), self.minimumWidth())
        layout = row.layout()
        height = row.sizeHint().height()
        if layout is not None and layout.hasHeightForWidth():
            height = max(height, layout.heightForWidth(width))
        item.setSizeHint(QSize(width, max(height, THUMB_H + 8)))

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        for i in range(self.count()):
            item = self.item(i)
            widget = self.itemWidget(item) if item is not None else None
            if item is not None and isinstance(widget, FileRow):
                self._update_size_hint(item, widget)

    # ------------------------------------------------------------------ API pública

    def add_file(self, path: str) -> None:
        """Agrega un archivo (si no estaba) en estado "En cola"."""
        if self._item_for(path) is not None:
            return
        item = QListWidgetItem()
        item.setData(PATH_ROLE, path)
        item.setData(Qt.ItemDataRole.ToolTipRole, path)
        item.setData(Qt.ItemDataRole.AccessibleTextRole, os.path.basename(path))
        item.setData(Qt.ItemDataRole.AccessibleDescriptionRole, S.FILE_STATUS_QUEUED)
        row = FileRow(path, self)
        self.addItem(item)
        self.setItemWidget(item, row)
        self._update_size_hint(item, row)
        if self.currentItem() is None:
            self.setCurrentItem(item)

    def set_status(self, path: str, text: str, state: FileState) -> None:
        """Cambia el estado visible (glifo + texto) y el texto accesible del ítem."""
        item = self._item_for(path)
        row = self._row_for(path)
        if item is None or row is None:
            return
        row.set_status(text, state)
        item.setData(Qt.ItemDataRole.AccessibleDescriptionRole, text)
        self._update_size_hint(item, row)

    def set_thumbnail(self, path: str, image: QImage) -> None:
        """Miniatura de la primera página (se ajusta a 48×68)."""
        row = self._row_for(path)
        if row is not None:
            row.set_thumbnail(image)

    def remove_file(self, path: str) -> None:
        item = self._item_for(path)
        if item is None:
            return
        row = self.row(item)
        taken = self.takeItem(row)
        del taken

    def paths(self) -> list[str]:
        return [str(self.item(i).data(PATH_ROLE)) for i in range(self.count())]

    def current_path(self) -> str | None:
        item = self.currentItem()
        return str(item.data(PATH_ROLE)) if item is not None else None

    def select_file(self, path: str) -> None:
        item = self._item_for(path)
        if item is not None:
            self.setCurrentItem(item)

    def status_text(self, path: str) -> str:
        row = self._row_for(path)
        return row.status_text if row is not None else ""

    def state_of(self, path: str) -> FileState | None:
        row = self._row_for(path)
        return row.state if row is not None else None

    def accessible_description(self, path: str) -> str:
        row = self._row_for(path)
        return row.accessibleDescription() if row is not None else ""


__all__ = ["STATE_GLYPHS", "THUMB_H", "THUMB_W", "FileListWidget", "FileRow", "FileState"]
