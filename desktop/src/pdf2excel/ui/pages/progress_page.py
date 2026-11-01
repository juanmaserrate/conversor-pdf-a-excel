"""Pantalla 3 "Convirtiendo": archivo actual, "Archivo i de n · página p de t", barra y Cancelar."""

from __future__ import annotations

import os
from typing import Literal

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QHBoxLayout, QLabel, QProgressBar, QPushButton, QVBoxLayout, QWidget

from pdf2excel.ui import strings_es as S

Mode = Literal["analyze", "export"]

# Etapas cuyo (actual, total) cuenta páginas: solo ahí tiene sentido "página p de t".
_PAGE_STAGES = frozenset({"classify", "ocr", "tables", "layout"})


class ProgressPage(QWidget):
    """Progreso del lote. ``cancel_requested`` al tocar Cancelar (o Esc); la ventana decide qué hacer."""

    cancel_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("ProgressPage")
        self.setAccessibleName(S.PROGRESS_TITLE)
        self._mode: Mode = "analyze"
        self._cancelling = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(48, 32, 48, 32)
        layout.setSpacing(12)
        layout.addStretch(2)

        self.title_label = QLabel(S.PROGRESS_TITLE, self)
        self.title_label.setObjectName("PageTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.title_label)

        self.file_label = QLabel("", self)
        self.file_label.setObjectName("ProgressFile")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.file_label.font()
        font.setBold(True)
        font.setPointSize(max(font.pointSize(), 10) + 1)
        self.file_label.setFont(font)
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

        self.counter_label = QLabel("", self)
        self.counter_label.setObjectName("Muted")
        self.counter_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.counter_label)

        self.bar = QProgressBar(self)
        self.bar.setObjectName("ProgressBar")
        self.bar.setRange(0, 1)
        self.bar.setValue(0)
        self.bar.setTextVisible(True)
        self.bar.setAccessibleName(S.PROGRESS_TITLE)
        layout.addWidget(self.bar)

        self.stage_label = QLabel("", self)
        self.stage_label.setObjectName("ProgressStage")
        self.stage_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stage_label.setWordWrap(True)
        layout.addWidget(self.stage_label)

        layout.addStretch(3)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.cancel_button = QPushButton(S.BTN_CANCEL, self)
        self.cancel_button.setAccessibleName(S.BTN_CANCEL)
        self.cancel_button.setAccessibleDescription(S.DONE_CANCELLED)
        self.cancel_button.setAutoDefault(False)
        self.cancel_button.clicked.connect(self._on_cancel_clicked)
        buttons.addWidget(self.cancel_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        self.setFocusProxy(self.cancel_button)

    # ------------------------------------------------------------------ API pública

    def set_mode(self, mode: Mode) -> None:
        """``analyze`` (antes de "Esto encontré") o ``export`` (guardando los Excel)."""
        self._mode = mode
        self.title_label.setText(S.ANALYZING if mode == "analyze" else S.PROGRESS_TITLE)
        self.setAccessibleName(self.title_label.text())

    def mode(self) -> Mode:
        return self._mode

    def set_progress(
        self, path: str, current: int, total: int, stage: str, file_index: int, file_count: int
    ) -> None:
        """Actualiza nombre, contador, barra (0..total) y texto de etapa."""
        self.file_label.setText(os.path.basename(path))
        self.file_label.setToolTip(path)
        page = current if stage in _PAGE_STAGES and total > 0 else None
        total_pages = total if page is not None else None
        self.counter_label.setText(S.progress_text(file_index, file_count, page, total_pages))
        total = max(1, int(total))
        self.bar.setRange(0, total)
        self.bar.setValue(max(0, min(int(current), total)))
        if not self._cancelling:
            self.stage_label.setText(S.STAGE_TEXT.get(stage, stage))
        self.setAccessibleDescription(f"{self.counter_label.text()} · {self.stage_label.text()}")

    def set_cancelling(self) -> None:
        """Después de tocar Cancelar: aviso y botón deshabilitado hasta que el lote termine."""
        self._cancelling = True
        self.stage_label.setText(S.PROGRESS_CANCELLING)
        self.cancel_button.setEnabled(False)

    def is_cancelling(self) -> bool:
        return self._cancelling

    def reset(self) -> None:
        self._cancelling = False
        self.cancel_button.setEnabled(True)
        self.file_label.setText("")
        self.file_label.setToolTip("")
        self.counter_label.setText("")
        self.stage_label.setText("")
        self.bar.setRange(0, 1)
        self.bar.setValue(0)

    def primary_action(self) -> None:
        """La pantalla de progreso no tiene acción primaria (Enter no hace nada)."""

    def escape_action(self) -> None:
        """Esc = Cancelar."""
        self._on_cancel_clicked()

    # ------------------------------------------------------------------ interno

    def _on_cancel_clicked(self) -> None:
        if self._cancelling:
            return
        self.set_cancelling()
        self.cancel_requested.emit()

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape:
            event.accept()
            self.escape_action()
            return
        super().keyPressEvent(event)


__all__ = ["ProgressPage"]
