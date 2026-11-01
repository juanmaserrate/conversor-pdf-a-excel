"""Pantalla 1 "Arrastrá": zona de arrastre grande, nota de privacidad y aviso discreto de versión nueva."""

from __future__ import annotations

from html import escape

from PySide6.QtCore import Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QKeyEvent
from PySide6.QtWidgets import QHBoxLayout, QLabel, QProgressBar, QToolButton, QVBoxLayout, QWidget

from pdf2excel.ui import strings_es as S
from pdf2excel.ui.drop_zone import DropZone


class DropPage(QWidget):
    """Pantalla vacía: ``files_chosen(list[str])`` con rutas absolutas de PDF (ya filtradas por DropZone).

    ``set_busy(True)`` muestra "Revisando los PDF…" con una barra indeterminada y deshabilita el arrastre
    (el análisis corto se hace acá, sin cambiar de pantalla). ``show_update`` muestra al pie la línea
    "Hay una versión nueva (x.y.z) · Descargar", descartable.
    """

    files_chosen = Signal(list)
    update_dismissed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DropPage")
        self.setAccessibleName(S.DROP_TITLE)
        self._busy = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 24, 32, 16)
        layout.setSpacing(12)

        self.drop_zone = DropZone(self)
        self.drop_zone.files_dropped.connect(self._on_files_dropped)
        layout.addWidget(self.drop_zone, 1)

        # Indicador de análisis en curso (oculto hasta set_busy(True)).
        self.busy_box = QWidget(self)
        self.busy_box.setAccessibleName(S.ANALYZING)
        busy_layout = QHBoxLayout(self.busy_box)
        busy_layout.setContentsMargins(0, 0, 0, 0)
        busy_layout.setSpacing(10)
        self.busy_label = QLabel(S.ANALYZING, self.busy_box)
        self.busy_label.setObjectName("BusyLabel")
        self.busy_bar = QProgressBar(self.busy_box)
        self.busy_bar.setRange(0, 0)  # indeterminada
        self.busy_bar.setTextVisible(False)
        self.busy_bar.setFixedHeight(10)
        self.busy_bar.setAccessibleName(S.ANALYZING)
        busy_layout.addStretch(1)
        busy_layout.addWidget(self.busy_label)
        busy_layout.addWidget(self.busy_bar, 1)
        busy_layout.addStretch(1)
        self.busy_box.hide()
        layout.addWidget(self.busy_box)

        self.privacy_label = QLabel(S.PRIVACY_NOTE, self)
        self.privacy_label.setObjectName("Muted")
        self.privacy_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.privacy_label.setWordWrap(True)
        font = self.privacy_label.font()
        font.setPointSize(max(8, font.pointSize() - 1))
        self.privacy_label.setFont(font)
        layout.addWidget(self.privacy_label)

        # Pie: aviso de actualización (oculto hasta show_update).
        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        footer.setSpacing(4)
        footer.addStretch(1)
        self.update_label = QLabel(self)
        self.update_label.setObjectName("UpdateNote")
        self.update_label.setTextFormat(Qt.TextFormat.RichText)
        self.update_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        self.update_label.setOpenExternalLinks(False)
        self.update_label.linkActivated.connect(self._open_link)
        self.update_label.hide()
        footer.addWidget(self.update_label)
        self.update_close = QToolButton(self)
        self.update_close.setText("✕")
        self.update_close.setAutoRaise(True)
        self.update_close.setToolTip(S.BTN_DISMISS)
        self.update_close.setAccessibleName(S.BTN_DISMISS)
        self.update_close.clicked.connect(self.hide_update)
        self.update_close.hide()
        footer.addWidget(self.update_close)
        footer.addStretch(1)
        layout.addLayout(footer)

        self._update_url = ""
        self.setFocusProxy(self.drop_zone)

    # ------------------------------------------------------------------ señales internas

    def _on_files_dropped(self, files: list[str]) -> None:
        if not self._busy and files:
            self.files_chosen.emit(list(files))

    @staticmethod
    def _open_link(link: str) -> None:
        QDesktopServices.openUrl(QUrl(link))

    # ------------------------------------------------------------------ API pública

    def set_busy(self, busy: bool) -> None:
        """Muestra/oculta "Revisando los PDF…" y habilita o no la zona de arrastre."""
        self._busy = busy
        self.busy_box.setVisible(busy)
        self.drop_zone.setEnabled(not busy)
        self.drop_zone.setAcceptDrops(not busy)
        self.drop_zone.setCursor(Qt.CursorShape.BusyCursor if busy else Qt.CursorShape.PointingHandCursor)

    def is_busy(self) -> bool:
        return self._busy

    def show_update(self, version: str, url: str) -> None:
        """Línea discreta al pie: "Hay una versión nueva (0.2.0) · Descargar" (el último tramo es el link)."""
        self._update_url = url
        text = S.UPDATE_AVAILABLE_FMT.format(v=version)
        head, sep, tail = text.rpartition(" · ")
        if sep:
            html = f'{escape(head)}{sep}<a href="{escape(url, quote=True)}">{escape(tail)}</a>'
        else:
            html = f'<a href="{escape(url, quote=True)}">{escape(text)}</a>'
        self.update_label.setText(html)
        self.update_label.setToolTip(url)
        self.update_label.setAccessibleName(text)
        self.update_label.show()
        self.update_close.show()

    def hide_update(self) -> None:
        self.update_label.hide()
        self.update_close.hide()
        self.update_dismissed.emit()

    def update_visible(self) -> bool:
        return self.update_label.isVisibleTo(self)

    def choose_files(self) -> None:
        """Abre el diálogo de archivos (Ctrl+O) si no hay un análisis en curso."""
        if not self._busy:
            self.drop_zone.choose_files()

    def primary_action(self) -> None:
        """Enter = elegir archivos."""
        self.choose_files()

    def escape_action(self) -> None:
        """Esc en la pantalla vacía no hace nada (la ventana principal decide si cancela un análisis)."""

    def focus_drop_zone(self) -> None:
        self.drop_zone.setFocus(Qt.FocusReason.OtherFocusReason)

    # ------------------------------------------------------------------ teclado

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape:
            event.ignore()  # que lo resuelva la ventana (cancelar análisis en curso)
            return
        super().keyPressEvent(event)


__all__ = ["DropPage"]
