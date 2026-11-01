"""Opciones avanzadas (engranaje) y su persistencia con ``QSettings``."""

from __future__ import annotations

import dataclasses
from pathlib import Path
from typing import Any

from PySide6.QtCore import QSettings, Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QLabel,
    QVBoxLayout,
    QWidget,
)

from pdf2excel import __version__, logging_setup
from pdf2excel.models import Options
from pdf2excel.ui import strings_es as S

SETTINGS_ORG = "ConversorPDF"
SETTINGS_APP = "ConversorPDFaExcel"

# Con `from __future__ import annotations` f.type es el string "bool".
_BOOL_FIELDS = frozenset(f.name for f in dataclasses.fields(Options) if f.type in ("bool", bool))


def _settings() -> QSettings:
    return QSettings(SETTINGS_ORG, SETTINGS_APP)


def _to_bool(value: Any, default: bool) -> bool:
    """QSettings en formato INI devuelve strings ("true"/"false"); en el registro, bool."""
    if isinstance(value, bool):
        return value
    if isinstance(value, int | float):
        return bool(value)
    if isinstance(value, str):
        text = value.strip().lower()
        if text in ("true", "1", "yes", "si", "sí", "on"):
            return True
        if text in ("false", "0", "no", "off", ""):
            return False
    return default


def load_options() -> Options:
    """Lee las opciones guardadas; cualquier clave ausente o rara toma el valor por defecto."""
    settings = _settings()
    defaults = Options()
    values: dict[str, Any] = {}
    for f in dataclasses.fields(Options):
        default = getattr(defaults, f.name)
        if not settings.contains(f.name):
            values[f.name] = default
            continue
        raw = settings.value(f.name)
        if f.name == "output_dir":
            text = str(raw).strip() if raw is not None else ""
            values[f.name] = Path(text) if text else None
        elif f.name in _BOOL_FIELDS or isinstance(default, bool):
            values[f.name] = _to_bool(raw, bool(default))
        elif f.name == "layout":
            text = str(raw).strip().lower() if raw is not None else ""
            values[f.name] = text if text in ("auto", "layout", "tables") else default
        elif isinstance(default, str):
            text = str(raw).strip() if raw is not None else ""
            values[f.name] = text or default
        else:
            values[f.name] = raw if raw is not None else default
    return Options(**values)


def save_options(opts: Options) -> None:
    """Guarda cada campo de ``Options`` bajo su nombre; ``output_dir`` como str (vacío = junto al PDF)."""
    settings = _settings()
    for f in dataclasses.fields(Options):
        value = getattr(opts, f.name)
        if f.name == "output_dir":
            settings.setValue(f.name, str(value) if value else "")
        elif isinstance(value, Path):
            settings.setValue(f.name, str(value))
        else:
            settings.setValue(f.name, value)
    settings.sync()


class SettingsDialog(QDialog):
    """Diálogo "Opciones avanzadas": idioma OCR + casillas. ``options()`` devuelve una ``Options`` nueva."""

    def __init__(self, opts: Options, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._initial = opts
        self.setWindowTitle(S.ADVANCED_TITLE)
        self.setModal(True)
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        intro = QLabel(S.ADVANCED_INTRO, self)
        intro.setObjectName("Muted")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)

        self.lang_combo = QComboBox(self)
        self.lang_combo.setAccessibleName(S.ADVANCED_OCR_LANG)
        for label, code in S.OCR_LANG_CHOICES:
            self.lang_combo.addItem(label, code)
        idx = self.lang_combo.findData(opts.ocr_lang)
        if idx < 0:
            # Idioma no listado (configurado a mano): lo mostramos tal cual para no perderlo.
            self.lang_combo.addItem(opts.ocr_lang, opts.ocr_lang)
            idx = self.lang_combo.count() - 1
        self.lang_combo.setCurrentIndex(idx)
        form.addRow(S.ADVANCED_OCR_LANG, self.lang_combo)

        self.layout_combo = QComboBox(self)
        self.layout_combo.setAccessibleName(S.ADVANCED_LAYOUT)
        self.layout_combo.setToolTip(S.ADVANCED_LAYOUT_TIP)
        for label, code in S.ADVANCED_LAYOUT_CHOICES:
            self.layout_combo.addItem(label, code)
        layout_idx = self.layout_combo.findData(opts.layout)
        self.layout_combo.setCurrentIndex(max(0, layout_idx))
        form.addRow(S.ADVANCED_LAYOUT, self.layout_combo)

        self.merge_check = self._check(S.ADVANCED_MERGE, opts.merge_continuations)
        self.include_text_check = self._check(S.ADVANCED_INCLUDE_TEXT, opts.include_text)
        self.one_sheet_check = self._check(S.ADVANCED_ONE_SHEET_PER_PAGE, opts.one_sheet_per_page)
        self.force_ocr_check = self._check(S.ADVANCED_FORCE_OCR, opts.force_ocr, S.ADVANCED_FORCE_OCR_TIP)
        self.origin_check = self._check(
            S.ADVANCED_ORIGIN_SHEET, opts.show_origin_sheet, S.ADVANCED_ORIGIN_SHEET_TIP
        )
        self.open_check = self._check(S.ADVANCED_OPEN_WHEN_DONE, opts.open_when_done)
        self.single_workbook_check = self._check(
            S.ADVANCED_SINGLE_WORKBOOK, opts.single_workbook, S.ADVANCED_SINGLE_WORKBOOK_TIP
        )
        for check in (
            self.merge_check,
            self.include_text_check,
            self.one_sheet_check,
            self.force_ocr_check,
            self.origin_check,
            self.open_check,
            self.single_workbook_check,
        ):
            form.addRow(check)
        layout.addLayout(form)

        line = QFrame(self)
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        log_dir = logging_setup.log_dir()
        self.version_label = QLabel(self)
        self.version_label.setObjectName("Muted")
        self.version_label.setTextFormat(Qt.TextFormat.RichText)
        self.version_label.setText(
            f"{S.VERSION_FMT.format(v=__version__)} · "
            f'<a href="{QUrl.fromLocalFile(str(log_dir)).toString()}">{S.LINK_LOG}</a>'
        )
        self.version_label.setToolTip(str(log_dir))
        self.version_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        self.version_label.setOpenExternalLinks(False)
        self.version_label.linkActivated.connect(self._open_link)
        layout.addWidget(self.version_label)

        self.buttons = QDialogButtonBox(self)
        self.ok_button = self.buttons.addButton(S.BTN_OK, QDialogButtonBox.ButtonRole.AcceptRole)
        self.cancel_button = self.buttons.addButton(S.BTN_CANCEL, QDialogButtonBox.ButtonRole.RejectRole)
        self.ok_button.setProperty("primary", True)
        self.ok_button.setDefault(True)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    @staticmethod
    def _open_link(link: str) -> None:
        QDesktopServices.openUrl(QUrl(link))

    def _check(self, text: str, checked: bool, tooltip: str | None = None) -> QCheckBox:
        box = QCheckBox(text, self)
        box.setChecked(checked)
        box.setAccessibleName(text)
        if tooltip:
            box.setToolTip(tooltip)
            box.setAccessibleDescription(tooltip)
        return box

    def options(self) -> Options:
        """``Options`` con lo elegido en el diálogo (conserva ``typed_values`` y ``output_dir`` iniciales)."""
        lang = self.lang_combo.currentData()
        layout_mode = self.layout_combo.currentData()
        return dataclasses.replace(
            self._initial,
            ocr_lang=str(lang) if lang else self._initial.ocr_lang,
            layout=layout_mode if layout_mode in ("auto", "layout", "tables") else self._initial.layout,
            merge_continuations=self.merge_check.isChecked(),
            include_text=self.include_text_check.isChecked(),
            one_sheet_per_page=self.one_sheet_check.isChecked(),
            force_ocr=self.force_ocr_check.isChecked(),
            show_origin_sheet=self.origin_check.isChecked(),
            open_when_done=self.open_check.isChecked(),
            single_workbook=self.single_workbook_check.isChecked(),
        )


__all__ = ["SETTINGS_APP", "SETTINGS_ORG", "SettingsDialog", "load_options", "save_options"]
