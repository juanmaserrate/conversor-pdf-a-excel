"""Modelo de solo lectura para la vista previa de una ``TypedTable`` (pantalla 2)."""

from __future__ import annotations

from collections.abc import Callable
from datetime import date, datetime
from typing import Any

from PySide6.QtCore import QAbstractTableModel, QLocale, QModelIndex, QObject, QPersistentModelIndex, Qt
from PySide6.QtGui import QColor, QFont

from pdf2excel.models import CellValue, ColumnInfo, TypedTable
from pdf2excel.ui import strings_es as S

REVIEW_COLOR = "#FFF2CC"
RIGHT_ALIGNED_KINDS = frozenset({"number", "percent", "date"})
DEFAULT_MAX_ROWS = 200

Formatter = Callable[[CellValue, ColumnInfo], str]
ModelIndex = QModelIndex | QPersistentModelIndex

_formatter: Formatter | None = None


def _fallback_format(value: CellValue, col: ColumnInfo) -> str:
    """Formato AR con ``QLocale("es_AR")`` si no se puede importar ``excel.writer``."""
    if value is None:
        return ""
    if isinstance(value, datetime | date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, bool):
        return "Sí" if value else "No"
    if isinstance(value, int | float):
        loc = QLocale("es_AR")
        loc.setNumberOptions(QLocale.NumberOption.DefaultNumberOptions)
        fmt = col.number_format
        if fmt is None:
            if col.kind == "percent":
                fmt = "0.00%"
            elif col.kind == "number":
                has_decimals = isinstance(value, float) and not float(value).is_integer()
                fmt = "#,##0.00" if has_decimals else "#,##0"
        if fmt == "0.00%":
            return f"{loc.toString(float(value) * 100, 'f', 2)} %"
        if fmt is not None and fmt.startswith('"$"'):
            return f"$ {loc.toString(float(value), 'f', 2)}"
        if fmt == "#,##0":
            return loc.toString(float(value), "f", 0)
        if fmt == "#,##0.00":
            return loc.toString(float(value), "f", 2)
        if isinstance(value, int):
            return loc.toString(float(value), "f", 0)
        return loc.toString(float(value), "f", 2)
    return str(value)


def _get_formatter() -> Formatter:
    """``excel.writer.format_preview_value`` si está disponible (import diferido), si no el fallback."""
    global _formatter
    if _formatter is None:
        try:
            from pdf2excel.excel.writer import format_preview_value

            _formatter = format_preview_value
        except Exception:
            _formatter = _fallback_format
    return _formatter


def format_cell(value: CellValue, col: ColumnInfo) -> str:
    """Texto de la celda tal como lo mostrará Excel en es-AR (``None`` → "")."""
    if value is None:
        return ""
    return _get_formatter()(value, col)


class TablePreviewModel(QAbstractTableModel):
    """Expone una ``TypedTable`` (primeras ``max_rows`` filas) para un ``QTableView`` de solo lectura."""

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._table: TypedTable | None = None
        self._max_rows = DEFAULT_MAX_ROWS
        self._bold = QFont()
        self._bold.setBold(True)
        self._review_color = QColor(REVIEW_COLOR)

    # ------------------------------------------------------------------ API pública

    def set_table(self, table: TypedTable | None, max_rows: int = DEFAULT_MAX_ROWS) -> None:
        self.beginResetModel()
        self._table = table
        self._max_rows = max(0, max_rows)
        self.endResetModel()

    def table(self) -> TypedTable | None:
        return self._table

    def total_rows(self) -> int:
        """Filas reales de la tabla (puede ser más que ``rowCount()``)."""
        return len(self._table.rows) if self._table is not None else 0

    def is_truncated(self) -> bool:
        return self.total_rows() > self.rowCount()

    # ------------------------------------------------------------------ QAbstractTableModel

    def rowCount(self, parent: ModelIndex = QModelIndex()) -> int:  # noqa: B008
        if parent.isValid() or self._table is None:
            return 0
        return min(len(self._table.rows), self._max_rows)

    def columnCount(self, parent: ModelIndex = QModelIndex()) -> int:  # noqa: B008
        if parent.isValid() or self._table is None:
            return 0
        return max(len(self._table.header), len(self._table.columns))

    def headerData(
        self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole
    ) -> Any:
        if self._table is None:
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if 0 <= section < len(self._table.header):
                    return self._table.header[section]
                col = self._column(section)
                return col.header if col is not None and col.header else str(section + 1)
            return str(section + 1)
        if role == Qt.ItemDataRole.TextAlignmentRole and orientation == Qt.Orientation.Horizontal:
            return Qt.AlignmentFlag.AlignCenter
        return None

    def flags(self, index: ModelIndex) -> Qt.ItemFlag:
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable

    def data(self, index: ModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        table = self._table
        if table is None or not index.isValid():
            return None
        r, c = index.row(), index.column()
        if r >= self.rowCount() or c >= self.columnCount():
            return None
        col = self._column(c)

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._value(r, c)
            if col is None:
                return "" if value is None else str(value)
            return format_cell(value, col)
        if role == Qt.ItemDataRole.TextAlignmentRole:
            if col is not None and col.kind in RIGHT_ALIGNED_KINDS:
                return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        if role == Qt.ItemDataRole.BackgroundRole:
            if (r, c) in table.review_cells:
                return self._review_color
            return None
        if role == Qt.ItemDataRole.FontRole:
            if r in table.total_rows:
                return self._bold
            return None
        if role == Qt.ItemDataRole.ToolTipRole:
            if (r, c) in table.review_cells:
                return S.REVIEW_TOOLTIP
            return None
        if role == Qt.ItemDataRole.AccessibleTextRole:
            value = self._value(r, c)
            text = format_cell(value, col) if col is not None else ("" if value is None else str(value))
            if (r, c) in table.review_cells:
                return f"{text} ({S.REVIEW_TOOLTIP})"
            return text
        return None

    # ------------------------------------------------------------------ helpers

    def _column(self, c: int) -> ColumnInfo | None:
        if self._table is None or c < 0 or c >= len(self._table.columns):
            return None
        return self._table.columns[c]

    def _value(self, r: int, c: int) -> CellValue:
        if self._table is None:
            return None
        row = self._table.rows[r]
        return row[c] if c < len(row) else None


__all__ = ["DEFAULT_MAX_ROWS", "REVIEW_COLOR", "TablePreviewModel", "format_cell"]
