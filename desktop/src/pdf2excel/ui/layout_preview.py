"""Vista previa aproximada del modo "Igual al PDF": una ``LayoutPage`` dibujada como grilla de celdas.

Es solo lectura. Las celdas combinadas usan ``setSpan``; un delegado dibuja relleno, bordes y texto con
los estilos del PDF (aproximados), y las imágenes se pintan como miniaturas sobre el viewport.
"""

from __future__ import annotations

import html

from PySide6.QtCore import QModelIndex, QPersistentModelIndex, QPoint, QRect, QRectF, QSize, Qt
from PySide6.QtGui import (
    QAbstractTextDocumentLayout,
    QBrush,
    QColor,
    QPainter,
    QPaintEvent,
    QPen,
    QPixmap,
    QResizeEvent,
    QTextDocument,
    QTextOption,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHeaderView,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from pdf2excel.models import BorderSide, LayoutCell, LayoutPage

_CELL_ROLE = int(Qt.ItemDataRole.UserRole) + 10
_MIN_ZOOM = 0.25
_MAX_ZOOM = 1.6
_MIN_FONT_PX = 6.0
_FONT_FALLBACK = "Segoe UI, Helvetica, Arial, sans-serif"


def _color(hex_rgb: str | None, default: str) -> QColor:
    text = (hex_rgb or default).lstrip("#")
    color = QColor(f"#{text}")
    return color if color.isValid() else QColor(f"#{default}")


def cell_html(cell: LayoutCell, zoom: float) -> str:
    """HTML mínimo con el texto de la celda (una ``<div>`` por línea, un ``<span>`` por tramo)."""
    parts: list[str] = []
    for line in cell.lines:
        spans: list[str] = []
        for run in line:
            if not run.text:
                continue
            size = max(_MIN_FONT_PX, run.size_pt * zoom * 1.33)
            style = f"font-size:{size:.1f}px;"
            if run.bold:
                style += "font-weight:bold;"
            if run.italic:
                style += "font-style:italic;"
            style += f"color:#{(run.color or '000000').lstrip('#')};"
            spans.append(f'<span style="{style}">{html.escape(run.text)}</span>')
        parts.append("<div>" + ("".join(spans) or "&nbsp;") + "</div>")
    return "".join(parts)


class _LayoutDelegate(QStyledItemDelegate):
    """Dibuja relleno + texto + bordes de cada ``LayoutCell`` (incluye el rango combinado completo)."""

    def __init__(self, view: LayoutPreview) -> None:
        super().__init__(view)
        self._view = view
        self._docs: dict[tuple[int, int, int], QTextDocument] = {}

    def clear_cache(self) -> None:
        self._docs.clear()

    def sizeHint(self, option: QStyleOptionViewItem, index: QModelIndex | QPersistentModelIndex) -> QSize:
        return QSize(0, 0)

    def paint(
        self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex | QPersistentModelIndex
    ) -> None:
        cell = index.data(_CELL_ROLE)
        if not isinstance(cell, LayoutCell):
            return
        rect: QRect = option.rect
        painter.save()
        if cell.fill:
            painter.fillRect(rect, _color(cell.fill, "FFFFFF"))
        if cell.lines:
            self._paint_text(painter, rect, cell)
        self._paint_borders(painter, rect, cell)
        painter.restore()

    def _paint_text(self, painter: QPainter, rect: QRect, cell: LayoutCell) -> None:
        zoom = self._view.zoom
        key = (cell.row, cell.col, int(zoom * 1000))
        doc = self._docs.get(key)
        inner = rect.adjusted(2, 1, -2, -1)
        if inner.width() < 4 or inner.height() < 4:
            return
        if doc is None:
            doc = QTextDocument()
            doc.setDocumentMargin(0)
            doc.setDefaultFont(self._view.font())
            doc.setDefaultStyleSheet(f"div {{ font-family: {_FONT_FALLBACK}; }}")
            doc.setHtml(cell_html(cell, zoom))
            self._docs[key] = doc
        option = QTextOption()
        option.setWrapMode(QTextOption.WrapMode.WordWrap if cell.wrap else QTextOption.WrapMode.NoWrap)
        align = {
            "left": Qt.AlignmentFlag.AlignLeft,
            "center": Qt.AlignmentFlag.AlignHCenter,
            "right": Qt.AlignmentFlag.AlignRight,
        }.get(cell.halign, Qt.AlignmentFlag.AlignLeft)
        option.setAlignment(align)
        doc.setDefaultTextOption(option)
        doc.setTextWidth(inner.width())
        height = doc.size().height()
        top = inner.top()
        if cell.valign == "center":
            top += int((inner.height() - height) / 2)
        elif cell.valign == "bottom":
            top += int(inner.height() - height)
        painter.save()
        painter.setClipRect(inner)
        painter.translate(inner.left(), max(inner.top(), top))
        ctx = QAbstractTextDocumentLayout.PaintContext()
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

    def _paint_borders(self, painter: QPainter, rect: QRect, cell: LayoutCell) -> None:
        zoom = self._view.zoom
        right = rect.right()
        bottom = rect.bottom()
        sides: list[tuple[BorderSide | None, QPoint, QPoint]] = [
            (cell.border_left, rect.topLeft(), QPoint(rect.left(), bottom)),
            (cell.border_top, rect.topLeft(), QPoint(right, rect.top())),
            (cell.border_right, QPoint(right, rect.top()), QPoint(right, bottom)),
            (cell.border_bottom, QPoint(rect.left(), bottom), QPoint(right, bottom)),
        ]
        for side, start, end in sides:
            if side is None:
                continue
            pen = QPen(_color(side.color, "000000"))
            pen.setWidthF(min(3.0, max(1.0, side.width_pt * zoom)))
            pen.setCapStyle(Qt.PenCapStyle.FlatCap)
            painter.setPen(pen)
            painter.drawLine(start, end)


class LayoutPreview(QTableWidget):
    """``QTableWidget`` de solo lectura que muestra una ``LayoutPage`` con proporciones del PDF."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(0, 0, parent)
        self.setObjectName("LayoutPreview")
        self._page: LayoutPage | None = None
        self.zoom = 1.0
        self._pixmaps: dict[int, QPixmap] = {}
        self._scaled: dict[tuple[int, int, int], QPixmap] = {}
        self._delegate = _LayoutDelegate(self)
        self.setItemDelegate(self._delegate)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.setShowGrid(False)
        self.setWordWrap(False)
        self.setCornerButtonEnabled(False)
        self.horizontalHeader().setVisible(False)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.horizontalHeader().setMinimumSectionSize(1)
        self.verticalHeader().setMinimumSectionSize(1)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.viewport().setStyleSheet("background: white;")
        self.setStyleSheet("QTableWidget#LayoutPreview { background: white; }")

    # ------------------------------------------------------------------ API

    def page(self) -> LayoutPage | None:
        return self._page

    def set_page(self, page: LayoutPage | None) -> None:
        """Dibuja ``page`` (o limpia la vista con ``None``)."""
        self.setUpdatesEnabled(False)
        try:
            self.clearSpans()
            self.clear()
            self._delegate.clear_cache()
            self._pixmaps.clear()
            self._scaled.clear()
            self._page = page
            if page is None:
                self.setRowCount(0)
                self.setColumnCount(0)
                return
            n_rows = len(page.row_heights_pt)
            n_cols = len(page.col_widths_pt)
            self.setRowCount(n_rows)
            self.setColumnCount(n_cols)
            self.zoom = self._fit_zoom()
            self._apply_sizes()
            self._fill_cells(page, n_rows, n_cols)
            self.setAccessibleName(f"Diseño de la página {page.page + 1}")
        finally:
            self.setUpdatesEnabled(True)
        self.viewport().update()

    # ------------------------------------------------------------------ internos

    def _fit_zoom(self) -> float:
        page = self._page
        if page is None or page.width_pt <= 0:
            return 1.0
        # El ancho de la página entra justo en el viewport (con lugar para la barra vertical).
        bar = self.style().pixelMetric(QStyle.PixelMetric.PM_ScrollBarExtent)
        available = max(120, self.width() - 2 * self.frameWidth() - bar - 1)
        return min(_MAX_ZOOM, max(_MIN_ZOOM, available / page.width_pt))

    def _apply_sizes(self) -> None:
        page = self._page
        if page is None:
            return
        for col, width in enumerate(page.col_widths_pt):
            self.setColumnWidth(col, max(1, round(width * self.zoom)))
        for row, height in enumerate(page.row_heights_pt):
            self.setRowHeight(row, max(1, round(height * self.zoom)))

    def _fill_cells(self, page: LayoutPage, n_rows: int, n_cols: int) -> None:
        occupied: set[tuple[int, int]] = set()
        for cell in page.cells:
            if not (0 <= cell.row < n_rows and 0 <= cell.col < n_cols):
                continue
            row_span = max(1, min(cell.row_span, n_rows - cell.row))
            col_span = max(1, min(cell.col_span, n_cols - cell.col))
            area = {(cell.row + r, cell.col + c) for r in range(row_span) for c in range(col_span)}
            if area & occupied:
                row_span = col_span = 1
                area = {(cell.row, cell.col)}
                if area & occupied:
                    continue
            occupied |= area
            if row_span > 1 or col_span > 1:
                self.setSpan(cell.row, cell.col, row_span, col_span)
            item = QTableWidgetItem()
            item.setData(_CELL_ROLE, cell)
            item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            if cell.text:
                item.setToolTip(cell.text)
            self.setItem(cell.row, cell.col, item)

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        if self._page is None:
            return
        zoom = self._fit_zoom()
        if abs(zoom - self.zoom) / self.zoom > 0.02:
            self.zoom = zoom
            self._delegate.clear_cache()
            self._scaled.clear()
            self._apply_sizes()

    def _pixmap_for(self, index: int) -> QPixmap | None:
        page = self._page
        if page is None:
            return None
        pix = self._pixmaps.get(index)
        if pix is None:
            pix = QPixmap()
            image = page.images[index]
            if not pix.loadFromData(image.data):
                return None
            self._pixmaps[index] = pix
        return pix

    def paintEvent(self, event: QPaintEvent) -> None:
        super().paintEvent(event)
        page = self._page
        if page is None or not page.images:
            return
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        for i, image in enumerate(page.images):
            if not (0 <= image.row < self.rowCount() and 0 <= image.col < self.columnCount()):
                continue
            pix = self._pixmap_for(i)
            if pix is None or pix.isNull():
                continue
            width = max(2, round(image.width_pt * self.zoom))
            height = max(2, round(image.height_pt * self.zoom))
            x = self.columnViewportPosition(image.col) + round(image.dx_pt * self.zoom)
            y = self.rowViewportPosition(image.row) + round(image.dy_pt * self.zoom)
            key = (i, width, height)
            scaled = self._scaled.get(key)
            if scaled is None:
                scaled = pix.scaled(
                    width,
                    height,
                    Qt.AspectRatioMode.IgnoreAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                self._scaled[key] = scaled
            painter.drawPixmap(x, y, scaled)
            painter.setPen(QPen(QColor("#c8c8c8"), 1))
            painter.setBrush(QBrush(Qt.BrushStyle.NoBrush))
            painter.drawRect(QRectF(x, y, width, height))
        painter.end()


__all__ = ["LayoutPreview", "cell_html"]
