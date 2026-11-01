"""Elección del modo de exportación: "Igual al PDF" (layout) o "Solo tablas" (LAYOUT_MODE.md §5). Sin Qt.

``suggest_mode`` mira lo que ``analyze`` ya sabe de un documento (tablas tipadas, ``LayoutPage`` y si hubo
OCR) y devuelve ``(modo, frase)``; la frase está en castellano llano para mostrarla bajo el selector de la UI.

La regla se calibró con los dos corpus de ``tests/corpus`` (listados, extractos y facturas de lista larga →
tablas; hojas de ruta, remitos, facturas con bloques y formularios → layout):

1. Sin diseño calculado → tablas.
2. Una tabla que sigue en otra página (listado, extracto) → tablas.
3. Escaneo con OCR y tablas → tablas (sin bordes confiables, la grilla sería ruido).
4. Hay ``MIN_OUTSIDE_BORDERED`` o más celdas con borde FUERA de las tablas detectadas (cajas de encabezado,
   campos para completar, firmas, líneas de escritura) → layout. Con o sin tablas: un formulario sin tablas
   también cae acá.
5. Tabla "con diseño": al menos ``MIN_TABLE_BANDS`` renglones rellenos de color debajo del encabezado
   (subtotales, totales) → layout.
6. Cualquier otra cosa (listados, texto corrido, una tabla con lo mínimo alrededor) → tablas.

``tabla_seccion_combinada`` (una sola tabla con encabezado agrupado y renglones de subtotal rellenos) da
``layout`` por la regla 5: en "Solo tablas" el encabezado agrupado y los subtotales pierden la forma que
les da el color, y el usuario espera verlos como en el PDF. Las tablas con solo un encabezado de color
(``header_dos_filas``) siguen en "tables".
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any

from pdf2excel.models import EffectiveMode, LayoutCell, LayoutPage

Bbox = tuple[float, float, float, float]

MIN_OUTSIDE_BORDERED = 6  # celdas con borde fuera de tablas para decir "esto es un formulario"
MIN_TABLE_BANDS = 3  # renglones rellenos (debajo del encabezado) para decir "tabla con diseño propio"

REASON_LAYOUT_FORM = (
    "Parece un formulario (varias cajas y datos sueltos): conviene respetar el diseño del PDF."
)
REASON_LAYOUT_MIXED = (
    "Mezcla tablas con cajas y datos sueltos: conviene respetar el diseño del PDF para que no se pierdan."
)
REASON_LAYOUT_BANDS = "Es una tabla con subtotales y totales de color: conviene respetar el diseño del PDF."
REASON_TABLES_LIST = "Parece un listado: conviene una tabla limpia."
REASON_TABLES_CONTINUES = "Parece un listado que sigue en varias páginas: conviene una tabla limpia."
REASON_TABLES_SCAN = "Es un escaneo: conviene una tabla limpia, sin intentar copiar el diseño."
REASON_TABLES_TEXT = "Es texto corrido, sin cajas ni tablas: conviene una hoja de texto."
REASON_TABLES_NONE = "No se pudo leer el diseño de este PDF: se exportan las tablas de datos."


@dataclass(frozen=True)
class ModeFeatures:
    """Números que usa la regla (se exponen para diagnóstico y para los tests)."""

    n_pages: int  # páginas con diseño calculado
    n_tables: int
    multipage_table: bool
    ocr_applied: bool
    outside_bordered: int  # celdas con algún borde fuera de las tablas, en todo el documento
    table_bands: int  # renglones rellenos de color dentro de tablas, sin contar el primero (encabezado)
    table_word_ratio: float  # palabras de las tablas / palabras visibles


def _sides(cell: LayoutCell) -> int:
    return sum(
        side is not None
        for side in (cell.border_left, cell.border_top, cell.border_right, cell.border_bottom)
    )


def _inside(cell: LayoutCell, boxes: Sequence[Bbox]) -> bool:
    cx = (cell.bbox[0] + cell.bbox[2]) / 2
    cy = (cell.bbox[1] + cell.bbox[3]) / 2
    return any(b[0] - 2 <= cx <= b[2] + 2 and b[1] - 2 <= cy <= b[3] + 2 for b in boxes)


def _table_words(report: Any) -> int:
    total = 0
    for table in getattr(report, "tables", ()):
        total += sum(len(str(c).split()) for c in table.header)
        total += sum(len(str(c).split()) for row in table.rows for c in row if c is not None)
    return total


def mode_features(report: Any, table_bboxes: Mapping[int, Sequence[Bbox]] | None = None) -> ModeFeatures:
    """Extrae los números de la regla. ``table_bboxes``: ``página (0-based) -> bboxes de tablas`` en el marco
    de ``LayoutCell.bbox``; sin ellos no se pueden distinguir celdas de tabla de celdas sueltas, y todo lo
    con borde cuenta como "fuera"."""
    layout_pages: list[LayoutPage] = list(getattr(report, "layout_pages", ()))
    tables = list(getattr(report, "tables", ()))
    outside = 0
    bands = 0
    for page in layout_pages:
        bboxes = list((table_bboxes or {}).get(page.page, ()))
        for cell in page.cells:
            if _sides(cell) and not (bboxes and _inside(cell, bboxes)):
                outside += 1
        for box in bboxes:
            filled_rows = {round(c.bbox[1]) for c in page.cells if c.fill and _inside(c, [box])}
            bands += max(0, len(filled_rows) - 1)
    words = sum(p.n_words for p in layout_pages)
    return ModeFeatures(
        n_pages=len(layout_pages),
        n_tables=len(tables),
        multipage_table=any(t.pages[1] > t.pages[0] for t in tables),
        ocr_applied=bool(getattr(report, "ocr_applied", False)),
        outside_bordered=outside,
        table_bands=bands,
        table_word_ratio=_table_words(report) / words if words else 0.0,
    )


def decide(f: ModeFeatures) -> tuple[EffectiveMode, str]:
    """La regla (ver el docstring del módulo) aplicada a los números."""
    if f.n_pages == 0:
        return "tables", REASON_TABLES_NONE
    if f.multipage_table:
        return "tables", REASON_TABLES_CONTINUES
    if f.ocr_applied and f.n_tables:
        return "tables", REASON_TABLES_SCAN
    if f.outside_bordered >= MIN_OUTSIDE_BORDERED:
        return "layout", REASON_LAYOUT_FORM if f.n_tables < 2 else REASON_LAYOUT_MIXED
    if f.table_bands >= MIN_TABLE_BANDS:
        return "layout", REASON_LAYOUT_BANDS
    if f.n_tables == 0:
        return "tables", REASON_TABLES_TEXT
    return "tables", REASON_TABLES_LIST


def suggest_mode(
    report_like: Any, table_bboxes: Mapping[int, Sequence[Bbox]] | None = None
) -> tuple[EffectiveMode, str]:
    """``("layout" | "tables", frase en castellano)`` para un reporte ya analizado."""
    return decide(mode_features(report_like, table_bboxes))
