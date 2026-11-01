"""Estilos del Excel de salida: colores, formatos numéricos y anchos (DESIGN.md §6).

Los `number_format` son neutrales: Excel los muestra según la configuración regional del usuario
(en es-AR, `#,##0.00` se ve como `1.234,56`).
"""

from __future__ import annotations

from typing import Any

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Colores (hex RGB sin '#').
HEADER_FILL = "1F4E78"
HEADER_FONT_COLOR = "FFFFFF"
REVIEW_FILL = "FFF2CC"
BORDER_COLOR = "BFBFBF"

# Formatos numéricos de openpyxl.
FMT_INT = "#,##0"
FMT_DEC = "#,##0.00"
FMT_MONEY = '"$" #,##0.00'
FMT_PCT = "0.00%"
FMT_DATE = "dd/mm/yyyy"

# Anchos de columna (en caracteres, unidad de Excel).
MAX_COL_WIDTH = 60
MIN_COL_WIDTH = 8
TEXT_COL_WIDTH = 100

# Texto del comentario que acompaña a cada celda marcada para revisar.
review_comment_text = "Revisar: no se pudo interpretar como número/fecha"
REVIEW_COMMENT_AUTHOR = "Conversor"


def thin_border() -> Border:
    """Borde fino gris en los cuatro lados."""
    side = Side(style="thin", color=BORDER_COLOR)
    return Border(left=side, right=side, top=side, bottom=side)


def total_row_border() -> Border:
    """Borde para filas de totales: fino a los lados y más grueso arriba."""
    side = Side(style="thin", color=BORDER_COLOR)
    top = Side(style="medium", color=BORDER_COLOR)
    return Border(left=side, right=side, top=top, bottom=side)


def header_style() -> dict[str, Any]:
    """Estilo de la fila de encabezado, como dict atributo -> valor para aplicar con `setattr`."""
    return {
        "font": Font(bold=True, color=HEADER_FONT_COLOR),
        "fill": PatternFill(fill_type="solid", fgColor=HEADER_FILL),
        "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "border": thin_border(),
    }


def review_fill() -> PatternFill:
    """Relleno amarillo de las celdas para revisar."""
    return PatternFill(fill_type="solid", fgColor=REVIEW_FILL)


def body_alignment(right: bool, wrap: bool = False) -> Alignment:
    """Alineación del cuerpo: derecha para números/fechas, izquierda para texto; siempre arriba."""
    return Alignment(horizontal="right" if right else "left", vertical="top", wrap_text=wrap)


__all__ = [
    "BORDER_COLOR",
    "FMT_DATE",
    "FMT_DEC",
    "FMT_INT",
    "FMT_MONEY",
    "FMT_PCT",
    "HEADER_FILL",
    "HEADER_FONT_COLOR",
    "MAX_COL_WIDTH",
    "MIN_COL_WIDTH",
    "REVIEW_COMMENT_AUTHOR",
    "REVIEW_FILL",
    "TEXT_COL_WIDTH",
    "body_alignment",
    "header_style",
    "review_comment_text",
    "review_fill",
    "thin_border",
    "total_row_border",
]
