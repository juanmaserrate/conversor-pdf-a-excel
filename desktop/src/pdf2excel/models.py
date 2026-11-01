"""Modelos compartidos entre core, excel, UI y jobs (contrato de DESIGN.md §4).

Solo dataclasses y alias de tipos: acá no va lógica. Todo lo demás importa de este módulo.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Literal

PageKind = Literal["digital", "scanned", "scanned_with_text", "garbage_text", "empty"]
ColumnKind = Literal["number", "date", "percent", "id", "text"]
Locale = Literal["ar", "us"]
CellValue = int | float | date | str | None
ProgressFn = Callable[[int, int, str], None]  # (actual, total, etapa) con etapa ∈ STAGES
STAGES = ("classify", "ocr", "tables", "layout", "typing", "excel")
LayoutMode = Literal["auto", "layout", "tables"]  # lo que elige el usuario ("auto" = decide la app)
EffectiveMode = Literal["layout", "tables"]  # lo que se usa de verdad al exportar
HAlign = Literal["left", "center", "right"]
VAlign = Literal["top", "center", "bottom"]
Bbox = tuple[float, float, float, float]  # (x0, y0, x1, y1) en puntos PDF, origen arriba a la izquierda


@dataclass
class PageInfo:
    index: int  # 0-based
    kind: PageKind
    n_chars: int
    image_coverage: float  # 0..1 fracción del área cubierta por imágenes
    rotation: int = 0  # 0/90/180/270 detectado por OSD (solo escaneos)


@dataclass
class RawTable:
    page: int  # primera página (0-based)
    page_end: int  # última página (== page si no se mergeó)
    bbox: tuple[float, float, float, float]
    rows: list[list[str]]  # rows[0] = encabezado; celdas ya limpias (clean_cell)
    source: str  # "hybrid_union" | "hybrid"
    quality: float  # table_quality(rows)
    title: str | None = None  # título detectado arriba de la tabla (title_near)


@dataclass
class ColumnInfo:
    kind: ColumnKind
    locale: Locale | None = None  # solo kind == "number"
    # openpyxl: "#,##0.00" | "#,##0" | '"$" #,##0.00' | "0.00%" | "dd/mm/yyyy"
    number_format: str | None = None
    header: str = ""


@dataclass
class TypedTable:
    name: str  # nombre de hoja propuesto (≤ 31 chars, válido)
    header: list[str]
    rows: list[list[CellValue]]  # SIN el encabezado
    columns: list[ColumnInfo]
    review_cells: set[tuple[int, int]] = field(default_factory=set)  # (fila_0based_en_rows, col)
    total_rows: set[int] = field(default_factory=set)  # filas Subtotal/IVA/TOTAL (negrita)
    pages: tuple[int, int] = (0, 0)  # (primera, última) 0-based
    source: str = ""
    quality: float = 0.0


@dataclass
class TextItem:
    page: int
    kind: Literal["kv", "text"]
    key: str | None
    value: str


# ---------------------------------------------------------------------------------------------------
# Modo "Igual al PDF" (docs/LAYOUT_MODE.md): una página PDF → una hoja con celdas posicionadas
# ---------------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class BorderSide:
    width_pt: float  # grosor en el PDF; el writer lo traduce a thin/medium/thick
    color: str = "000000"  # "RRGGBB"


@dataclass
class TextRun:
    """Un tramo de texto con un solo estilo (una celda "ETIQUETA / Valor" tiene runs de estilos distintos)."""

    text: str
    bold: bool = False
    italic: bool = False
    size_pt: float = 8.0  # tamaño en el PDF (el writer lo traduce a pt de Excel, mínimo legible)
    color: str | None = None  # "RRGGBB"; None = automático
    font: str | None = None  # familia sugerida ("Arial", "Times New Roman", "Courier New"…); None = Calibri


@dataclass
class LayoutCell:
    """Una celda (o rango combinado) de la hoja. ``row``/``col`` son 0-based y apuntan a la esquina
    superior izquierda; ``row_span``/``col_span`` ≥ 1 son los tamaños del rango combinado."""

    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    # Cada elemento es una línea visual (saltos de línea en Excel); [] = caja vacía (campo sin completar).
    lines: list[list[TextRun]] = field(default_factory=list)
    halign: HAlign = "left"
    valign: VAlign = "center"
    wrap: bool = True
    fill: str | None = None  # "RRGGBB"; None = sin relleno
    border_left: BorderSide | None = None
    border_top: BorderSide | None = None
    border_right: BorderSide | None = None
    border_bottom: BorderSide | None = None
    # Si no es None se escribe este valor tipado (número/fecha) en vez del texto.
    value: CellValue = None
    number_format: str | None = None
    bbox: Bbox = (0.0, 0.0, 0.0, 0.0)  # rectángulo de la celda en puntos PDF (preview y diagnóstico)

    @property
    def text(self) -> str:
        """Texto plano de la celda (líneas unidas con ``\n``)."""
        return "\n".join("".join(run.text for run in line) for line in self.lines)


@dataclass
class LayoutImage:
    """Imagen (logo, mapa, foto) anclada a una celda con un desplazamiento en puntos."""

    data: bytes  # PNG o JPEG ya recomprimido (lado mayor ≤ 1600 px)
    format: Literal["png", "jpeg"]
    bbox: Bbox  # posición real en la página, en puntos PDF
    row: int  # celda ancla (0-based)
    col: int
    dx_pt: float = 0.0  # desplazamiento desde la esquina superior izquierda de la celda ancla
    dy_pt: float = 0.0
    width_pt: float = 0.0
    height_pt: float = 0.0


@dataclass
class LayoutPage:
    page: int  # 0-based
    width_pt: float
    height_pt: float
    col_widths_pt: list[float]  # ancho de cada columna de la grilla, en puntos PDF
    row_heights_pt: list[float]  # alto de cada fila de la grilla, en puntos PDF
    cells: list[LayoutCell] = field(default_factory=list)
    images: list[LayoutImage] = field(default_factory=list)
    hidden_texts: list[str] = field(default_factory=list)  # textos descartados por estar tapados/invisibles
    n_words: int = 0  # palabras visibles de la página (métrica de fidelidad)
    warnings: list[str] = field(default_factory=list)

    @property
    def landscape(self) -> bool:
        return self.width_pt > self.height_pt


@dataclass
class Options:
    typed_values: bool = True  # única opción visible: "Convertir números y fechas (recomendado)"
    merge_continuations: bool = True
    include_text: bool = True
    one_sheet_per_page: bool = False
    force_ocr: bool = False
    ocr_lang: str = "spa+eng"
    output_dir: Path | None = None  # None = junto al PDF
    open_when_done: bool = True
    show_origin_sheet: bool = False
    single_workbook: bool = False  # un solo libro para todo el lote
    # "Igual al PDF" (una hoja por página con cajas, bordes, colores e imágenes en su lugar) o "Solo tablas".
    # "auto": la app decide por documento (ConversionReport.suggested_mode); el usuario puede cambiarlo.
    layout: LayoutMode = "auto"
    layout_images: bool = True  # incluir logos/mapas/fotos en el modo "Igual al PDF"


@dataclass
class ConversionReport:
    source: Path
    pages: list[PageInfo]
    tables: list[TypedTable]
    text_items: list[TextItem]
    warnings: list[str] = field(default_factory=list)
    timings: dict[str, float] = field(default_factory=dict)
    ocr_applied: bool = False
    output_path: Path | None = None  # lo completa export()
    app_version: str = ""
    # Modo "Igual al PDF": una LayoutPage por página con contenido (vacío si el documento no la calculó).
    layout_pages: list[LayoutPage] = field(default_factory=list)
    suggested_mode: EffectiveMode = "tables"  # lo que recomienda la app para ESTE documento
    mode: EffectiveMode = "tables"  # lo que usa export(): Options.layout, o suggested_mode si es "auto"
    mode_reason: str = ""  # una frase en castellano que explica la sugerencia (la UI la muestra)

    @property
    def n_rows(self) -> int:
        """Total de filas de datos (sin encabezados) en todas las tablas."""
        return sum(len(t.rows) for t in self.tables)

    @property
    def n_review(self) -> int:
        """Total de celdas marcadas para revisar en todas las tablas."""
        return sum(len(t.review_cells) for t in self.tables)


__all__ = [
    "STAGES",
    "Bbox",
    "BorderSide",
    "CellValue",
    "ColumnInfo",
    "ColumnKind",
    "ConversionReport",
    "EffectiveMode",
    "HAlign",
    "LayoutCell",
    "LayoutImage",
    "LayoutMode",
    "LayoutPage",
    "Locale",
    "Options",
    "PageInfo",
    "PageKind",
    "ProgressFn",
    "RawTable",
    "TextItem",
    "TextRun",
    "TypedTable",
    "VAlign",
]
