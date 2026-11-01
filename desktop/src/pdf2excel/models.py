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
STAGES = ("classify", "ocr", "tables", "typing", "excel")


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
    "CellValue",
    "ColumnInfo",
    "ColumnKind",
    "ConversionReport",
    "Locale",
    "Options",
    "PageInfo",
    "PageKind",
    "ProgressFn",
    "RawTable",
    "TextItem",
    "TypedTable",
]
