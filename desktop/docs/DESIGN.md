# Conversor PDF a Excel — Diseño final (síntesis)

Aplicación de escritorio para un equipo no técnico (Argentina) que convierte **cualquier PDF** (remitos,
facturas, listas de precios, reportes de stock, extractos bancarios; con y sin bordes; multipágina;
escaneados y mixtos) a un Excel **prolijo**: tablas con tipos reales (números y fechas argentinas como
números y fechas), encabezados fijos, filtros, nada inventado y todo lo dudoso marcado para revisar.

Este documento sintetiza tres propuestas evaluadas por tres jueces. Parte de la propuesta mejor puntuada
(P0, calidad de extracción) e injerta la UI de P1 y el empaquetado/CI de P2. Las contradicciones se
resolvieron con experimentos reproducibles (sección 11).

---

## 0. Decisiones en una página

| Tema | Decisión | Origen | Evidencia |
|---|---|---|---|
| Geometría de tablas | `pymupdf 2.0` `find_tables(strategy="lines")` **con** `pymupdf-layout` (modelo ONNX) | P0/P1 | Sin el modelo: 0 tablas en remito sin bordes, extracto bancario y todos los escaneados (3 jueces lo verificaron) |
| Texto de celdas | Palabras (`get_text("words")`) asignadas a la celda por su centro, agrupadas en renglones por centro-y (tolerancia 45 % de la altura). **Nunca** `Table.extract()` | P0 | `extract()` rompe "350.00" → "35000\n."; la agrupación por centro-y subió el OCR de 96,2 a 98,1 % |
| Candidatos por página | Solo **2**: `hybrid_union` (`union=True`, documento fresco) y `hybrid` (`union=False`). **Sin camelot** | P0 recortado por los 3 jueces | Sin camelot las métricas son idénticas: 100/100/98,3/98,1/93,3 (sección 11.a). Ahorra pandas+opencv-full+pypdfium2 ≈ 170 MB |
| Selección | `table_quality()` sin ground truth + `0,7·calidad + 0,3·cobertura` | P0 | Acierta en 11/11 PDFs del corpus |
| Totales sin grilla / columnas fantasma | `extend_tail` + `merge_ghost_columns` | P0 | 95,3 → 100 % (factura) y 77 → 98 % (factura OCR) |
| Tipado AR | Por **columna**, no por celda: identificadores por encabezado → texto; fecha ≥ 90 %; número ≥ 50 %; empate AR/US → AR; negativos `(1.234,56)` y `1.234,56-`; "21%" → 0,21; "1.234" solo no vota | P0 + P2 + P1 | test_typing_ar; caso "Stock = 24/0/112/s/d" exige umbral 50 % |
| OCR | **Propio, sin Ghostscript ni ocrmypdf**: render 300 dpi → deskew por perfil de proyección → quitar líneas de grilla (OpenCV) → `tesseract` CLI `--psm 4 -c textonly_pdf=1` → capa de texto invisible sobre la imagen → mismo pipeline de tablas. Orientación con `--psm 0` (osd) | P1/P2 corregido con el deskew de P0 | **98,8 %** factura escaneada (ocrmypdf 98,1 %), 91,7 % lista mixta (ocrmypdf 93,3 %), 1,4–1,9 s/pág vs ~3 s/pág (sección 11.b) |
| Tesseract en Windows | `tesseract.exe` (build UB-Mannheim 5.3.x) vendoreado en `resources/tesseract/` + `tessdata_fast` spa+eng+osd | P1 | El Tesseract embebido en MuPDF tarda 3,5–10 s/pág en A4 y da 0 tablas en escaneos difíciles (jueces 0 y 2) |
| UI | Flujo de 4 pantallas: *Arrastrá* → *Esto encontré* (vista previa tipada) → *Convirtiendo* → *Listo*. Una acción primaria por pantalla, una sola opción visible, engranaje para lo avanzado | P1 | Captura `4_revision.png` con análisis real; 3 jueces la prefieren |
| Errores | `UserError(code, title, detail, action)` mapeado desde excepciones técnicas; lote nunca se frena; nunca se sobrescribe (`nombre (2).xlsx`) | P1/P2 | — |
| Concurrencia | `QThreadPool` + `QRunnable` + señales; cancelación cooperativa por `threading.Event`. **Sin** `ProcessPool` | P0/P2 (3 jueces) | Gap máximo del event loop 67–82 ms medido; ProcessPool en PyInstaller/Windows es la zona de más bugs |
| Empaquetado | PyInstaller 6.22 **onedir** + `--collect-data pymupdf` + `prune_qt.py`; `PySide6-Essentials`; Inno Setup por usuario sin UAC con "Enviar a"; zip portable; firma condicional; `--selftest` del bundle en CI | P2 + fix de P1 | .spec, prune, iss y workflow ya probados en Linux; onefile descartado (3–3,6 s de arranque) |
| Tests | Corpus sintético con ground truth + scorer con umbrales que rompen el build; métrica que también penaliza filas/columnas espurias | P0 ampliado | — |
| Licencia | `desktop/` bajo **AGPL-3.0** (PyMuPDF es AGPL; el repo es público). Sin Ghostscript → una sola AGPL | P2 | — |

---

## 1. Objetivo, alcance y no-objetivos

**Objetivo**: que una persona sin conocimientos técnicos arrastre uno o varios PDF, vea qué se encontró y
obtenga un `.xlsx` que se abre en Excel con tablas limpias y tipadas. Windows primero (instalador y
portable); macOS/Linux con el mismo código.

**Alcance v1**: PDFs digitales (con y sin bordes, multipágina), escaneados (OCR en castellano e inglés),
mixtos; texto fuera de tablas; lotes; cancelación; vista previa; errores en castellano llano.

**No-objetivos v1**: editar la tabla dentro de la app (eso es Excel), fórmulas, auto-updater completo
(solo aviso de versión nueva), fotos de celular con perspectiva fuerte (se convierten pero se marcan).

---

## 2. Stack y justificación

| Componente | Paquete / versión fijada | Por qué |
|---|---|---|
| Lenguaje | Python 3.13 (3.12 si falta algún wheel en Windows) | Todo lo de abajo está verificado en 3.13 en esta máquina |
| UI | `PySide6-Essentials==6.12.*` (QtWidgets; nada de QML/WebEngine) | Look nativo, drag&drop del explorador, QFileDialog nativo, hilos maduros, offscreen para tests, hooks oficiales de PyInstaller. Essentials evita WebEngine/Charts/3D (−150 MB) |
| PDF + geometría | `pymupdf==2.0`, `pymupdf-layout==2.0`, `onnxruntime==1.31.1` | Único motor que detecta tablas sin bordes y sobre capa OCR (modelo ONNX ≈ 30 MB, 0,4–0,7 s/pág) |
| Preprocesado OCR | `opencv-python-headless==5.0.*`, `numpy` | Otsu, deskew, quitar líneas de grilla (sin esto tesseract devuelve "COCINE CI CIC") |
| OCR | `tesseract` 5.3.x CLI (vendoreado en Windows) + `tessdata_fast` spa, eng, osd | Medido 98,8 % en escaneo difícil, 1,4–1,9 s/pág; `--psm 0` detecta 0/90/180/270 |
| Excel | `openpyxl==3.1.5` | Tipos reales, `number_format`, estilos, freeze panes, autofiltro; sin pandas |
| Soporte | `platformdirs` (config, logs, caché OCR) | Rutas correctas en los 3 SO |
| Empaquetado | `pyinstaller==6.22.*`, Inno Setup 6, 7-Zip | Ver sección 8 |
| Dev | `pytest`, `pytest-qt`, `reportlab` (corpus), `ruff`, `mypy`, `pillow` | — |

**Descartados y por qué**: `camelot-py` (nunca ganó en el selector; +170 MB), `tabula-py` (Java),
`pdfplumber` en runtime (`extract_tables(text)` 15–45 % de celdas; `get_text("words")` de pymupdf cubre
el texto), `ocrmypdf` + Ghostscript (misma calidad que el OCR propio pero 2× más lento, +55 MB, segunda
AGPL, usa `ProcessPoolExecutor` internamente), Tesseract embebido en MuPDF (sin `psm`, sin confianza,
3,5–10 s/pág en A4, "Image too small to scale"), `rapidocr` (5 s/pág, pierde dígitos sueltos y acentos),
tkinter (no está en este Python; sin drag&drop nativo), Flet/Tauri/Electron (segundo toolchain, sin drop
nativo o 600 MB), PyInstaller onefile (3–3,6 s de arranque, antivirus), Nuitka (fase 2 si el tamaño molesta).

---

## 3. Arquitectura

```
desktop/
  pyproject.toml                 # name="pdf2excel"; deps fijadas; extras [dev]; script "conversor-pdf-a-excel"
  README.md                      # para el equipo (instalación, uso, qué hacer si algo falla)
  LICENSE                        # AGPL-3.0
  docs/DESIGN.md                 # este documento
  src/pdf2excel/
    __init__.py                  # __version__ = "0.1.0" (única fuente de versión)
    __main__.py                  # python -m pdf2excel → app.main() ; con argumentos → cli.main()
    app.py                       # QApplication, freeze_support(), High-DPI, QLocale("es_AR"), qtbase_es, theme.qss, MainWindow
    cli.py                       # `pdf2excel archivo.pdf [...]` y `--selftest <pdf>` (CI). Sin Qt.
    models.py                    # dataclasses compartidas (sección 4). Sin lógica.
    errors.py                    # UserError + from_exception()
    resources_path.py            # resource_dir(), tessdata_dir(), tesseract_dir()  (sys._MEIPASS si frozen)
    logging_setup.py             # log rotativo en platformdirs.user_log_dir("ConversorPDFaExcel")
    update_check.py              # GET latest.json en un hilo → UpdateInfo | None
    core/                        # NO importa Qt. Testeable con pytest puro; lo usa la CLI.
      document.py                # open_document(), classify_page(), needs_ocr()
      ocr.py                     # find_tesseract(), ocr_available(), preprocess(), ocr_document()
      cells.py                   # words_in_rect(), clean_cell(), extend_tail(), merge_ghost_columns()
      candidates.py              # page_candidates()  → {"hybrid_union": [...], "hybrid": [...]}
      quality.py                 # table_quality(), select_best()
      merge.py                   # merge_multipage()
      typing_ar.py               # infer_column_locale(), parse_number(), parse_date(), parse_percent(), type_table()
      text_blocks.py             # text_outside_tables(), title_near()
      pipeline.py                # analyze(), export(), convert()  (orquestación)
    excel/
      styles.py                  # colores, fuentes, number formats
      naming.py                  # unique_output_path(), safe_sheet_name()
      writer.py                  # write_workbook()
    ui/
      strings_es.py              # TODOS los textos visibles (rioplatense, "vos")
      theme.qss                  # colores/acento; sigue el modo oscuro del sistema
      drop_zone.py               # DropZone(QFrame)  files_dropped = Signal(list)
      preview_model.py           # TablePreviewModel(QAbstractTableModel) solo lectura
      file_list.py               # FileListWidget (miniatura + nombre + estado)
      pages/drop_page.py         # Pantalla 1
      pages/review_page.py       # Pantalla 2 "Esto encontré"
      pages/progress_page.py     # Pantalla 3
      pages/done_page.py         # Pantalla 4
      settings_dialog.py         # engranaje: opciones avanzadas
      main_window.py             # QMainWindow + QStackedWidget; solo coordina señales
    jobs/
      worker.py                  # AnalyzeJob / ExportJob (QRunnable) + JobSignals; cancel por threading.Event
  resources/
    icons/app.ico app.icns app.png app.svg
    tessdata/                    # spa, eng, osd (tessdata_fast) — lo baja packaging/fetch_tesseract.py; no se commitea
    tesseract/                   # Windows: tesseract.exe + DLLs (UB-Mannheim) — ídem
  packaging/
    conversor.spec  prune_qt.py  installer.iss  file_version_info.txt  stamp_version.py  fetch_tesseract.py
  tests/
    corpus/generate.py  corpus/ground_truth.json  corpus/*.pdf (generados en CI)
    scorer.py  conftest.py  test_*.py
  ../.github/workflows/build-desktop.yml
```

**Flujo de datos**

```
drop/elegir ──► AnalyzeJob (QThreadPool) ──► core.pipeline.analyze(path, opts, progress, cancel)
                                              │ open_document → classify_page (por página)
                                              │ si alguna página necesita OCR → core.ocr.ocr_document → PDF temporal cacheado
                                              │ por página: words + page_candidates → select_best
                                              │ merge_multipage (todo el documento)
                                              │ type_table (por tabla) → text_outside_tables (por página)
                                              └► ConversionReport (en memoria, sin escribir)
ReviewPage muestra tablas tipadas ──► "Guardar Excel" ──► ExportJob ──► core.pipeline.export(report, opts)
                                                                        │ excel.naming.unique_output_path
                                                                        └► excel.writer.write_workbook → DonePage
```

**Concurrencia**: UI en el hilo principal; cada archivo es un `QRunnable` en un `QThreadPool` con
`maxThreadCount=2` (pymupdf/onnx/opencv liberan el GIL en buena parte; tesseract es un subproceso). Las
señales Qt cruzan hilos de forma segura (`Qt.QueuedConnection`). Cancelación: `cancel.is_set()` entre
páginas y `Popen.kill()` del tesseract en curso. La UI nunca llama a `core` directamente. Medido en los
prototipos: gap máximo del event loop 67–82 ms durante extracción real. `multiprocessing.freeze_support()`
va primero en `main()` por si en el futuro se usan procesos.

---

## 4. Modelos e interfaces compartidas (contrato exacto)

`src/pdf2excel/models.py` (sin lógica; todo lo demás importa de acá):

```python
from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, Literal

PageKind = Literal["digital", "scanned", "scanned_with_text", "garbage_text", "empty"]
ColumnKind = Literal["number", "date", "percent", "id", "text"]
Locale = Literal["ar", "us"]
CellValue = int | float | date | str | None
ProgressFn = Callable[[int, int, str], None]      # (actual, total, etapa) etapa ∈ STAGES
STAGES = ("classify", "ocr", "tables", "typing", "excel")

@dataclass
class PageInfo:
    index: int                   # 0-based
    kind: PageKind
    n_chars: int
    image_coverage: float        # 0..1 fracción del área cubierta por imágenes
    rotation: int = 0            # 0/90/180/270 detectado por OSD (solo escaneos)

@dataclass
class RawTable:
    page: int                    # primera página (0-based)
    page_end: int                # última página (== page si no se mergeó)
    bbox: tuple[float, float, float, float]
    rows: list[list[str]]        # rows[0] = encabezado; celdas ya limpias (clean_cell)
    source: str                  # "hybrid_union" | "hybrid"
    quality: float               # table_quality(rows)
    title: str | None = None     # título detectado arriba de la tabla (title_near)

@dataclass
class ColumnInfo:
    kind: ColumnKind
    locale: Locale | None = None         # solo kind == "number"
    number_format: str | None = None     # openpyxl: "#,##0.00" | "#,##0" | '"$" #,##0.00' | "0.00%" | "dd/mm/yyyy"
    header: str = ""

@dataclass
class TypedTable:
    name: str                            # nombre de hoja propuesto (≤ 31 chars, válido)
    header: list[str]
    rows: list[list[CellValue]]          # SIN el encabezado
    columns: list[ColumnInfo]
    review_cells: set[tuple[int, int]] = field(default_factory=set)   # (fila_0based_en_rows, col)
    total_rows: set[int] = field(default_factory=set)                 # filas Subtotal/IVA/TOTAL (negrita)
    pages: tuple[int, int] = (0, 0)      # (primera, última) 0-based
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
    typed_values: bool = True            # única opción visible: "Convertir números y fechas (recomendado)"
    merge_continuations: bool = True
    include_text: bool = True
    one_sheet_per_page: bool = False
    force_ocr: bool = False
    ocr_lang: str = "spa+eng"
    output_dir: Path | None = None       # None = junto al PDF
    open_when_done: bool = True
    show_origin_sheet: bool = False

@dataclass
class ConversionReport:
    source: Path
    pages: list[PageInfo]
    tables: list[TypedTable]
    text_items: list[TextItem]
    warnings: list[str] = field(default_factory=list)
    timings: dict[str, float] = field(default_factory=dict)
    ocr_applied: bool = False
    output_path: Path | None = None      # lo completa export()
    app_version: str = ""
```

`src/pdf2excel/errors.py`:

```python
ErrorCode = Literal["encrypted", "damaged", "not_pdf", "no_permission", "file_in_use",
                    "ocr_unavailable", "ocr_failed", "cancelled", "unknown"]

class UserError(Exception):
    def __init__(self, code: ErrorCode, title: str, detail: str = "", action: str | None = None,
                 cause: BaseException | None = None): ...
    # action ∈ {"retry", "save_as", "choose_folder", "password", "skip", None}

def from_exception(exc: BaseException, path: Path | None = None) -> UserError:
    """pymupdf.FileDataError → damaged | doc.needs_pass → encrypted | PermissionError sobre .xlsx → file_in_use
       PermissionError sobre carpeta → no_permission | FileNotFoundError(tesseract) → ocr_unavailable | resto → unknown"""
```

Los títulos/detalles en castellano viven en `ui/strings_es.py` (`ERRORS[code] = (title, detail, action_label)`);
`UserError.title/detail` llevan un texto por defecto en castellano para la CLI.

**Firmas públicas del core** (las que usan pipeline, UI y tests):

```python
# core/document.py
def open_document(path: Path, password: str | None = None) -> pymupdf.Document   # UserError encrypted/damaged/not_pdf
def classify_page(page: pymupdf.Page) -> PageInfo
def needs_ocr(info: PageInfo, force: bool = False) -> bool

# core/ocr.py
def find_tesseract() -> Path | None            # resources/tesseract/tesseract(.exe) si frozen → PATH
def ocr_available() -> bool
def deskew_angle(gray: np.ndarray) -> float    # grados, perfil de proyección, ±4°, paso 0,1
def remove_grid_lines(gray: np.ndarray) -> np.ndarray
def detect_orientation(gray: np.ndarray) -> int   # 0/90/180/270 via tesseract --psm 0 (osd); 0 si falla
def ocr_document(src: Path, pages: list[int], lang: str, progress: ProgressFn, cancel: threading.Event) -> Path
    # devuelve PDF temporal (caché platformdirs.user_cache_dir/ocr/<sha1>.pdf) con TODAS las páginas:
    # las de `pages` reemplazadas por imagen deskewed + capa de texto invisible; las demás copiadas.

# core/cells.py
def words_in_rect(words: list[tuple], rect) -> str
def clean_cell(s: str | None) -> str
def extend_tail(page, table, rows: list[list[str]], words, max_gap_pt: float = 20) -> int
def merge_ghost_columns(rows: list[list[str]]) -> list[list[str]]

# core/candidates.py
def page_candidates(path: Path, page_no: int) -> dict[str, list[RawTable]]   # {"hybrid_union": [...], "hybrid": [...]}

# core/quality.py
def table_quality(rows: list[list[str]]) -> float
def select_best(cands: dict[str, list[RawTable]]) -> tuple[str, list[RawTable]]

# core/merge.py
def merge_multipage(tables: list[RawTable]) -> list[RawTable]

# core/typing_ar.py
def infer_column_locale(values: list[str]) -> Locale | None
def parse_number(s: str, locale: Locale) -> int | float | None
def parse_percent(s: str, locale: Locale) -> float | None
def parse_date(s: str) -> date | None
def is_identifier_header(header: str) -> bool
def type_table(raw: RawTable, typed_values: bool = True) -> TypedTable

# core/text_blocks.py
def text_outside_tables(page, table_bboxes: list[tuple]) -> list[TextItem]
def title_near(page, bbox: tuple, max_gap_pt: float = 40) -> str | None

# core/pipeline.py
def analyze(src: Path, opts: Options, progress: ProgressFn | None = None,
            cancel: threading.Event | None = None, password: str | None = None) -> ConversionReport
def export(report: ConversionReport, opts: Options) -> Path
def convert(src: Path, opts: Options, progress=None, cancel=None, password=None) -> ConversionReport

# excel/naming.py
def unique_output_path(src: Path, output_dir: Path | None) -> Path     # nombre.xlsx, nombre (2).xlsx, ...
def safe_sheet_name(title: str | None, index: int, used: set[str]) -> str

# excel/writer.py
def write_workbook(report: ConversionReport, dst: Path, opts: Options) -> None
```

---

## 5. Pipeline de extracción

### 5.1 Clasificación de página (`classify_page`)
- `n_chars = len(page.get_text().strip())`; `image_coverage` = unión de los rects de `page.get_image_info()` / área de la página.
- Reglas: `n_chars < 20 and image_coverage > 0.5` → **scanned**; `image_coverage > 0.5 and n_chars ≥ 20` →
  **scanned_with_text** (ya tiene OCR; se usa tal cual); `n_chars == 0 and image_coverage == 0` → **empty**;
  texto con > 30 % de caracteres fuera de `[A-Za-z0-9áéíóúñÁÉÍÓÚÑüÜ$%.,;:/()\-–—°º#" \n]` (fuentes sin
  ToUnicode) → **garbage_text** (se trata como escaneo); resto → **digital**.
- `needs_ocr(info, force)` = `force or info.kind in ("scanned", "garbage_text")`.
- Verificado: `remito_grid_scan` → scanned, `*_ocr` → scanned_with_text, vectoriales → digital. 3–30 ms/pág.

### 5.2 OCR (`ocr_document`) — solo si hace falta, una vez por documento
Por cada página que necesita OCR:
1. `page.get_pixmap(dpi=300, colorspace=csGRAY)` → `np.ndarray`.
2. **Orientación**: `detect_orientation` (`tesseract --psm 0 -l osd` sobre un render a 150 dpi) → `np.rot90`.
3. **Deskew**: `deskew_angle` maximiza la varianza del perfil horizontal de la imagen binarizada (Otsu) a 1/4
   de escala, ángulos −4°…+4° en pasos de 0,1°; si |θ| ≥ 0,2° se rota con `warpAffine` (INTER_CUBIC,
   borde blanco). Es el paso que faltaba en las réplicas de los jueces (22–28 % sin deskew).
4. **Imagen para OCR**: `remove_grid_lines` = Otsu invertido → apertura morfológica con kernels
   `(k,1)` y `(1,k)`, `k = max(40, ancho//30)` → dilatar 3×3 → pintar blanco. Sin esto tesseract devuelve
   basura en tablas con bordes ("favor CECI CO ESA"). La imagen que se guarda en el PDF es la deskewed
   **con** líneas (el layout ML las aprovecha).
5. `tesseract img out -l spa+eng --psm 4 --dpi 300 -c textonly_pdf=1 pdf` (subproceso, `CREATE_NO_WINDOW`
   en Windows, `TESSDATA_PREFIX` apuntando a `resources/tessdata`). psm 4 ganó a 6 y 3 (98,8 vs 98,4 vs 95,0 %).
6. Página nueva en el documento de salida: `insert_image` (JPEG q80 de la imagen deskewed) +
   `show_pdf_page(rect, pdf_textonly, 0)` → capa de texto invisible alineada.
7. Páginas que no necesitan OCR se copian con `insert_pdf`. Salida cacheada por sha1 del archivo origen en
   `user_cache_dir("ConversorPDFaExcel")/ocr/`. Progreso: `("ocr", i, n)`. Cancelación: `kill()` del subproceso.
- Medido: 1,4–1,9 s/pág (4 CPUs). Si `find_tesseract()` es `None`: warning `ocr_unavailable`, se convierten
  las páginas con texto y la UI avisa (no bloquea).

### 5.3 Candidatos por página (`page_candidates`)
Para cada página (del PDF original o del OCR'd):
- `words = page.get_text("words")`.
- **a. hybrid_union**: `doc_fresh = pymupdf.open(path)` (documento aparte: `union=True` sobre una página
  cuyo layout ya se calculó devuelve 0 tablas — bug documentado) → `page.find_tables(strategy="lines", union=True)`.
  `union` fusiona cajas del layout ML con líneas: arregla columnas fantasma del OCR (5×11 → 5×6).
- **b. hybrid**: `find_tables(strategy="lines", union=False)` (en el extracto multipágina `union` partió la
  tabla; por eso van los dos).
- Para cada tabla: `rows = [[clean_cell(words_in_rect(words, cell)) for cell in row.cells] for row in table.rows]`
  → `extend_tail` → descartar filas vacías → `merge_ghost_columns` → descartar columnas vacías →
  `RawTable(..., quality=table_quality(rows), title=title_near(page, bbox))`.
- `words_in_rect`: palabra pertenece a la celda si su centro cae en el rect; renglones por centro-y con
  tolerancia 45 % de la altura; ordenadas por x; wrap → un renglón con espacio.
- `clean_cell`: quita tokens que son solo `|¦!` (barras alucinadas por OCR), NBSP → espacio, colapsa
  espacios, quita guiones de corte a fin de línea, `strip(" |")`.
- `extend_tail`: renglones debajo del bbox con gap < 20 pt cuyas palabras caen todas dentro de los rangos x de
  las columnas y con ≥ 2 celdas no vacías → se agregan (Subtotal / IVA / TOTAL con SPAN y sin grilla).
  *Cambio en integración:* además, al menos una celda del renglón debe parsear como número, porcentaje o
  fecha (`typing_ar`); si no, se corta. Motivo: en `hard_lista_mixta` absorbía el párrafo de prosa que sigue
  a la tabla (8,3 % de filas espurias → 0 %); los totales reales siempre traen un importe.
- `merge_ghost_columns`: columnas adyacentes mutuamente excluyentes donde una tiene ≤ 2 celdas → se unen.

### 5.4 Puntaje y selección (`table_quality`, `select_best`)
`table_quality(rows)` ∈ [0,1], sin ground truth, sobre el cuerpo (sin encabezado):
```
q = 0,05·(1−vacías) + 0,15·encabezado_ok + 0,35·tipado + 0,15·(1−barras) + 0,10·(1−largas) + 0,20·(1−rotas) − 0,15·fantasma
```
- `encabezado_ok`: fracción de celdas de la fila 1 no vacías y no numéricas.
- `tipado`: fracción de columnas "tipables" (≥ 50 % de celdas que parsean como número con el locale de la
  columna, o todas fechas) ponderada por su tasa de parseo; 0,5 si no hay columnas candidatas.
- `barras`: celdas que son solo `|.-`; `largas`: celdas > 60 chars; `rotas`: filas con ≤ 1 celda llena;
  `fantasma`: pares de columnas adyacentes mutuamente excluyentes.
- Lección medida: penalizar fuerte las celdas vacías hacía perder los totales; Débito/Crédito vacías son legítimas.

`select_best(cands)`: por candidato, `calidad = Σ q·filas / Σ filas`, `cobertura = celdas no vacías / máximo
entre candidatos`; clave `0,7·calidad + 0,3·cobertura`; sin tablas → −1. Devuelve `(nombre, tablas)`.
Test `test_quality_selection` verifica que el elegido es el de mayor celdas % contra GT (hoy 11/11).

### 5.5 Merge multipágina (`merge_multipage`)
Tablas de páginas consecutivas con el mismo número de columnas y:
- **Regla 1**: encabezado idéntico (case-insensitive, espacios colapsados) → se concatena descartando el
  encabezado repetido (factura 35+27 → 64 filas; extracto 52+21 → 72).
- **Regla 2** (continuación sin encabezado): límites x de columnas coinciden ±3 pt, la primera fila de la
  segunda tabla no parece encabezado (≥ 50 % de celdas numéricas/fechas) → se concatena entera.
- Filas "Transporte"/"Subtotal transporte" al final/inicio se marcan en `total_rows`, no se duplican.
- Con `opts.merge_continuations=False` o `one_sheet_per_page=True` no se mergea.

### 5.6 Tipado argentino por columna (`type_table`)
Orden por columna (sobre celdas no vacías del cuerpo):
1. **Identificador** → `kind="id"`, texto siempre: `is_identifier_header` matchea
   `comprobante|c[oó]digo|cuit|cuil|cbu|dni|n[°ºo]\b|nro|referencia|\bid\b|art[ií]culo|sku|factura|remito`,
   o todas las celdas son dígitos de ≥ 5 cifras con la misma longitud, o patrón CUIT `\d{2}-\d{8}-\d`.
   Códigos tipo "A-100", "0001-00004521" nunca pierden ceros ni guiones.
2. **Fecha** si ≥ 90 % parsea `dd/mm/aaaa`, `dd-mm-aa`, `dd.mm.aaaa` (año de 2 dígitos: 00–69 → 20xx,
   70–99 → 19xx; `32/13/2026` no parsea). Tolera "Saldo final:" suelto en la columna.
3. **Número**: `infer_column_locale` sobre toda la columna. Evidencia AR: termina en `,dd` o usa `.` como
   miles junto con `,`; evidencia US: termina en `.dd` sin `,ddd`. "1.234" solo **no vota**; sin
   separadores ("21", "1250") no vota pero parsea en cualquier locale. Empate o sin evidencia → **AR**.
   La columna es numérica si ≥ 50 % de sus celdas parsean con ese locale (caso Stock = 24/0/112/s/d →
   3 números + 1 texto). Se aceptan `$`, `ARS`, `U$S`, `USD`, espacios, `+`/`−`, negativos entre
   paréntesis `(1.234,56)` y con signo al final `1.234,56-` (extractos). Si ≥ 50 % de las celdas traen
   símbolo de moneda → `number_format='"$" #,##0.00'`; si todas son enteras → `"#,##0"`; si no → `"#,##0.00"`.
4. **Porcentaje**: `^-?\d+([.,]\d+)?\s?%$` → fracción (`21%` → 0,21), `number_format="0.00%"`.
5. Lo demás → `kind="text"`.
Cada celda se parsea con el tipo de su columna; **lo que no parsea queda como texto y va a `review_cells`**.
Nunca se inventa un número. Con `typed_values=False` todo queda texto (sin review).
Mejora medible en backlog: reparación de formato minoritario ("977.50" con 2 decimales en columna AR → 977,50)
y corrección de patrones OCR en columnas de código ("A131" → "A-131" si la mayoría es `A-ddd`).

### 5.7 Texto fuera de tablas (`text_outside_tables`, `title_near`)
- Bloques de `page.get_text("blocks")` que no se solapan > 50 % con un bbox de tabla.
- Una línea `"Cliente: Bar El Catorce — Fecha: 15/10/2026"` se parte por ` — ` / ` | ` y por `clave: valor`
  en `TextItem(kind="kv")`; el resto `kind="text"`.
- `title_near`: el bloque de una sola línea, ≤ 60 chars, inmediatamente arriba del bbox (gap ≤ 40 pt) que no
  es un `kv` → título de la tabla ("1. Bebidas con alcohol").
- Encabezados/pies repetidos en ≥ 70 % de las páginas en la misma posición se emiten una sola vez.

### 5.8 Métricas actuales (selección automática, sin camelot, contra ground truth)
| PDF | celdas % | filas exactas % | tiempo |
|---|---|---|---|
| hard_factura_ar (2 pág, totales sin grilla, wrap) | 100 | 100 | 1,4 s |
| hard_extracto_banco (2 pág, sin bordes, negativos) | 100 | 100 | 1,5 s |
| hard_lista_mixta (2 tablas en una página) | 98,3 (único mismatch GT `None` vs "s/d") | 90,9 | 0,5 s |
| hard_factura_ar_scan → OCR propio | **98,8** | 93,8 | 3,8 s OCR + 1,6 s |
| hard_lista_mixta_scan → OCR propio | 91,7 | 54,5 | 1,4 s OCR + 0,6 s |
| remito_grid / nogrid / *_scan / *_ocr | 7×5 celdas limpias | — | < 1 s |

---

## 6. Excel de salida (especificación exacta)

- Un `.xlsx` por PDF. Nombre `<pdf>.xlsx` junto al PDF (o en la carpeta elegida). Si existe →
  `<pdf> (2).xlsx`, `(3)`…: **nunca se sobrescribe**. Opción avanzada "Un solo libro para todo el lote".
- **Una hoja por tabla** (mergeada). Nombre: `title` detectado, si no `"Tabla N"`; `safe_sheet_name` trunca a
  31 chars, quita `[]:*?/\`, evita duplicados con sufijo ` (2)`. Con `one_sheet_per_page` se agrega
  ` p.N`.
- **Fila 1 = encabezado**: negrita, fondo `1F4E78`, texto blanco, centrado, `wrap_text`; `freeze_panes="A2"`;
  `auto_filter.ref` sobre toda la tabla; bordes finos `BFBFBF`; filas alternadas sin relleno (limpio).
- **Tipos reales** (verificados por relectura con openpyxl): fechas → `datetime` con `dd/mm/yyyy`; importes →
  `float` con `#,##0.00` (o `"$" #,##0.00`); enteros → `int` con `#,##0`; porcentajes → fracción con `0.00%`;
  identificadores y texto → `str`. Los `number_format` son neutrales: Excel en Windows es-AR muestra
  `1.234,56` solo. Negativos conservan el signo; vacías quedan `None` (no 0, no "").
- Alineación: números y fechas a la derecha, texto a la izquierda, vertical arriba. Filas en `total_rows` en
  negrita con borde superior. **Sin fórmulas**: si el PDF dice TOTAL se escribe el valor leído.
- **Celdas para revisar** (`review_cells`): texto tal cual, relleno amarillo `FFF2CC`, comentario
  "Revisar: no se pudo interpretar como número/fecha". Hoja **"Revisión"** (Hoja | Celda | Valor | Motivo)
  solo si hay algo que revisar. Si hubo OCR, la primera fila de "Revisión" dice "Este PDF era un escaneo: los
  números pueden tener errores de lectura".
- Anchos: `max(8, longitud del renglón más largo + 2)` acotado a 60; texto > 60 con wrap.
- Hoja **"Datos del documento"** (si `include_text`): pares clave/valor (Cliente, Fecha, CUIT…) con la fecha
  tipada, y debajo el texto libre (Página | Texto, ancho 100, wrap). Si el PDF **no tiene tablas**, el libro
  tiene solo esta hoja (renombrada "Texto") y la UI avisa "No encontramos tablas: te dejamos el texto".
- Hoja **"Origen"** (oculta; visible con `show_origin_sheet`): archivo, páginas, por tabla: estrategia,
  puntaje, páginas, OCR sí/no, versión de la app. Para reportar problemas.
- Propiedades: `creator="Conversor PDF a Excel <versión>"`, `title=<nombre del PDF>`. Primera hoja activa.

---

## 7. Interfaz de usuario

Una ventana (`QMainWindow`, 860×560, mínimo 640×440), sin menú ni barra de herramientas, `QStackedWidget`
con 4 pantallas, **una sola acción primaria por pantalla** (Enter la ejecuta; Esc vuelve/cancela). Fuente del
sistema (Segoe UI en Windows) 10–11 pt. Tema claro, acento `1F6FEB`; sigue el modo oscuro del sistema.
Ícono de engranaje chico arriba a la derecha → opciones avanzadas. Todos los textos en `ui/strings_es.py`.

**Pantalla 1 — Vacío** (`DropPage`): zona punteada grande con ícono PDF: **"Arrastrá tus PDF acá"** /
"o hacé clic para elegirlos" (acepta varios archivos y carpetas; filtro `*.pdf`; Ctrl+O). Abajo, chico:
"Los archivos se convierten en tu computadora, no se suben a ningún lado." Si `update_check` encontró
versión nueva: línea discreta al pie "Hay una versión nueva (0.2.0) · Descargar" (descartable).

**Pantalla 2 — Esto encontré** (`ReviewPage`): aparece apenas termina el análisis del primer archivo (los
demás se agregan a medida que terminan). Izquierda: lista de archivos con miniatura de la 1.ª página
(`get_pixmap(dpi=40)`) y estado: "3 tablas · 3 pág." / "1 tabla · 2 pág. · escaneado, se leyó con OCR" /
"Sin tablas: va el texto completo" / "⚠ No se pudo leer: protegido con contraseña". Derecha: vista previa
de solo lectura (`TablePreviewModel`, primeras 200 filas) con combo "Tabla 1 de 3", números alineados a la
derecha y formateados en AR, fechas `dd/mm/aaaa`, celdas para revisar en amarillo; debajo "4 celdas
marcadas para revisar". Única opción visible: casilla **"Convertir números y fechas (recomendado)"**
(marcada). Selector "Guardar: Junto a cada PDF ▾ / Elegir carpeta…". Botones: "Volver" y **"Guardar Excel"**.
La vista previa no es un editor: existe para generar confianza antes de guardar.

**Pantalla 3 — Convirtiendo** (`ProgressPage`): nombre del archivo, "Archivo 2 de 5 · página 12 de 50",
barra determinada, mensaje de etapa: "Leyendo el escaneo (OCR)… puede tardar ~2 s por página" /
"Buscando tablas…" / "Armando el Excel…"; botón **"Cancelar"** (lo ya convertido se conserva).
La misma pantalla se usa durante el análisis previo a "Esto encontré" cuando hay OCR (si no, el análisis
tarda < 3 s y se muestra un spinner en la pantalla 1).

**Pantalla 4 — Listo** (`DonePage`): "¡Listo! Guardamos 5 archivos" + lista "✔ remito.pdf → remito.xlsx
(3 tablas, 128 filas)" / "⚠ factura.pdf: no se pudo leer, está protegido con contraseña" / "⚠ 4 celdas
para revisar en amarillo". Botones: **"Abrir Excel"** (si fue 1 archivo o lote ≤ 3; `QDesktopServices.openUrl`),
"Abrir carpeta", "Convertir otros".

**Errores** (`UserError` → `QMessageBox` o fila de la lista, sin trazas, con acción):
- `encrypted`: "Este PDF está protegido con contraseña. Ingresala para continuar." [campo + Continuar / Saltar]
- `damaged`/`not_pdf`: "No pudimos leer remito.pdf: parece estar dañado o no es un PDF." [Saltar]
- `file_in_use`: "No se pudo guardar remito.xlsx porque está abierto en Excel. Cerralo y tocá Reintentar." [Reintentar / Guardar con otro nombre]
- `no_permission`: "No tenés permiso para escribir en esta carpeta. Elegí otra." [Elegir carpeta]
- `ocr_unavailable`: "El reconocimiento de texto (OCR) no está disponible en esta instalación; convertimos solo las páginas con texto." (aviso no bloqueante)
- `unknown`: "No pudimos convertir este archivo." + "Ver detalles" (abre el log). Un archivo fallido nunca frena el lote.

**Opciones avanzadas** (engranaje; nadie las necesita el día 1): idioma OCR (Español+Inglés / Español /
Inglés), "Unir tablas que continúan en otra página" (sí), "Incluir texto fuera de tablas" (sí), "Una hoja por
página" (no), "Forzar OCR en todas las páginas" (no; tooltip "Usalo si el PDF tiene texto pero sale con
caracteres raros"), "Mostrar hoja Origen" (no), "Abrir Excel al terminar" (sí), "Un solo libro para el lote" (no).
Se guardan con `QSettings`.

**Accesibilidad mínima**: todo por teclado (Tab en orden lógico, Enter = primaria, Esc = volver/cancelar,
Ctrl+O), foco visible, `setAccessibleName/Description` en DropZone, botones y estados; contraste AA (texto
`1F2328` sobre `FFFFFF`; botón primario blanco sobre `1F6FEB`); High-DPI de Qt 6 respeta el escalado de
Windows; nada se transmite solo por color (ícono ✔/⚠ + palabra).

**Integración Windows**: tarea del instalador "Enviar a → Conversor PDF a Excel": la app recibe rutas por
`argv`, las carga y va directo a "Esto encontré".

---

## 8. Empaquetado y CI

- **PyInstaller 6.22 onedir** (`packaging/conversor.spec`): entry `src/pdf2excel/__main__.py`,
  `collect_data_files("pymupdf")` (**obligatorio**: sin eso falla con `FileNotFoundError …/layout/resources/onnx/*.yaml`),
  `collect_all("pymupdf_layout")`, hiddenimports `onnxruntime`, `openpyxl.cell._writer`; `datas`:
  `resources/tessdata → tessdata`, `resources/tesseract → tesseract` (Windows), `resources/icons`;
  `excludes`: Qt no usado (WebEngine, Quick, Qml, Pdf, Multimedia, Charts, 3D, Sql, Test…), `pandas`,
  `matplotlib`, `scipy`, `tkinter`, `IPython`, `camelot`, `pdfplumber`, `pymupdf4llm`; `console=False`,
  `optimize=1`, `version=file_version_info.txt` (idioma 040A es-AR), ícono.
- **`packaging/prune_qt.py`** post-build: borra Qml/Quick/VirtualKeyboard/Pdf/Wayland/EglFS, plugins no
  usados, traducciones que no sean es/en; `opengl32sw.dll` en Windows. Medido −31,8 MB en Linux.
- **Tesseract**: `packaging/fetch_tesseract.py` descarga (con verificación sha256 fijada) el build
  UB-Mannheim 5.3.x portable para Windows a `resources/tesseract/` y `tessdata_fast` `spa`, `eng`, `osd` a
  `resources/tessdata/` (≈ 45 MB + 16 MB). Nada de esto se commitea. En macOS/Linux se usa el tesseract del
  sistema (brew/apt); si falta, la app avisa y convierte solo PDFs con texto.
- **Tamaño estimado Windows**: Qt mínimo ≈ 68 MB + pymupdf ≈ 38 + layout+onnxruntime ≈ 60 + numpy ≈ 30 +
  opencv-headless ≈ 82 (sin ffmpeg) + tesseract+tessdata ≈ 60 + Python/stdlib ≈ 25 → **≈ 330–360 MB onedir,
  instalador LZMA2 ≈ 120–140 MB**. Arranque: precalentar el modelo ONNX en un hilo al abrir (≈ 0,5 s).
- **Instalador Inno Setup 6** (`packaging/installer.iss`): `PrivilegesRequired=lowest` (instala en
  `%LOCALAPPDATA%\Programs` sin UAC; diálogo permite "para todos"), `Spanish.isl`, acceso directo opcional,
  tarea "Enviar a", `CloseApplications=yes`, misma `AppId` ⇒ actualizar es correr el setup encima,
  `[InstallDelete] _internal`. Además **zip portable** (7z -mx=7) del onedir.
- **Firma**: `signtool` sobre `.exe` y setup, timestamp DigiCert, certificado en secrets
  (`WIN_CERT_PFX_B64`, `WIN_CERT_PASSWORD`); si no hay, el paso se salta y el README explica el aviso de
  SmartScreen ("Más información → Ejecutar de todas formas"). macOS: `codesign` + `notarytool` solo con secrets.
- **Actualizaciones**: sin auto-updater. `update_check.py` hace GET a `releases/latest/download/latest.json`
  en un hilo (timeout 3 s, falla en silencio) y la UI muestra una línea discreta; el setup actualiza encima.
- **CI** (`.github/workflows/build-desktop.yml`): `test` (ubuntu-24.04: apt tesseract-ocr-spa, xvfb, libegl;
  `pip install -e .[dev]`; `ruff` + `mypy`; `QT_QPA_PLATFORM=offscreen pytest`, corpus regenerado) →
  `build-windows` (windows-2025: `fetch_tesseract.py`, `stamp_version.py` desde el tag, `pyinstaller`,
  `prune_qt.py`, **smoke** `ConversorPDFaExcel.exe --selftest tests/corpus/remito_grid_scan.pdf` que falla el
  build si el OCR dentro del bundle no anda, firma condicional, ISCC, 7z, artefactos) → `build-macos`
  (macos-14: .app + .dmg) → `release` en tags `v*` (setup, zip, dmg, `latest.json`). Caché de pip y de los
  binarios de tesseract. Tiempo estimado 10–15 min.
- **Reproducibilidad**: deps fijadas exactas, Python fijo, `--clean`, `pip-audit` semanal y Dependabot;
  cada bump de pymupdf/PySide6 corre el corpus y el smoke del bundle (ahí es donde se rompen las cosas).

---

## 9. Tests y corpus sintético

**Corpus** (`tests/corpus/generate.py`, reportlab, determinista con seed; se regenera en CI; GT en
`ground_truth.json` con valores ya tipados `{"v": 350.0, "t": "float"}`):
- remito con bordes / sin bordes; `prueba.pdf` (ya existentes, se copian).
- `hard_factura_ar`: 2 páginas, `repeatRows`, celdas con wrap, importes `$ 1.234,56`, Subtotal/IVA/TOTAL con
  SPAN y sin grilla, encabezado y pie.
- `hard_extracto_banco`: 2 páginas sin bordes, 70 movimientos, negativos, Débito/Crédito mutuamente vacías,
  "Saldo final" suelto, fechas `dd/mm/aa`.
- `hard_lista_mixta`: dos tablas en la misma página (grilla gris + solo línea bajo el encabezado), "s/d", fechas.
- Nuevos: continuación **sin** repetir encabezado; encabezado de 2 filas; apaisado; columna de porcentajes;
  números US (`1,234.56`); negativos `1.234,56-` y `(1.234,56)`; PDF con contraseña; PDF sin tablas; PDF
  vacío; mixto (pág. 1 texto + pág. 2 escaneada); fuentes sin ToUnicode (texto basura).
- Versiones `_scan` de cada uno (render 200 dpi, rotación ±1,2°, blur 0,6, ruido, JPEG q70; una rotada 90°;
  una a 150 dpi) y `_ocr` (salida del OCR propio, como fixture de regresión de la geometría).
- `tests/corpus_real/` ignorada por git para PDFs reales anonimizados del equipo (mismos tests, local).

**Scorer** (`tests/scorer.py`): `score(extracted, gt) -> Score(cells_pct, rows_pct, spurious_rows_pct,
spurious_cols, tables_detected)`. Celdas: por fila GT, mejor fila extraída, igualdad tras normalizar
(números redondeados a 2 decimales, fechas ISO, texto minúsculas con espacios colapsados). **Nuevo**:
`spurious_rows_pct` = filas extraídas que no matchean ninguna fila GT (penaliza ruido de encabezado de
página); `spurious_cols` = columnas extra.

**Tests (pytest)** y umbrales que rompen el build:
- `test_corpus_accuracy`: vectoriales ≥ 99 % celdas, 100 % tablas detectadas, espurias ≤ 1 %; `_scan` → OCR
  propio ≥ 92 % celdas (factura) / ≥ 88 % (lista mixta); multipágina = 1 tabla mergeada.
- `test_quality_selection`: en cada página el candidato elegido es el de mayor celdas % contra GT.
- `test_typing_ar`: tabla paramétrica ≥ 60 casos (`"1.234,56"`→1234.56, `"$ 1.234"`→1234, `"-62.175,43"`,
  `"1.234,56-"`→−1234.56, `"(1.234,56)"`→−1234.56, `"350.00"` en columna US→350.0, `"1.234"` solo en columna
  AR→1234, `"21%"`→0.21, `"15/10/2026"`→date, `"7/1/25"`→2025-01-07, `"32/13/2026"`→texto, `"778563"` bajo
  "Comprobante"→texto, `"20-12345678-9"`→texto, `"A-100"`→texto, Stock 24/0/112/s/d → 3 int + 1 str en review).
- `test_cells`, `test_quality`, `test_merge` (reglas 1 y 2), `test_document` (PageKind por archivo),
  `test_ocr` (slow; deskew ±1,2° → |error| < 0,2°; OSD 90°; remito_grid_scan → 7×5).
- `test_excel_roundtrip`: relectura con openpyxl: `data_type` 'n'/'d'/'s', `number_format`, `freeze_panes`,
  autofiltro, nombres de hoja válidos, ningún string numérico en columnas tipadas, `(2).xlsx` sin sobrescribir.
- `test_ui` (pytest-qt, offscreen): drop simulado con `QDropEvent` (acepta .pdf, filtra .jpg), cambio de
  páginas del stack, worker + `QSignalSpy`, cancelación, event loop con gap < 150 ms durante una conversión real.
- `--selftest` del bundle en CI (Windows) sobre un PDF escaneado.
- CI publica la tabla archivo × (celdas %, filas %, espurias %, tablas, fuente, segundos) en el summary.

---

## 10. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| **AGPL-3.0 de PyMuPDF/pymupdf-layout**: distribuir el .exe obliga a publicar el código bajo AGPL | `desktop/` con `LICENSE` AGPL-3.0 y link al repo en "Acerca de"; sin Ghostscript (una sola AGPL). Si se quiere cerrar el código: licencia comercial Artifex |
| Dependencia del modelo de layout (ONNX): un upgrade mueve la geometría | Versiones exactas; corpus con umbrales en CI antes de actualizar |
| `--collect-data pymupdf` olvidado o cambio en cómo pymupdf autoimporta `layout` | Smoke `--selftest` del bundle en CI |
| OCR en fotos de celular (perspectiva, sombras, 150 dpi) | Deskew + Otsu; aviso en UI cuando la página es escaneo; celdas no parseadas en amarillo; medir con corpus real; backlog: binarización adaptativa y corrección de perspectiva |
| Tesseract alucina `|`, confunde `,`/`.`, `1`→`l`, `0`→`O` | `clean_cell`, locale por columna, review_cells; backlog: reparación de formato minoritario y patrones de código |
| Vendorear `tesseract.exe` (binario externo, `TESSDATA_PREFIX`) | `fetch_tesseract.py` con sha256; `find_tesseract()` con degradación clara (`ocr_unavailable`); smoke en CI |
| Encabezados de 2 filas, continuación sin header, tablas rotadas: no medidos aún | Están en el corpus nuevo; no prometer "cualquier PDF" hasta medirlos |
| Ambigüedad "1.234" en columnas con 1–2 valores | Empate → AR; locale elegido visible en hoja Origen |
| Puntaje con pesos ajustados a mano sobre 11 PDFs (sobreajuste) | `test_quality_selection` + corpus real del equipo; solo 2 candidatos reduce el espacio de error |
| Tamaño 330–360 MB / instalador 120–140 MB | onedir + prune; opencv-headless sin ffmpeg; plan B: apertura morfológica en numpy (−80 MB) |
| SmartScreen "Editor desconocido" sin firma | Certificado OV o Azure Trusted Signing; README; distribuir por LAN |
| `.xlsx` abierto en Excel (PermissionError) | `file_in_use` → Reintentar / Guardar con otro nombre |
| PDFs grandes (200+ páginas escaneadas): minutos | Progreso por página, cancelación, OCR cacheado por hash, 2 archivos en paralelo |
| Qt 6.12 + Python 3.13 muy nuevos | Pinneo; actualizar solo con pipeline verde |

---

## 11. Experimentos de la síntesis

Directorio: `/tmp/claude-0/-home-user/932cdd7a-5305-5ce6-bcbe-bcac3f3b947f/scratchpad/synth/ocr_exp.py`
(usa `proto.py`/`selector.py` de P0 y su `ground_truth.json`).

**(a) Selector sin camelot** (solo `hybrid_union` + `hybrid`): hard_factura_ar 100/100 (1,39 s),
hard_extracto_banco 100/100 (1,48 s), hard_lista_mixta 98,3/90,9 (0,53 s), hard_factura_ar_ocr 98,1/90,6,
hard_lista_mixta_ocr 93,3/63,6; remitos y `prueba`/`escaneado_ocr` con celdas limpias. **Idéntico a P0 con
4 candidatos** ⇒ camelot se elimina.

**(b) OCR propio sin Ghostscript** (render 300 dpi → deskew por perfil de proyección → quitar líneas →
`tesseract --psm N -c textonly_pdf=1` → capa de texto → pipeline):

| Config | remito_grid_scan | remito_nogrid_scan | hard_factura_ar_scan | hard_lista_mixta_scan |
|---|---|---|---|---|
| psm 4, líneas fuera, deskew | 7×5 ✓ | 7×5 ✓ | **98,8 % / 93,8 %** (1,9 s/pág) | **91,7 %** / 54,5 % (1,4 s/pág) |
| psm 6, líneas fuera, deskew | 7×5 ✓ | 7×5 ✓ | 98,4 % / 92,2 % (1,8 s/pág) | 90,0 % |
| psm 3, líneas fuera, deskew | 7×5 ✓ | 7×5 ✓ | 95,0 % / 76,6 % (2,3 s/pág) | 91,7 % |
| psm 6, **con** líneas, deskew | 1×4 basura | 7×5 ✓ | 3,4 % | 90,0 % |
| Referencia ocrmypdf (P0) | 7×5 ✓ | 7×5 ✓ | 98,1 % / 90,6 % (~3 s/pág) | 93,3 % / 63,6 % |

Ángulos detectados: 0,6° / −0,2° / −0,9° (los escaneos sintéticos tienen ±1,2° y el render agrega blur).
Quitar las líneas es imprescindible en tablas con bordes; el deskew es lo que separa este resultado del
22–28 % que midieron los jueces sobre el sándwich de P1. OSD (`--psm 0 -l osd`) detectó correctamente 0°,
90° y 180° sobre la factura escaneada a 150 dpi.

---

## 12. Plan de trabajo

Diez ítems con archivos disjuntos (ver `work_items` devueltos por el arquitecto): base (modelos, errores,
esqueleto), detección de tablas, tipado+merge+texto, clasificación+OCR, pipeline+CLI, Excel, componentes de
UI + worker, ensamblado de UI, corpus + tests de calidad, empaquetado + CI. Los ítems de UI dependen de las
interfaces del core (sección 4), nunca al revés. Orden sugerido: base → (core en paralelo, Excel, UI
componentes, corpus, packaging) → pipeline → UI ensamblado → integración con el corpus real del equipo.

Esfuerzo estimado: 6–7 semanas de una persona senior, o ~4 semanas con dos (core + UI/packaging en
paralelo). MVP usable sin instalador al final de la semana 3.

---

## 13. Preguntas abiertas

1. ¿La app se distribuye solo dentro del equipo o también a terceros? Define si alcanza con AGPL o hace
   falta licencia comercial de Artifex.
2. ¿Hay presupuesto para certificado de firma de código (OV ≈ USD 200/año o Azure Trusted Signing)?
3. ¿Qué tan frecuentes son fotos de celular (vs. escáner)? Define si la corrección de perspectiva entra en v1.
4. ¿Hace falta un solo libro para todo el lote como opción por defecto en algún flujo del equipo?
5. ¿Python 3.13 o 3.12 para el build de Windows? Decidir en la primera corrida de CI según wheels disponibles.
