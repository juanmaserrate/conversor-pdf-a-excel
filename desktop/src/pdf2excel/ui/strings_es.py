"""TODOS los textos visibles de la app, en castellano rioplatense ("vos").

Solo constantes en mayúsculas y unas pocas funciones de formato. Nada de Qt acá: lo usan tanto los
widgets como los tests. Los textos de error siguen DESIGN.md §7; ``{name}`` se reemplaza por el nombre
del archivo.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pdf2excel.errors import ErrorCode, UserError

if TYPE_CHECKING:
    from pdf2excel.models import ConversionReport

# --------------------------------------------------------------------------- generales

APP_TITLE = "Conversor PDF a Excel"
PRIVACY_NOTE = "Los archivos se convierten en tu computadora, no se suben a ningún lado."
UPDATE_AVAILABLE_FMT = "Hay una versión nueva ({v}) · Descargar"
VERSION_FMT = "Versión {v}"
LINK_LOG = "Ver carpeta de logs"
ABOUT_LICENSE = "Software libre bajo licencia AGPL-3.0"

# --------------------------------------------------------------------------- pantalla 1: arrastrá

DROP_TITLE = "Arrastrá tus PDF acá"
DROP_SUBTITLE = "o hacé clic para elegirlos"
FILE_DIALOG_TITLE = "Elegí los PDF"
FILE_DIALOG_FILTER = "Archivos PDF (*.pdf)"
FOLDER_DIALOG_TITLE = "Elegí la carpeta donde guardar los Excel"
ANALYZING = "Revisando los PDF…"

# --------------------------------------------------------------------------- botones

BTN_SAVE = "Guardar Excel"
BTN_BACK = "Volver"
BTN_CANCEL = "Cancelar"
BTN_OK = "Aceptar"
BTN_OPEN_EXCEL = "Abrir Excel"
BTN_OPEN_FOLDER = "Abrir carpeta"
BTN_CONVERT_MORE = "Convertir otros"
BTN_REMOVE = "Quitar"
BTN_ADD_MORE = "Agregar más PDF"
BTN_SETTINGS = "Opciones avanzadas"
BTN_CONTINUE = "Continuar"
BTN_SKIP = "Saltar"
BTN_RETRY = "Reintentar"
BTN_SAVE_AS = "Guardar con otro nombre"
BTN_CHOOSE_FOLDER = "Elegir carpeta"
BTN_DETAILS = "Ver detalles"
BTN_DISMISS = "Cerrar"

# --------------------------------------------------------------------------- pantalla 2: esto encontré

REVIEW_TITLE = "Esto encontré"
REVIEW_SUBTITLE = "Mirá la vista previa y, si está bien, guardá el Excel."
FORMAT_LABEL = "Formato del Excel:"
FORMAT_LAYOUT = "Igual al PDF"
FORMAT_TABLES = "Solo tablas de datos"
FORMAT_LAYOUT_TIP = "Una hoja por página, con las cajas, los textos y las imágenes en su lugar"
FORMAT_TABLES_TIP = "Cada tabla en su hoja, limpia, lista para filtrar y sumar"
FORMAT_LAYOUT_DISABLED_TIP = (
    "Este PDF tiene demasiadas páginas para armar el diseño de cada una. "
    'Si lo necesitás, elegí "Siempre igual al PDF" en Opciones avanzadas y volvé a cargarlo.'
)
LAYOUT_PAGE_COMBO_FMT = "Página {i} de {n}"
LAYOUT_PREVIEW_NOTE = "Vista aproximada: en el Excel cada página queda en su propia hoja"
LAYOUT_PREVIEW_EMPTY = "No hay diseño para mostrar"
OPT_TYPED = "Convertir números y fechas (recomendado)"
SAVE_LABEL = "Guardar:"
SAVE_NEXT_TO_PDF = "Junto a cada PDF"
SAVE_CHOOSE = "Elegir carpeta…"
SAVE_IN_FOLDER_FMT = "En {folder}"
TABLE_COMBO_FMT = "Tabla {i} de {n}"
TABLE_PAGES_FMT = "páginas {a} a {b}"
TABLE_PAGE_FMT = "página {p}"
NO_TABLES = "Sin tablas: va el texto completo"
NO_TABLES_LONG = "No encontramos tablas: te dejamos el texto"
REVIEW_CELLS_FMT = "{n} celdas marcadas para revisar en amarillo"
REVIEW_CELL_ONE = "1 celda marcada para revisar en amarillo"
REVIEW_NONE = "Todas las celdas se interpretaron bien"
REVIEW_TOOLTIP = "Revisar: no se pudo interpretar"
PREVIEW_TRUNCATED_FMT = "Se muestran las primeras {n} filas de {total}"
PREVIEW_EMPTY = "Elegí un archivo de la lista para ver la vista previa"
SCANNED_NOTE = "Este PDF era un escaneo: revisá los números marcados en amarillo"
FILE_STATUS_QUEUED = "En cola"
FILE_STATUS_WORKING = "Revisando…"
FILE_STATUS_EXPORTING = "Guardando…"
FILE_STATUS_OCR = "Leyendo el escaneo (OCR)…"
FILE_STATUS_OCR_NOTE = "escaneado, se leyó con OCR"
PAGES_ABBR = "pág."

# --------------------------------------------------------------------------- pantalla 3: convirtiendo

PROGRESS_TITLE = "Convirtiendo…"
PROGRESS_FILE_FMT = "Archivo {i} de {n} · página {p} de {t}"
PROGRESS_FILE_SHORT_FMT = "Archivo {i} de {n}"
PROGRESS_CANCELLING = "Cancelando… lo ya convertido se conserva"
STAGE_TEXT: dict[str, str] = {
    "classify": "Revisando el PDF…",
    "ocr": "Leyendo el escaneo (OCR)… puede tardar ~2 s por página",
    "tables": "Buscando tablas…",
    "layout": "Armando el diseño de la hoja…",
    "typing": "Interpretando números y fechas…",
    "excel": "Armando el Excel…",
}

# --------------------------------------------------------------------------- pantalla 4: listo

DONE_TITLE_FMT = "¡Listo! Guardamos {n} archivo(s)"
DONE_TITLE_ONE = "¡Listo! Guardamos 1 archivo"
DONE_TITLE_MANY_FMT = "¡Listo! Guardamos {n} archivos"
DONE_TITLE_BATCH_FMT = "¡Listo! Guardamos 1 archivo con {n} PDF"  # libro único del lote
DONE_NOTHING = "No se guardó ningún archivo"
DONE_REVIEW_FMT = "⚠ {n} celdas para revisar en amarillo"

# Confirmación al cerrar la ventana con una conversión en curso
CLOSE_CONFIRM = "¿Cancelar la conversión en curso?"
DONE_CANCELLED = "Se canceló la conversión. Lo que ya se había convertido se conserva."

# --------------------------------------------------------------------------- errores (DESIGN.md §7)

# code -> (título, detalle, etiqueta de la acción o None). {name} = nombre del archivo.
ERRORS: dict[ErrorCode, tuple[str, str, str | None]] = {
    "encrypted": (
        "Este PDF está protegido con contraseña",
        "Ingresala para continuar.",
        BTN_CONTINUE,
    ),
    "damaged": (
        "No pudimos leer {name}",
        "Parece estar dañado o no es un PDF.",
        BTN_SKIP,
    ),
    "not_pdf": (
        "No pudimos leer {name}",
        "Parece estar dañado o no es un PDF.",
        BTN_SKIP,
    ),
    "file_in_use": (
        "No se pudo guardar el Excel",
        "{name} está abierto en Excel. Cerralo y tocá Reintentar.",
        BTN_RETRY,
    ),
    "no_permission": (
        "No tenés permiso para escribir en esta carpeta",
        "Elegí otra.",
        BTN_CHOOSE_FOLDER,
    ),
    "ocr_unavailable": (
        "El reconocimiento de texto (OCR) no está disponible en esta instalación",
        "Convertimos solo las páginas con texto.",
        None,
    ),
    "ocr_failed": (
        "No se pudo leer el escaneo (OCR)",
        "Probá de nuevo; si sigue fallando, mirá el log.",
        BTN_RETRY,
    ),
    "cancelled": (
        "Conversión cancelada",
        "Lo que ya se había convertido se conserva.",
        None,
    ),
    "unknown": (
        "No pudimos convertir este archivo",
        "Podés ver los detalles técnicos en el log para reportar el problema.",
        BTN_DETAILS,
    ),
}

# Frase corta (en minúscula, sigue a "⚠ archivo.pdf: ") para la lista de la pantalla 4.
ERROR_SHORT: dict[ErrorCode, str] = {
    "encrypted": "no se pudo leer, está protegido con contraseña",
    "damaged": "no se pudo leer, parece dañado o no es un PDF",
    "not_pdf": "no se pudo leer, no es un PDF",
    "file_in_use": "no se pudo guardar el Excel, está abierto en otro programa",
    "no_permission": "no se pudo guardar, no tenés permiso en esa carpeta",
    "ocr_unavailable": "el OCR no está disponible, se convirtieron solo las páginas con texto",
    "ocr_failed": "falló la lectura del escaneo (OCR)",
    "cancelled": "se canceló antes de terminar",
    "unknown": "no se pudo convertir (mirá el log para los detalles)",
}

PASSWORD_PROMPT_FMT = "Contraseña de {name}:"
PASSWORD_PLACEHOLDER = "Escribí la contraseña y tocá Continuar"
PASSWORD_WRONG = "La contraseña no es correcta. Probá de nuevo."

# --------------------------------------------------------------------------- opciones avanzadas

ADVANCED_TITLE = "Opciones avanzadas"
ADVANCED_INTRO = "Nadie las necesita el primer día: los valores por defecto andan bien."
ADVANCED_OCR_LANG = "Idioma del OCR"
ADVANCED_LAYOUT = "Formato del Excel"
ADVANCED_LAYOUT_CHOICES: tuple[tuple[str, str], ...] = (
    ("Automático (recomendado)", "auto"),
    ("Siempre igual al PDF", "layout"),
    ("Siempre solo tablas", "tables"),
)
ADVANCED_LAYOUT_TIP = 'Automático decide por cada PDF; después podés cambiarlo en la pantalla "Esto encontré"'
ADVANCED_MERGE = "Unir tablas que continúan en otra página"
ADVANCED_INCLUDE_TEXT = "Incluir texto fuera de tablas"
ADVANCED_ONE_SHEET_PER_PAGE = "Una hoja por página"
ADVANCED_FORCE_OCR = "Forzar OCR en todas las páginas"
ADVANCED_FORCE_OCR_TIP = "Usalo si el PDF tiene texto pero sale con caracteres raros"
ADVANCED_ORIGIN_SHEET = "Mostrar hoja Origen"
ADVANCED_ORIGIN_SHEET_TIP = "Una hoja extra con datos técnicos para reportar problemas"
ADVANCED_OPEN_WHEN_DONE = "Abrir Excel al terminar"
ADVANCED_SINGLE_WORKBOOK = "Un solo libro para el lote"
ADVANCED_SINGLE_WORKBOOK_TIP = "Todas las tablas de todos los PDF en un único archivo .xlsx"

# (etiqueta visible, código de idioma de tesseract)
OCR_LANG_CHOICES: tuple[tuple[str, str], ...] = (
    ("Español + Inglés", "spa+eng"),
    ("Español", "spa"),
    ("Inglés", "eng"),
)

# --------------------------------------------------------------------------- funciones de formato


def plural(n: int, singular: str, plural_form: str) -> str:
    """``plural(3, "tabla", "tablas")`` → ``"3 tablas"``; ``plural(1, ...)`` → ``"1 tabla"``."""
    return f"{n} {singular if n == 1 else plural_form}"


def file_status(report: ConversionReport) -> str:
    """Estado corto de un archivo analizado para la lista de la pantalla 2.

    "3 tablas · 3 pág." / "1 tabla · 2 pág. · escaneado, se leyó con OCR" / NO_TABLES.
    """
    if report.mode == "layout" and report.layout_pages:
        parts = [plural(len(report.layout_pages), "hoja igual al PDF", "hojas iguales al PDF")]
        if report.ocr_applied:
            parts.append(FILE_STATUS_OCR_NOTE)
        return " · ".join(parts)
    n_tables = len(report.tables)
    if n_tables == 0:
        return NO_TABLES
    parts = [plural(n_tables, "tabla", "tablas"), f"{len(report.pages)} {PAGES_ABBR}"]
    if report.ocr_applied:
        parts.append(FILE_STATUS_OCR_NOTE)
    return " · ".join(parts)


def done_line(report: ConversionReport) -> str:
    """Renglón de la pantalla 4: "✔ remito.pdf → remito.xlsx (3 tablas, 128 filas)"."""
    src_name = report.source.name
    out_name = report.output_path.name if report.output_path else f"{report.source.stem}.xlsx"
    n_tables = len(report.tables)
    if report.mode == "layout" and report.layout_pages:
        detail = plural(len(report.layout_pages), "hoja igual al PDF", "hojas iguales al PDF")
    elif n_tables == 0:
        detail = "sin tablas, texto completo"
    else:
        detail = f"{plural(n_tables, 'tabla', 'tablas')}, {plural(report.n_rows, 'fila', 'filas')}"
    return f"✔ {src_name} → {out_name} ({detail})"


def error_line(name: str, err: UserError) -> str:
    """Renglón de error para la lista: "⚠ factura.pdf: no se pudo leer, está protegido con contraseña"."""
    short = ERROR_SHORT.get(err.code, ERROR_SHORT["unknown"])
    return f"⚠ {name}: {short}"


def error_texts(err: UserError, name: str = "") -> tuple[str, str, str | None]:
    """(título, detalle, etiqueta de acción) de ``ERRORS`` con ``{name}`` ya reemplazado."""
    title, detail, action = ERRORS.get(err.code, ERRORS["unknown"])
    shown = name or "el archivo"
    return title.format(name=shown), detail.format(name=shown), action


def done_title(n_saved: int, n_books: int | None = None) -> str:
    """ "¡Listo! Guardamos 5 archivos" / "¡Listo! Guardamos 1 archivo" / DONE_NOTHING.

    ``n_books`` = libros .xlsx distintos; con un solo libro para varios PDF (lote en un único libro):
    "¡Listo! Guardamos 1 archivo con 4 PDF".
    """
    if n_saved <= 0:
        return DONE_NOTHING
    if n_books == 1 and n_saved > 1:
        return DONE_TITLE_BATCH_FMT.format(n=n_saved)
    if n_saved == 1:
        return DONE_TITLE_ONE
    return DONE_TITLE_MANY_FMT.format(n=n_saved)


def review_cells_text(n: int) -> str:
    """ "4 celdas marcadas para revisar en amarillo" (singular y cero contemplados)."""
    if n <= 0:
        return REVIEW_NONE
    if n == 1:
        return REVIEW_CELL_ONE
    return REVIEW_CELLS_FMT.format(n=n)


def progress_text(i: int, n: int, page: int | None = None, total_pages: int | None = None) -> str:
    """ "Archivo 2 de 5 · página 12 de 50" (sin la parte de páginas si no se conoce)."""
    if page is None or total_pages is None or total_pages <= 0:
        return PROGRESS_FILE_SHORT_FMT.format(i=i, n=n)
    return PROGRESS_FILE_FMT.format(i=i, n=n, p=page, t=total_pages)
