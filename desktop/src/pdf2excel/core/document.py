"""Apertura segura del PDF y clasificación de páginas (DESIGN.md §5.1).

Funciones puras sobre ``pymupdf``: nada de Qt. Todo error de apertura se convierte en un ``UserError``
con código estable (``not_pdf`` / ``damaged`` / ``encrypted``) para que la UI y la CLI lo muestren en
castellano llano.
"""

from __future__ import annotations

import hashlib
import logging
import re
from pathlib import Path

import pymupdf

from pdf2excel.errors import UserError, from_exception
from pdf2excel.models import PageInfo, PageKind

log = logging.getLogger(__name__)

# Umbrales de clasificación (DESIGN.md §5.1)
MIN_TEXT_CHARS = 20  # menos que esto es "sin texto útil"
SCANNED_COVERAGE = 0.5  # fracción de la página cubierta por imágenes para considerarla escaneo
# Desvío de DESIGN §5.1: una página SIN texto pero con una imagen que cubre más que un logo (> 5 %) es un
# escaneo chico (ticket/remito pegado en una hoja A4, foto insertada), no una página "digital" vacía.
SMALL_SCAN_COVERAGE = 0.05
GARBAGE_FRACTION = 0.30  # fracción de caracteres "raros" para considerar el texto basura

# Caracteres esperables en texto castellano/inglés de documentos comerciales. Lo que queda afuera en
# más del 30 % de los casos suele venir de fuentes sin ToUnicode (texto basura → se trata como escaneo).
_ALLOWED_CHAR_RE = re.compile(r"[A-Za-z0-9áéíóúñÁÉÍÓÚÑüÜ$%.,;:/()\-–—°º#\"' \n\t]")

_PDF_MAGIC = b"%PDF-"
_HEADER_SNIFF_BYTES = 1024  # la spec de PDF tolera basura antes del encabezado (hasta 1 KB)

_SHA1_CHUNK = 1 << 20


def _looks_like_pdf(path: Path) -> bool:
    """True si los primeros bytes contienen ``%PDF-`` (acepta basura previa, como la spec)."""
    try:
        with path.open("rb") as fh:
            head = fh.read(_HEADER_SNIFF_BYTES)
    except OSError:
        return False
    return _PDF_MAGIC in head


def open_document(path: Path, password: str | None = None) -> pymupdf.Document:
    """Abre ``path`` como PDF o lanza un ``UserError``.

    - Extensión distinta de ``.pdf`` o contenido que no empieza con ``%PDF-`` → ``not_pdf``.
    - ``pymupdf.FileDataError`` / ``EmptyFileError`` (archivo roto) → ``damaged``.
    - PDF cifrado sin contraseña, o con una contraseña incorrecta → ``encrypted`` (acción ``password``).
    """
    path = Path(path)
    if path.suffix.lower() != ".pdf":
        raise UserError(
            "not_pdf",
            detail=f"{path.name} no tiene extensión .pdf; solo se pueden convertir archivos PDF.",
            action="skip",
        )
    if path.is_file() and not _looks_like_pdf(path):
        raise UserError(
            "not_pdf",
            detail=f"{path.name} no es un archivo PDF (aunque tenga esa extensión).",
            action="skip",
        )
    try:
        doc = pymupdf.open(path, filetype="pdf")
    except (pymupdf.FileDataError, pymupdf.EmptyFileError) as exc:
        raise UserError(
            "damaged",
            detail=f"No pudimos leer {path.name}: parece estar dañado o no es un PDF.",
            action="skip",
            cause=exc,
        ) from exc
    except Exception as exc:  # FileNotFoundError de pymupdf, permisos, etc.
        raise from_exception(exc, path) from exc

    if not doc.is_pdf:
        doc.close()
        raise UserError("not_pdf", detail=f"{path.name} no es un archivo PDF.", action="skip")

    if doc.needs_pass:
        authenticated = bool(password) and doc.authenticate(password or "") > 0
        if not authenticated:
            doc.close()
            detail = (
                "La contraseña no es correcta. Probá de nuevo o saltá este archivo."
                if password
                else f"{path.name} está protegido con contraseña. Ingresala para continuar."
            )
            raise UserError("encrypted", detail=detail, action="password")
    return doc


def _image_coverage(page: pymupdf.Page) -> float:
    """Fracción del área de la página cubierta por imágenes (suma de áreas recortadas, acotada a 1)."""
    page_rect = page.rect
    page_area = page_rect.width * page_rect.height
    if page_area <= 0:
        return 0.0
    try:
        infos = page.get_image_info()
    except Exception:  # páginas raras: lo tratamos como "sin imágenes"
        log.debug("get_image_info falló en la página %s", page.number, exc_info=True)
        return 0.0
    covered = 0.0
    for info in infos:
        bbox = info.get("bbox")
        if not bbox:
            continue
        rect = pymupdf.Rect(bbox) & page_rect
        if rect.is_empty or rect.is_infinite:
            continue
        covered += rect.width * rect.height
    return min(1.0, covered / page_area)


def _garbage_fraction(text: str) -> float:
    """Fracción de caracteres fuera del alfabeto esperable (0 si el texto está vacío)."""
    if not text:
        return 0.0
    allowed = len(_ALLOWED_CHAR_RE.findall(text))
    return 1.0 - allowed / len(text)


def classify_page(page: pymupdf.Page) -> PageInfo:
    """Clasifica la página: ``digital`` / ``scanned`` / ``scanned_with_text`` / ``garbage_text`` / ``empty``.

    Reglas (DESIGN.md §5.1): pocas letras + mucha imagen → escaneo; letras + mucha imagen → escaneo que ya
    tiene capa OCR; ni letras ni imagen → vacía; texto con > 30 % de caracteres raros → basura (fuentes sin
    ToUnicode, se trata como escaneo); el resto es digital. Además (desvío documentado): sin texto útil y
    con una imagen que cubre > 5 % de la página → escaneo (si no, nunca pasaría por OCR).
    """
    text = page.get_text().strip()
    n_chars = len(text)
    coverage = _image_coverage(page)

    kind: PageKind
    if n_chars < MIN_TEXT_CHARS and coverage > SCANNED_COVERAGE:
        kind = "scanned"
    elif n_chars < MIN_TEXT_CHARS and coverage > SMALL_SCAN_COVERAGE:
        kind = "scanned"  # escaneo chico pegado en la hoja: sin OCR no aporta nada
    elif coverage > SCANNED_COVERAGE:
        kind = "scanned_with_text"
    elif n_chars == 0 and coverage == 0:
        kind = "empty"
    elif n_chars >= MIN_TEXT_CHARS and _garbage_fraction(text) > GARBAGE_FRACTION:
        kind = "garbage_text"
    else:
        kind = "digital"
    return PageInfo(index=page.number, kind=kind, n_chars=n_chars, image_coverage=coverage)


def classify_document(doc: pymupdf.Document) -> list[PageInfo]:
    """``classify_page`` para todas las páginas, en orden."""
    return [classify_page(doc[i]) for i in range(doc.page_count)]


def needs_ocr(info: PageInfo, force: bool = False) -> bool:
    """True si la página hay que pasarla por OCR (escaneo sin texto o texto basura), o si se fuerza."""
    return force or info.kind in ("scanned", "garbage_text")


def file_sha1(path: Path) -> str:
    """SHA-1 hexadecimal del contenido del archivo (clave de la caché de OCR)."""
    digest = hashlib.sha1()  # no es criptográfico: solo identifica el archivo
    with Path(path).open("rb") as fh:
        while chunk := fh.read(_SHA1_CHUNK):
            digest.update(chunk)
    return digest.hexdigest()


__all__ = [
    "GARBAGE_FRACTION",
    "MIN_TEXT_CHARS",
    "SCANNED_COVERAGE",
    "SMALL_SCAN_COVERAGE",
    "classify_document",
    "classify_page",
    "file_sha1",
    "needs_ocr",
    "open_document",
]
