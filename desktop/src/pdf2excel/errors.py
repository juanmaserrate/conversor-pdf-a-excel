"""Errores para humanos: ``UserError`` y ``from_exception()``.

Toda excepción técnica que llegue a la UI o a la CLI pasa por ``from_exception`` y se convierte en un
``UserError`` con un código estable, un título y un detalle en castellano llano y, si corresponde, una
acción sugerida. La UI puede reemplazar los textos por los de ``ui/strings_es.py`` usando ``code``.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from typing import Literal

ErrorCode = Literal[
    "encrypted",
    "damaged",
    "not_pdf",
    "no_permission",
    "file_in_use",
    "ocr_unavailable",
    "ocr_failed",
    "cancelled",
    "unknown",
]
ErrorAction = Literal["retry", "save_as", "choose_folder", "password", "skip"] | None

DEFAULT_TITLES: dict[ErrorCode, str] = {
    "encrypted": "El PDF está protegido con contraseña",
    "damaged": "No pudimos leer el PDF: parece estar dañado",
    "not_pdf": "El archivo no es un PDF",
    "no_permission": "No tenés permiso para escribir en esa carpeta",
    "file_in_use": "El Excel está abierto en otro programa",
    "ocr_unavailable": "El reconocimiento de texto (OCR) no está disponible",
    "ocr_failed": "Falló el reconocimiento de texto (OCR)",
    "cancelled": "Conversión cancelada",
    "unknown": "No pudimos convertir este archivo",
}

DEFAULT_DETAILS: dict[ErrorCode, str] = {
    "encrypted": "Ingresá la contraseña para continuar o saltá este archivo.",
    "damaged": "Probá abrirlo con un lector de PDF; si tampoco abre, pedí que te lo vuelvan a enviar.",
    "not_pdf": "Solo se pueden convertir archivos .pdf.",
    "no_permission": "Elegí otra carpeta para guardar el Excel.",
    "file_in_use": "Cerralo (por ejemplo en Excel) y tocá Reintentar, o guardalo con otro nombre.",
    "ocr_unavailable": "Convertimos solo las páginas que ya tienen texto. Las páginas escaneadas se omiten.",
    "ocr_failed": "El programa de OCR devolvió un error. Probá de nuevo; si sigue fallando, mirá el log.",
    "cancelled": "Lo que ya se había convertido se conserva.",
    "unknown": "Podés ver los detalles técnicos en el log para reportar el problema.",
}

# Acción sugerida por defecto para cada código (la usa from_exception; UserError no la impone).
DEFAULT_ACTIONS: dict[ErrorCode, ErrorAction] = {
    "encrypted": "password",
    "damaged": "skip",
    "not_pdf": "skip",
    "no_permission": "choose_folder",
    "file_in_use": "retry",
    "ocr_unavailable": None,
    "ocr_failed": "retry",
    "cancelled": None,
    "unknown": None,
}


class UserError(Exception):
    """Error presentable al usuario: código estable + textos en castellano + acción sugerida."""

    code: ErrorCode
    title: str
    detail: str
    action: ErrorAction
    cause: BaseException | None

    def __init__(
        self,
        code: ErrorCode,
        title: str = "",
        detail: str = "",
        action: ErrorAction = None,
        cause: BaseException | None = None,
    ) -> None:
        self.code = code
        self.title = title or DEFAULT_TITLES.get(code, DEFAULT_TITLES["unknown"])
        self.detail = detail or DEFAULT_DETAILS.get(code, "")
        self.action = action
        self.cause = cause
        if cause is not None and self.__cause__ is None:
            self.__cause__ = cause
        super().__init__(self.title)

    def __str__(self) -> str:
        return f"{self.title}. {self.detail}".strip() if self.detail else self.title

    def __repr__(self) -> str:
        return f"UserError(code={self.code!r}, title={self.title!r}, action={self.action!r})"


_DAMAGED_CLASS_NAMES = frozenset({"FileDataError", "EmptyFileError"})


def _is_pymupdf_open_error(exc: BaseException) -> bool:
    """True si ``exc`` es pymupdf.FileDataError / EmptyFileError (sin importar pymupdf si no está cargado)."""
    mod = sys.modules.get("pymupdf")
    if mod is not None:
        classes = tuple(
            c for c in (getattr(mod, "FileDataError", None), getattr(mod, "EmptyFileError", None)) if c
        )
        if classes and isinstance(exc, classes):
            return True
    return type(exc).__name__ in _DAMAGED_CLASS_NAMES


def _exception_path(exc: BaseException, path: Path | None) -> Path | None:
    """Ruta más específica disponible: la de la excepción (OSError.filename) o la que nos pasaron."""
    filename = getattr(exc, "filename", None)
    if isinstance(filename, str | bytes | Path) and filename:
        return Path(filename.decode(errors="replace") if isinstance(filename, bytes) else filename)
    return path


def from_exception(exc: BaseException, path: Path | None = None) -> UserError:
    """Convierte cualquier excepción en un ``UserError`` según las reglas de DESIGN.md §4.

    - ``UserError`` → se devuelve tal cual.
    - ``pymupdf.FileDataError`` / ``EmptyFileError`` / ``RuntimeError("cannot open ...")`` → ``damaged``
      (o ``not_pdf`` si la ruta no termina en .pdf).
    - ``ValueError`` con "not a PDF" o ruta con extensión distinta de .pdf → ``not_pdf``.
    - ``PermissionError`` sobre un .xlsx → ``file_in_use`` (acción "retry"); otro → ``no_permission``
      (acción "choose_folder").
    - ``FileNotFoundError`` que menciona tesseract → ``ocr_unavailable``.
    - ``subprocess.CalledProcessError`` / ``TimeoutExpired`` → ``ocr_failed``.
    - Resto → ``unknown`` con el detalle técnico. Siempre se guarda ``cause``.
    """
    if isinstance(exc, UserError):
        return exc

    msg = str(exc)
    msg_lower = msg.lower()
    target = _exception_path(exc, path)
    name = target.name if target is not None else ""
    is_pdf_ext = target is None or target.suffix.lower() == ".pdf"

    def make(code: ErrorCode, detail: str = "", action: ErrorAction | Literal[False] = False) -> UserError:
        chosen: ErrorAction = DEFAULT_ACTIONS[code] if action is False else action
        return UserError(code, detail=detail, action=chosen, cause=exc)

    if _is_pymupdf_open_error(exc) or (isinstance(exc, RuntimeError) and "cannot open" in msg_lower):
        if not is_pdf_ext:
            return make("not_pdf", f"{name} no tiene extensión .pdf; solo se pueden convertir archivos PDF.")
        detail = f"No pudimos leer {name}: parece estar dañado o no es un PDF." if name else ""
        return make("damaged", detail)

    if isinstance(exc, ValueError) and ("not a pdf" in msg_lower or not is_pdf_ext):
        detail = f"{name} no es un archivo PDF." if name else ""
        return make("not_pdf", detail)

    if isinstance(exc, PermissionError):
        if target is not None and target.suffix.lower() == ".xlsx":
            detail = f"No se pudo guardar {name} porque está abierto en Excel. Cerralo y tocá Reintentar."
            return make("file_in_use", detail)
        folder = target if target is not None else None
        detail = f"No tenés permiso para escribir en {folder}. Elegí otra carpeta." if folder else ""
        return make("no_permission", detail)

    if isinstance(exc, FileNotFoundError) and "tesseract" in (msg_lower + " " + name.lower()):
        return make("ocr_unavailable")

    if isinstance(exc, subprocess.CalledProcessError | subprocess.TimeoutExpired):
        return make("ocr_failed")

    detail = f"{type(exc).__name__}: {msg}" if msg else type(exc).__name__
    return make("unknown", detail)


__all__ = [
    "DEFAULT_ACTIONS",
    "DEFAULT_DETAILS",
    "DEFAULT_TITLES",
    "ErrorAction",
    "ErrorCode",
    "UserError",
    "from_exception",
]
