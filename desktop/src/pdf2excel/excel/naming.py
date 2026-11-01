"""Nombres seguros para el archivo de salida y las hojas del libro (DESIGN.md §6).

- `unique_output_path`: nunca se sobrescribe un .xlsx existente (`nombre (2).xlsx`, `(3)`...).
- `safe_sheet_name`: nombre de hoja válido para Excel (≤ 31 chars, sin `[]:*?/\\`), único dentro del libro
  y nunca igual a una hoja reservada ("Revisión", "Datos del documento", "Texto", "Origen").
"""

from __future__ import annotations

import re
from pathlib import Path

from pdf2excel.errors import UserError

MAX_SHEET_NAME = 31
MAX_SUFFIX = 999

# Hojas que escribe el writer y que una tabla nunca puede usar como nombre.
RESERVED_SHEET_NAMES: frozenset[str] = frozenset({"Revisión", "Datos del documento", "Texto", "Origen"})

_FORBIDDEN_CHARS = re.compile(r"[\[\]:*?/\\]")
_SPACES = re.compile(r"\s+")


def unique_output_path(src: Path, output_dir: Path | None) -> Path:
    """`<dir>/<src.stem>.xlsx`; si existe, `<stem> (2).xlsx`, `(3)`... hasta 999.

    `output_dir=None` significa junto al PDF. No crea carpetas ni archivos.
    """
    directory = Path(output_dir) if output_dir is not None else Path(src).parent
    stem = Path(src).stem or "salida"
    candidate = directory / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate
    for n in range(2, MAX_SUFFIX + 1):
        candidate = directory / f"{stem} ({n}).xlsx"
        if not candidate.exists():
            return candidate
    raise UserError(
        "unknown",
        title="No se pudo elegir un nombre para el Excel",
        detail=(
            f"Ya hay demasiados archivos llamados '{stem}' en la carpeta. "
            "Elegí otra carpeta o renombrá algunos."
        ),
        action="choose_folder",
    )


def clean_sheet_title(title: str | None) -> str:
    """Quita caracteres prohibidos y apóstrofos en los extremos, colapsa espacios y recorta."""
    if not title:
        return ""
    cleaned = _FORBIDDEN_CHARS.sub(" ", str(title))
    cleaned = _SPACES.sub(" ", cleaned).strip().strip("'").strip()
    return cleaned


def safe_sheet_name(title: str | None, index: int, used: set[str]) -> str:
    """Nombre de hoja válido y único; agrega el resultado a `used`.

    Base = `title` limpio o `"Tabla {index}"`; se trunca a 31 caracteres. Si ya está en `used` (sin
    distinguir mayúsculas) o es un nombre reservado, se agrega ` (2)`, ` (3)`... recortando la base para
    no pasar de 31.
    """
    base = clean_sheet_title(title) or f"Tabla {index}"
    base = base[:MAX_SHEET_NAME].rstrip()
    taken = {u.casefold() for u in used} | {r.casefold() for r in RESERVED_SHEET_NAMES}

    name = base
    n = 2
    while name.casefold() in taken:
        suffix = f" ({n})"
        name = base[: MAX_SHEET_NAME - len(suffix)].rstrip() + suffix
        n += 1
    used.add(name)
    return name


__all__ = [
    "MAX_SHEET_NAME",
    "RESERVED_SHEET_NAMES",
    "clean_sheet_title",
    "safe_sheet_name",
    "unique_output_path",
]
