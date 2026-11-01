"""Estampa la versión de un tag (``v1.2.3``) en los tres lugares que la necesitan.

- ``src/pdf2excel/__init__.py``          → ``__version__ = "1.2.3"`` (única fuente en runtime)
- ``packaging/installer.iss``           → ``#define AppVersion "1.2.3"``
- ``packaging/file_version_info.txt``   → ``filevers``/``prodvers`` y ``FileVersion``/``ProductVersion``

Uso (desde ``desktop/``)::

    python packaging/stamp_version.py v1.2.3

Solo stdlib: lo corre CI en el runner de Windows antes de PyInstaller. Imprime la versión normalizada.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent  # desktop/

INIT_FILE = ROOT / "src" / "pdf2excel" / "__init__.py"
ISS_FILE = ROOT / "packaging" / "installer.iss"
VERSION_INFO_FILE = ROOT / "packaging" / "file_version_info.txt"

_VERSION_RE = re.compile(r"^(?P<num>\d+(?:\.\d+){0,3})(?P<suffix>(?:[-+][0-9A-Za-z.+-]+)?)$")


def normalize(tag_or_version: str) -> str:
    """``v1.2.3`` / ``refs/tags/v1.2.3`` / ``1.2.3`` → ``1.2.3``. Conserva sufijos (``1.2.3-rc1``).

    Lanza ``ValueError`` si no parece una versión (p. ej. el nombre de una rama).
    """
    value = tag_or_version.strip()
    if value.startswith("refs/tags/"):
        value = value[len("refs/tags/") :]
    if value[:1] in ("v", "V"):
        value = value[1:]
    if not _VERSION_RE.match(value):
        raise ValueError(f"{tag_or_version!r} no es una versión válida (se espera algo como v1.2.3)")
    return value


def numeric_tuple(version: str) -> tuple[int, int, int, int]:
    """``"1.2.3-rc1"`` → ``(1, 2, 3, 0)``: cuatro enteros para ``VSVersionInfo``."""
    match = _VERSION_RE.match(version)
    if not match:
        raise ValueError(f"{version!r} no es una versión válida")
    nums = [int(p) for p in match.group("num").split(".")]
    nums = [*nums, 0, 0, 0, 0][:4]
    return nums[0], nums[1], nums[2], nums[3]


def _sub_exactly(pattern: str, replacement: str, text: str, what: str, count: int | None = None) -> str:
    new_text, n = re.subn(pattern, replacement, text, flags=re.MULTILINE)
    if n == 0 or (count is not None and n != count):
        raise ValueError(f"No se encontró {what} para reemplazar (coincidencias: {n})")
    return new_text


def stamp_init(text: str, version: str) -> str:
    return _sub_exactly(
        r'^(__version__\s*=\s*)["\'][^"\']*["\']',
        rf'\g<1>"{version}"',
        text,
        "__version__ en __init__.py",
        count=1,
    )


def stamp_iss(text: str, version: str) -> str:
    return _sub_exactly(
        r'^(#define\s+AppVersion\s+)"[^"]*"',
        rf'\g<1>"{version}"',
        text,
        "#define AppVersion en installer.iss",
        count=1,
    )


def stamp_version_info(text: str, version: str) -> str:
    nums = numeric_tuple(version)
    tuple_text = f"({nums[0]}, {nums[1]}, {nums[2]}, {nums[3]})"
    text = _sub_exactly(r"filevers=\([^)]*\)", f"filevers={tuple_text}", text, "filevers", count=1)
    text = _sub_exactly(r"prodvers=\([^)]*\)", f"prodvers={tuple_text}", text, "prodvers", count=1)
    text = _sub_exactly(
        r"StringStruct\('FileVersion',\s*'[^']*'\)",
        f"StringStruct('FileVersion', '{version}')",
        text,
        "FileVersion",
        count=1,
    )
    return _sub_exactly(
        r"StringStruct\('ProductVersion',\s*'[^']*'\)",
        f"StringStruct('ProductVersion', '{version}')",
        text,
        "ProductVersion",
        count=1,
    )


def main(tag_or_version: str, root: Path | None = None) -> str:
    """Normaliza, reescribe los tres archivos y devuelve (e imprime) la versión."""
    base = Path(root) if root is not None else ROOT
    version = normalize(tag_or_version)
    targets = {
        base / "src" / "pdf2excel" / "__init__.py": stamp_init,
        base / "packaging" / "installer.iss": stamp_iss,
        base / "packaging" / "file_version_info.txt": stamp_version_info,
    }
    for path, stamper in targets.items():
        original = path.read_text(encoding="utf-8")
        path.write_text(stamper(original, version), encoding="utf-8", newline="\n")
    print(version)
    return version


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Uso: python packaging/stamp_version.py v1.2.3")
    try:
        main(sys.argv[1])
    except (ValueError, OSError) as exc:
        sys.exit(f"stamp_version: {exc}")
