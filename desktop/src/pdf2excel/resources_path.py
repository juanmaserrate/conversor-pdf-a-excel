"""Rutas a los recursos empaquetados (íconos, tessdata, tesseract) en desarrollo y dentro del bundle."""

from __future__ import annotations

import os
import sys
from pathlib import Path

_ENV_OVERRIDE = "PDF2EXCEL_RESOURCES"


def resource_dir() -> Path:
    """Carpeta ``resources/``.

    - Congelado con PyInstaller: ``sys._MEIPASS`` (los ``datas`` del .spec se copian a la raíz).
    - En desarrollo (src-layout): ``<repo>/desktop/resources``.
    - ``PDF2EXCEL_RESOURCES`` en el entorno la pisa (útil para tests y layouts raros).
    """
    override = os.environ.get(_ENV_OVERRIDE)
    if override:
        return Path(override)
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", None)
        if base:
            return Path(base)
        return Path(sys.executable).resolve().parent
    # src/pdf2excel/resources_path.py → parents[2] == desktop/
    return Path(__file__).resolve().parents[2] / "resources"


def tessdata_dir() -> Path:
    """Carpeta con ``spa.traineddata``, ``eng.traineddata`` y ``osd.traineddata`` (tessdata_fast)."""
    return resource_dir() / "tessdata"


def tesseract_dir() -> Path:
    """Carpeta con ``tesseract.exe`` + DLLs vendoreados (solo Windows)."""
    return resource_dir() / "tesseract"


def icon_path(name: str) -> Path:
    """Ruta a ``resources/icons/<name>`` (p. ej. ``icon_path("app.png")``)."""
    return resource_dir() / "icons" / name


__all__ = ["icon_path", "resource_dir", "tessdata_dir", "tesseract_dir"]
