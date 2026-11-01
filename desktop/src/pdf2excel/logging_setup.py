"""Log rotativo de la app en la carpeta de logs del usuario (platformdirs)."""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

import platformdirs

from pdf2excel import APP_ID, __version__

LOG_FILE_NAME = "app.log"
MAX_BYTES = 1_000_000  # 1 MB
BACKUP_COUNT = 3
_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


def log_dir() -> Path:
    """Carpeta de logs del usuario, p. ej. ``%LOCALAPPDATA%\\ConversorPDFaExcel\\Logs`` en Windows."""
    return Path(platformdirs.user_log_dir(APP_ID, appauthor=False))


def log_path() -> Path:
    """Ruta completa del archivo de log actual."""
    return log_dir() / LOG_FILE_NAME


def setup_logging(level: int = logging.INFO, directory: Path | None = None) -> Path:
    """Configura el logger raíz con un ``RotatingFileHandler`` (1 MB × 3) y devuelve la ruta del log.

    Es idempotente: llamarla dos veces no duplica handlers. En desarrollo (no congelado) también
    escribe a stderr. ``directory`` permite elegir otra carpeta (tests); por defecto ``log_dir()``.
    """
    target_dir = directory if directory is not None else log_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    path = target_dir / LOG_FILE_NAME

    root = logging.getLogger()
    root.setLevel(level)
    formatter = logging.Formatter(_FORMAT)

    already = any(
        isinstance(h, RotatingFileHandler) and Path(h.baseFilename) == path.resolve() for h in root.handlers
    )
    if not already:
        file_handler = RotatingFileHandler(
            path, maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)

    frozen = bool(getattr(sys, "frozen", False))
    has_stream = any(
        isinstance(h, logging.StreamHandler) and not isinstance(h, RotatingFileHandler) for h in root.handlers
    )
    if not frozen and not has_stream and sys.stderr is not None:
        stream_handler = logging.StreamHandler(sys.stderr)
        stream_handler.setFormatter(formatter)
        root.addHandler(stream_handler)

    # Silenciar libs ruidosas.
    for noisy in ("PIL", "onnxruntime", "urllib3"):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    logging.getLogger(__name__).info("%s %s — log en %s", APP_ID, __version__, path)
    return path


__all__ = ["BACKUP_COUNT", "LOG_FILE_NAME", "MAX_BYTES", "log_dir", "log_path", "setup_logging"]
