"""Chequeo de actualizaciones: un GET a ``latest.json`` del último release, en un hilo, sin bloquear la UI.

No hay auto-updater: si existe una versión más nueva la UI muestra una línea discreta con el link al
instalador y el setup se instala "encima" (misma ``AppId`` en Inno Setup). Este módulo no importa Qt y
nunca lanza: cualquier problema de red, de formato o de parseo devuelve ``None``.

El archivo ``latest.json`` lo genera el job ``release`` del workflow y tiene la forma
``{"version": "0.2.0", "url": "<setup .exe>", "notes_url": "<página del release>"}``.
"""

from __future__ import annotations

import json
import logging
import os
import re
import threading
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlsplit

from pdf2excel import APP_ID, __version__

log = logging.getLogger(__name__)

#: Dueño del repo en GitHub (la URL se arma con él). Se puede pisar por completo con la variable de
#: entorno ``PDF2EXCEL_UPDATE_URL`` (útil para pruebas o para un mirror interno).
REPO_OWNER = "juanmaserrate"
REPO_NAME = "conversor-pdf-a-excel"
UPDATE_URL_ENV = "PDF2EXCEL_UPDATE_URL"
LATEST_URL = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/releases/latest/download/latest.json"

DEFAULT_TIMEOUT = 3.0
_MAX_BODY = 64 * 1024  # latest.json pesa ~200 bytes; cualquier cosa mayor es sospechosa
_VERSION_RE = re.compile(r"^\s*v?(\d+(?:\.\d+)*)")


@dataclass(frozen=True)
class UpdateInfo:
    """Versión nueva disponible: número, link de descarga y link a las notas del release."""

    version: str
    url: str
    notes_url: str


def latest_url() -> str:
    """URL de ``latest.json``: ``PDF2EXCEL_UPDATE_URL`` si está definida, si no ``LATEST_URL``."""
    override = os.environ.get(UPDATE_URL_ENV, "").strip()
    return override or LATEST_URL


def version_tuple(version: str) -> tuple[int, ...]:
    """``"v1.2.3"`` → ``(1, 2, 3)``. Ignora sufijos (``1.2.3-rc1`` → ``(1, 2, 3)``).

    Si no hay nada numérico devuelve ``(0,)`` para que la comparación nunca lance.
    """
    match = _VERSION_RE.match(version or "")
    if not match:
        return (0,)
    parts = tuple(int(p) for p in match.group(1).split("."))
    # Normalizar ceros finales: (1, 2) == (1, 2, 0)
    while len(parts) > 1 and parts[-1] == 0:
        parts = parts[:-1]
    return parts


def is_newer(candidate: str, current: str) -> bool:
    """True si ``candidate`` es estrictamente mayor que ``current`` comparando tuplas de enteros."""
    return version_tuple(candidate) > version_tuple(current)


def _fetch_json(url: str, timeout: float) -> dict[str, Any]:
    scheme = urlsplit(url).scheme.lower()
    if scheme not in ("http", "https"):
        raise ValueError(f"esquema no permitido en la URL de actualizaciones: {scheme!r}")
    request = urllib.request.Request(
        url,
        headers={"User-Agent": f"{APP_ID}/{__version__}", "Accept": "application/json"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:  # esquema validado arriba
        raw = response.read(_MAX_BODY + 1)
    if len(raw) > _MAX_BODY:
        raise ValueError("latest.json demasiado grande")
    data = json.loads(raw.decode("utf-8"))
    if not isinstance(data, dict):
        raise ValueError("latest.json no es un objeto JSON")
    return data


def parse_latest(data: dict[str, Any], current: str) -> UpdateInfo | None:
    """Convierte el contenido de ``latest.json`` en ``UpdateInfo`` si es más nuevo que ``current``."""
    version = data.get("version")
    url = data.get("url")
    if not isinstance(version, str) or not isinstance(url, str) or not version.strip() or not url.strip():
        return None
    if urlsplit(url).scheme.lower() not in ("http", "https"):
        return None
    if not is_newer(version, current):
        return None
    notes_url = data.get("notes_url")
    if not isinstance(notes_url, str) or not notes_url.strip():
        notes_url = url
    return UpdateInfo(version=version.strip().lstrip("vV"), url=url.strip(), notes_url=notes_url.strip())


def check_for_update(current: str = __version__, timeout: float = DEFAULT_TIMEOUT) -> UpdateInfo | None:
    """Devuelve ``UpdateInfo`` si hay una versión más nueva que ``current``; ``None`` si no o si algo falla.

    Bloquea hasta ``timeout`` segundos (conexión y lectura). Pensado para correr en un hilo aparte:
    ver :func:`check_for_update_async`. Nunca lanza excepciones.
    """
    try:
        data = _fetch_json(latest_url(), timeout)
        return parse_latest(data, current)
    except Exception as exc:  # por diseño: el chequeo falla en silencio
        log.debug("Chequeo de actualizaciones sin resultado: %s", exc)
        return None


def check_for_update_async(
    callback: Callable[[UpdateInfo | None], None],
    current: str = __version__,
    timeout: float = DEFAULT_TIMEOUT,
) -> threading.Thread:
    """Corre :func:`check_for_update` en un hilo daemon y llama a ``callback(result)`` desde ese hilo.

    El hilo ya está arrancado cuando se devuelve. Ojo: ``callback`` corre en el hilo secundario; la UI
    debe cruzar al hilo principal (por ejemplo emitiendo una señal Qt) antes de tocar widgets.
    """

    def _run() -> None:
        result = check_for_update(current, timeout)
        try:
            callback(result)
        except Exception:  # un callback roto no debe tirar el hilo con traza al stderr
            log.exception("El callback del chequeo de actualizaciones lanzó una excepción")

    thread = threading.Thread(target=_run, name="update-check", daemon=True)
    thread.start()
    return thread


__all__ = [
    "DEFAULT_TIMEOUT",
    "LATEST_URL",
    "REPO_NAME",
    "REPO_OWNER",
    "UPDATE_URL_ENV",
    "UpdateInfo",
    "check_for_update",
    "check_for_update_async",
    "is_newer",
    "latest_url",
    "parse_latest",
    "version_tuple",
]
