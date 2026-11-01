"""Arranque de la interfaz gráfica: ``main(argv) -> int``.

Orden (DESIGN.md §3 y §7): ``freeze_support()`` primero, log, ``QApplication`` con nombre/ícono, locale
es-AR y traducciones de Qt (diálogos nativos en castellano), tema, precalentado del modelo de layout en un
hilo, ventana principal con los PDF recibidos por ``argv`` ("Enviar a") y chequeo de versión nueva en
segundo plano. ``PDF2EXCEL_AUTOQUIT_MS`` cierra la app a los N ms (smoke tests sin pantalla).
"""

from __future__ import annotations

import logging
import multiprocessing
import os
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QLibraryInfo, QLocale, QObject, Qt, QTimer, QTranslator, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from pdf2excel import APP_ID
from pdf2excel.logging_setup import setup_logging
from pdf2excel.resources_path import icon_path
from pdf2excel.ui import strings_es as S

log = logging.getLogger(__name__)

ORGANIZATION_NAME = "ConversorPDF"
AUTOQUIT_ENV = "PDF2EXCEL_AUTOQUIT_MS"
THEME_FILE = Path(__file__).resolve().parent / "ui" / "theme.qss"
WARM_UP_JOIN_SECONDS = 10.0


class UpdateBridge(QObject):
    """Cruza el resultado del chequeo de versión (hilo secundario) al hilo de la UI con una señal."""

    found = Signal(str, str)  # (versión, url de descarga)

    def deliver(self, info: object) -> None:
        version = getattr(info, "version", None)
        url = getattr(info, "url", None)
        if isinstance(version, str) and isinstance(url, str) and version and url:
            self.found.emit(version, url)


def install_translator(app: QApplication) -> QTranslator | None:
    """Carga ``qtbase_es`` de las traducciones de Qt si existe (diálogos nativos en castellano)."""
    translator = QTranslator(app)
    directory = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    if directory and translator.load("qtbase_es", directory):
        app.installTranslator(translator)
        return translator
    log.debug("Sin traducción qtbase_es en %s", directory)
    return None


def apply_theme(app: QApplication, theme_file: Path = THEME_FILE) -> bool:
    try:
        app.setStyleSheet(theme_file.read_text(encoding="utf-8"))
        return True
    except OSError as exc:
        log.warning("No se pudo cargar el tema %s: %s", theme_file, exc)
        return False


def _warm_up() -> None:
    # Import diferido dentro del hilo: así la ventana aparece sin esperar a pymupdf/onnx.
    from pdf2excel.core import candidates

    candidates.warm_up()


def start_warm_up() -> threading.Thread:
    thread = threading.Thread(target=_warm_up, name="layout-warm-up", daemon=True)
    thread.start()
    return thread


def start_update_check(bridge: UpdateBridge) -> bool:
    """Chequeo de versión nueva en un hilo (opcional: si falta ``update_check`` no pasa nada)."""
    try:
        from pdf2excel.update_check import check_for_update_async
    except ImportError:
        return False
    check_for_update_async(bridge.deliver)
    return True


def autoquit_ms() -> int | None:
    raw = os.environ.get(AUTOQUIT_ENV, "").strip()
    if not raw:
        return None
    try:
        return max(0, int(raw))
    except ValueError:
        return None


def pdf_args(argv: list[str]) -> list[str]:
    """Rutas ``.pdf`` recibidas por línea de comandos ("Enviar a"), absolutas y sin duplicados."""
    seen: set[str] = set()
    files: list[str] = []
    for arg in argv:
        if not arg.lower().endswith(".pdf"):
            continue
        absolute = os.path.abspath(arg)
        if absolute not in seen:
            seen.add(absolute)
            files.append(absolute)
    return files


def main(argv: list[str] | None = None) -> int:
    multiprocessing.freeze_support()  # primero, por si en el futuro se usan procesos (DESIGN §3)
    try:
        setup_logging()
    except Exception:  # sin carpeta de logs no se cae la app
        logging.basicConfig(level=logging.INFO)
        log.exception("No se pudo configurar el log en disco")
    args = list(sys.argv[1:] if argv is None else argv)

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    existing = QApplication.instance()
    app = existing if isinstance(existing, QApplication) else QApplication(sys.argv[:1])
    app.setApplicationName(APP_ID)
    app.setApplicationDisplayName(S.APP_TITLE)
    app.setOrganizationName(ORGANIZATION_NAME)
    icon_file = icon_path("app.png")
    if icon_file.is_file():
        app.setWindowIcon(QIcon(str(icon_file)))

    QLocale.setDefault(QLocale("es_AR"))
    translator = install_translator(app)  # noqa: F841 - mantener la referencia viva mientras corre la app
    apply_theme(app)
    warm_up_thread = start_warm_up()

    from pdf2excel.ui.main_window import MainWindow  # después del tema y el locale

    window = MainWindow(initial_files=pdf_args(args))
    bridge = UpdateBridge(window)
    bridge.found.connect(window.drop_page.show_update)
    start_update_check(bridge)
    window.show()

    quit_ms = autoquit_ms()
    if quit_ms is not None:
        QTimer.singleShot(quit_ms, app.quit)

    code = int(app.exec())
    # No cortar el precalentado a mitad de camino (onnxruntime en un hilo daemon al salir).
    warm_up_thread.join(timeout=WARM_UP_JOIN_SECONDS)
    return code


__all__ = [
    "AUTOQUIT_ENV",
    "ORGANIZATION_NAME",
    "UpdateBridge",
    "apply_theme",
    "autoquit_ms",
    "install_translator",
    "main",
    "pdf_args",
    "start_update_check",
    "start_warm_up",
]


if __name__ == "__main__":
    raise SystemExit(main())
