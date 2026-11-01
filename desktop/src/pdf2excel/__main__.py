"""Punto de entrada: ``python -m pdf2excel`` / ``conversor-pdf-a-excel`` / el ejecutable empaquetado.

Sin argumentos (o solo con rutas de PDF, como hace "Enviar a") abre la interfaz gráfica con esos archivos
precargados. Con ``--cli``, ``--selftest``, ``--help`` o cualquier otra opción, o cuando no hay pantalla y
hay PDFs que convertir, delega en ``pdf2excel.cli``.
"""

from __future__ import annotations

import multiprocessing
import os
import sys

CLI_FLAGS = frozenset({"--cli", "--selftest", "--help", "-h", "--version"})


def has_display() -> bool:
    """True si hay dónde abrir una ventana: siempre en Windows/macOS; en Linux, DISPLAY/Wayland/offscreen."""
    if sys.platform.startswith("win") or sys.platform == "darwin":
        return True
    if os.environ.get("QT_QPA_PLATFORM"):
        return True
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def wants_cli(argv: list[str]) -> bool:
    """La GUI solo recibe rutas: cualquier opción con guion, o PDFs sin pantalla, van a la CLI."""
    if any(arg in CLI_FLAGS for arg in argv):
        return True
    if any(arg.startswith("-") and arg != "-" for arg in argv):
        return True
    return bool(argv) and not has_display()


def main(argv: list[str] | None = None) -> int:
    multiprocessing.freeze_support()  # primero, por si en el futuro se usan procesos (DESIGN §3)
    args = sys.argv[1:] if argv is None else list(argv)

    if wants_cli(args):
        from pdf2excel.cli import main as cli_main

        return cli_main(args)

    try:
        from pdf2excel.app import main as gui_main
    except ImportError as exc:
        from pdf2excel.cli import main as cli_main

        if args:  # hay PDFs pero no hay GUI: convertirlos igual
            return cli_main(["--cli", *args])
        stream = sys.stderr
        if stream is not None:
            stream.write(
                f"La interfaz gráfica no está disponible ({exc}). Usá: pdf2excel --cli archivo.pdf\n"
            )
        return 2
    return int(gui_main(args))


if __name__ == "__main__":
    raise SystemExit(main())
