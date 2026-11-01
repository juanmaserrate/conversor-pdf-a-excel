"""Línea de comandos (sin Qt): ``pdf2excel [--cli] PDF [PDF...] [opciones]`` y ``--selftest PDF``.

Pensada para lotes por script y para el smoke del bundle en CI. Imprime una línea por archivo
(``✔ remito.pdf → remito.xlsx (3 tablas, 128 filas, 2,3 s)`` o ``⚠ remito.pdf: <motivo>``), o un
resumen JSON con ``--json``. Sale con 0 si todos los archivos se convirtieron y con 1 si alguno falló.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import tempfile
import time
import traceback
from collections.abc import Callable
from pathlib import Path
from typing import Any

from pdf2excel import APP_NAME, __version__
from pdf2excel.errors import UserError, from_exception
from pdf2excel.models import ConversionReport, Options

log = logging.getLogger(__name__)

PROG = "pdf2excel"
OK_MARK = "✔"
WARN_MARK = "⚠"
_ASCII_MARKS = {OK_MARK: "OK", WARN_MARK: "!!", "→": "->"}


# ---------------------------------------------------------------------------------------------------
# Salida tolerante: stdout puede no existir (bundle console=False) o no soportar UTF-8 (cp1252)
# ---------------------------------------------------------------------------------------------------


def _emit(text: str, *, err: bool = False) -> None:
    stream = sys.stderr if err else sys.stdout
    if stream is None:
        return
    try:
        stream.write(text + "\n")
    except UnicodeEncodeError:
        plain = text
        for mark, ascii_mark in _ASCII_MARKS.items():
            plain = plain.replace(mark, ascii_mark)
        encoding = getattr(stream, "encoding", None) or "ascii"
        stream.write(plain.encode(encoding, errors="replace").decode(encoding, errors="replace") + "\n")
    stream.flush()


def _seconds(value: float) -> str:
    """``2,3 s`` con coma decimal (castellano)."""
    return f"{value:.1f}".replace(".", ",") + " s"


def _plural(n: int, singular: str, plural: str) -> str:
    return f"{n} {singular if n == 1 else plural}"


def summary_line(report: ConversionReport, elapsed: float) -> str:
    """``✔ nombre.pdf → nombre.xlsx (3 tablas, 128 filas, 2 celdas para revisar, 2,3 s)``."""
    parts = [
        _plural(len(report.tables), "tabla", "tablas"),
        _plural(report.n_rows, "fila", "filas"),
    ]
    if report.n_review:
        parts.append(_plural(report.n_review, "celda para revisar", "celdas para revisar"))
    if report.ocr_applied:
        parts.append("OCR")
    parts.append(_seconds(elapsed))
    out_name = report.output_path.name if report.output_path is not None else "(sin guardar)"
    return f"{OK_MARK} {Path(report.source).name} → {out_name} ({', '.join(parts)})"


def error_line(src: Path, err: UserError) -> str:
    return f"{WARN_MARK} {Path(src).name}: {err.title}"


def report_to_json(report: ConversionReport, elapsed: float) -> dict[str, Any]:
    return {
        "file": str(report.source),
        "ok": True,
        "output": str(report.output_path) if report.output_path is not None else None,
        "pages": len(report.pages),
        "tables": len(report.tables),
        "rows": report.n_rows,
        "review_cells": report.n_review,
        "ocr": report.ocr_applied,
        "warnings": list(report.warnings),
        "seconds": round(elapsed, 2),
    }


def error_to_json(src: Path, err: UserError, elapsed: float) -> dict[str, Any]:
    return {
        "file": str(src),
        "ok": False,
        "error": {"code": err.code, "title": err.title, "detail": err.detail, "action": err.action},
        "seconds": round(elapsed, 2),
    }


# ---------------------------------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROG,
        description=f"{APP_NAME} {__version__}: convierte PDFs (digitales o escaneados) a .xlsx prolijos.",
        epilog="Sin argumentos se abre la interfaz gráfica.",
    )
    parser.add_argument("pdfs", nargs="*", metavar="PDF", help="archivos PDF a convertir")
    parser.add_argument("--cli", action="store_true", help="forzar el modo línea de comandos (sin ventana)")
    parser.add_argument(
        "-o", "--output-dir", metavar="DIR", help="carpeta de salida (por defecto, junto a cada PDF)"
    )
    parser.add_argument("--no-types", action="store_true", help="dejar números y fechas como texto")
    parser.add_argument("--force-ocr", action="store_true", help="pasar todas las páginas por OCR")
    parser.add_argument("--no-text", action="store_true", help="no incluir el texto fuera de tablas")
    parser.add_argument(
        "--one-sheet-per-page", action="store_true", help="una hoja por página (no une continuaciones)"
    )
    parser.add_argument("--no-merge", action="store_true", help="no unir tablas que continúan en otra página")
    parser.add_argument(
        "--lang", default="spa+eng", metavar="IDIOMAS", help="idiomas de OCR para tesseract (spa+eng)"
    )
    parser.add_argument("--origin-sheet", action="store_true", help="mostrar la hoja Origen (diagnóstico)")
    parser.add_argument("--single-workbook", action="store_true", help="un solo libro para todo el lote")
    parser.add_argument("--json", action="store_true", help="imprimir un resumen JSON por archivo")
    parser.add_argument("-v", "--verbose", action="store_true", help="mostrar el log en la consola")
    parser.add_argument(
        "--selftest", metavar="PDF", help="convertir PDF a una carpeta temporal y verificar (CI)"
    )
    parser.add_argument(
        "--selftest-report",
        metavar="ARCHIVO",
        help="escribir además el resultado del selftest en este archivo (bundle sin consola)",
    )
    parser.add_argument("--version", action="version", version=f"{PROG} {__version__}")
    return parser


def options_from_args(args: argparse.Namespace) -> Options:
    return Options(
        typed_values=not args.no_types,
        merge_continuations=not args.no_merge,
        include_text=not args.no_text,
        one_sheet_per_page=args.one_sheet_per_page,
        force_ocr=args.force_ocr,
        ocr_lang=args.lang,
        output_dir=Path(args.output_dir).expanduser() if args.output_dir else None,
        open_when_done=False,
        show_origin_sheet=args.origin_sheet,
        single_workbook=args.single_workbook,
    )


def _setup_logging(verbose: bool) -> None:
    try:
        from pdf2excel.logging_setup import setup_logging

        setup_logging(logging.INFO if verbose else logging.WARNING)
    except Exception:  # sin carpeta de logs (permisos, disco lleno): la CLI sigue igual
        logging.basicConfig(level=logging.INFO if verbose else logging.WARNING)


# ---------------------------------------------------------------------------------------------------
# Selftest (smoke del bundle en CI)
# ---------------------------------------------------------------------------------------------------


def _check_bundled_tesseract() -> None:
    """Si el bundle trae ``resources/tesseract/``, el OCR tiene que usar ESE ejecutable (no uno del PATH)."""
    from pdf2excel.core.ocr import TESSERACT_ENV, find_tesseract
    from pdf2excel.resources_path import tesseract_dir

    vendored = tesseract_dir()
    if not vendored.is_dir() or os.environ.get(TESSERACT_ENV):
        return
    found = find_tesseract()
    if found is None or vendored.resolve() not in found.resolve().parents:
        raise RuntimeError(f"tesseract vendoreado en {vendored} pero find_tesseract() devolvió {found}")


def selftest(pdf: Path, report_path: Path | None = None) -> int:
    """Convierte ``pdf`` a un tmpdir, relee el .xlsx y exige una hoja con ≥ 2 filas y ≥ 2 columnas.

    Usa una caché de OCR vacía (``PDF2EXCEL_CACHE_DIR`` apuntando al tmpdir) para que el OCR corra de
    verdad: con la caché del usuario poblada pasaría aunque tesseract estuviera roto. Si el bundle trae
    tesseract vendoreado, verifica que sea el que se usa. Imprime ``selftest ok: <hojas> hojas, <filas>
    filas, ocr=<bool>, <segundos>s`` y devuelve 0; cualquier fallo imprime el traceback y devuelve 1.

    El bundle de Windows se construye sin consola (stdout va a la nada), así que el mismo texto se
    escribe también en ``report_path`` si se indica: es lo que lee el CI.
    """
    lines: list[str] = []

    def out(text: str, *, err: bool = False) -> None:
        _emit(text, err=err)
        lines.append(text)

    try:
        return _selftest(pdf, out)
    finally:
        if report_path is not None:
            Path(report_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def _selftest(pdf: Path, out: Callable[..., None]) -> int:
    t0 = time.perf_counter()
    try:
        import openpyxl

        from pdf2excel.core.ocr import CACHE_DIR_ENV
        from pdf2excel.core.pipeline import convert
        from pdf2excel.excel.naming import RESERVED_SHEET_NAMES

        pdf = Path(pdf)
        if not pdf.is_file():
            raise FileNotFoundError(f"No existe el archivo {pdf}")
        if "pymupdf.layout" not in sys.modules:
            raise RuntimeError("pymupdf.layout no se importó: el modelo de layout no está disponible")
        _check_bundled_tesseract()
        with tempfile.TemporaryDirectory(prefix="pdf2excel-selftest-") as tmp:
            previous_cache = os.environ.get(CACHE_DIR_ENV)
            os.environ[CACHE_DIR_ENV] = str(Path(tmp) / "cache")
            try:
                report = convert(pdf, Options(output_dir=Path(tmp), open_when_done=False))
            finally:
                if previous_cache is None:
                    os.environ.pop(CACHE_DIR_ENV, None)
                else:
                    os.environ[CACHE_DIR_ENV] = previous_cache
            if report.output_path is None or not report.output_path.is_file():
                raise RuntimeError("convert() no dejó ningún .xlsx")
            if not report.tables:
                raise RuntimeError(f"No se encontraron tablas en {pdf.name} (warnings={report.warnings})")
            wb = openpyxl.load_workbook(report.output_path)
            table_sheets = [
                ws
                for ws in wb.worksheets
                if ws.title not in RESERVED_SHEET_NAMES and ws.max_row >= 2 and ws.max_column >= 2
            ]
            if not table_sheets:
                raise RuntimeError(f"Ninguna hoja con ≥ 2 filas y ≥ 2 columnas: {wb.sheetnames}")
            n_sheets = len(wb.sheetnames)
        elapsed = time.perf_counter() - t0
        out(f"selftest ok: {n_sheets} hojas, {report.n_rows} filas, ocr={report.ocr_applied}, {elapsed:.1f}s")
        return 0
    except Exception:
        out("selftest FALLO:", err=True)
        out(traceback.format_exc().rstrip(), err=True)
        return 1


# ---------------------------------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------------------------------


def _convert_all(pdfs: list[Path], opts: Options, as_json: bool) -> int:
    from pdf2excel.core.pipeline import analyze, convert, export_batch

    results: list[dict[str, Any]] = []
    analyzed: list[tuple[Path, ConversionReport, float]] = []
    failed = 0
    for pdf in pdfs:
        t0 = time.perf_counter()
        try:
            if opts.single_workbook:
                report = analyze(pdf, opts)
                analyzed.append((pdf, report, time.perf_counter() - t0))
                continue
            report = convert(pdf, opts)
        except UserError as err:
            failed += 1
            elapsed = time.perf_counter() - t0
            if as_json:
                results.append(error_to_json(pdf, err, elapsed))
            else:
                _emit(error_line(pdf, err))
            continue
        except Exception as exc:  # defensa: analyze/convert ya convierten, pero nunca cortamos el lote
            failed += 1
            user_err = from_exception(exc, pdf)
            log.exception("Fallo inesperado con %s", pdf)
            if as_json:
                results.append(error_to_json(pdf, user_err, time.perf_counter() - t0))
            else:
                _emit(error_line(pdf, user_err))
            continue
        elapsed = time.perf_counter() - t0
        if as_json:
            results.append(report_to_json(report, elapsed))
        else:
            _emit(summary_line(report, elapsed))

    if opts.single_workbook and analyzed:
        try:
            export_batch([r for _, r, _ in analyzed], opts)
        except UserError as err:
            failed += len(analyzed)
            for pdf, _, elapsed in analyzed:
                if as_json:
                    results.append(error_to_json(pdf, err, elapsed))
                else:
                    _emit(error_line(pdf, err))
        else:
            for _, report, elapsed in analyzed:
                if as_json:
                    results.append(report_to_json(report, elapsed))
                else:
                    _emit(summary_line(report, elapsed))

    if as_json:
        _emit(json.dumps(results, ensure_ascii=False, indent=2))
    return 1 if failed else 0


def main(argv: list[str] | None = None) -> int:
    """Punto de entrada de la CLI. Devuelve el código de salida (0 ok, 1 algún archivo falló, 2 uso)."""
    parser = build_parser()
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    _setup_logging(args.verbose)

    if args.selftest:
        report_path = Path(args.selftest_report) if args.selftest_report else None
        return selftest(Path(args.selftest), report_path)

    if not args.pdfs:
        parser.print_usage(sys.stderr)
        _emit(f"{PROG}: indicá al menos un PDF (o ejecutá sin argumentos para abrir la ventana).", err=True)
        return 2

    pdfs = [Path(p).expanduser() for p in args.pdfs]
    opts = options_from_args(args)
    return _convert_all(pdfs, opts, args.json)


__all__ = [
    "build_parser",
    "error_line",
    "error_to_json",
    "main",
    "options_from_args",
    "report_to_json",
    "selftest",
    "summary_line",
]
