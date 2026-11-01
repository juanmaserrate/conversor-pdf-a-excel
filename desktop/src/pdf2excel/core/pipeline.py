"""Orquestación del core (DESIGN.md §3 y §5): ``analyze`` → ``export`` (+ ``convert`` y ``export_batch``).

Sin Qt: lo usan la CLI, los tests y los jobs de la UI. ``analyze`` no escribe nada (devuelve un
``ConversionReport`` en memoria para la vista previa); ``export`` arma el .xlsx sin sobrescribir nunca.
Progreso por etapa (``STAGES``) y cancelación cooperativa por ``threading.Event``: cualquier excepción
técnica se convierte en ``UserError`` con ``errors.from_exception``.
"""

from __future__ import annotations

import contextlib
import logging
import tempfile
import threading
import time
from dataclasses import replace
from datetime import datetime
from pathlib import Path

import pymupdf

from pdf2excel import __version__
from pdf2excel.core.candidates import page_candidates
from pdf2excel.core.document import classify_document, needs_ocr, open_document
from pdf2excel.core.layout import extract_layout_page
from pdf2excel.core.merge import merge_multipage
from pdf2excel.core.mode import suggest_mode
from pdf2excel.core.ocr import ocr_available, ocr_document
from pdf2excel.core.quality import select_best
from pdf2excel.core.text_blocks import HiddenSpans, repeated_lines, text_outside_tables, title_near
from pdf2excel.core.typing_ar import type_table
from pdf2excel.core.visibility import hidden_spans
from pdf2excel.errors import UserError, from_exception
from pdf2excel.excel.naming import unique_output_path
from pdf2excel.excel.writer import write_workbook
from pdf2excel.models import (
    ConversionReport,
    EffectiveMode,
    LayoutPage,
    Options,
    PageInfo,
    ProgressFn,
    RawTable,
    TextItem,
    TypedTable,
)

log = logging.getLogger(__name__)

# Avisos (ConversionReport.warnings) que la UI traduce en strings_es.
WARN_NO_TABLES = "no_tables"
WARN_SCANNED = "scanned"
WARN_REVIEW = "review"
WARN_OCR_UNAVAILABLE = "ocr_unavailable"
WARN_LAYOUT_PAGE = "layout_page_failed"
SCANNED_LAYOUT_REASON = (
    "Es un escaneo: el diseño se arma con el texto que se pudo leer (OCR). No se pueden reproducir las cajas "
    "ni los colores, y puede haber errores de lectura. Si preferís una tabla limpia, elegí «Solo tablas»."
)

# Más páginas que esto = listado grande: el diseño por página solo se calcula si el usuario lo pidió.
LAYOUT_MAX_PAGES = 40
_HIDDEN_CACHE_PAGES = 50

BATCH_STEM = "Conversion"  # nombre base del libro único del lote: "Conversion 2026-11-01.xlsx"

Bbox = tuple[float, float, float, float]
_PDF_ENCRYPT_NONE = int(getattr(pymupdf, "PDF_ENCRYPT_NONE", 1))  # los stubs de pymupdf no lo declaran


class _Context:
    """Progreso + cancelación: cada aviso de progreso chequea el ``Event`` justo después."""

    def __init__(self, progress: ProgressFn | None, cancel: threading.Event | None) -> None:
        self.progress = progress
        self.cancel = cancel

    def check(self) -> None:
        if self.cancel is not None and self.cancel.is_set():
            raise UserError("cancelled")

    def report(self, current: int, total: int, stage: str) -> None:
        if self.progress is not None:
            self.progress(current, total, stage)
        self.check()


def _is_encrypted(doc: pymupdf.Document) -> bool:
    """True si el PDF venía cifrado (ya autenticado).

    No se usa ``doc.needs_pass``: en pymupdf 2.0 leerlo DESPUÉS de ``authenticate()`` deja la extracción de
    texto vacía para ese documento. ``metadata["encryption"]`` describe el cifrado sin ese efecto.
    """
    try:
        return bool((doc.metadata or {}).get("encryption"))
    except Exception:  # metadata inaccesible: lo tratamos como no cifrado
        return False


def _decrypted_copy(doc: pymupdf.Document, src: Path, workdir: Path) -> Path:
    """Copia sin cifrado de un PDF ya autenticado (``page_candidates`` reabre el archivo sin contraseña)."""
    target = workdir / f"{src.stem}.decrypted.pdf"
    doc.save(target, encryption=_PDF_ENCRYPT_NONE)
    return target


def _close_quietly(doc: pymupdf.Document) -> None:
    """Cierra el documento sin enmascarar una excepción en curso (p. ej. si ya estaba cerrado)."""
    with contextlib.suppress(Exception):
        doc.close()


def _normalized(text: str) -> str:
    return " ".join(text.split())


def _dedupe_repeated(pages_items: list[list[TextItem]]) -> list[TextItem]:
    """Concatena los ítems de todas las páginas emitiendo encabezados/pies repetidos una sola vez."""
    repeated = repeated_lines(pages_items)
    seen: set[str] = set()
    out: list[TextItem] = []
    for items in pages_items:
        for item in items:
            if item.kind == "text":
                key = _normalized(item.value)
                if key in repeated:
                    if key in seen:
                        continue
                    seen.add(key)
            out.append(item)
    return out


def _run_ocr(
    src: Path,
    doc: pymupdf.Document,
    pages: list[PageInfo],
    to_ocr: list[int],
    opts: Options,
    ctx: _Context,
    password: str | None,
    warnings: list[str],
) -> tuple[pymupdf.Document, Path, bool]:
    """OCR de ``to_ocr`` si hay tesseract; si no, aviso y se sigue solo con las páginas con texto.

    Devuelve ``(doc abierto sobre el PDF de trabajo, ruta de trabajo, ocr_applied)``. El ``doc`` de
    entrada se cierra si se reemplaza.
    """
    if not ocr_available():
        log.warning("%s: %d página(s) necesitan OCR pero no hay tesseract", src.name, len(to_ocr))
        warnings.append(WARN_OCR_UNAVAILABLE)
        return doc, src, False

    work_path = ocr_document(
        src, to_ocr, lang=opts.ocr_lang, progress=ctx.progress, cancel=ctx.cancel, password=password
    )
    ctx.check()
    doc.close()
    work_doc = open_document(work_path)
    fresh = classify_document(work_doc)
    for i in to_ocr:
        if i < len(pages):
            pages[i].kind = "scanned_with_text"
            if i < len(fresh):
                pages[i].n_chars = fresh[i].n_chars
    return work_doc, work_path, True


def _hidden(doc: pymupdf.Document, index: int, cache: dict[int, HiddenSpans]) -> HiddenSpans:
    """Spans tapados/invisibles de la página, calculados una sola vez por análisis."""
    if index in cache:
        return cache[index]
    spans = hidden_spans(doc[index])
    if len(cache) < _HIDDEN_CACHE_PAGES:  # en documentos largos se recalcula: no guardar el trace de todo
        cache[index] = spans
    return spans


def _page_tables(
    doc: pymupdf.Document, work_path: Path, index: int, cache: dict[int, HiddenSpans]
) -> list[RawTable]:
    """Candidatos → mejor candidato → títulos (``title_near``) para una página."""
    cands = page_candidates(work_path, index)
    source, tables = select_best(cands)
    page = doc[index]
    hidden = _hidden(doc, index, cache) if tables else None
    for table in tables:
        table.title = title_near(page, table.bbox, hidden=hidden)
    if tables:
        log.debug("página %d: %d tabla(s) de %s", index + 1, len(tables), source)
    return tables


def _layout_stage(
    doc: pymupdf.Document,
    pages: list[PageInfo],
    skipped_pages: set[int],
    opts: Options,
    ctx: _Context,
    warnings: list[str],
) -> list[LayoutPage]:
    """Una ``LayoutPage`` por página con contenido; vacío si el modo es "tables" o el listado es enorme."""
    n_pages = len(pages)
    if opts.layout == "tables" or (n_pages > LAYOUT_MAX_PAGES and opts.layout != "layout"):
        return []
    out: list[LayoutPage] = []
    for i in range(n_pages):
        ctx.check()
        if i not in skipped_pages:
            try:
                out.append(
                    extract_layout_page(
                        doc[i],
                        embed_images=opts.layout_images,
                        scanned=pages[i].kind == "scanned_with_text",
                        typed=opts.typed_values,
                    )
                )
            except Exception:  # un diseño imposible no debe tirar el modo tablas
                log.exception("No se pudo armar el diseño de la página %d", i + 1)
                if WARN_LAYOUT_PAGE not in warnings:
                    warnings.append(WARN_LAYOUT_PAGE)
        ctx.report(i + 1, n_pages, "layout")
    return out


def _unrotated_bboxes(doc: pymupdf.Document, by_page: dict[int, list[Bbox]]) -> dict[int, list[Bbox]]:
    """Bboxes de tabla (marco de la página rotada, el de ``find_tables``) → marco de ``LayoutCell.bbox``."""
    out: dict[int, list[Bbox]] = {}
    for index, boxes in by_page.items():
        if not boxes:
            continue
        matrix = doc[index].derotation_matrix
        rects = [pymupdf.Rect(b) * matrix for b in boxes]
        out[index] = [(r.x0, r.y0, r.x1, r.y1) for r in rects]
    return out


def _choose_mode(report: ConversionReport, opts: Options, table_bboxes: dict[int, list[Bbox]]) -> None:
    """Completa ``suggested_mode``/``mode``/``mode_reason`` según ``Options.layout`` y lo que se encontró."""
    n_pages = len(report.pages)
    suggested, reason = suggest_mode(report, table_bboxes)
    if opts.layout == "tables":
        suggested, reason = "tables", "Elegiste exportar solo las tablas de datos."
    elif n_pages > LAYOUT_MAX_PAGES and opts.layout != "layout":
        suggested, reason = (
            "tables",
            "Es un documento muy largo (parece un listado): conviene una tabla limpia.",
        )
    mode: EffectiveMode = suggested if opts.layout == "auto" else opts.layout
    if mode == "layout" and not report.layout_pages:
        mode = "tables"  # nada que dibujar (p. ej. escaneo sin OCR): las tablas son lo único posible
        if opts.layout == "layout":
            reason = "No se pudo reproducir el diseño de este PDF: se exportan las tablas de datos."
    if mode == "layout" and any(p.kind == "scanned_with_text" for p in report.pages):
        reason = SCANNED_LAYOUT_REASON  # el escaneo no trae cajas ni colores: avisar qué se puede esperar
    report.suggested_mode = suggested
    report.mode = mode
    report.mode_reason = reason


def _analyze(
    src: Path, opts: Options, ctx: _Context, password: str | None, stack: contextlib.ExitStack
) -> ConversionReport:
    t_start = time.perf_counter()
    timings: dict[str, float] = {}
    warnings: list[str] = []

    # 1. Apertura y clasificación.
    doc = open_document(src, password)
    doc_holder: list[pymupdf.Document] = [doc]  # el doc puede reemplazarse tras el OCR
    stack.callback(lambda: _close_quietly(doc_holder[0]))
    pages = classify_document(doc)
    n_pages = len(pages)
    timings["classify"] = time.perf_counter() - t_start
    ctx.report(n_pages, n_pages, "classify")

    work_path = src
    if _is_encrypted(doc):
        # Autenticado con contraseña: page_candidates reabre el archivo → copia sin cifrar en un tmpdir.
        workdir = Path(
            stack.enter_context(tempfile.TemporaryDirectory(prefix="pdf2excel-", ignore_cleanup_errors=True))
        )
        work_path = _decrypted_copy(doc, src, workdir)

    # 2. OCR (una vez por documento) solo si hace falta.
    ocr_applied = False
    skipped_pages: set[int] = set()
    to_ocr = [p.index for p in pages if needs_ocr(p, opts.force_ocr)]
    if to_ocr:
        t_ocr = time.perf_counter()
        doc, work_path, ocr_applied = _run_ocr(src, doc, pages, to_ocr, opts, ctx, password, warnings)
        doc_holder[0] = doc
        if ocr_applied:
            timings["ocr"] = time.perf_counter() - t_ocr
        else:
            # Sin tesseract se saltean solo las páginas que de verdad no tienen texto confiable (escaneo,
            # texto basura); con force_ocr las digitales siguen con su texto propio.
            skipped_pages = {p.index for p in pages if p.index in to_ocr and needs_ocr(p)}
    skipped_pages |= {p.index for p in pages if p.kind == "empty"}

    # 3. Tablas por página.
    t_tables = time.perf_counter()
    all_tables: list[RawTable] = []
    bboxes_by_page: dict[int, list[Bbox]] = {}
    hidden_cache: dict[int, HiddenSpans] = {}
    for i in range(n_pages):
        ctx.check()
        tables = [] if i in skipped_pages else _page_tables(doc, work_path, i, hidden_cache)
        all_tables.extend(tables)
        bboxes_by_page[i] = [t.bbox for t in tables]
        ctx.report(i + 1, n_pages, "tables")
    timings["tables"] = time.perf_counter() - t_tables

    # 3b. Diseño por página (modo "Igual al PDF"); un "Solo tablas" explícito no paga este costo.
    t_layout = time.perf_counter()
    layout_pages = _layout_stage(doc, pages, skipped_pages, opts, ctx, warnings)
    if opts.layout != "tables":
        timings["layout"] = time.perf_counter() - t_layout

    # 4. Merge multipágina.
    raw = (
        merge_multipage(all_tables)
        if opts.merge_continuations and not opts.one_sheet_per_page
        else all_tables
    )

    # 5. Tipado por columna.
    t_typing = time.perf_counter()
    typed: list[TypedTable] = []
    if raw:
        for k, table in enumerate(raw, start=1):
            typed.append(type_table(table, opts.typed_values))
            ctx.report(k, len(raw), "typing")
    else:
        ctx.report(1, 1, "typing")
    timings["typing"] = time.perf_counter() - t_typing

    # 6. Texto fuera de tablas (encabezados/pies repetidos una sola vez).
    text_items: list[TextItem] = []
    if opts.include_text:
        t_text = time.perf_counter()
        pages_items = [
            []
            if i in skipped_pages
            else text_outside_tables(doc[i], bboxes_by_page.get(i, []), _hidden(doc, i, hidden_cache))
            for i in range(n_pages)
        ]
        text_items = _dedupe_repeated(pages_items)
        timings["text"] = time.perf_counter() - t_text
        ctx.check()

    # 7. Avisos y reporte.
    if not typed:
        warnings.append(WARN_NO_TABLES)
    if ocr_applied:
        warnings.append(WARN_SCANNED)
    if any(t.review_cells for t in typed):
        warnings.append(WARN_REVIEW)
    timings["total"] = time.perf_counter() - t_start
    log.info(
        "%s: %d pág., %d tabla(s), %d fila(s), ocr=%s, %.2f s",
        src.name,
        n_pages,
        len(typed),
        sum(len(t.rows) for t in typed),
        ocr_applied,
        timings["total"],
    )
    report = ConversionReport(
        source=src,
        pages=pages,
        tables=typed,
        text_items=text_items,
        warnings=warnings,
        timings=timings,
        ocr_applied=ocr_applied,
        output_path=None,
        app_version=__version__,
        layout_pages=layout_pages,
    )
    _choose_mode(report, opts, _unrotated_bboxes(doc, bboxes_by_page))
    return report


def analyze(
    src: Path,
    opts: Options,
    progress: ProgressFn | None = None,
    cancel: threading.Event | None = None,
    password: str | None = None,
) -> ConversionReport:
    """Analiza ``src`` sin escribir nada: páginas, OCR si hace falta, tablas tipadas y texto.

    ``progress(actual, total, etapa)`` recibe las etapas ``classify`` → (``ocr``) → ``tables`` → (``layout``)
    → ``typing``.
    Si ``cancel`` se activa, lanza ``UserError("cancelled")`` en el siguiente punto de control. Cualquier
    otra excepción se convierte con ``from_exception``.
    """
    src = Path(src)
    ctx = _Context(progress, cancel)
    ctx.check()
    try:
        with contextlib.ExitStack() as stack:
            return _analyze(src, opts, ctx, password, stack)
    except UserError:
        raise
    except Exception as exc:
        log.exception("Fallo analizando %s", src)
        raise from_exception(exc, src) from exc


def export(report: ConversionReport, opts: Options, dst: Path | None = None) -> Path:
    """Escribe el .xlsx (``<pdf>.xlsx``, ``<pdf> (2).xlsx``…) y completa ``report.output_path``.

    ``dst`` es un destino explícito ("Guardar con otro nombre"): se respeta tal cual, sin pasar por
    ``unique_output_path`` (el diálogo nativo ya confirmó el nombre con el usuario).
    """
    dst = Path(dst) if dst is not None else unique_output_path(report.source, opts.output_dir)
    try:
        write_workbook(report, dst, opts)
    except UserError:
        raise
    except Exception as exc:
        raise from_exception(exc, dst) from exc
    report.output_path = dst
    log.info("%s → %s", Path(report.source).name, dst)
    return dst


def convert(
    src: Path,
    opts: Options,
    progress: ProgressFn | None = None,
    cancel: threading.Event | None = None,
    password: str | None = None,
) -> ConversionReport:
    """``analyze`` + ``export`` en un paso (CLI y lotes sin vista previa)."""
    ctx = _Context(progress, cancel)
    report = analyze(src, opts, progress, cancel, password)
    ctx.check()
    t_excel = time.perf_counter()
    export(report, opts)
    report.timings["excel"] = time.perf_counter() - t_excel
    ctx.report(1, 1, "excel")
    return report


def _batch_tables(report: ConversionReport, page_offset: int) -> list[TypedTable]:
    stem = Path(report.source).stem
    out: list[TypedTable] = []
    for i, table in enumerate(report.tables, start=1):
        name = f"{stem} · {table.name or f'Tabla {i}'}"
        first, last = table.pages
        out.append(replace(table, name=name, pages=(first + page_offset, last + page_offset)))
    return out


def _batch_text(report: ConversionReport, page_offset: int) -> list[TextItem]:
    if not report.text_items:
        return []
    header = TextItem(page=page_offset, kind="kv", key="Archivo", value=Path(report.source).name)
    return [header, *(replace(it, page=it.page + page_offset) for it in report.text_items)]


def _uses_layout(report: ConversionReport) -> bool:
    """¿Este reporte se exporta como "Igual al PDF"? (necesita páginas de diseño que dibujar)."""
    return report.mode == "layout" and bool(report.layout_pages)


def export_batch(reports: list[ConversionReport], opts: Options) -> Path:
    """Un solo libro ``Conversion <fecha>.xlsx`` con todas las tablas del lote (``Options.single_workbook``).

    Las hojas se llaman ``<pdf> · <hoja>``; el texto de cada PDF va precedido por un ``Archivo: nombre``.
    Cada reporte respeta su ``mode``: los que están en "layout" aportan una hoja por página
    (``<pdf> · Página N``) y el resto sus tablas, así que el lote puede mezclar modos.
    Se guarda en ``opts.output_dir`` o, si no hay, en la carpeta del primer PDF. Completa
    ``output_path`` en cada reporte.
    """
    if not reports:
        raise UserError("unknown", detail="No hay nada para exportar: el lote está vacío.")
    first_source = Path(reports[0].source)
    pages: list[PageInfo] = []
    tables: list[TypedTable] = []
    text_items: list[TextItem] = []
    warnings: list[str] = []
    timings: dict[str, float] = {}
    offset = 0
    layout_reports: list[tuple[str, ConversionReport]] = []
    for report in reports:
        pages.extend(replace(p, index=p.index + offset) for p in report.pages)
        if _uses_layout(report):
            layout_reports.append((Path(report.source).stem, report))
        else:
            tables.extend(_batch_tables(report, offset))
            text_items.extend(_batch_text(report, offset))
        warnings.extend(w for w in report.warnings if w not in warnings)
        for stage, seconds in report.timings.items():
            timings[stage] = timings.get(stage, 0.0) + seconds
        offset += len(report.pages)

    stamp = datetime.now().strftime("%Y-%m-%d")
    directory = Path(opts.output_dir) if opts.output_dir is not None else first_source.parent
    dst = unique_output_path(directory / f"{BATCH_STEM} {stamp}.pdf", None)
    combined = ConversionReport(
        source=directory / dst.stem,  # el writer solo usa el nombre (title / hoja Origen); no se abre
        pages=pages,
        tables=tables,
        text_items=text_items,
        warnings=warnings,
        timings=timings,
        ocr_applied=any(r.ocr_applied for r in reports),
        output_path=None,
        app_version=__version__,
        mode="layout" if len(layout_reports) == len(reports) else "tables",
    )
    try:
        write_workbook(combined, dst, opts, layout_reports)
    except UserError:
        raise
    except Exception as exc:
        raise from_exception(exc, dst) from exc
    for report in reports:
        report.output_path = dst
    log.info("lote de %d archivo(s) → %s", len(reports), dst)
    return dst


__all__ = [
    "BATCH_STEM",
    "WARN_NO_TABLES",
    "WARN_OCR_UNAVAILABLE",
    "WARN_REVIEW",
    "WARN_SCANNED",
    "analyze",
    "convert",
    "export",
    "export_batch",
]
