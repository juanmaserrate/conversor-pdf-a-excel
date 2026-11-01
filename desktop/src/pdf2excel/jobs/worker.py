"""Trabajos en segundo plano: ``AnalyzeJob`` / ``ExportJob`` (``QRunnable``) y ``JobQueue`` (``QThreadPool``).

Es el ÚNICO lugar de la capa Qt que llama al core, y lo hace con importes diferidos dentro de ``run()``
para que la ventana aparezca sin cargar pymupdf/onnx. Nunca toca widgets: todo sale por ``JobSignals``
(conexiones encoladas al hilo principal). Cancelación cooperativa con un ``threading.Event`` por lote.
"""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal

from pdf2excel.errors import UserError, from_exception
from pdf2excel.models import ConversionReport, Options

log = logging.getLogger(__name__)

DoneCallback = Callable[[], None]


class JobSignals(QObject):
    """Señales de los trabajos. Vive en el hilo principal; los jobs emiten desde el pool."""

    progress = Signal(str, int, int, str)  # (path, actual, total, etapa)
    analyzed = Signal(str, object)  # (path, ConversionReport)
    exported = Signal(str, object)  # (path, ConversionReport con output_path)
    failed = Signal(str, object)  # (path, UserError)
    all_done = Signal()


def _emit_quietly(signal: Any, *args: object) -> None:
    """``signal.emit(*args)`` tolerando que ``JobSignals`` ya no exista.

    Pasa cuando la app cerró con un trabajo que no honró la cancelación a tiempo: la ventana (y con ella la
    cola) ya se destruyó. No es un fallo del archivo, así que no se loguea como tal.
    """
    try:
        signal.emit(*args)
    except RuntimeError as exc:  # "Signal source has been deleted"
        log.debug("Señal descartada (la cola ya no existe): %s", exc)


class _BaseJob(QRunnable):
    """Esqueleto común: contabilidad de inicio/fin y conversión de excepciones a ``UserError``."""

    def __init__(
        self,
        path: str,
        signals: JobSignals,
        cancel: threading.Event,
        on_start: DoneCallback | None = None,
        on_done: DoneCallback | None = None,
    ) -> None:
        super().__init__()
        self.setAutoDelete(True)
        self.path = path
        self.signals = signals
        self.cancel = cancel
        self._on_start = on_start
        self._on_done = on_done

    def run(self) -> None:
        if self._on_start is not None:
            self._on_start()
        try:
            if self.cancel.is_set():
                # Se canceló el lote antes de empezar: no hay nada que informar.
                return
            self._work()
        except UserError as err:
            if err.code != "cancelled":
                log.info("%s: %s (%s)", self.path, err.code, err.title)
            self._emit(self.signals.failed, self.path, err)
        except BaseException as exc:
            if self.cancel.is_set():
                # La app cerró (o canceló) con el trabajo a medio camino: no es un fallo del archivo.
                log.info("%s: interrumpido tras la cancelación (%r)", self.path, exc)
                self._emit(self.signals.failed, self.path, UserError("cancelled", cause=exc))
            else:
                log.exception("Fallo inesperado procesando %s", self.path)
                self._emit(self.signals.failed, self.path, from_exception(exc, Path(self.path)))
        finally:
            if self._on_done is not None:
                self._on_done()

    def _emit(self, signal: Any, *args: object) -> None:
        _emit_quietly(signal, *args)

    def _work(self) -> None:  # pragma: no cover - lo implementan las subclases
        raise NotImplementedError


class AnalyzeJob(_BaseJob):
    """Analiza un PDF (sin escribir nada) y emite ``analyzed(path, report)``."""

    def __init__(
        self,
        path: str,
        opts: Options,
        signals: JobSignals,
        cancel: threading.Event,
        password: str | None = None,
        on_start: DoneCallback | None = None,
        on_done: DoneCallback | None = None,
    ) -> None:
        super().__init__(path, signals, cancel, on_start, on_done)
        self.opts = opts
        self.password = password

    def _work(self) -> None:
        from pdf2excel.core.pipeline import analyze  # import diferido: no cargar el motor al arrancar

        path = self.path
        progress = self.signals.progress

        def report_progress(current: int, total: int, stage: str) -> None:
            _emit_quietly(progress, path, int(current), int(total), str(stage))

        report = analyze(Path(path), self.opts, report_progress, self.cancel, self.password)
        self._emit(self.signals.analyzed, path, report)


class ExportJob(_BaseJob):
    """Escribe el Excel de un ``ConversionReport`` ya analizado y emite ``exported(path, report)``.

    ``dst`` es un destino explícito ("Guardar con otro nombre"): se pasa tal cual a ``pipeline.export``.
    """

    def __init__(
        self,
        report: ConversionReport,
        opts: Options,
        signals: JobSignals,
        cancel: threading.Event,
        on_start: DoneCallback | None = None,
        on_done: DoneCallback | None = None,
        dst: Path | None = None,
    ) -> None:
        super().__init__(str(report.source), signals, cancel, on_start, on_done)
        self.report = report
        self.opts = opts
        self.dst = dst

    def _work(self) -> None:
        from pdf2excel.core.pipeline import export  # import diferido

        self._emit(self.signals.progress, self.path, 0, 1, "excel")
        if self.dst is not None:
            output = export(self.report, self.opts, dst=self.dst)
        else:
            output = export(self.report, self.opts)
        if output is not None:
            self.report.output_path = Path(output)
        self._emit(self.signals.progress, self.path, 1, 1, "excel")
        self._emit(self.signals.exported, self.path, self.report)


class BatchExportJob(ExportJob):
    """Un solo libro para todo el lote (``Options.single_workbook``): ``pipeline.export_batch`` una vez.

    Emite ``progress``/``exported`` por cada reporte, así la ventana trata el lote igual que a N exports.
    """

    def __init__(
        self,
        reports: list[ConversionReport],
        opts: Options,
        signals: JobSignals,
        cancel: threading.Event,
        on_start: DoneCallback | None = None,
        on_done: DoneCallback | None = None,
    ) -> None:
        if not reports:
            raise ValueError("BatchExportJob necesita al menos un reporte")
        super().__init__(reports[0], opts, signals, cancel, on_start, on_done)
        self.reports = list(reports)

    def _work(self) -> None:
        from pdf2excel.core.pipeline import export_batch  # import diferido

        for report in self.reports:
            self._emit(self.signals.progress, str(report.source), 0, 1, "excel")
        export_batch(self.reports, self.opts)
        for report in self.reports:
            self._emit(self.signals.progress, str(report.source), 1, 1, "excel")
            self._emit(self.signals.exported, str(report.source), report)


class JobQueue(QObject):
    """Cola de trabajos sobre un ``QThreadPool`` propio (2 hilos por defecto).

    Cuenta los trabajos pendientes y emite ``signals.all_done`` cuando el último termina.
    ``cancel_all()`` marca el ``Event`` del lote, vacía la cola y prepara un ``Event`` nuevo.
    """

    def __init__(self, max_threads: int = 2, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.signals = JobSignals(self)
        self._pool = QThreadPool(self)
        self._pool.setMaxThreadCount(max(1, max_threads))
        self._cancel = threading.Event()
        self._lock = threading.Lock()
        self._pending = 0  # encolados o corriendo
        self._running = 0  # corriendo

    # ------------------------------------------------------------------ contabilidad

    def _on_start(self) -> None:
        with self._lock:
            self._running += 1

    def _on_done(self) -> None:
        with self._lock:
            self._running = max(0, self._running - 1)
            finished = self._pending > 0
            self._pending = max(0, self._pending - 1)
            emit_done = finished and self._pending == 0
        if emit_done:
            self.signals.all_done.emit()

    def _start(self, jobs: list[QRunnable]) -> None:
        if not jobs:
            # Nada que hacer: avisamos igual para que la UI no quede esperando.
            self.signals.all_done.emit()
            return
        with self._lock:
            self._pending += len(jobs)
        for job in jobs:
            self._pool.start(job)

    # ------------------------------------------------------------------ API pública

    def submit_analyze(
        self, paths: list[str], opts: Options, passwords: dict[str, str] | None = None
    ) -> None:
        pwds = passwords or {}
        cancel = self._cancel
        jobs: list[QRunnable] = [
            AnalyzeJob(
                p,
                opts,
                self.signals,
                cancel,
                password=pwds.get(p),
                on_start=self._on_start,
                on_done=self._on_done,
            )
            for p in paths
        ]
        self._start(jobs)

    def submit_export(self, reports: list[ConversionReport], opts: Options, dst: Path | None = None) -> None:
        """Un ``ExportJob`` por reporte; ``dst`` (destino explícito) solo tiene sentido con un reporte."""
        cancel = self._cancel
        jobs: list[QRunnable] = [
            ExportJob(r, opts, self.signals, cancel, on_start=self._on_start, on_done=self._on_done, dst=dst)
            for r in reports
        ]
        self._start(jobs)

    def submit_batch_export(self, reports: list[ConversionReport], opts: Options) -> None:
        """Un único ``BatchExportJob`` para todo el lote (libro único); sin reportes emite ``all_done``."""
        if not reports:
            self._start([])
            return
        job = BatchExportJob(
            reports, opts, self.signals, self._cancel, on_start=self._on_start, on_done=self._on_done
        )
        self._start([job])

    def cancel_all(self) -> None:
        """Cancela el lote: los que corren paran en el próximo chequeo; los encolados no arrancan."""
        with self._lock:
            self._cancel.set()
            self._pool.clear()  # los QRunnable encolados (autoDelete) se descartan y nunca corren
            had_pending = self._pending > 0
            self._pending = self._running
            emit_done = had_pending and self._pending == 0
            self._cancel = threading.Event()  # listo para el próximo lote
        if emit_done:
            self.signals.all_done.emit()

    def is_busy(self) -> bool:
        with self._lock:
            return self._pending > 0

    def pending_count(self) -> int:
        with self._lock:
            return self._pending

    def wait_for_done(self, msecs: int = -1) -> bool:
        """Bloquea hasta que terminen los trabajos que están corriendo (para cerrar la app o tests)."""
        return self._pool.waitForDone(msecs)


__all__ = ["AnalyzeJob", "BatchExportJob", "ExportJob", "JobQueue", "JobSignals"]
