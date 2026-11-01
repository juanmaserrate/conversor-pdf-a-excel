"""OCR propio, sin Ghostscript ni ocrmypdf (DESIGN.md §5.2).

Por cada página escaneada: render a 300 dpi → orientación (``tesseract --psm 0 -l osd``) → deskew por perfil
de proyección → quitar líneas de grilla (OpenCV) → ``tesseract --psm 4 -c textonly_pdf=1`` → página nueva
con la imagen deskewed + capa de texto invisible. El resultado se cachea por SHA-1 del archivo origen.

Este módulo no importa Qt. La cancelación es cooperativa (``threading.Event``) y mata el subproceso de
tesseract en curso.
"""

from __future__ import annotations

import contextlib
import hashlib
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from collections.abc import Sequence
from pathlib import Path

import cv2
import numpy as np
import platformdirs
import pymupdf

from pdf2excel import APP_ID
from pdf2excel.core.document import file_sha1, open_document
from pdf2excel.errors import UserError
from pdf2excel.models import PageInfo, ProgressFn
from pdf2excel.resources_path import tessdata_dir, tesseract_dir

log = logging.getLogger(__name__)

TESSERACT_ENV = "PDF2EXCEL_TESSERACT"  # ruta explícita al ejecutable (tiene prioridad)
CACHE_DIR_ENV = "PDF2EXCEL_CACHE_DIR"  # carpeta de caché de OCR alternativa (selftest, tests, CI)
CACHE_MAX_BYTES = 500 * 1024 * 1024  # la caché de OCR no pasa de 500 MB…
CACHE_MAX_AGE_DAYS = 30  # …ni guarda PDFs que nadie tocó en 30 días
DEFAULT_LANG = "spa+eng"
DEFAULT_DPI = 300
DEFAULT_PSM = 4  # ganó a 6 y 3 en el experimento (98,8 vs 98,4 vs 95,0 %)
OSD_DPI = 150
MIN_DESKEW_DEG = 0.2  # por debajo de esto no vale la pena rotar (el blur empeora más de lo que corrige)
JPEG_QUALITY = 80
_POLL_SECONDS = 0.2
_OSD_TIMEOUT = 60.0
_WINDOWS = sys.platform.startswith("win")


# ---------------------------------------------------------------------------------------------------
# Localización de tesseract
# ---------------------------------------------------------------------------------------------------


def find_tesseract() -> Path | None:
    """Ruta al ejecutable de tesseract o ``None`` si no hay.

    Orden: variable de entorno ``PDF2EXCEL_TESSERACT`` (si está definida es la única que se mira),
    el binario vendoreado en ``resources/tesseract/`` y, por último, el ``PATH`` del sistema.
    """
    override = os.environ.get(TESSERACT_ENV)
    if override is not None and override.strip():
        candidate = Path(override.strip())
        return candidate if candidate.is_file() else None

    exe_name = "tesseract.exe" if _WINDOWS else "tesseract"
    bundled = tesseract_dir() / exe_name
    if bundled.is_file():
        return bundled

    found = shutil.which("tesseract")
    return Path(found) if found else None


def ocr_available() -> bool:
    """True si hay un tesseract utilizable."""
    return find_tesseract() is not None


def tessdata_prefix() -> Path | None:
    """``resources/tessdata`` si trae ``spa.traineddata``; si no ``None`` (se usa la tessdata del sistema)."""
    directory = tessdata_dir()
    return directory if (directory / "spa.traineddata").is_file() else None


def _tesseract_env() -> dict[str, str]:
    env = dict(os.environ)
    prefix = tessdata_prefix()
    if prefix is not None:
        env["TESSDATA_PREFIX"] = str(prefix)
    env["OMP_THREAD_LIMIT"] = "2"  # dos archivos en paralelo en el QThreadPool: no saturar la CPU
    return env


def _creation_flags() -> int:
    """``CREATE_NO_WINDOW`` en Windows (no abrir una consola por cada tesseract); 0 en el resto."""
    return int(getattr(subprocess, "CREATE_NO_WINDOW", 0)) if _WINDOWS else 0


def _check_cancel(cancel: threading.Event | None) -> None:
    if cancel is not None and cancel.is_set():
        raise UserError("cancelled")


# ---------------------------------------------------------------------------------------------------
# Imagen: render, deskew, líneas
# ---------------------------------------------------------------------------------------------------


def render_gray(page: pymupdf.Page, dpi: int = DEFAULT_DPI) -> np.ndarray:
    """Render en escala de grises (uint8, alto × ancho) de la página a ``dpi``."""
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8)
    if pix.n != 1:  # por las dudas: algunos renders traen más canales
        arr = arr.reshape(pix.height, pix.width, pix.n)[:, :, 0]
    return arr.reshape(pix.height, pix.width).copy()


def _binarize_inv(gray: np.ndarray) -> np.ndarray:
    """Otsu invertido: tinta = 255, papel = 0."""
    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    return bw


def deskew_angle(gray: np.ndarray, max_deg: float = 4.0, step: float = 0.1) -> float:
    """Ángulo (grados) que endereza los renglones: el que maximiza la varianza del perfil de filas.

    Trabaja a 1/4 de escala sobre la imagen binarizada (Otsu). Rotar la imagen con ``rotate_gray`` por el
    ángulo devuelto deja el texto horizontal: una página inclinada +1,2° devuelve ≈ −1,2.
    """
    if gray.size == 0:
        return 0.0
    small = cv2.resize(gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    bw = _binarize_inv(small)
    h, w = bw.shape
    center = (w / 2.0, h / 2.0)
    best_angle = 0.0
    best_var = -1.0
    n_steps = round(2 * max_deg / step)
    for i in range(n_steps + 1):
        angle = -max_deg + i * step
        matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
        rotated = cv2.warpAffine(bw, matrix, (w, h), flags=cv2.INTER_NEAREST, borderValue=0)
        variance = float(rotated.sum(axis=1, dtype=np.float64).var())
        if variance > best_var:
            best_var = variance
            best_angle = angle
    return round(best_angle, 3)


def rotate_gray(gray: np.ndarray, angle: float) -> np.ndarray:
    """Rota la imagen ``angle`` grados (sentido antihorario positivo) con borde blanco, cúbica."""
    h, w = gray.shape[:2]
    matrix = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), angle, 1.0)
    return cv2.warpAffine(gray, matrix, (w, h), flags=cv2.INTER_CUBIC, borderValue=255)


def remove_grid_lines(gray: np.ndarray) -> np.ndarray:
    """Copia de la imagen con las líneas largas (bordes de tabla) pintadas de blanco.

    Sin esto tesseract devuelve basura en tablas con bordes ("favor CECI CO ESA"). Apertura morfológica con
    kernels ``(k,1)`` y ``(1,k)``, ``k = max(40, ancho // 30)``; la máscara se dilata 3×3.
    """
    bw = _binarize_inv(gray)
    k = max(40, gray.shape[1] // 30)
    horizontal = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (k, 1)))
    vertical = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, k)))
    mask = cv2.dilate(cv2.bitwise_or(horizontal, vertical), np.ones((3, 3), np.uint8))
    out = gray.copy()
    out[mask > 0] = 255
    return out


def _write_png(path: Path, img: np.ndarray) -> None:
    ok, buf = cv2.imencode(".png", img, [cv2.IMWRITE_PNG_COMPRESSION, 1])
    if not ok:
        raise UserError("ocr_failed", detail="No se pudo codificar la imagen de la página como PNG.")
    path.write_bytes(buf.tobytes())


def _encode_jpeg(img: np.ndarray, quality: int = JPEG_QUALITY) -> bytes:
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, quality])
    if not ok:
        raise UserError("ocr_failed", detail="No se pudo codificar la imagen de la página como JPEG.")
    return buf.tobytes()


# ---------------------------------------------------------------------------------------------------
# Tesseract
# ---------------------------------------------------------------------------------------------------


def detect_orientation(gray: np.ndarray, src_dpi: int = DEFAULT_DPI) -> int:
    """Rotación (0/90/180/270) que hay que aplicar en sentido horario para enderezar la página.

    Usa ``tesseract --psm 0 -l osd`` sobre un render a ~150 dpi. Cualquier problema (sin tesseract, sin
    ``osd.traineddata``, página sin texto) devuelve 0: nunca frena la conversión.
    """
    exe = find_tesseract()
    if exe is None or gray.size == 0:
        return 0
    scale = OSD_DPI / float(src_dpi) if src_dpi > 0 else 0.5
    small = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) if scale < 1 else gray
    try:
        with tempfile.TemporaryDirectory(prefix="pdf2excel-osd-") as tmp:
            png = Path(tmp) / "osd.png"
            _write_png(png, small)
            cmd = [str(exe), str(png), "-", "--psm", "0", "-l", "osd", "--dpi", str(OSD_DPI)]
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=_OSD_TIMEOUT,
                env=_tesseract_env(),
                creationflags=_creation_flags(),
                check=False,
            )
        for line in (result.stdout + "\n" + result.stderr).splitlines():
            if line.startswith("Rotate:"):
                value = int(line.split(":", 1)[1].strip())
                return value % 360 if value % 90 == 0 else 0
    except Exception:
        log.debug("detect_orientation falló; se asume 0°", exc_info=True)
    return 0


def run_tesseract_textonly_pdf(
    png: Path,
    out_base: Path,
    lang: str,
    psm: int = DEFAULT_PSM,
    dpi: int = DEFAULT_DPI,
    timeout: float = 120.0,
    cancel: threading.Event | None = None,
) -> Path:
    """Corre ``tesseract png out_base -l lang --psm N --dpi D -c textonly_pdf=1 pdf`` y devuelve el PDF.

    Hace ``poll`` cada 0,2 s para honrar ``cancel`` (mata el proceso → ``UserError("cancelled")``).
    Código de salida ≠ 0 o timeout → ``UserError("ocr_failed")`` con el stderr recortado.
    """
    exe = find_tesseract()
    if exe is None:
        raise UserError("ocr_unavailable")
    _check_cancel(cancel)
    cmd = [
        str(exe),
        str(png),
        str(out_base),
        "-l",
        lang,
        "--psm",
        str(psm),
        "--dpi",
        str(dpi),
        "-c",
        "textonly_pdf=1",
        "pdf",
    ]
    log.debug("tesseract: %s", " ".join(cmd))
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=_tesseract_env(),
            creationflags=_creation_flags(),
        )
    except FileNotFoundError as exc:
        raise UserError("ocr_unavailable", cause=exc) from exc
    except OSError as exc:
        raise UserError("ocr_failed", detail=f"No se pudo ejecutar tesseract: {exc}", cause=exc) from exc

    started = time.monotonic()
    try:
        while proc.poll() is None:
            if cancel is not None and cancel.is_set():
                _kill(proc)
                raise UserError("cancelled")
            if time.monotonic() - started > timeout:
                _kill(proc)
                raise UserError(
                    "ocr_failed",
                    detail=f"tesseract no terminó en {timeout:.0f} s; la página puede ser demasiado grande.",
                )
            time.sleep(_POLL_SECONDS)
        stderr = _read_stream(proc)
    finally:
        _close_streams(proc)

    if proc.returncode != 0:
        detail = stderr.strip()[:300] or f"tesseract terminó con código {proc.returncode}."
        raise UserError("ocr_failed", detail=detail)

    pdf = out_base.parent / (out_base.name + ".pdf")
    if not pdf.is_file():
        raise UserError("ocr_failed", detail="tesseract terminó bien pero no generó el PDF de texto.")
    return pdf


def _kill(proc: subprocess.Popen[bytes]) -> None:
    try:
        proc.kill()
        proc.wait(timeout=5)
    except Exception:
        log.debug("No se pudo matar tesseract (pid %s)", proc.pid, exc_info=True)


def _read_stream(proc: subprocess.Popen[bytes]) -> str:
    try:
        _, err = proc.communicate(timeout=5)
    except Exception:
        return ""
    return err.decode("utf-8", errors="replace") if err else ""


def _close_streams(proc: subprocess.Popen[bytes]) -> None:
    for stream in (proc.stdout, proc.stderr):
        if stream is not None:
            with contextlib.suppress(Exception):
                stream.close()


# ---------------------------------------------------------------------------------------------------
# Página y documento
# ---------------------------------------------------------------------------------------------------


def ocr_page_into(
    src_page: pymupdf.Page,
    dst_doc: pymupdf.Document,
    lang: str = DEFAULT_LANG,
    psm: int = DEFAULT_PSM,
    dpi: int = DEFAULT_DPI,
    cancel: threading.Event | None = None,
    *,
    workdir: Path,
) -> PageInfo:
    """Agrega a ``dst_doc`` una página nueva con la imagen enderezada + la capa de texto invisible.

    Pasos: render gris → orientación OSD (``np.rot90``) → deskew → PNG sin líneas de grilla → tesseract →
    ``insert_image`` (JPEG q80 de la imagen deskewed **con** líneas: el layout ML las aprovecha) →
    ``show_pdf_page`` del PDF ``textonly``. Devuelve el ``PageInfo`` de la página nueva.
    """
    index = src_page.number
    t0 = time.perf_counter()
    gray = render_gray(src_page, dpi=dpi)
    _check_cancel(cancel)

    rot = detect_orientation(gray, src_dpi=dpi)
    if rot:
        gray = np.ascontiguousarray(np.rot90(gray, k=-(rot // 90)))  # OSD "Rotate: N" = N° horario
    _check_cancel(cancel)

    angle = deskew_angle(gray)
    if abs(angle) >= MIN_DESKEW_DEG:
        gray = rotate_gray(gray, angle)
    _check_cancel(cancel)

    ocr_img = remove_grid_lines(gray)
    png = Path(workdir) / f"p{index}.png"
    _write_png(png, ocr_img)
    text_pdf = run_tesseract_textonly_pdf(
        png, Path(workdir) / f"t{index}", lang=lang, psm=psm, dpi=dpi, cancel=cancel
    )

    width, height = src_page.rect.width, src_page.rect.height
    if rot in (90, 270):
        width, height = height, width
    new_page = dst_doc.new_page(width=width, height=height)
    new_page.insert_image(new_page.rect, stream=_encode_jpeg(gray))
    with pymupdf.open(text_pdf) as text_doc:
        new_page.show_pdf_page(new_page.rect, text_doc, 0)

    n_chars = len(new_page.get_text().strip())
    log.info(
        "OCR página %d: rot=%d° deskew=%.1f° chars=%d en %.2f s",
        index + 1,
        rot,
        angle,
        n_chars,
        time.perf_counter() - t0,
    )
    return PageInfo(index=index, kind="scanned_with_text", n_chars=n_chars, image_coverage=1.0, rotation=rot)


def default_cache_dir() -> Path:
    """``user_cache_dir("ConversorPDFaExcel")/ocr``, o ``$PDF2EXCEL_CACHE_DIR/ocr`` si está definida.

    La variable sirve para que ``--selftest`` y los tests nunca toquen (ni aprovechen) la caché real del
    usuario: un selftest con caché poblada pasaría aunque tesseract estuviera roto.
    """
    override = os.environ.get(CACHE_DIR_ENV, "").strip()
    if override:
        return Path(override) / "ocr"
    return Path(platformdirs.user_cache_dir(APP_ID, appauthor=False)) / "ocr"


def prune_cache(
    cache_dir: Path | None = None,
    max_bytes: int = CACHE_MAX_BYTES,
    max_age_days: float = CACHE_MAX_AGE_DAYS,
    now: float | None = None,
) -> int:
    """Borra de la caché los PDFs más viejos que ``max_age_days`` y, si sigue pesando más de ``max_bytes``,
    los menos recientes hasta bajar del tope. Devuelve cuántos archivos borró. Nunca lanza.

    Se llama al final de cada ``ocr_document``: para un equipo que convierte escaneos a diario la caché
    crecería sin límite (cientos de KB a MB por página).
    """
    directory = cache_dir if cache_dir is not None else default_cache_dir()
    try:
        entries = [p for p in directory.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    except OSError:
        return 0
    moment = time.time() if now is None else now
    removed = 0
    stats: list[tuple[float, int, Path]] = []
    for path in entries:
        try:
            st = path.stat()
        except OSError:
            continue
        stats.append((st.st_mtime, st.st_size, path))
    keep: list[tuple[float, int, Path]] = []
    for mtime, size, path in stats:
        if moment - mtime > max_age_days * 86400:
            removed += _unlink_quietly(path)
        else:
            keep.append((mtime, size, path))
    total = sum(size for _m, size, _p in keep)
    for _mtime, size, path in sorted(keep):  # más viejos primero
        if total <= max_bytes:
            break
        if _unlink_quietly(path):
            removed += 1
            total -= size
    if removed:
        log.info("caché de OCR: %d archivo(s) borrado(s) de %s", removed, directory)
    return removed


def _unlink_quietly(path: Path) -> int:
    try:
        path.unlink()
    except OSError:
        return 0
    return 1


def _safe_token(value: str) -> str:
    return "".join(c if (c.isalnum() or c in "+-_") else "_" for c in value) or "x"


def cache_path_for(src: Path, pages: Sequence[int], lang: str, cache_dir: Path | None = None) -> Path:
    """Ruta del PDF cacheado: ``<sha1(src)>-<lang>-<hash de páginas>.pdf``."""
    directory = cache_dir if cache_dir is not None else default_cache_dir()
    pages_key = ",".join(str(p) for p in sorted(set(pages)))
    pages_hash = hashlib.sha1(pages_key.encode()).hexdigest()[:10]  # identificación, no seguridad
    return directory / f"{file_sha1(src)}-{_safe_token(lang)}-{pages_hash}.pdf"


def ocr_document(
    src: Path,
    pages: list[int],
    lang: str = DEFAULT_LANG,
    progress: ProgressFn | None = None,
    cancel: threading.Event | None = None,
    cache_dir: Path | None = None,
    password: str | None = None,
) -> Path:
    """Devuelve un PDF con TODAS las páginas de ``src``: las de ``pages`` con OCR, las demás copiadas.

    El resultado se cachea en ``cache_dir`` (por defecto ``user_cache_dir/ocr``) por SHA-1 del origen,
    idioma y conjunto de páginas; si ya existe se devuelve sin trabajar. ``progress(i, n, "ocr")`` se
    llama por página OCR; ``cancel`` se chequea entre páginas y mata el tesseract en curso.
    """
    src = Path(src)
    _check_cancel(cancel)
    if find_tesseract() is None:
        raise UserError("ocr_unavailable")

    target = cache_path_for(src, pages, lang, cache_dir)
    if target.is_file() and target.stat().st_size > 0:
        log.info("OCR cacheado para %s → %s", src.name, target)
        return target
    target.parent.mkdir(parents=True, exist_ok=True)

    wanted = sorted({p for p in pages if p >= 0})
    total = len(wanted)
    t_start = time.perf_counter()
    doc = open_document(src, password=password)
    try:
        wanted = [p for p in wanted if p < doc.page_count]
        wanted_set = set(wanted)
        total = len(wanted)
        with (
            pymupdf.open() as out,
            tempfile.TemporaryDirectory(prefix="pdf2excel-ocr-") as tmp,
        ):
            workdir = Path(tmp)
            done = 0
            for page_no in range(doc.page_count):
                _check_cancel(cancel)
                if page_no in wanted_set:
                    ocr_page_into(doc[page_no], out, lang=lang, cancel=cancel, workdir=workdir)
                    done += 1
                    if progress is not None:
                        progress(done, total, "ocr")
                else:
                    out.insert_pdf(doc, from_page=page_no, to_page=page_no)
            _check_cancel(cancel)
            # Guardado atómico: otro hilo puede estar leyendo la caché.
            partial = target.with_name(target.name + f".{os.getpid()}.{threading.get_ident()}.part")
            try:
                out.save(partial, garbage=3, deflate=True)
                os.replace(partial, target)
            finally:
                if partial.exists():
                    partial.unlink(missing_ok=True)
    finally:
        doc.close()

    elapsed = time.perf_counter() - t_start
    log.info(
        "OCR de %s: %d página(s) en %.2f s (%.2f s/pág) → %s",
        src.name,
        total,
        elapsed,
        elapsed / total if total else 0.0,
        target,
    )
    prune_cache(target.parent)  # el recién creado tiene el mtime más nuevo: nunca se borra a sí mismo
    return target


__all__ = [
    "CACHE_DIR_ENV",
    "CACHE_MAX_AGE_DAYS",
    "CACHE_MAX_BYTES",
    "DEFAULT_DPI",
    "DEFAULT_LANG",
    "DEFAULT_PSM",
    "TESSERACT_ENV",
    "cache_path_for",
    "default_cache_dir",
    "deskew_angle",
    "detect_orientation",
    "find_tesseract",
    "ocr_available",
    "ocr_document",
    "ocr_page_into",
    "prune_cache",
    "remove_grid_lines",
    "render_gray",
    "rotate_gray",
    "run_tesseract_textonly_pdf",
    "tessdata_prefix",
]
