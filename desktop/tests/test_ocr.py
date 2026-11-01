"""Tests de core/ocr.py (WI-04). Lentos: requieren tesseract (spa, eng, osd) instalado.

Se saltean completos si ``ocr_available()`` es False. Las imágenes sintéticas se generan renderizando
texto real con pymupdf (fuente Helvetica) para que OSD y tesseract tengan algo legible.
"""

from __future__ import annotations

import os
import shutil
import stat
import sys
import threading
import time
from pathlib import Path

import cv2
import numpy as np
import pymupdf
import pytest

from pdf2excel.core import ocr
from pdf2excel.core.document import classify_page
from pdf2excel.core.ocr import (
    deskew_angle,
    detect_orientation,
    find_tesseract,
    ocr_available,
    ocr_document,
    remove_grid_lines,
    render_gray,
    rotate_gray,
    run_tesseract_textonly_pdf,
)
from pdf2excel.errors import UserError

pytestmark = [pytest.mark.slow, pytest.mark.skipif(not ocr_available(), reason="tesseract no disponible")]

CORPUS_DIRS = [
    Path("/tmp/claude-0/-home-user/932cdd7a-5305-5ce6-bcbe-bcac3f3b947f/scratchpad/t"),
    Path(__file__).parent / "corpus",
    Path(__file__).parent / "corpus" / "_generated",
]
# Objetivo de DESIGN.md §5.2 (1,4–1,9 s/pág medido). Con PDF2EXCEL_STRICT_TIMING=1 se exige al pie de la
# letra; si no, se tolera 3× para máquinas de CI compartidas.
MAX_SECONDS_PER_PAGE = 2.5
TIMING_FACTOR = 1.0 if os.environ.get("PDF2EXCEL_STRICT_TIMING") else 3.0
POSIX = not sys.platform.startswith("win")


def _corpus_copy(tmp_path: Path, name: str) -> Path:
    for directory in CORPUS_DIRS:
        src = directory / name
        if src.is_file():
            dst = tmp_path / name
            shutil.copyfile(src, dst)
            return dst
    pytest.skip(f"PDF de corpus no disponible: {name}")


def _text_page_gray(dpi: int = 300) -> np.ndarray:
    """Render gris de una página A4 con ~25 renglones de texto castellano real."""
    doc = pymupdf.open()
    page = doc.new_page()
    body = (
        "El remito número 4521 fue entregado al cliente Bar El Catorce el 15 de octubre "
        "por la mañana, con diez cajas de mezcal joven y cinco de tequila reposado."
    )
    y = 90.0
    for i in range(25):
        page.insert_text((50, y), f"{i + 1:02d}. {body[: 95 + (i % 5) * 3]}", fontsize=10, fontname="helv")
        y += 26
    gray = render_gray(page, dpi=dpi)
    doc.close()
    return gray


@pytest.fixture(scope="module")
def text_gray() -> np.ndarray:
    return _text_page_gray()


# ------------------------------------------------------------------------------------------------
# Localización de tesseract
# ------------------------------------------------------------------------------------------------


def test_find_tesseract_returns_existing_executable() -> None:
    exe = find_tesseract()
    assert exe is not None
    assert exe.is_file()


def test_env_override_missing_disables_ocr(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ocr.TESSERACT_ENV, "/no/existe/tesseract")
    assert find_tesseract() is None
    assert ocr_available() is False


def test_env_override_existing_wins(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    fake = tmp_path / "mi_tesseract"
    fake.write_text("#!/bin/sh\nexit 0\n")
    monkeypatch.setenv(ocr.TESSERACT_ENV, str(fake))
    assert find_tesseract() == fake


def test_tessdata_prefix_none_without_bundled(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("PDF2EXCEL_RESOURCES", str(tmp_path))
    assert ocr.tessdata_prefix() is None
    (tmp_path / "tessdata").mkdir()
    (tmp_path / "tessdata" / "spa.traineddata").write_bytes(b"x")
    assert ocr.tessdata_prefix() == tmp_path / "tessdata"


# ------------------------------------------------------------------------------------------------
# Imagen
# ------------------------------------------------------------------------------------------------


def test_render_gray_shape_and_dtype() -> None:
    doc = pymupdf.open()
    page = doc.new_page()  # A4
    gray = render_gray(page, dpi=72)
    assert gray.dtype == np.uint8
    assert gray.shape == (842, 595)
    assert gray.min() == 255  # página en blanco


def test_deskew_angle_recovers_1_2_degrees(text_gray: np.ndarray) -> None:
    assert abs(deskew_angle(text_gray)) < 0.2  # la imagen derecha da ~0
    skewed = rotate_gray(text_gray, 1.2)
    angle = deskew_angle(skewed)
    assert abs(angle - (-1.2)) <= 0.2
    # aplicar la corrección deja la imagen derecha otra vez
    fixed = rotate_gray(skewed, angle)
    assert abs(deskew_angle(fixed)) <= 0.2


def test_deskew_angle_negative_skew(text_gray: np.ndarray) -> None:
    skewed = rotate_gray(text_gray, -2.5)
    assert abs(deskew_angle(skewed) - 2.5) <= 0.2


def test_detect_orientation(text_gray: np.ndarray) -> None:
    assert detect_orientation(text_gray) == 0
    assert detect_orientation(np.ascontiguousarray(np.rot90(text_gray))) == 90


def test_detect_orientation_180(text_gray: np.ndarray) -> None:
    assert detect_orientation(np.ascontiguousarray(np.rot90(text_gray, 2))) == 180


def test_detect_orientation_without_tesseract_is_zero(
    monkeypatch: pytest.MonkeyPatch, text_gray: np.ndarray
) -> None:
    monkeypatch.setenv(ocr.TESSERACT_ENV, "/no/existe/tesseract")
    assert detect_orientation(text_gray) == 0


def test_remove_grid_lines_keeps_text(text_gray: np.ndarray) -> None:
    h, w = text_gray.shape
    with_grid = text_gray.copy()
    grid_mask = np.zeros_like(text_gray, dtype=bool)
    for x in range(100, w - 100, 300):
        cv2.line(with_grid, (x, 50), (x, h - 50), 0, 4)
        grid_mask[50 : h - 50, x - 2 : x + 2] = True
    for y in range(60, h - 60, 260):
        cv2.line(with_grid, (50, y), (w - 50, y), 0, 4)
        grid_mask[y - 2 : y + 2, 50 : w - 50] = True

    cleaned = remove_grid_lines(with_grid)

    grid_before = int((with_grid[grid_mask] < 128).sum())
    grid_after = int((cleaned[grid_mask] < 128).sum())
    assert grid_before > 0
    assert grid_after <= 0.10 * grid_before, (grid_before, grid_after)

    text_mask = (text_gray < 128) & ~grid_mask
    text_before = int(text_mask.sum())
    text_after = int((cleaned < 128)[text_mask].sum())
    assert text_before > 0
    assert text_after >= 0.90 * text_before, (text_before, text_after)


# ------------------------------------------------------------------------------------------------
# tesseract: errores y cancelación con un ejecutable falso
# ------------------------------------------------------------------------------------------------


def _fake_tesseract(tmp_path: Path, body: str) -> Path:
    script = tmp_path / "tesseract_falso.sh"
    script.write_text("#!/bin/sh\n" + body)
    script.chmod(script.stat().st_mode | stat.S_IXUSR)
    return script


@pytest.mark.skipif(not POSIX, reason="el ejecutable falso es un script sh")
def test_run_tesseract_nonzero_exit_is_ocr_failed(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv(
        ocr.TESSERACT_ENV, str(_fake_tesseract(tmp_path, "echo 'Error opening data file' >&2\nexit 1\n"))
    )
    png = tmp_path / "p.png"
    png.write_bytes(b"")
    with pytest.raises(UserError) as exc_info:
        run_tesseract_textonly_pdf(png, tmp_path / "out", "spa")
    assert exc_info.value.code == "ocr_failed"
    assert "Error opening data file" in exc_info.value.detail


@pytest.mark.skipif(not POSIX, reason="el ejecutable falso es un script sh")
def test_run_tesseract_cancel_kills_process(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv(ocr.TESSERACT_ENV, str(_fake_tesseract(tmp_path, "sleep 30\n")))
    cancel = threading.Event()
    threading.Timer(0.5, cancel.set).start()
    png = tmp_path / "p.png"
    png.write_bytes(b"")
    t0 = time.perf_counter()
    with pytest.raises(UserError) as exc_info:
        run_tesseract_textonly_pdf(png, tmp_path / "out", "spa", cancel=cancel)
    assert exc_info.value.code == "cancelled"
    assert time.perf_counter() - t0 < 5.0


@pytest.mark.skipif(not POSIX, reason="el ejecutable falso es un script sh")
def test_run_tesseract_timeout_is_ocr_failed(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv(ocr.TESSERACT_ENV, str(_fake_tesseract(tmp_path, "sleep 30\n")))
    png = tmp_path / "p.png"
    png.write_bytes(b"")
    with pytest.raises(UserError) as exc_info:
        run_tesseract_textonly_pdf(png, tmp_path / "out", "spa", timeout=0.5)
    assert exc_info.value.code == "ocr_failed"


def test_run_tesseract_unavailable(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv(ocr.TESSERACT_ENV, "/no/existe/tesseract")
    with pytest.raises(UserError) as exc_info:
        run_tesseract_textonly_pdf(tmp_path / "p.png", tmp_path / "out", "spa")
    assert exc_info.value.code == "ocr_unavailable"


# ------------------------------------------------------------------------------------------------
# ocr_document sobre el remito escaneado
# ------------------------------------------------------------------------------------------------


def _assert_remito_table(out_pdf: Path) -> None:
    with pymupdf.open(out_pdf) as doc:
        assert doc.page_count == 1
        page = doc[0]
        words = page.get_text("words")
        assert len(words) >= 30, len(words)
        tables = page.find_tables()
        assert len(tables.tables) >= 1
        table = tables.tables[0]
        assert (table.row_count, table.col_count) == (7, 5)
        first_cell = table.rows[1].cells[0]
        assert page.get_text("text", clip=first_cell).strip() == "A-100"
        info = classify_page(page)
        assert info.kind == "scanned_with_text"


def test_ocr_document_remito_grid_scan(tmp_path: Path) -> None:
    src = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    cache = tmp_path / "cache"
    calls: list[tuple[int, int, str]] = []

    t0 = time.perf_counter()
    out = ocr_document(src, [0], cache_dir=cache, progress=lambda a, b, s: calls.append((a, b, s)))
    elapsed = time.perf_counter() - t0

    assert out.parent == cache
    assert out.suffix == ".pdf"
    assert calls == [(1, 1, "ocr")]
    _assert_remito_table(out)
    assert elapsed <= MAX_SECONDS_PER_PAGE * TIMING_FACTOR, f"{elapsed:.2f} s para 1 página A4"

    # segunda llamada: caché (sin tesseract)
    t0 = time.perf_counter()
    again = ocr_document(src, [0], cache_dir=cache)
    assert time.perf_counter() - t0 < 0.1
    assert again == out


def test_ocr_document_nogrid_scan(tmp_path: Path) -> None:
    src = _corpus_copy(tmp_path, "remito_nogrid_scan.pdf")
    out = ocr_document(src, [0], cache_dir=tmp_path / "cache")
    _assert_remito_table(out)


def test_ocr_document_cache_key_depends_on_lang_and_pages(tmp_path: Path) -> None:
    src = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    cache = tmp_path / "cache"
    a = ocr.cache_path_for(src, [0], "spa+eng", cache)
    b = ocr.cache_path_for(src, [0], "spa", cache)
    c = ocr.cache_path_for(src, [0, 1], "spa+eng", cache)
    d = ocr.cache_path_for(src, [0, 0], "spa+eng", cache)
    assert len({a, b, c}) == 3
    assert a == d
    assert a.parent == cache


def test_ocr_document_cancelled_before_start(tmp_path: Path) -> None:
    src = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    cancel = threading.Event()
    cancel.set()
    with pytest.raises(UserError) as exc_info:
        ocr_document(src, [0], cache_dir=tmp_path / "cache", cancel=cancel)
    assert exc_info.value.code == "cancelled"
    assert not list(tmp_path.rglob("*.pdf.*.part"))  # no quedan parciales
    cache = tmp_path / "cache"
    assert not cache.exists() or not list(cache.glob("*.pdf"))


def test_ocr_document_unavailable(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    src = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    monkeypatch.setenv(ocr.TESSERACT_ENV, "/no/existe")
    with pytest.raises(UserError) as exc_info:
        ocr_document(src, [0], cache_dir=tmp_path / "cache")
    assert exc_info.value.code == "ocr_unavailable"


def test_ocr_document_copies_pages_not_in_list(tmp_path: Path) -> None:
    """Documento mixto: página 0 digital (se copia), página 1 escaneo (OCR)."""
    scan = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    mixed = pymupdf.open()
    mixed.new_page().insert_text((72, 72), "Página digital con texto propio que se copia tal cual.")
    with pymupdf.open(scan) as scan_doc:
        mixed.insert_pdf(scan_doc)
    src = tmp_path / "mixto.pdf"
    mixed.save(src)
    mixed.close()

    out = ocr_document(src, [1], cache_dir=tmp_path / "cache")
    with pymupdf.open(out) as doc:
        assert doc.page_count == 2
        assert "Página digital" in doc[0].get_text()
        assert classify_page(doc[0]).kind == "digital"
        assert classify_page(doc[1]).kind == "scanned_with_text"
        assert "A-100" in doc[1].get_text()


def test_ocr_page_into_rotated_scan(tmp_path: Path) -> None:
    """Un escaneo girado 90° se endereza (OSD) y la página de salida queda vertical."""
    scan = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    with pymupdf.open(scan) as src_doc:
        gray = render_gray(src_doc[0], dpi=150)
    rotated = np.ascontiguousarray(np.rot90(gray))  # apaisado, texto girado 90° antihorario
    ok, png = cv2.imencode(".png", rotated)
    assert ok
    landscape = pymupdf.open()
    page = landscape.new_page(width=842, height=595)
    page.insert_image(page.rect, stream=png.tobytes())
    src = tmp_path / "girado.pdf"
    landscape.save(src)
    landscape.close()

    out = ocr_document(src, [0], cache_dir=tmp_path / "cache")
    with pymupdf.open(out) as doc:
        page = doc[0]
        assert page.rect.width < page.rect.height  # volvió a vertical
        text = page.get_text()
        assert "A-100" in text
        assert "Mezcal" in text


def _open_fds_under(directory: Path) -> list[str]:
    """Rutas de los descriptores abiertos por este proceso dentro de ``directory`` (solo Linux)."""
    found = []
    for fd in os.listdir("/proc/self/fd"):
        try:
            target = os.readlink(f"/proc/self/fd/{fd}")
        except OSError:
            continue
        if target.startswith(str(directory)):
            found.append(target)
    return found


@pytest.mark.skipif(not Path("/proc/self/fd").exists(), reason="necesita /proc/self/fd (Linux)")
def test_ocr_page_into_does_not_keep_temp_files_open(tmp_path: Path) -> None:
    """Regresión Windows: show_pdf_page dejaba t<N>.pdf abierto desde el destino y el tmpdir no se podía
    borrar (WinError 32). En Linux se ve como un descriptor abierto dentro del workdir."""
    scan = _corpus_copy(tmp_path, "remito_grid_scan.pdf")
    workdir = tmp_path / "work"
    workdir.mkdir()
    out = pymupdf.open()
    with pymupdf.open(scan) as src_doc:
        ocr.ocr_page_into(src_doc[0], out, workdir=workdir)
    assert _open_fds_under(workdir) == []
    assert len(out[0].get_text().strip()) > 50  # la capa de texto sí quedó incrustada
