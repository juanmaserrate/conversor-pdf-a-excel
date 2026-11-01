"""Tests de core/document.py (WI-04): apertura segura y clasificación de páginas.

Los PDFs del corpus de referencia se copian desde el scratchpad si están disponibles; los casos
sintéticos (cifrado, en blanco, texto basura, imagen chica) se generan con pymupdf en ``tmp_path``.
"""

from __future__ import annotations

import hashlib
import shutil
from pathlib import Path

import pymupdf
import pytest

from pdf2excel.core.document import (
    classify_document,
    classify_page,
    file_sha1,
    needs_ocr,
    open_document,
)
from pdf2excel.errors import UserError
from pdf2excel.models import PageInfo

CORPUS_DIRS = [
    Path("/tmp/claude-0/-home-user/932cdd7a-5305-5ce6-bcbe-bcac3f3b947f/scratchpad/t"),
    Path(__file__).parent / "corpus",
]


def _find_corpus_pdf(name: str) -> Path | None:
    for directory in CORPUS_DIRS:
        candidate = directory / name
        if candidate.is_file():
            return candidate
    return None


@pytest.fixture
def corpus(tmp_path: Path):
    """Copia un PDF del corpus a ``tmp_path`` (nunca se toca el original); skip si no está."""

    def _get(name: str) -> Path:
        src = _find_corpus_pdf(name)
        if src is None:
            pytest.skip(f"PDF de corpus no disponible: {name}")
        dst = tmp_path / name
        shutil.copyfile(src, dst)
        return dst

    return _get


def _pdf_with_text(path: Path, text: str, *, fontname: str = "helv") -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    y = 72.0
    for line in text.splitlines() or [text]:
        page.insert_text((72, y), line, fontsize=11, fontname=fontname)
        y += 16
    doc.save(path)
    doc.close()
    return path


def _gray_png_bytes(width: int, height: int, value: int = 200) -> bytes:
    pix = pymupdf.Pixmap(pymupdf.csGRAY, pymupdf.IRect(0, 0, width, height), 0)
    pix.clear_with(value)
    return pix.tobytes("png")


# ------------------------------------------------------------------------------------------------
# Clasificación sobre el corpus
# ------------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("name", "kind"),
    [
        ("remito_grid.pdf", "digital"),
        ("remito_nogrid.pdf", "digital"),
        ("prueba.pdf", "digital"),
        ("remito_grid_scan.pdf", "scanned"),
        ("remito_nogrid_scan.pdf", "scanned"),
        ("remito_grid_ocr.pdf", "scanned_with_text"),
    ],
)
def test_classify_corpus(corpus, name: str, kind: str) -> None:
    doc = open_document(corpus(name))
    infos = classify_document(doc)
    assert len(infos) == doc.page_count
    assert infos[0].index == 0
    assert infos[0].kind == kind
    if kind == "digital":
        assert infos[0].image_coverage == 0.0
        assert infos[0].n_chars >= 20
    else:
        assert infos[0].image_coverage > 0.5
    doc.close()


def test_classify_scanned_has_no_text(corpus) -> None:
    doc = open_document(corpus("remito_grid_scan.pdf"))
    info = classify_page(doc[0])
    assert info.n_chars < 20
    assert info.rotation == 0
    assert needs_ocr(info)


# ------------------------------------------------------------------------------------------------
# Casos sintéticos
# ------------------------------------------------------------------------------------------------


def test_blank_page_is_empty(tmp_path: Path) -> None:
    doc = pymupdf.open()
    doc.new_page()
    path = tmp_path / "blanco.pdf"
    doc.save(path)
    doc.close()

    opened = open_document(path)
    info = classify_page(opened[0])
    assert info.kind == "empty"
    assert info.n_chars == 0
    assert info.image_coverage == 0.0
    assert not needs_ocr(info)


def test_short_text_without_images_is_digital(tmp_path: Path) -> None:
    path = _pdf_with_text(tmp_path / "corto.pdf", "Hola")
    doc = open_document(path)
    assert classify_page(doc[0]).kind == "digital"


def test_small_image_with_text_is_digital(tmp_path: Path) -> None:
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Remito N° 0001-00004521 — Cliente: Bar El Catorce", fontsize=11)
    page.insert_image(pymupdf.Rect(400, 40, 500, 120), stream=_gray_png_bytes(100, 80))  # logo chico
    path = tmp_path / "logo.pdf"
    doc.save(path)
    doc.close()

    info = classify_page(open_document(path)[0])
    assert info.kind == "digital"
    assert 0.0 < info.image_coverage < 0.1


@pytest.mark.parametrize("rect", [(40, 60, 555, 400), (100, 100, 300, 350)])  # ~35 % y ~10 % de la hoja
def test_small_image_without_text_is_scanned(tmp_path: Path, rect: tuple[int, int, int, int]) -> None:
    """Ticket/remito chico escaneado y pegado en una A4: sin texto y con imagen → escaneo (desvío §5.1)."""
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(pymupdf.Rect(*rect), stream=_gray_png_bytes(200, 150))
    path = tmp_path / "ticket.pdf"
    doc.save(path)
    doc.close()

    info = classify_page(open_document(path)[0])
    assert info.kind == "scanned"
    assert 0.05 < info.image_coverage < 0.5
    assert needs_ocr(info)


def test_tiny_logo_without_text_is_not_scanned(tmp_path: Path) -> None:
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(pymupdf.Rect(400, 40, 500, 120), stream=_gray_png_bytes(100, 80))  # < 5 %
    path = tmp_path / "solo_logo.pdf"
    doc.save(path)
    doc.close()

    info = classify_page(open_document(path)[0])
    assert info.kind == "digital" and info.image_coverage < 0.05


def test_full_page_image_without_text_is_scanned(tmp_path: Path) -> None:
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, stream=_gray_png_bytes(200, 283))
    path = tmp_path / "escaneo.pdf"
    doc.save(path)
    doc.close()

    info = classify_page(open_document(path)[0])
    assert info.kind == "scanned"
    assert info.image_coverage > 0.9


def test_full_page_image_with_text_is_scanned_with_text(tmp_path: Path) -> None:
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, stream=_gray_png_bytes(200, 283))
    page.insert_text((72, 72), "Texto de una capa OCR ya existente sobre la imagen", fontsize=11)
    path = tmp_path / "ocr.pdf"
    doc.save(path)
    doc.close()

    info = classify_page(open_document(path)[0])
    assert info.kind == "scanned_with_text"
    assert not needs_ocr(info)


def test_image_bbox_outside_page_is_clipped(tmp_path: Path) -> None:
    """Una imagen que se sale de la página no puede dar cobertura > 1."""
    doc = pymupdf.open()
    page = doc.new_page()
    big = pymupdf.Rect(-300, -300, page.rect.width + 300, page.rect.height + 300)
    page.insert_image(big, stream=_gray_png_bytes(50, 50))
    path = tmp_path / "grande.pdf"
    doc.save(path)
    doc.close()

    info = classify_page(open_document(path)[0])
    assert info.image_coverage <= 1.0
    assert info.kind == "scanned"


class _FakePage:
    """Lo mínimo que usa classify_page: get_text(), get_image_info(), rect, number."""

    def __init__(self, text: str, images: list[tuple[float, float, float, float]] | None = None) -> None:
        self._text = text
        self._images = images or []
        self.rect = pymupdf.Rect(0, 0, 595, 842)
        self.number = 3

    def get_text(self) -> str:
        return self._text

    def get_image_info(self) -> list[dict]:
        return [{"bbox": bbox} for bbox in self._images]


def test_garbage_text_detected() -> None:
    # Fuentes sin ToUnicode suelen dar símbolos de uso privado o letras de otros alfabetos.
    garbage = " ЖЩДЛ ∆∑∏  ⌂⌐¬ ЖЩДЛ ∆∑∏ ⌂⌐¬  ЖЩДЛ" * 3
    info = classify_page(_FakePage(garbage))  # type: ignore[arg-type]
    assert info.kind == "garbage_text"
    assert info.index == 3
    assert needs_ocr(info)


def test_spanish_text_with_symbols_is_not_garbage() -> None:
    text = (
        "Remito N° 0001-00004521 — Cliente: Bar El Catorce\n"
        "Código Producto Cantidad Precio unitario Total\n"
        "A-100 Mezcal joven 10 $ 350,00 $ 3.500,00 (21%) 25º; aprox. 1/2 'ok' \"dos\"\n"
        "Señor/a: Ñandú, pingüino, canción, árbol, éxito, índice, óptimo, único.\n"
    )
    info = classify_page(_FakePage(text))  # type: ignore[arg-type]
    assert info.kind == "digital"


def test_coverage_is_sum_of_areas_capped_at_one() -> None:
    page = _FakePage("", images=[(0, 0, 595, 842), (0, 0, 595, 842)])
    info = classify_page(page)  # type: ignore[arg-type]
    assert info.image_coverage == 1.0
    assert info.kind == "scanned"


# ------------------------------------------------------------------------------------------------
# needs_ocr
# ------------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("kind", "expected"),
    [
        ("digital", False),
        ("scanned", True),
        ("scanned_with_text", False),
        ("garbage_text", True),
        ("empty", False),
    ],
)
def test_needs_ocr_by_kind(kind: str, expected: bool) -> None:
    info = PageInfo(index=0, kind=kind, n_chars=0, image_coverage=0.0)  # type: ignore[arg-type]
    assert needs_ocr(info) is expected


def test_needs_ocr_force_overrides_digital() -> None:
    info = PageInfo(index=0, kind="digital", n_chars=500, image_coverage=0.0)
    assert needs_ocr(info) is False
    assert needs_ocr(info, force=True) is True


# ------------------------------------------------------------------------------------------------
# open_document: errores
# ------------------------------------------------------------------------------------------------


def test_encrypted_pdf_requires_password(tmp_path: Path) -> None:
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 72), "Secreto corporativo")
    path = tmp_path / "cifrado.pdf"
    aes256 = pymupdf.PDF_ENCRYPT_AES_256  # type: ignore[attr-defined]  # constante dinámica de pymupdf
    doc.save(path, encryption=aes256, owner_pw="dueño", user_pw="clave123")
    doc.close()

    with pytest.raises(UserError) as exc_info:
        open_document(path)
    assert exc_info.value.code == "encrypted"
    assert exc_info.value.action == "password"

    with pytest.raises(UserError) as exc_info:
        open_document(path, password="incorrecta")
    assert exc_info.value.code == "encrypted"
    assert exc_info.value.action == "password"

    opened = open_document(path, password="clave123")
    assert opened.page_count == 1
    assert "Secreto" in opened[0].get_text()
    assert classify_page(opened[0]).kind == "digital"
    opened.close()


def test_text_file_renamed_pdf(tmp_path: Path) -> None:
    path = tmp_path / "falso.pdf"
    path.write_text("Esto es un archivo de texto, no un PDF.\n" * 40, encoding="utf-8")
    with pytest.raises(UserError) as exc_info:
        open_document(path)
    assert exc_info.value.code in ("damaged", "not_pdf")
    assert exc_info.value.action == "skip"


def test_truncated_pdf_is_damaged(tmp_path: Path) -> None:
    good = _pdf_with_text(tmp_path / "bueno.pdf", "Documento completo con texto suficiente.")
    data = good.read_bytes()
    broken = tmp_path / "roto.pdf"
    broken.write_bytes(data[: len(data) // 3])
    with pytest.raises(UserError) as exc_info:
        open_document(broken)
    assert exc_info.value.code in ("damaged", "not_pdf")


def test_wrong_extension_is_not_pdf(tmp_path: Path) -> None:
    good = _pdf_with_text(tmp_path / "bueno.pdf", "Hola mundo")
    renamed = tmp_path / "documento.txt"
    shutil.copyfile(good, renamed)
    with pytest.raises(UserError) as exc_info:
        open_document(renamed)
    assert exc_info.value.code == "not_pdf"


def test_png_renamed_pdf_is_not_pdf(tmp_path: Path) -> None:
    """pymupdf abre PNGs como documentos: el chequeo de contenido tiene que frenarlo."""
    path = tmp_path / "imagen.pdf"
    path.write_bytes(_gray_png_bytes(40, 40))
    with pytest.raises(UserError) as exc_info:
        open_document(path)
    assert exc_info.value.code in ("not_pdf", "damaged")


def test_missing_file_is_user_error(tmp_path: Path) -> None:
    with pytest.raises(UserError):
        open_document(tmp_path / "no_existe.pdf")


def test_open_document_accepts_str_like_paths(tmp_path: Path) -> None:
    path = _pdf_with_text(tmp_path / "ok.pdf", "Texto de prueba suficiente para digital.")
    doc = open_document(Path(str(path)))
    assert doc.page_count == 1
    doc.close()


# ------------------------------------------------------------------------------------------------
# file_sha1
# ------------------------------------------------------------------------------------------------


def test_file_sha1_matches_hashlib(tmp_path: Path) -> None:
    path = tmp_path / "datos.bin"
    payload = bytes(range(256)) * 5000  # > 1 MB para pasar por varios chunks
    path.write_bytes(payload)
    assert file_sha1(path) == hashlib.sha1(payload).hexdigest()
    assert len(file_sha1(path)) == 40
