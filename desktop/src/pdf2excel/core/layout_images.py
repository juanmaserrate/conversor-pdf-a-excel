"""Imágenes (logos, mapas, fotos) del modo "Igual al PDF": se recomprimen para no inflar el libro."""

from __future__ import annotations

import contextlib
from typing import Any

import pymupdf

from pdf2excel.core.layout_geom import _Tf
from pdf2excel.models import LayoutImage

MIN_IMAGE_PT = 4.0
MAX_SIDE_PX = 1600
TARGET_DPI = 200.0
MIN_SIDE_PX = 160
JPEG_QUALITY = 85
BACKGROUND_FRACTION = 0.85  # una imagen que cubre esto de la página es un fondo / escaneo


def extract_images(
    page: Any,
    image_info: list[dict[str, Any]],
    matrix: Any,
    page_size: tuple[float, float],
    scanned: bool,
) -> list[LayoutImage]:
    """Imágenes de la página como ``LayoutImage`` (fila/columna se completan después, con la grilla).

    ``scanned``: la página es un escaneo, así que la imagen de fondo no se incluye (el texto OCR ya es el
    contenido). Una misma imagen usada varias veces se recomprime una sola vez.
    """
    tf = _Tf(matrix)
    page_area = max(1.0, page_size[0] * page_size[1])
    cache: dict[int, tuple[bytes, str]] = {}
    out: list[LayoutImage] = []
    for info in sorted(image_info, key=lambda i: (i["bbox"][1], i["bbox"][0])):
        x0, y0, x1, y1 = tf.rect(info["bbox"])
        w, h = x1 - x0, y1 - y0
        if w < MIN_IMAGE_PT or h < MIN_IMAGE_PT:
            continue
        coverage = (w * h) / page_area
        if scanned and coverage >= 0.5:
            continue
        if coverage >= BACKGROUND_FRACTION:
            continue
        xref = int(info.get("xref", 0) or 0)
        key = xref if xref else -len(cache) - 1
        try:
            if key in cache and xref:
                data, fmt = cache[key]
            else:
                data, fmt = _encode(page, (x0, y0, x1, y1), xref, w, h, bool(page.rotation))
                if xref:
                    cache[key] = (data, fmt)
        except (RuntimeError, ValueError, OSError):
            continue
        out.append(
            LayoutImage(
                data=data,
                format="png" if fmt == "png" else "jpeg",
                bbox=(x0, y0, x1, y1),
                row=0,
                col=0,
                width_pt=w,
                height_pt=h,
            )
        )
    return out


def _target_size(src_w: int, src_h: int, shown_w: float, shown_h: float) -> tuple[int, int]:
    want_w = max(MIN_SIDE_PX, round(shown_w / 72.0 * TARGET_DPI))
    want_h = max(MIN_SIDE_PX, round(shown_h / 72.0 * TARGET_DPI))
    factor = min(1.0, want_w / max(1, src_w), want_h / max(1, src_h), MAX_SIDE_PX / max(src_w, src_h, 1))
    return max(1, round(src_w * factor)), max(1, round(src_h * factor))


def _encode(
    page: Any,
    visual_box: tuple[float, float, float, float],
    xref: int,
    shown_w: float,
    shown_h: float,
    rotated: bool,
) -> tuple[bytes, str]:
    doc = page.parent
    if rotated or not xref:
        # página girada o imagen en línea: se rasteriza el recuadro de la imagen
        clip = pymupdf.Rect(visual_box)  # get_pixmap interpreta el clip en el marco visual (rotado)
        zoom = min(TARGET_DPI / 72.0, MAX_SIDE_PX / max(clip.width, clip.height, 1.0))
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip, alpha=False)
        return pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY), "jpeg"
    pix = pymupdf.Pixmap(doc, xref)
    smask = _smask_xref(doc, xref)
    if pix.colorspace is None or pix.colorspace.n not in (1, 3) or pix.alpha:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    if smask:
        with contextlib.suppress(RuntimeError, ValueError):
            pix = pymupdf.Pixmap(pix, pymupdf.Pixmap(doc, smask))
    tw, th = _target_size(pix.width, pix.height, shown_w, shown_h)
    if (tw, th) != (pix.width, pix.height):
        pix = pymupdf.Pixmap(pix, tw, th)
    if pix.alpha:
        return pix.tobytes("png"), "png"
    return pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY), "jpeg"


def _smask_xref(doc: Any, xref: int) -> int:
    """xref de la máscara de transparencia de la imagen (0 si no tiene), sin decodificar la imagen."""
    kind, value = doc.xref_get_key(xref, "SMask")
    if kind != "xref":
        return 0
    try:
        return int(str(value).split()[0])
    except (ValueError, IndexError):
        return 0
