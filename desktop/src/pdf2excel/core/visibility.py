"""Qué texto de una página PDF se ve de verdad (LAYOUT_MODE.md §2 y §3.1).

Un PDF puede dibujar texto y después taparlo con un relleno opaco (así imprime el navegador los campos "N/A"
ocultos), sacarlo de la página, achicarlo a tamaño cero, pintarlo del color del fondo o marcarlo como
invisible. Este módulo decide qué spans quedan ocultos usando el orden de dibujo (``seqno``) que comparten
``page.get_texttrace()`` y ``page.get_drawings()``. Sin Qt.

Todas las coordenadas están en el marco de ``page.get_text()`` (página SIN rotar, origen arriba a la
izquierda), igual que las palabras de PyMuPDF.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import Any

import pymupdf

COVER_MIN_FRACTION = 0.9  # fracción del bbox del span que debe cubrir el relleno que lo tapa
OPAQUE_ALPHA = 0.99
MIN_FONT_SIZE = 0.5  # por debajo de este tamaño el texto no se lee
BG_COLOR_TOL = 0.04  # diferencia máxima por canal para decir "color igual al fondo"
SCANNED_IMAGE_COVERAGE = 0.5
_BAND = 24.0  # alto de las franjas del índice espacial de rellenos

Color = tuple[float, float, float]
Box = tuple[float, float, float, float]


@dataclass(frozen=True)
class _Cover:
    seqno: int
    box: Box
    color: Color
    alpha: float = (
        1.0  # < OPAQUE_ALPHA: translúcido; no tapa texto, pero tiñe el fondo que ve el texto de encima
    )


def normalize_text(text: str) -> str:
    """Normaliza ligaduras y espacios raros para comparar texto de spans con palabras de PyMuPDF."""
    out = unicodedata.normalize("NFKC", text) if any(ord(ch) > 0xFB00 for ch in text) else text
    return out.replace(" ", " ").strip()


def color_to_rgb(color: Any) -> Color | None:
    """Convierte un color de PyMuPDF (gris, RGB o CMYK, en 0..1) a RGB; ``None`` si no hay color."""
    if color is None:
        return None
    try:
        values = [float(v) for v in color]
    except TypeError:
        return None
    if len(values) == 1:
        return (values[0], values[0], values[0])
    if len(values) == 3:
        return (values[0], values[1], values[2])
    if len(values) == 4:
        c, m, y, k = values
        return ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return None


def _area(box: Box) -> float:
    return max(0.0, box[2] - box[0]) * max(0.0, box[3] - box[1])


def _intersection(a: Box, b: Box) -> float:
    return _area((max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])))


def _is_rectangular(drawing: dict[str, Any]) -> bool:
    """¿El relleno es (casi) un rectángulo? Rechaza trazados arbitrarios (curvas, triángulos, mapas)."""
    items = drawing["items"]
    kinds = {item[0] for item in items}
    if kinds <= {"re", "qu"}:
        return True
    rect = drawing["rect"]
    bbox_area = rect.width * rect.height
    if bbox_area <= 0:
        return False
    points: list[tuple[float, float]] = []
    for item in items:
        if item[0] in ("l", "c"):
            points.append((item[1].x, item[1].y))
        elif item[0] in ("re", "qu"):
            return False
    if len(points) < 3:
        return False
    twice_area = 0.0
    for (x0, y0), (x1, y1) in zip(points, points[1:] + points[:1], strict=True):
        twice_area += x0 * y1 - x1 * y0
    return abs(twice_area) / 2.0 >= 0.9 * bbox_area


def opaque_covers(drawings: list[dict[str, Any]]) -> list[_Cover]:
    """Rellenos rectangulares (cualquier color) con su orden de dibujo; los translúcidos llevan ``alpha``."""
    covers: list[_Cover] = []
    for d in drawings:
        fill = d.get("fill")
        if fill is None or d["type"] not in ("f", "fs"):
            continue
        raw_alpha = d.get("fill_opacity")
        alpha = 1.0 if raw_alpha is None else float(raw_alpha)
        if alpha <= 0.05:
            continue
        rect = d["rect"]
        if rect.width < 2.0 or rect.height < 2.0:
            continue  # demasiado fino para tapar texto: es un borde
        rgb = color_to_rgb(fill)
        if rgb is None or not _is_rectangular(d):
            continue
        covers.append(_Cover(int(d.get("seqno", 0)), (rect.x0, rect.y0, rect.x1, rect.y1), rgb, alpha))
    return covers


class _CoverIndex:
    """Índice por franjas horizontales de los rellenos opacos, para no comparar cada span con todos."""

    def __init__(self, covers: list[_Cover]) -> None:
        self._bands: dict[int, list[_Cover]] = {}
        for cover in covers:
            first = int(cover.box[1] // _BAND)
            last = int(cover.box[3] // _BAND)
            for band in range(first, last + 1):
                self._bands.setdefault(band, []).append(cover)

    def candidates(self, box: Box) -> list[_Cover]:
        seen: dict[int, _Cover] = {}
        for band in range(int(box[1] // _BAND), int(box[3] // _BAND) + 1):
            for cover in self._bands.get(band, ()):
                seen[id(cover)] = cover
        return list(seen.values())


class _WordIndex:
    """Cajas de las palabras que MuPDF considera visibles (ya recortadas por los clips), por franjas."""

    def __init__(self, words: list[tuple]) -> None:
        self._bands: dict[int, list[Box]] = {}
        for w in words:
            box: Box = (w[0], w[1], w[2], w[3])
            self._bands.setdefault(int(((box[1] + box[3]) / 2.0) // _BAND), []).append(box)

    def has_word_in(self, box: Box) -> bool:
        """``True`` si alguna palabra visible corresponde al span.

        Dos casos, los dos reales:

        * span largo con varias palabras (un subtítulo entero): el centro de alguna palabra cae dentro
          del span;
        * palabra partida en varios spans (los navegadores guardan un glifo por span: "1", "/", "2") y
          MuPDF la une en una sola ("1/"): el centro del span cae dentro de la caja de la palabra.
        Si no se cumple ninguno, MuPDF descartó el texto por un clip (está recortado, no se ve).
        """
        pad = 1.0
        cx, cy = (box[0] + box[2]) / 2.0, (box[1] + box[3]) / 2.0
        first = int((box[1] - pad) // _BAND) - 1
        last = int((box[3] + pad) // _BAND) + 1
        for band in range(first, last + 1):
            for word in self._bands.get(band, ()):
                wx, wy = (word[0] + word[2]) / 2.0, (word[1] + word[3]) / 2.0
                if box[0] - pad <= wx <= box[2] + pad and box[1] - pad <= wy <= box[3] + pad:
                    return True
                if word[0] - pad <= cx <= word[2] + pad and word[1] - pad <= cy <= word[3] + pad:
                    return True
        return False


def span_text(trace: dict[str, Any]) -> str:
    return "".join(chr(ch[0]) for ch in trace.get("chars", ()) if ch[0] >= 32 or ch[0] in (9, 10, 13))


def looks_scanned(page: pymupdf.Page, trace: list[dict[str, Any]], images: list[dict[str, Any]]) -> bool:
    """Escaneo con capa OCR: hay texto invisible (render mode 3) y una imagen que cubre media página."""
    if not any(t.get("type") == 3 for t in trace):
        return False
    page_area = page.rect.width * page.rect.height
    if page_area <= 0:
        return False
    best = max((_area(tuple(img["bbox"])) for img in images), default=0.0)
    return best >= SCANNED_IMAGE_COVERAGE * page_area


def classify_spans(
    page: pymupdf.Page,
    trace: list[dict[str, Any]],
    drawings: list[dict[str, Any]],
    images: list[dict[str, Any]],
    scanned: bool,
) -> list[bool]:
    """Para cada span de ``trace``: ``True`` si está oculto (no se ve en la página impresa)."""
    page_box_rect = page.rect * page.derotation_matrix
    page_box: Box = (page_box_rect.x0, page_box_rect.y0, page_box_rect.x1, page_box_rect.y1)
    covers = opaque_covers(drawings)
    index = _CoverIndex(covers)
    image_boxes: list[Box] = [tuple(img["bbox"]) for img in images]
    word_index = _WordIndex(page.get_text("words"))
    hidden: list[bool] = []
    for t in trace:
        text = span_text(t)
        bbox = tuple(t["bbox"])
        box: Box = (bbox[0], bbox[1], bbox[2], bbox[3])
        area = _area(box)
        if not text.strip():
            hidden.append(False)
            continue
        if t.get("size", 1.0) < MIN_FONT_SIZE or area <= 0.0:
            hidden.append(True)
            continue
        if t.get("opacity", 1.0) <= 0.0:
            hidden.append(True)
            continue
        if t.get("type") == 3 and not scanned:
            hidden.append(True)
            continue
        if _intersection(box, page_box) < 0.2 * area:
            hidden.append(True)
            continue
        if t.get("type") != 3 and not word_index.has_word_in(box):
            # MuPDF aplica los clips al armar las palabras: si no queda ninguna, el texto está recortado
            hidden.append(True)
            continue
        seq = int(t.get("seqno", 0))
        covered = False
        under: _Cover | None = None
        cx, cy = (box[0] + box[2]) / 2.0, (box[1] + box[3]) / 2.0
        for cover in index.candidates(box):
            if cover.seqno > seq:
                if cover.alpha < OPAQUE_ALPHA:
                    continue
                if _intersection(box, cover.box) >= COVER_MIN_FRACTION * area:
                    covered = True
                    break
            elif (
                cover.box[0] <= cx <= cover.box[2]
                and cover.box[1] <= cy <= cover.box[3]
                and (under is None or cover.seqno > under.seqno)
            ):
                under = cover
        if covered:
            hidden.append(True)
            continue
        text_rgb = color_to_rgb(t.get("color"))
        if text_rgb is not None and t.get("type") != 3:
            if under is not None:
                background: Color | None = under.color
            elif any(_intersection(box, ib) > 0.2 * area for ib in image_boxes):
                background = None  # texto sobre una foto: no se puede saber el color del fondo
            else:
                background = (1.0, 1.0, 1.0)
            if background is not None and under is not None and under.alpha < OPAQUE_ALPHA:
                # fondo translúcido: se mezcla con el blanco de la hoja
                a = under.alpha
                background = (
                    a * background[0] + (1 - a),
                    a * background[1] + (1 - a),
                    a * background[2] + (1 - a),
                )
            if background is not None and all(
                abs(a - b) <= BG_COLOR_TOL for a, b in zip(text_rgb, background, strict=False)
            ):
                hidden.append(True)
                continue
        hidden.append(False)
    return hidden


def _load(page: pymupdf.Page) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    return page.get_texttrace(), page.get_drawings(), page.get_image_info()


def hidden_spans(
    page: pymupdf.Page, scanned: bool | None = None
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """``(spans_ocultos, spans_visibles)`` de la página (entradas de ``get_texttrace``)."""
    trace, drawings, images = _load(page)
    is_scanned = looks_scanned(page, trace, images) if scanned is None else scanned
    flags = classify_spans(page, trace, drawings, images, is_scanned)
    hid = [t for t, h in zip(trace, flags, strict=True) if h]
    vis = [t for t, h in zip(trace, flags, strict=True) if not h and span_text(t).strip()]
    return hid, vis


def hidden_span_rects(page: pymupdf.Page, scanned: bool | None = None) -> list[pymupdf.Rect]:
    """bbox de los spans de texto que NO se ven: tapados por un relleno opaco dibujado después, fuera de la
    página, de tamaño ~0, con opacidad 0, del color del fondo o invisibles (render mode 3, salvo que la
    página sea un escaneo con capa OCR; ``scanned=None`` lo detecta)."""
    hid, _vis = hidden_spans(page, scanned)
    return [pymupdf.Rect(t["bbox"]) for t in hid]


def _center(word: tuple) -> tuple[float, float]:
    return (word[0] + word[2]) / 2.0, (word[1] + word[3]) / 2.0


def _inside(point: tuple[float, float], box: Box) -> bool:
    return box[0] <= point[0] <= box[2] and box[1] <= point[1] <= box[3]


def drop_hidden_words(
    words: list[tuple], hidden: list[dict[str, Any]], visible: list[dict[str, Any]]
) -> list[tuple]:
    """Quita de ``words`` los que pertenecen a spans ocultos.

    Primero se empareja cada palabra de un span oculto por texto y posición (así un texto tapado y vuelto a
    escribir encima se queda solo con el de arriba); lo que no empareja se descarta si su centro cae en un
    span oculto y en ningún span visible.
    """
    if not hidden:
        return words
    removed: set[int] = set()
    hidden_boxes: list[Box] = []
    for t in hidden:
        bb = t["bbox"]
        box: Box = (bb[0], bb[1], bb[2], bb[3])
        hidden_boxes.append(box)
        for token in normalize_text(span_text(t)).split():
            for i, w in enumerate(words):
                if i in removed:
                    continue
                if normalize_text(w[4]) == token and _inside(_center(w), box):
                    removed.add(i)
                    break
    visible_boxes: list[Box] = [(v["bbox"][0], v["bbox"][1], v["bbox"][2], v["bbox"][3]) for v in visible]
    for i, w in enumerate(words):
        if i in removed:
            continue
        c = _center(w)
        if any(_inside(c, hb) for hb in hidden_boxes) and not any(_inside(c, vb) for vb in visible_boxes):
            removed.add(i)
    return [w for i, w in enumerate(words) if i not in removed]


def visible_words(page: pymupdf.Page, clip: Any = None, scanned: bool | None = None) -> list[tuple]:
    """Igual que ``page.get_text("words", clip=clip)`` pero sin las palabras ocultas."""
    words = page.get_text("words", clip=clip)
    hid, vis = hidden_spans(page, scanned)
    return drop_hidden_words(words, hid, vis)


__all__ = [
    "classify_spans",
    "color_to_rgb",
    "drop_hidden_words",
    "hidden_span_rects",
    "hidden_spans",
    "looks_scanned",
    "normalize_text",
    "opaque_covers",
    "span_text",
    "visible_words",
]
