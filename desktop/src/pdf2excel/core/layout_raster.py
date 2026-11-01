"""Cajas, líneas, rellenos y peso del texto de una página ESCANEADA (modo "Igual al PDF").

Un escaneo no trae vectores: las cajas y las líneas de la tabla son píxeles. Acá se detectan con OpenCV sobre
un render de la página (aperturas morfológicas con núcleos largos para líneas, regiones de color liso para
rellenos) y se devuelven como los mismos ``Segment``/``Box`` que produce ``extract_primitives`` con PDFs
digitales, así el resto del extractor no distingue de dónde vienen. También estima, por palabra del OCR, si el
trazo es negrita y de qué color es la tinta. Sin Qt.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any

import cv2
import numpy as np

from pdf2excel.core.layout_geom import Box, Primitives, Segment, merge_segments
from pdf2excel.core.layout_text import Word, group_lines

RENDER_DPI = 200
LINE_MIN_PT = 20.0  # una línea de borde mide al menos esto; las letras no llegan
LINE_INK = 232  # gris por debajo del cual un píxel puede ser línea (los bordes finos salen gris claro)
LINE_CONTRAST = 22  # cuánto más oscuro que su entorno (0-255) debe ser un píxel de línea
MAX_LINE_THICKNESS_PT = 4.5
EDGE_MARGIN = 0.012  # fracción de la página: lo que está pegado al borde es sombra del escáner
MAX_SEGMENTS = 700
FILL_MIN_PT = 6.0
FILL_MIN_EXTENT = 0.82
PAPER_MIN_CHANNEL = 240  # un color con todos los canales ≥ esto es papel, no relleno
MAX_FILL_BINS = 8
UNIFY_GAP = 1.2  # hueco (× tamaño) que corta un tramo de palabras que comparten negrita
SIDE_BEARING = 1.04  # la tinta de una palabra mide ~4 % menos que la suma de los avances de sus letras
STROKE_BASE_PX = 1.35
STROKE_PER_PX = 0.05
BOLD_RATIO = 1.2  # grosor relativo (ver _stroke_score) a partir del cual es negrita
BOLD_RATIO_SHORT = 1.3  # ídem para palabras de menos de 4 letras (más ruidosas)
GRAY_TEXT_MIN = 90  # tinta más clara que esto (0-255) se escribe con color; más oscura es negro


def _hex(rgb: tuple[int, int, int] | np.ndarray) -> str:
    r, g, b = (int(v) for v in rgb)
    return f"{r:02X}{g:02X}{b:02X}"


def render_page(page: Any, dpi: int = RENDER_DPI) -> tuple[np.ndarray, float]:
    """Render RGB (uint8) de la página en su marco visual y los píxeles por punto."""
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 1:
        arr = np.repeat(arr, 3, axis=2)
    elif pix.n > 3:
        arr = arr[:, :, :3]
    return np.ascontiguousarray(arr), pix.width / max(1.0, float(page.rect.width))


def _line_segments(rgb: np.ndarray, gray: np.ndarray, scale: float) -> list[Segment]:
    h, w = gray.shape
    # "más oscuro que lo que lo rodea": así una línea gris sobre un relleno gris no se confunde con el relleno
    side = max(5, round(MAX_LINE_THICKNESS_PT * 1.2 * scale)) | 1
    background = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, np.ones((side, side), np.uint8))
    ink = ((background.astype(np.int16) - gray.astype(np.int16) >= LINE_CONTRAST) & (gray < LINE_INK)).astype(
        np.uint8
    ) * 255
    k = max(8, round(LINE_MIN_PT * scale))
    max_thick = MAX_LINE_THICKNESS_PT * scale
    out: list[Segment] = []
    for orient, kernel in (("h", (k, 1)), ("v", (1, k))):
        opened = cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, kernel))
        n, labels, stats, _ = cv2.connectedComponentsWithStats(opened, connectivity=8)
        for i in range(1, n):
            x, y, cw, ch, _area = (int(v) for v in stats[i])
            length, thick = (cw, ch) if orient == "h" else (ch, cw)
            if thick > max_thick or length < k:
                continue
            if orient == "h" and (y < EDGE_MARGIN * h or y + ch > (1 - EDGE_MARGIN) * h):
                continue
            if orient == "v" and (x < EDGE_MARGIN * w or x + cw > (1 - EDGE_MARGIN) * w):
                continue
            mask = labels[y : y + ch, x : x + cw] == i
            color = _hex(np.median(rgb[y : y + ch, x : x + cw][mask], axis=0))
            width_pt = max(0.5, thick / scale)
            if orient == "h":
                out.append(Segment("h", (y + ch / 2.0) / scale, x / scale, (x + cw) / scale, width_pt, color))
            else:
                out.append(Segment("v", (x + cw / 2.0) / scale, y / scale, (y + ch) / scale, width_pt, color))
    if len(out) > MAX_SEGMENTS:
        out.sort(key=lambda s: s.b - s.a, reverse=True)
        out = out[:MAX_SEGMENTS]
    return merge_segments(out, gap=3.0)


def _fill_boxes(rgb: np.ndarray, scale: float, page_area_pt: float) -> list[Box]:
    """Regiones rectangulares de color liso (encabezado gris de una tabla, franjas, avisos)."""
    small = rgb.min(axis=2) < PAPER_MIN_CHANNEL
    dark = rgb.max(axis=2) <= 170  # tinta y líneas oscuras: no son relleno
    cand = small & ~dark
    if not cand.any():
        return []
    quant = (rgb.astype(np.int32) // 16) * (cand[:, :, None])
    keys = quant[:, :, 0] * 4096 + quant[:, :, 1] * 64 + quant[:, :, 2]
    keys = np.where(cand, keys + 1, 0)
    values, counts = np.unique(keys[keys > 0], return_counts=True)
    order = np.argsort(counts)[::-1][:MAX_FILL_BINS]
    min_px = (FILL_MIN_PT * scale) ** 2
    open_k = np.ones((max(3, round(1.2 * scale)),) * 2, np.uint8)
    close_k = np.ones((max(5, round(3.5 * scale)),) * 2, np.uint8)
    boxes: list[Box] = []
    for idx in order:
        if counts[idx] < min_px:
            continue
        mask = (keys == values[idx]).astype(np.uint8) * 255
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, open_k)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, close_k)
        n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=4)
        for i in range(1, n):
            x, y, cw, ch, area = (int(v) for v in stats[i])
            if cw < FILL_MIN_PT * scale or ch < FILL_MIN_PT * scale:
                continue
            if area < FILL_MIN_EXTENT * cw * ch:
                continue
            if cw * ch / (scale * scale) >= 0.85 * page_area_pt:
                continue
            region = rgb[y : y + ch, x : x + cw][labels[y : y + ch, x : x + cw] == i]
            color = _hex(np.median(region, axis=0))
            boxes.append(Box(x / scale, y / scale, (x + cw) / scale, (y + ch) / scale, color))
    for seq, b in enumerate(sorted(boxes, key=lambda b: (b.y0, b.x0)), start=1):
        b.seq = seq
    return boxes


def raster_primitives(rgb: np.ndarray, scale: float, page_size: tuple[float, float]) -> Primitives:
    """Segmentos de borde y rellenos que se ven en el render de una página escaneada."""
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    prims = Primitives()
    prims.segments = _line_segments(rgb, gray, scale)
    fills = _fill_boxes(rgb, scale, page_size[0] * page_size[1])
    prims.fills = fills
    prims.boxes = [Box(f.x0, f.y0, f.x1, f.y1, f.fill, f.seq) for f in fills]
    return prims


def _stroke_score(crop: np.ndarray, size_px: float) -> float | None:
    """Grosor del trazo (2 × área / perímetro de la tinta, en px) sobre el que tendría esa letra en fina. Con
    antialias y JPEG el trazo fino de una letra chica no baja de ~1.4 px, así que el grosor esperado es
    ``STROKE_BASE_PX + STROKE_PER_PX × tamaño``; ~1.0 es letra normal y ~1.4 negrita. None si no hay tinta."""
    lo, hi = float(crop.min()), float(np.percentile(crop, 95))
    if hi - lo < 40:
        return None
    ink = (crop < (lo + hi) / 2.0).astype(np.uint8)
    area = int(ink.sum())
    if area < 12:
        return None
    eroded = cv2.erode(ink, cv2.getStructuringElement(cv2.MORPH_CROSS, (3, 3)))
    boundary = area - int(eroded.sum())
    if boundary <= 0:
        return None
    return (2.0 * area / boundary) / (STROKE_BASE_PX + STROKE_PER_PX * size_px)


def scan_width_ratio(words: list[Word], rgb: np.ndarray, scale: float) -> float:
    """Cuánto más ancha es la tinta que la caja que da el OCR (la capa de texto de tesseract suele quedar
    ~10 % más angosta). Se mide sobre las palabras donde la tinta se separa limpia de las vecinas y se
    devuelve la mediana de la página, limitada a 1.0-1.3; sin datos suficientes, 1.0."""
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape
    ratios: list[float] = []
    for word in words:
        if len(word.text) < 5 or word.vertical:
            continue
        pad = max(1.5, 0.3 * word.size)
        xa, xb = max(0, int((word.x0 - pad) * scale)), min(w, int((word.x1 + pad) * scale) + 1)
        ya, yb = max(0, int(word.y0 * scale)), min(h, int(word.y1 * scale) + 1)
        if xb - xa < 8 or yb - ya < 4:
            continue
        cols = (gray[ya:yb, xa:xb] < 150).any(axis=0)
        gap = max(2, round(0.2 * word.size * scale))
        runs: list[list[int]] = []
        start = None
        for i, on in enumerate(cols):
            if on and start is None:
                start = i
            elif not on and start is not None:
                runs.append([start, i])
                start = None
        if start is not None:
            runs.append([start, len(cols)])
        merged: list[list[int]] = []
        for run in runs:
            if merged and run[0] - merged[-1][1] <= gap:
                merged[-1][1] = run[1]
            else:
                merged.append(run)
        centre = (word.x0 + word.x1) / 2.0 * scale - xa
        inked = next((r for r in merged if r[0] <= centre < r[1]), None)
        if inked is None:
            continue
        # la tinta no debe pegarse a los bordes de la ventana (se habría juntado con la palabra vecina)
        if inked[0] <= 0 or inked[1] >= len(cols):
            continue
        ratio = (inked[1] - inked[0]) / scale / max(1.0, word.x1 - word.x0)
        if 0.95 <= ratio <= 1.4:
            ratios.append(ratio)
    if len(ratios) < 5:
        return 1.0
    return float(min(1.3, max(1.0, float(np.median(ratios)) * SIDE_BEARING)))


def _unify_bold(words: list[Word], bold: dict[int, bool], position: dict[int, int]) -> None:
    """Las palabras de un mismo tramo del renglón (sin huecos grandes) comparten la negrita del voto
    mayoritario (por letras): un título no sale mitad negrita porque una palabra se leyó fina."""
    for line in group_lines(words):
        cluster: list[int] = []
        prev: Word | None = None
        for wd in [*line, None]:
            if wd is not None and (prev is None or wd.x0 - prev.x1 <= UNIFY_GAP * max(wd.size, prev.size)):
                cluster.append(position[id(wd)])
            else:
                if cluster:
                    weight = sum(len(words[j].text) * (1 if bold.get(j) else -1) for j in cluster)
                    if weight != 0:
                        for j in cluster:
                            bold[j] = weight > 0
                cluster = [position[id(wd)]] if wd is not None else []
            prev = wd


def apply_scan_styles(words: list[Word], rgb: np.ndarray, scale: float) -> list[Word]:
    """Negrita y color de cada palabra del OCR a partir de los píxeles (el OCR solo da el texto y la caja).
    Una palabra corta no decide sola: hereda la negrita de la vecina del mismo renglón si están pegadas."""
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape
    ratio: dict[int, float | None] = {}
    tone: dict[int, tuple[int, str | None]] = {}
    for k, word in enumerate(words):
        x0, y0 = max(0, int(word.x0 * scale)), max(0, int(word.y0 * scale))
        x1, y1 = min(w, int(word.x1 * scale) + 1), min(h, int(word.y1 * scale) + 1)
        if x1 - x0 < 4 or y1 - y0 < 4:
            ratio[k] = None
            continue
        crop = gray[y0:y1, x0:x1]
        ratio[k] = _stroke_score(crop, max(1.0, word.size * scale))
        order = np.argsort(crop.ravel())[: max(1, crop.size // 12)]
        level = int(np.mean(crop.ravel()[order]))
        color = None
        if level >= GRAY_TEXT_MIN:
            color = _hex(np.mean(rgb[y0:y1, x0:x1].reshape(-1, 3)[order], axis=0))
        tone[k] = (level, color)
    position = {id(wd): k for k, wd in enumerate(words)}
    bold: dict[int, bool] = {}
    for k, word in enumerate(words):
        r = ratio.get(k)
        if r is not None and len(word.text) >= 4:
            bold[k] = r >= BOLD_RATIO
    for line in group_lines(words):
        idx = [position[id(wd)] for wd in line]
        for pos, k in enumerate(idx):
            if k in bold:
                continue
            word = words[k]
            near = [
                j
                for j in (idx[pos - 1] if pos > 0 else None, idx[pos + 1] if pos + 1 < len(idx) else None)
                if j is not None
                and j in bold
                and abs(words[j].x0 - word.x1) <= 1.2 * word.size
                and abs(word.x0 - words[j].x1) <= 1.2 * word.size
            ]
            r = ratio.get(k)
            if near:
                bold[k] = (
                    any(bold[j] for j in near) if r is None else (r >= BOLD_RATIO_SHORT or bold[near[0]])
                )
            else:
                bold[k] = bool(r is not None and r >= BOLD_RATIO_SHORT and len(word.text) >= 2)
    _unify_bold(words, bold, position)
    for k, word in enumerate(words):
        is_bold = bold.get(k, False)
        color = tone.get(k, (0, None))[1]
        for i, ch in enumerate(word.chars):
            word.chars[i] = replace(ch, style=replace(ch.style, bold=is_bold, color=color))
    return words
