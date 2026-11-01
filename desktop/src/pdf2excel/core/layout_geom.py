"""Geometría del modo "Igual al PDF": bordes, cajas y rellenos a partir de ``page.get_drawings()``.

Todo en puntos PDF, marco visual (ya aplicada la rotación de la página), origen arriba a la izquierda.
Sin Qt. Lo usa ``core/layout.py``.
"""

from __future__ import annotations

from bisect import bisect_left
from dataclasses import dataclass, field
from typing import Any

THIN_MAX = 1.6  # un relleno más fino que esto (pt) es un borde dibujado como rectángulo
RULE_MAX = 3.5  # un relleno hasta este grosor y muy alargado también es una línea
RULE_RATIO = 10.0
MIN_SEGMENT = 2.0  # largo mínimo de un tramo de borde
MIN_BOX = 3.0  # lado mínimo de una caja
MERGE_GAP = 1.5  # huecos de hasta este largo entre tramos colineales se unen
WHITE_LEVEL = 0.97  # canales >= este valor = blanco/sin relleno
MAX_MARK = 24.0  # lado máximo (pt) de un tilde/cruz dibujado
MIN_MARK = 2.5
MAX_FLOOD_UNITS = 60000
MAX_FILLS = 1500  # más rellenos que esto es arte vectorial (mapas, gráficos): no se usan para colorear celdas
BG_PAGE_FRACTION = 0.85  # cajas que cubren esto de la página son fondo/marco: no son una celda

Rect4 = tuple[float, float, float, float]


@dataclass
class Segment:
    orient: str  # "h" | "v"
    pos: float  # y si es horizontal, x si es vertical
    a: float  # inicio
    b: float  # fin
    width: float
    color: str


@dataclass
class Box:
    x0: float
    y0: float
    x1: float
    y1: float
    fill: str | None = None
    seq: int = 0

    @property
    def rect(self) -> Rect4:
        return (self.x0, self.y0, self.x1, self.y1)

    @property
    def area(self) -> float:
        return max(0.0, self.x1 - self.x0) * max(0.0, self.y1 - self.y0)


@dataclass
class Mark:
    """Tilde o cruz dibujado con trazos diagonales (no es texto: el PDF lo dibuja como vector)."""

    rect: Rect4
    glyph: str  # "✔" | "✗"
    color: str


@dataclass
class Primitives:
    segments: list[Segment] = field(default_factory=list)
    boxes: list[Box] = field(default_factory=list)  # cajas dibujadas (trazo o relleno)
    fills: list[Box] = field(default_factory=list)  # rellenos de color (para colorear celdas)
    whites: list[Box] = field(default_factory=list)  # rellenos blancos (anulan, por orden, un color de abajo)
    marks: list[Mark] = field(default_factory=list)  # tildes y cruces dibujados con trazos (casillas)


def rgb_hex(color: Any, alpha: float | None = None) -> str | None:
    """Color de PyMuPDF (gris/RGB/CMYK) → ``"RRGGBB"``; mezcla con blanco si ``alpha`` < 1."""
    if color is None:
        return None
    values = [float(v) for v in color]
    if len(values) == 1:
        r = g = b = values[0]
    elif len(values) == 3:
        r, g, b = values
    elif len(values) == 4:
        c, m, y, k = values
        r, g, b = (1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)
    else:
        return None
    if alpha is not None and alpha < 1.0:
        r, g, b = (1 - alpha) + alpha * r, (1 - alpha) + alpha * g, (1 - alpha) + alpha * b
    return "{:02X}{:02X}{:02X}".format(*(max(0, min(255, round(v * 255))) for v in (r, g, b)))


def is_whiteish(hex_color: str | None) -> bool:
    if hex_color is None:
        return True
    return all(int(hex_color[i : i + 2], 16) / 255.0 >= WHITE_LEVEL for i in (0, 2, 4))


def cluster_values(values: list[float], tol: float) -> list[float]:
    """Agrupa valores ordenados en cadenas con huecos ≤ ``tol`` (ancho máximo 2×tol) y devuelve la media."""
    if not values:
        return []
    out: list[float] = []
    group: list[float] = []
    for v in sorted(values):
        if group and (v - group[-1] > tol or v - group[0] > 2.0 * tol):
            out.append(sum(group) / len(group))
            group = []
        group.append(v)
    if group:
        out.append(sum(group) / len(group))
    return out


def nearest_index(lines: list[float], value: float) -> int:
    """Índice de la línea más cercana a ``value`` en una lista ordenada."""
    i = bisect_left(lines, value)
    if i <= 0:
        return 0
    if i >= len(lines):
        return len(lines) - 1
    return i if lines[i] - value < value - lines[i - 1] else i - 1


class _Tf:
    """Transformación afín (a, b, c, d, e, f) de PyMuPDF; ``None`` = identidad."""

    def __init__(self, matrix: Any) -> None:
        m = (matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f) if matrix is not None else None
        self.identity = m is None or m == (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        self.m = m

    def pt(self, x: float, y: float) -> tuple[float, float]:
        if self.identity or self.m is None:
            return x, y
        a, b, c, d, e, f = self.m
        return a * x + c * y + e, b * x + d * y + f

    def rect(self, r: Any) -> Rect4:
        if self.identity or self.m is None:
            return (r[0], r[1], r[2], r[3])
        xs, ys = zip(
            *(self.pt(r[0], r[1]), self.pt(r[2], r[1]), self.pt(r[2], r[3]), self.pt(r[0], r[3])), strict=True
        )
        return (min(xs), min(ys), max(xs), max(ys))


def _add_rect_segments(out: list[Segment], r: Rect4, width: float, color: str) -> None:
    x0, y0, x1, y1 = r
    out.append(Segment("h", y0, x0, x1, width, color))
    out.append(Segment("h", y1, x0, x1, width, color))
    out.append(Segment("v", x0, y0, y1, width, color))
    out.append(Segment("v", x1, y0, y1, width, color))


def kinds_are_rects(items: list[Any]) -> bool:
    return all(it[0] in ("re", "qu") for it in items)


def _thin_segment(r: Rect4, color: str) -> Segment | None:
    w, h = r[2] - r[0], r[3] - r[1]
    if w >= h:
        if w < MIN_SEGMENT:
            return None
        return Segment("h", (r[1] + r[3]) / 2.0, r[0], r[2], max(h, 0.25), color)
    if h < MIN_SEGMENT:
        return None
    return Segment("v", (r[0] + r[2]) / 2.0, r[1], r[3], max(w, 0.25), color)


def _is_thin(r: Rect4) -> bool:
    w, h = r[2] - r[0], r[3] - r[1]
    small, big = min(w, h), max(w, h)
    if small <= THIN_MAX and big >= MIN_SEGMENT:
        return True
    return small <= RULE_MAX and big >= RULE_RATIO * small


def _poly_area_ratio(items: list[Any], rect: Rect4) -> float:
    pts = [(it[1].x, it[1].y) for it in items if it[0] in ("l", "c")]
    bbox_area = (rect[2] - rect[0]) * (rect[3] - rect[1])
    if len(pts) < 3 or bbox_area <= 0:
        return 0.0
    twice = 0.0
    for (x0, y0), (x1, y1) in zip(pts, pts[1:] + pts[:1], strict=True):
        twice += x0 * y1 - x1 * y0
    return abs(twice) / 2.0 / bbox_area


def extract_primitives(
    drawings: list[dict[str, Any]], matrix: Any, page_rect: Rect4, covered: set[int] | None = None
) -> Primitives:
    """Segmentos de borde, cajas y rellenos de la página. ``covered``: ``seqno`` de dibujos a ignorar."""
    tf = _Tf(matrix)
    prims = Primitives()
    page_area = max(1.0, (page_rect[2] - page_rect[0]) * (page_rect[3] - page_rect[1]))

    def add_box(r: Rect4, fill: str | None, seq: int) -> None:
        w, h = r[2] - r[0], r[3] - r[1]
        if w < MIN_BOX or h < MIN_BOX or w * h >= BG_PAGE_FRACTION * page_area:
            return
        prims.boxes.append(Box(r[0], r[1], r[2], r[3], fill, seq))

    stray: list[tuple[tuple[float, float], tuple[float, float], str]] = []
    for d in drawings:
        seq = int(d.get("seqno", 0))
        kind = d["type"]
        items = d["items"]
        if not items or (covered is not None and seq in covered):
            continue
        fill_hex: str | None = None
        if kind in ("f", "fs"):
            fill_hex = rgb_hex(d.get("fill"), d.get("fill_opacity"))
            if is_whiteish(fill_hex):
                fill_hex = None
                if kinds_are_rects(items):
                    for it in items:
                        r = tf.rect(it[1] if it[0] == "re" else it[1].rect)
                        if (
                            (r[2] - r[0]) >= MIN_BOX
                            and (r[3] - r[1]) >= MIN_BOX
                            and (r[2] - r[0]) * (r[3] - r[1]) < BG_PAGE_FRACTION * page_area
                        ):
                            prims.whites.append(Box(r[0], r[1], r[2], r[3], None, seq))
        stroke_hex = rgb_hex(d.get("color")) if kind in ("s", "fs") else None
        width = float(d.get("width") or 0.5)
        kinds = {it[0] for it in items}
        # --- relleno de color
        if fill_hex is not None:
            if kinds <= {"re", "qu"}:
                for it in items:
                    r = tf.rect(it[1] if it[0] == "re" else it[1].rect)
                    if _is_thin(r):
                        seg = _thin_segment(r, fill_hex)
                        if seg is not None:
                            prims.segments.append(seg)
                    elif (r[2] - r[0]) >= MIN_BOX and (r[3] - r[1]) >= MIN_BOX:
                        if (r[2] - r[0]) * (r[3] - r[1]) < BG_PAGE_FRACTION * page_area:
                            prims.fills.append(Box(r[0], r[1], r[2], r[3], fill_hex, seq))
                            add_box(r, fill_hex, seq)
            else:
                r = tf.rect(d["rect"])
                if _is_thin(r) and _poly_area_ratio(items, r) >= 0.8:
                    # línea dibujada como un relleno fino hecho con trazos ("l"), no con "re"
                    seg = _thin_segment(r, fill_hex)
                    if seg is not None:
                        prims.segments.append(seg)
                elif (
                    not _is_thin(r)
                    and (r[2] - r[0]) >= MIN_BOX
                    and (r[3] - r[1]) >= MIN_BOX
                    and _poly_area_ratio(items, r) >= 0.8
                    and (r[2] - r[0]) * (r[3] - r[1]) < BG_PAGE_FRACTION * page_area
                ):
                    prims.fills.append(Box(r[0], r[1], r[2], r[3], fill_hex, seq))
                    add_box(r, fill_hex, seq)
        # --- trazo
        if stroke_hex is None:
            continue
        mark = _classify_mark(items, tf.rect(d["rect"]), stroke_hex)
        if mark is not None:
            prims.marks.append(mark)
            continue
        if len(items) == 1 and items[0][0] == "l":
            # los renglones de un tilde pueden venir como dibujos sueltos de una línea cada uno
            a, b = items[0][1], items[0][2]
            p0, p1 = tf.pt(a.x, a.y), tf.pt(b.x, b.y)
            if (
                abs(p1[0] - p0[0]) > 0.6
                and abs(p1[1] - p0[1]) > 0.6
                and max(abs(p1[0] - p0[0]), abs(p1[1] - p0[1])) <= MAX_MARK
            ):
                stray.append((p0, p1, stroke_hex))
                continue
        has_curves = "c" in kinds
        n_lines = sum(1 for it in items if it[0] == "l")
        if has_curves and kinds <= {"l", "c"} and n_lines >= 4 and _has_both_axes(items):
            r = tf.rect(d["rect"])
            if (r[2] - r[0]) >= MIN_BOX and (r[3] - r[1]) >= MIN_BOX:
                # caja con esquinas redondeadas: lados completos del rectángulo
                _add_rect_segments(prims.segments, r, width, stroke_hex)
                add_box(r, fill_hex, seq)
                continue
        for it in items:
            op = it[0]
            if op == "re" or (op == "qu" and _axis_aligned(it[1])):
                r = tf.rect(it[1] if op == "re" else it[1].rect)
                w, h = r[2] - r[0], r[3] - r[1]
                if min(w, h) < 1.0:
                    seg = _thin_segment(r, stroke_hex)
                    if seg is not None:
                        seg.width = max(seg.width, width)
                        prims.segments.append(seg)
                else:
                    _add_rect_segments(prims.segments, r, width, stroke_hex)
                    add_box(r, fill_hex, seq)
            elif op == "l":
                x0, y0 = tf.pt(it[1].x, it[1].y)
                x1, y1 = tf.pt(it[2].x, it[2].y)
                if abs(y1 - y0) <= 0.6 and abs(x1 - x0) >= MIN_SEGMENT:
                    prims.segments.append(
                        Segment("h", (y0 + y1) / 2.0, min(x0, x1), max(x0, x1), width, stroke_hex)
                    )
                elif abs(x1 - x0) <= 0.6 and abs(y1 - y0) >= MIN_SEGMENT:
                    prims.segments.append(
                        Segment("v", (x0 + x1) / 2.0, min(y0, y1), max(y0, y1), width, stroke_hex)
                    )
    prims.marks.extend(_group_stray_marks(stray))
    if not prims.fills or len(prims.whites) > MAX_FILLS:
        prims.whites = []  # sin colores debajo no hay nada que anular
    if len(prims.fills) > MAX_FILLS:
        fill_seqs = {f.seq for f in prims.fills}
        prims.boxes = [b for b in prims.boxes if b.fill is None or b.seq not in fill_seqs]
        prims.fills = []
    prims.segments = merge_segments(prims.segments)
    return prims


def _classify_mark(items: list[Any], r: Rect4, color: str) -> Mark | None:
    """¿El trazo es un tilde (dos renglones en V) o una cruz dibujados con líneas diagonales y chicos?"""
    w, h = r[2] - r[0], r[3] - r[1]
    if not (MIN_MARK <= w <= MAX_MARK and MIN_MARK <= h <= MAX_MARK):
        return None
    lines = [it for it in items if it[0] == "l"]
    if len(lines) < 2 or len(lines) != len(items) or len(lines) > 4:
        return None
    diagonal = [it for it in lines if abs(it[2].x - it[1].x) > 0.6 and abs(it[2].y - it[1].y) > 0.6]
    if len(diagonal) != len(lines):
        return None
    if len(lines) == 2:
        a, b = lines
        shared = [
            (p, q)
            for p in (a[1], a[2])
            for q in (b[1], b[2])
            if abs(p.x - q.x) < 0.5 and abs(p.y - q.y) < 0.5
        ]
        if shared:
            return Mark(r, "✔", color)
    return Mark(r, "✗", color)


def _group_stray_marks(stray: list[tuple[tuple[float, float], tuple[float, float], str]]) -> list[Mark]:
    """Junta renglones diagonales sueltos que se tocan (extremo con extremo o cruzados) en tildes y cruces."""
    n = len(stray)
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def box(i: int) -> Rect4:
        (x0, y0), (x1, y1), _c = stray[i]
        return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))

    for i in range(n):
        bi = box(i)
        for j in range(i + 1, n):
            bj = box(j)
            if (
                min(bi[2], bj[2]) - max(bi[0], bj[0]) >= -0.6
                and min(bi[3], bj[3]) - max(bi[1], bj[1]) >= -0.6
                and stray[i][2] == stray[j][2]
            ):
                parent[find(i)] = find(j)
    groups: dict[int, list[int]] = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)
    marks: list[Mark] = []
    for members in groups.values():
        if len(members) > 4:
            continue
        boxes = [box(i) for i in members]
        r = (
            min(b[0] for b in boxes),
            min(b[1] for b in boxes),
            max(b[2] for b in boxes),
            max(b[3] for b in boxes),
        )
        if not (MIN_MARK <= r[2] - r[0] <= MAX_MARK and MIN_MARK <= r[3] - r[1] <= MAX_MARK):
            continue
        glyph = "✗"
        if len(members) == 2:
            a, b = stray[members[0]], stray[members[1]]
            shared = [
                (p, q)
                for p in (a[0], a[1])
                for q in (b[0], b[1])
                if abs(p[0] - q[0]) < 0.5 and abs(p[1] - q[1]) < 0.5
            ]
            if shared:
                glyph = "✔"
        marks.append(Mark(r, glyph, stray[members[0]][2]))
    return marks


def _has_both_axes(items: list[Any]) -> bool:
    horizontal = vertical = False
    for it in items:
        if it[0] != "l":
            continue
        dx, dy = abs(it[2].x - it[1].x), abs(it[2].y - it[1].y)
        horizontal = horizontal or (dy <= 0.6 and dx >= MIN_SEGMENT)
        vertical = vertical or (dx <= 0.6 and dy >= MIN_SEGMENT)
    return horizontal and vertical


def _axis_aligned(quad: Any) -> bool:
    pts = [quad.ul, quad.ur, quad.lr, quad.ll]
    return all(
        abs(a.x - b.x) < 0.5 or abs(a.y - b.y) < 0.5 for a, b in zip(pts, pts[1:] + pts[:1], strict=True)
    )


def merge_segments(segments: list[Segment], gap: float = MERGE_GAP) -> list[Segment]:
    """Une tramos colineales contiguos (mismo color y grosor parecido)."""
    out: list[Segment] = []
    for orient in ("h", "v"):
        group = sorted((s for s in segments if s.orient == orient), key=lambda s: (s.pos, s.a))
        # agrupar por posición (±0.6)
        lines: list[list[Segment]] = []
        for s in group:
            if lines and s.pos - lines[-1][0].pos <= 0.6:
                lines[-1].append(s)
            else:
                lines.append([s])
        for line in lines:
            line.sort(key=lambda s: (s.color, round(s.width * 2), s.a))
            current: Segment | None = None
            for s in line:
                if (
                    current is not None
                    and s.color == current.color
                    and abs(s.width - current.width) <= 0.5
                    and s.a <= current.b + gap
                    and s.b >= current.a - gap
                ):
                    current.a = min(current.a, s.a)
                    current.b = max(current.b, s.b)
                    current.width = max(current.width, s.width)
                else:
                    if current is not None:
                        out.append(current)
                    current = Segment(s.orient, s.pos, s.a, s.b, s.width, s.color)
            if current is not None:
                out.append(current)
    out.sort(key=lambda s: (s.orient, s.pos, s.a))
    return out


def find_closed_boxes(segments: list[Segment], tol: float = 2.0) -> list[Rect4]:
    """Regiones rectangulares cerradas por segmentos (celdas de una tabla dibujada con líneas)."""
    xs_raw: list[float] = []
    ys_raw: list[float] = []
    for s in segments:
        if s.orient == "v":
            xs_raw.append(s.pos)
            ys_raw.extend((s.a, s.b))
        else:
            ys_raw.append(s.pos)
            xs_raw.extend((s.a, s.b))
    xs = cluster_values(xs_raw, tol)
    ys = cluster_values(ys_raw, tol)
    nx, ny = len(xs), len(ys)
    if nx < 2 or ny < 2 or (nx - 1) * (ny - 1) > MAX_FLOOD_UNITS:
        return []
    horiz = [bytearray(nx - 1) for _ in range(ny)]  # horiz[i][c]: borde en ys[i] sobre la unidad c
    vert = [bytearray(ny - 1) for _ in range(nx)]  # vert[j][r]: borde en xs[j] junto a la unidad r
    for s in segments:
        if s.orient == "h":
            i = nearest_index(ys, s.pos)
            for c in range(nearest_index(xs, s.a), nearest_index(xs, s.b)):
                horiz[i][c] = 1
        else:
            j = nearest_index(xs, s.pos)
            for r in range(nearest_index(ys, s.a), nearest_index(ys, s.b)):
                vert[j][r] = 1
    label = [[False] * (nx - 1) for _ in range(ny - 1)]
    found: list[Rect4] = []
    for r0 in range(ny - 1):
        for c0 in range(nx - 1):
            if label[r0][c0]:
                continue
            label[r0][c0] = True
            stack = [(r0, c0)]
            count = 0
            rmin = rmax = r0
            cmin = cmax = c0
            while stack:
                r, c = stack.pop()
                count += 1
                rmin, rmax = min(rmin, r), max(rmax, r)
                cmin, cmax = min(cmin, c), max(cmax, c)
                if c + 1 < nx - 1 and not label[r][c + 1] and not vert[c + 1][r]:
                    label[r][c + 1] = True
                    stack.append((r, c + 1))
                if c > 0 and not label[r][c - 1] and not vert[c][r]:
                    label[r][c - 1] = True
                    stack.append((r, c - 1))
                if r + 1 < ny - 1 and not label[r + 1][c] and not horiz[r + 1][c]:
                    label[r + 1][c] = True
                    stack.append((r + 1, c))
                if r > 0 and not label[r - 1][c] and not horiz[r][c]:
                    label[r - 1][c] = True
                    stack.append((r - 1, c))
            n_rows, n_cols = rmax - rmin + 1, cmax - cmin + 1
            if count != n_rows * n_cols:
                continue
            top = sum(horiz[rmin][c] for c in range(cmin, cmax + 1)) / n_cols
            bottom = sum(horiz[rmax + 1][c] for c in range(cmin, cmax + 1)) / n_cols
            left = sum(vert[cmin][r] for r in range(rmin, rmax + 1)) / n_rows
            right = sum(vert[cmax + 1][r] for r in range(rmin, rmax + 1)) / n_rows
            if min(top, bottom, left, right) < 0.6 or (top + bottom + left + right) / 4.0 < 0.85:
                continue
            x0, x1, y0, y1 = xs[cmin], xs[cmax + 1], ys[rmin], ys[rmax + 1]
            if x1 - x0 >= MIN_BOX and y1 - y0 >= MIN_BOX:
                found.append((x0, y0, x1, y1))
    return found


def find_open_boxes(segments: list[Segment], tol: float = 1.5) -> list[Rect4]:
    """Cajas con tres lados dibujados (dos verticales y una horizontal que las une), p. ej. un recuadro de
    observaciones sin borde superior. El lado que falta no se dibuja."""
    buckets: dict[int, list[Segment]] = {}
    for s in segments:
        if s.orient == "v" and s.b - s.a >= MIN_BOX:
            buckets.setdefault(round(s.pos / 4.0), []).append(s)

    def near(x: float) -> list[Segment]:
        key = round(x / 4.0)
        found = [v for k in (key - 1, key, key + 1) for v in buckets.get(k, ()) if abs(v.pos - x) <= tol]
        return found

    found_boxes: list[Rect4] = []
    for h in segments:
        if h.orient != "h" or h.b - h.a < 2 * MIN_BOX:
            continue
        lefts, rights = near(h.a), near(h.b)
        for lv in lefts:
            for rv in rights:
                for end_l, end_r, other_l, other_r, upward in (
                    (lv.b, rv.b, lv.a, rv.a, True),  # la horizontal está abajo
                    (lv.a, rv.a, lv.b, rv.b, False),  # la horizontal está arriba
                ):
                    if (
                        abs(end_l - h.pos) <= tol
                        and abs(end_r - h.pos) <= tol
                        and abs(other_l - other_r) <= tol
                    ):
                        y_far = (other_l + other_r) / 2.0
                        y0, y1 = (y_far, h.pos) if upward else (h.pos, y_far)
                        if y1 - y0 >= MIN_BOX:
                            found_boxes.append((h.a, y0, h.b, y1))
    return found_boxes


def iou_close(a: Rect4, b: Rect4, tol: float = 1.5) -> bool:
    return all(abs(p - q) <= tol for p, q in zip(a, b, strict=True))
