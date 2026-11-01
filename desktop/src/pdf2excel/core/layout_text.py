"""Texto del modo "Igual al PDF": caracteres → palabras → renglones → párrafos, con su estilo.

Parte de los spans de ``page.get_texttrace()`` (que trae fuente, tamaño, color y posición de cada carácter) en
vez de ``get_text("words")``, así el estilo de cada palabra viene sin adivinar. Sin Qt.
"""

from __future__ import annotations

import itertools
from dataclasses import dataclass, field
from typing import Any

from pdf2excel.core.layout_geom import Rect4, _Tf, rgb_hex
from pdf2excel.core.visibility import span_text  # noqa: F401  (re-export para layout.py)
from pdf2excel.models import TextRun

WORD_GAP = 0.3  # hueco entre caracteres (× tamaño) a partir del cual empieza otra palabra
LINE_TOL = 0.4  # tolerancia de línea base (× tamaño) para juntar palabras en un renglón
SEGMENT_GAP = 1.5  # hueco horizontal (× tamaño) que separa dos tramos del mismo renglón
PARAGRAPH_GAP = 1.6  # separación de líneas base (× tamaño) dentro de un párrafo
TABULAR_GAP = 1.3  # ídem si el renglón forma parte de una fila de columnas (listados)
SIZE_RATIO = 1.25  # cambio de tamaño que corta un párrafo
_LIGATURES = {
    0xFB00: "ff",
    0xFB01: "fi",
    0xFB02: "fl",
    0xFB03: "ffi",
    0xFB04: "ffl",
    0xFB05: "st",
    0xFB06: "st",
}
_SPACES = {32, 160, 9, 0x2002, 0x2003, 0x2009, 0x200A, 0x202F}


@dataclass(frozen=True)
class Style:
    bold: bool = False
    italic: bool = False
    size: float = 8.0
    color: str | None = None
    font: str | None = None


@dataclass
class Char:
    x0: float
    y0: float
    x1: float
    y1: float
    oy: float
    text: str
    style: Style


@dataclass
class Word:
    chars: list[Char]
    x0: float
    y0: float
    x1: float
    y1: float
    oy: float
    size: float
    vertical: bool = False
    seq: int = 0
    spaced: bool = (
        False  # el PDF trae un espacio de verdad entre la palabra anterior y esta (mismo texto corrido)
    )

    @property
    def text(self) -> str:
        return "".join(c.text for c in self.chars)

    @property
    def center(self) -> tuple[float, float]:
        return (self.x0 + self.x1) / 2.0, (self.y0 + self.y1) / 2.0


@dataclass
class Segment:
    """Tramo de un renglón: palabras consecutivas sin huecos grandes."""

    words: list[Word]
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0
    oy: float = 0.0
    size: float = 8.0
    siblings: int = 1  # cuántos tramos tiene su renglón (>1: fila de columnas, no un párrafo corrido)

    def __post_init__(self) -> None:
        self.x0 = min(w.x0 for w in self.words)
        self.x1 = max(w.x1 for w in self.words)
        self.y0 = min(w.y0 for w in self.words)
        self.y1 = max(w.y1 for w in self.words)
        self.size = max(w.size for w in self.words)
        self.oy = self.words[0].oy


@dataclass
class Block:
    """Párrafo de texto suelto: tramos apilados con el mismo borde (izquierdo, derecho o centrado)."""

    segs: list[Segment]
    mode: str | None = None  # "left" | "right" | "center" una vez que tiene 2 renglones
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0
    words: list[Word] = field(default_factory=list)

    def refresh(self) -> None:
        self.x0 = min(s.x0 for s in self.segs)
        self.x1 = max(s.x1 for s in self.segs)
        self.y0 = min(s.y0 for s in self.segs)
        self.y1 = max(s.y1 for s in self.segs)
        self.words = [w for s in self.segs for w in s.words]

    @property
    def rect(self) -> Rect4:
        return (self.x0, self.y0, self.x1, self.y1)


OCR_DOUBT_WARNING = "Escaneo leído con OCR: las celdas con una nota tienen texto que parece mal leído."


def doubtful_ocr_tokens(text: str) -> list[str]:
    """Palabras que un OCR probablemente leyó mal, por la forma de las mayúsculas.

    El texto que sale de un escaneo no trae confianza por palabra, pero un OCR en letra chica y gris deja
    marcas típicas: una palabra solo de letras que arranca en minúscula y tiene una mayúscula adentro
    (`curT` por `CUIT`) o que alterna minúscula-mayúscula varias veces. Un apellido como `McDonald` o `DeLuca`
    (arranca en mayúscula) no se marca."""
    out: list[str] = []
    for tok in text.replace("\n", " ").split():
        word = tok.strip(".,;:()[]¿?¡!\"'")
        if len(word) < 3 or not word.isalpha():
            continue
        flips = sum(1 for a, b in itertools.pairwise(word) if a.islower() and b.isupper())
        if flips >= 2 or (word[0].islower() and flips >= 1):
            out.append(word)
    return out


def font_family(name: str) -> str | None:
    low = name.lower()
    if "arial" in low or "helvetica" in low:
        return "Arial"
    if "times" in low:
        return "Times New Roman"
    if "courier" in low:
        return "Courier New"
    return None


def span_style(trace: dict[str, Any]) -> Style:
    font = str(trace.get("font", ""))
    low = font.lower()
    flags = int(trace.get("flags", 0))
    bold = bool(flags & 16) or any(k in low for k in ("bold", "black", "heavy", "semibold"))
    italic = bool(flags & 2) or "italic" in low or "oblique" in low
    color = rgb_hex(trace.get("color"))
    if color == "000000":
        color = None
    return Style(bold, italic, round(float(trace.get("size", 8.0)) * 2) / 2.0, color, font_family(font))


def _char_text(code: int) -> str:
    if code in _LIGATURES:
        return _LIGATURES[code]
    return chr(code)


def spans_to_words(spans: list[dict[str, Any]], matrix: Any) -> list[Word]:
    """Palabras visibles (en marco visual) a partir de los spans de ``get_texttrace`` que no están ocultos."""
    tf = _Tf(matrix)
    chars: list[Char] = []
    vertical_words: list[Word] = []
    seen: set[tuple[str, int, int]] = set()
    for span in spans:
        style = span_style(span)
        dx, dy = span.get("dir", (1.0, 0.0))
        if tf.m is not None and not tf.identity:
            a, b, c, d = tf.m[0], tf.m[1], tf.m[2], tf.m[3]
            dx, dy = a * dx + c * dy, b * dx + d * dy
        is_vertical = abs(dy) > 0.3
        span_chars: list[Char] = []
        for ch in span.get("chars", ()):
            code = ch[0]
            if code < 32 and code not in _SPACES:
                continue
            x0, y0, x1, y1 = tf.rect(ch[3])
            _ox, oy = tf.pt(ch[2][0], ch[2][1])
            text = " " if code in _SPACES else _char_text(code)
            key = (text, round(x0 * 2), round(oy * 2))
            if key in seen and text != " ":
                continue  # texto repetido encima de sí mismo (falsa negrita / sombra)
            seen.add(key)
            span_chars.append(Char(x0, y0, x1, y1, oy, text, style))
        if is_vertical:
            # los espacios internos se conservan ("VERTICAL TEXT"); solo se recortan los de las puntas
            while span_chars and span_chars[0].text == " ":
                span_chars.pop(0)
            while span_chars and span_chars[-1].text == " ":
                span_chars.pop()
            if span_chars:
                vertical_words.append(_make_word(span_chars, vertical=True))
        else:
            chars.extend(span_chars)
    words = _chars_to_words(chars)
    words.extend(vertical_words)
    for i, w in enumerate(words):
        w.seq = i
    return words


def _make_word(chars: list[Char], vertical: bool = False) -> Word:
    return Word(
        chars=chars,
        x0=min(c.x0 for c in chars),
        y0=min(c.y0 for c in chars),
        x1=max(c.x1 for c in chars),
        y1=max(c.y1 for c in chars),
        oy=chars[0].oy,
        size=max(c.style.size for c in chars),
        vertical=vertical,
    )


def _chars_to_words(chars: list[Char]) -> list[Word]:
    if not chars:
        return []
    ordered = sorted(chars, key=lambda c: (c.oy, c.x0))
    rows: list[list[Char]] = []
    for ch in ordered:
        if rows and abs(ch.oy - rows[-1][0].oy) <= LINE_TOL * max(ch.style.size, 1.0) * 0.75:
            rows[-1].append(ch)
        else:
            rows.append([ch])
    words: list[Word] = []
    spaced_starts: set[int] = set()
    for row in rows:
        row.sort(key=lambda c: c.x0)
        current: list[Char] = []
        prev: Char | None = None
        after_space = False
        for ch in row:
            if ch.text == " ":
                if current:
                    words.append(_make_word(current))
                    current = []
                prev = ch
                after_space = bool(words)
                continue
            if current and prev is not None and prev.text != " ":
                gap = ch.x0 - prev.x1
                if gap > WORD_GAP * max(ch.style.size, prev.style.size):
                    words.append(_make_word(current))
                    current = []
            pending_space = not current and after_space
            current.append(ch)
            prev = ch
            if pending_space:
                spaced_starts.add(id(ch))
            after_space = False
        if current:
            words.append(_make_word(current))
    for w in words:
        w.spaced = id(w.chars[0]) in spaced_starts
    return words


def group_lines(words: list[Word]) -> list[list[Word]]:
    """Agrupa palabras en renglones por línea base y las ordena de izquierda a derecha."""
    lines: list[list[Word]] = []
    for w in sorted(words, key=lambda w: (w.oy, w.x0)):
        if lines and not w.vertical and not lines[-1][0].vertical:
            ref = lines[-1][0]
            if abs(w.oy - ref.oy) <= LINE_TOL * max(w.size, ref.size):
                lines[-1].append(w)
                continue
        lines.append([w])
    for line in lines:
        line.sort(key=lambda w: w.x0)
    return lines


def split_segments(
    line: list[Word], gap_factor: float = SEGMENT_GAP, cuts: list[float] | None = None
) -> list[Segment]:
    """Parte un renglón en tramos por huecos grandes y por ``cuts`` (posiciones x donde empieza otro campo,
    p. ej. el comienzo de la línea de subrayado sobre la que se escribe un valor)."""
    segs: list[Segment] = []
    current: list[Word] = []
    for w in line:
        cut_here = bool(current) and any(current[-1].x0 < c <= w.x0 + 1.0 for c in cuts or ())
        if current and (
            w.x0 - current[-1].x1 > gap_factor * max(w.size, current[-1].size) or w.vertical or cut_here
        ):
            segs.append(Segment(current))
            current = []
        current.append(w)
    if current:
        segs.append(Segment(current))
    for seg in segs:
        seg.siblings = len(segs)
    return segs


def line_runs(words: list[Word]) -> list[TextRun]:
    """Un renglón → runs por estilo (palabras separadas por un espacio)."""
    runs: list[TextRun] = []
    last_style: Style | None = None
    buffer = ""
    for i, w in enumerate(words):
        if i > 0 and last_style is not None:
            # el espacio hereda el estilo de la palabra anterior
            buffer += " "
        for ch in w.chars:
            if last_style is not None and ch.style != last_style and buffer:
                runs.append(_run(buffer, last_style))
                buffer = ""
            last_style = ch.style
            buffer += ch.text
    if buffer and last_style is not None:
        runs.append(_run(buffer, last_style))
    # recortar espacios de los extremos
    while runs and not runs[-1].text.strip():
        runs.pop()
    if runs:
        runs[-1].text = runs[-1].text.rstrip()
        runs[0].text = runs[0].text.lstrip()
    return [r for r in runs if r.text]


def _run(text: str, style: Style) -> TextRun:
    return TextRun(text, style.bold, style.italic, style.size, style.color, style.font)


def build_blocks(segs: list[Segment]) -> list[Block]:
    """Junta tramos en párrafos: renglones consecutivos con el mismo borde y tamaño parecido."""
    blocks: list[Block] = []
    for seg in sorted(segs, key=lambda s: (s.oy, s.x0)):
        best: Block | None = None
        best_d = 1e9
        best_mode: str | None = None
        for blk in blocks:
            last = blk.segs[-1]
            d = seg.oy - last.oy
            size = max(seg.size, last.size)
            gap_limit = TABULAR_GAP if max(seg.siblings, last.siblings) > 1 else PARAGRAPH_GAP
            if d <= 0.5 * min(seg.size, last.size) or d > gap_limit * size:
                continue
            ratio = max(seg.size, last.size) / max(1.0, min(seg.size, last.size))
            if ratio > SIZE_RATIO or seg.x1 < blk.x0 or seg.x0 > blk.x1:
                continue
            tol = max(1.5, 0.5 * size)
            modes = {
                "left": abs(seg.x0 - last.x0) <= tol,
                "right": abs(seg.x1 - last.x1) <= tol,
                "center": abs((seg.x0 + seg.x1) - (last.x0 + last.x1)) <= 2 * tol,
            }
            if len(blk.segs) == 1 and blk.mode is None:
                mode = next((m for m in ("left", "right", "center") if modes[m]), None)
            else:
                mode = blk.mode if blk.mode and modes[blk.mode] else None
            if mode is None:
                continue
            if d < best_d:
                best, best_d, best_mode = blk, d, mode
        if best is None:
            blk = Block([seg])
            blk.refresh()
            blocks.append(blk)
        else:
            best.segs.append(seg)
            best.mode = best_mode
            best.refresh()
    return blocks
