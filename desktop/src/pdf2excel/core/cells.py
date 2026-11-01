"""Texto de celdas a partir de palabras, limpieza y reparación geométrica de tablas (DESIGN.md §5.3).

Funciones puras sobre la salida de ``page.get_text("words")`` y los objetos ``Table`` de pymupdf.
Nunca se usa ``Table.extract()`` (rompe "350.00" → "35000\\n."): el texto de cada celda se arma con las
palabras cuyo centro cae dentro del rectángulo de la celda.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from typing import Any

import pymupdf

# (x0, y0, x1, y1, text, block_no, line_no, word_no) tal como lo devuelve page.get_text("words")
Word = tuple[float, float, float, float, str, int, int, int]
RectLike = Any  # pymupdf.Rect | tuple[float, float, float, float] | pymupdf.IRect

_PIPE_TOKEN_RE = re.compile(r"(?<!\S)[|¦!]+(?!\S)")  # tokens que son solo barras (bordes alucinados por OCR)
_SPACES_RE = re.compile(r"[ \t   ]+")
_HYPHEN_BREAK_RE = re.compile(r"(?<=[^\W\d_])-$")  # letra seguida de guión al final de un renglón


def _center(w: Sequence[float]) -> tuple[float, float]:
    return (w[0] + w[2]) / 2.0, (w[1] + w[3]) / 2.0


def words_in_rect(words: list[Word] | list[tuple], rect: RectLike) -> str:
    """Texto de la celda `rect`: palabras cuyo centro cae en el rect, agrupadas en renglones.

    Los renglones se forman por centro-y con tolerancia ``0.45 * max(h_palabra, h_renglón)`` (robusto a OCR),
    las palabras de cada renglón se ordenan por x y los renglones se unen con ``"\\n"``.
    """
    r = pymupdf.Rect(rect)
    selected: list[tuple] = []
    for w in words:
        cx, cy = _center(w)
        if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1:
            selected.append(w)
    if not selected:
        return ""
    selected.sort(key=lambda w: (_center(w)[1], w[0]))

    # cada renglón: [centro_y, altura, palabras]
    lines: list[list[Any]] = []
    for w in selected:
        cy = _center(w)[1]
        h = float(w[3] - w[1])
        if lines and abs(cy - lines[-1][0]) <= 0.45 * max(h, lines[-1][1]):
            lines[-1][2].append(w)
            lines[-1][1] = max(lines[-1][1], h)
        else:
            lines.append([cy, h, [w]])
    return "\n".join(" ".join(str(x[4]) for x in sorted(ws, key=lambda x: x[0])) for _, _, ws in lines)


def _join_wrapped_lines(lines: list[str]) -> str:
    """Une renglones de una celda con espacio; "premi-" + "um" → "premium" (guión de corte)."""
    out = ""
    for line in lines:
        if not out:
            out = line
        elif _HYPHEN_BREAK_RE.search(out) and line[:1].islower():
            out = out[:-1] + line
        else:
            out = f"{out} {line}"
    return out


def clean_cell(s: str | None) -> str:
    """Normaliza el texto de una celda.

    Reemplaza tokens que son solo ``|¦!`` por espacio, NBSP → espacio, colapsa espacios, une renglones
    (wrap de celda → un renglón), quita guiones de corte a fin de renglón y hace ``strip(" |")``.
    """
    if not s:
        return ""
    s = _PIPE_TOKEN_RE.sub(" ", str(s))
    s = _SPACES_RE.sub(" ", s)
    lines = [ln.strip() for ln in s.split("\n")]
    lines = [ln for ln in lines if ln]
    return _join_wrapped_lines(lines).strip(" |")


def _column_x_ranges(table: Any) -> list[tuple[float, float]]:
    """Límites x por columna: unión de las celdas no nulas de esa columna ±2 pt; (0, 0) si no hay ninguna."""
    ranges: list[tuple[float, float]] = []
    for ci in range(table.col_count):
        cells = [row.cells[ci] for row in table.rows if ci < len(row.cells) and row.cells[ci]]
        if cells:
            ranges.append((min(c[0] for c in cells) - 2.0, max(c[2] for c in cells) + 2.0))
        else:
            ranges.append((0.0, 0.0))
    return ranges


def _in_some_range(x: float, ranges: list[tuple[float, float]]) -> bool:
    return any(x0 <= x <= x1 for x0, x1 in ranges)


MAX_TAIL_ROWS = 50  # tope de renglones que puede agregar extend_tail (los totales son 1-5 filas)
_TAIL_LINE_TOL_PT = 3.0  # palabras con |Δy0| < 3 pt forman el mismo renglón
_MAX_TAIL_WORDS_PER_CELL = 6  # más palabras en una celda de texto = prosa, no un total
_KV_CELL_RE = re.compile(r"^[^:]{1,40}:\s*\S")  # "Vencimiento: 15/11/2026", "Observaciones: …"


def extend_tail(
    page: Any,
    table: Any,
    rows: list[list[str]],
    words: list[Word] | list[tuple],
    max_gap_pt: float = 20,
) -> int:
    """Agrega a `rows` los renglones de palabras debajo del bbox que caen en las columnas de la tabla.

    Recupera Subtotal / IVA / TOTAL dibujados sin grilla: a partir de ``table.bbox[3]`` toma renglones
    (palabras agrupadas por y con tolerancia 3 pt) con gap < `max_gap_pt`, cuyas palabras caen TODAS dentro
    de algún rango x de columna, que llenan ≥ 2 celdas y donde al menos una celda es un número, porcentaje o
    fecha **en una columna que ya era numérica/fecha en el cuerpo** (así no se traga la prosa que sigue a la
    tabla). Renglones ``clave: valor`` o con celdas de muchas palabras cortan la cola. Devuelve la cantidad
    de filas agregadas. `page` se recibe por contrato (no hace falta para la geometría).
    """
    added, _bottom = extend_tail_with_bottom(page, table, rows, words, max_gap_pt)
    return added


def extend_tail_with_bottom(
    page: Any,
    table: Any,
    rows: list[list[str]],
    words: list[Word] | list[tuple],
    max_gap_pt: float = 20,
) -> tuple[int, float]:
    """Como ``extend_tail`` pero devuelve también el ``y1`` del último renglón agregado (o del bbox).

    Sirve para extender ``RawTable.bbox`` con la cola y que ``text_outside_tables`` no repita esas líneas.
    Garantiza progreso estricto: cada iteración consume palabras (nunca vuelve a mirar las ya usadas), el
    cursor ``y`` avanza al menos 1 pt y hay un tope de ``MAX_TAIL_ROWS`` filas.
    """
    del page  # por contrato; la geometría sale de `table` y `words`
    ranges = _column_x_ranges(table)
    bottom = float(table.bbox[3])
    if not ranges:
        return 0, bottom
    value_columns = _value_columns(rows)
    y = bottom
    # Índices de palabras candidatas (debajo del bbox), ordenados por y0; los consumidos salen de la lista.
    pending = sorted((i for i, w in enumerate(words) if w[1] >= y - 1), key=lambda i: words[i][1])
    added = 0
    while pending and added < MAX_TAIL_ROWS:
        near = [i for i in pending if y - 1 <= words[i][1] < y + max_gap_pt]
        if not near:
            break
        y0 = min(words[i][1] for i in near)
        line_idx = [i for i in pending if abs(words[i][1] - y0) < _TAIL_LINE_TOL_PT]
        line = sorted((words[i] for i in line_idx), key=lambda w: w[0])
        if not all(_in_some_range(_center(w)[0], ranges) for w in line):
            break
        cells = [" ".join(str(w[4]) for w in line if x0 <= _center(w)[0] <= x1) for x0, x1 in ranges]
        if not _tail_row_ok(cells, value_columns):
            break
        rows.append(cells)
        added += 1
        consumed = set(line_idx)
        pending = [i for i in pending if i not in consumed]
        # Avance estricto: palabras de alto < 1 pt (OCR, marcas de agua) no pueden dejar `y` donde estaba.
        y = max(max(w[3] for w in line), y0 + 1.0, y + 1.0)
        bottom = max(bottom, y)
    return added, bottom


def _cell_has_value(text: str) -> bool:
    """¿La celda parsea como número, porcentaje o fecha (en cualquier locale)?"""
    from pdf2excel.core import typing_ar  # import diferido: el resto de cells no depende del tipado

    text = str(text).strip()
    if not text:
        return False
    if typing_ar.parse_date(text) is not None:
        return True
    return any(
        typing_ar.parse_number(text, locale) is not None or typing_ar.parse_percent(text, locale) is not None
        for locale in ("ar", "us")
    )


def _value_columns(rows: Sequence[Sequence[str]]) -> set[int] | None:
    """Columnas cuyo cuerpo (sin encabezado) es mayormente numérico/fecha.

    ``None`` si no hay cuerpo o ninguna columna es numérica: entonces vale cualquier columna (tablas de una
    fila de texto con un TOTAL debajo).
    """
    body = [r for r in rows[1:] if any(str(c).strip() for c in r)]
    if not body:
        return None
    ncols = max(len(r) for r in body)
    out: set[int] = set()
    for c in range(ncols):
        values = [str(r[c]).strip() for r in body if c < len(r) and str(r[c]).strip()]
        if values and sum(1 for v in values if _cell_has_value(v)) >= 0.5 * len(values):
            out.add(c)
    return out or None


def _tail_row_ok(cells: Sequence[str], value_columns: set[int] | None) -> bool:
    """Criterio de cola: ≥ 2 celdas, un valor en una columna numérica, sin ``clave: valor`` ni prosa."""
    filled = [(c, str(text).strip()) for c, text in enumerate(cells) if str(text).strip()]
    if len(filled) < 2:
        return False
    has_value = False
    for c, text in filled:
        if _KV_CELL_RE.match(text):
            return False  # "Vencimiento: 15/11/2026" es texto suelto, no una fila de la tabla
        if _cell_has_value(text):
            if value_columns is None or c in value_columns:
                has_value = True
        elif len(text.split()) > _MAX_TAIL_WORDS_PER_CELL:
            return False  # prosa que cae en las columnas
    return has_value


def _tail_has_value(cells: Sequence[str]) -> bool:
    """Alguna celda parsea como número, porcentaje o fecha (los totales siempre traen un importe)."""
    return any(_cell_has_value(c) for c in cells)


def _pad_rows(rows: list[list[str]]) -> tuple[list[list[str]], int]:
    n = max((len(r) for r in rows), default=0)
    return [list(r) + [""] * (n - len(r)) for r in rows], n


def merge_ghost_columns(rows: list[list[str]]) -> list[list[str]]:
    """Une pares de columnas adyacentes mutuamente excluyentes y elimina columnas totalmente vacías.

    Dos columnas se unen cuando ninguna fila tiene ambas llenas y alguna de las dos tiene ≤ 2 celdas no
    vacías (típico: "Subtotal" en una columna fantasma al lado de la etiqueta real).
    """
    if not rows:
        return rows
    rows, n = _pad_rows(rows)
    c = 0
    while c < n - 1:
        a = [bool(r[c].strip()) for r in rows]
        b = [bool(r[c + 1].strip()) for r in rows]
        exclusive = not any(x and y for x, y in zip(a, b, strict=True))
        if any(a) and any(b) and exclusive and (sum(a) <= 2 or sum(b) <= 2):
            for r in rows:
                r[c] = f"{r[c]} {r[c + 1]}".strip()
                del r[c + 1]
            n -= 1
        else:
            c += 1
    keep = [c for c in range(n) if any(r[c].strip() for r in rows)]
    return [[r[c] for c in keep] for r in rows]


def drop_empty_rows(rows: list[list[str]]) -> list[list[str]]:
    """Descarta las filas sin ninguna celda con contenido."""
    return [list(r) for r in rows if any(str(c).strip() for c in r)]


__all__ = [
    "MAX_TAIL_ROWS",
    "Word",
    "clean_cell",
    "drop_empty_rows",
    "extend_tail",
    "extend_tail_with_bottom",
    "merge_ghost_columns",
    "words_in_rect",
]
