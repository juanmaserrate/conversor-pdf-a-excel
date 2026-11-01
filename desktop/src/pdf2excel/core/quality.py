"""Puntaje de calidad de tabla sin ground truth y selección entre candidatos (DESIGN.md §5.4)."""

from __future__ import annotations

import re
from collections.abc import Callable
from datetime import date
from functools import lru_cache
from typing import Any

from pdf2excel.models import RawTable

_BARS_RE = re.compile(r"[|¦!.\-]*")
_NUMERIC_HEADER_RE = re.compile(r"^[\d$.,\- ]+$")
_LONG_CELL_CHARS = 60

# Pesos de la fórmula (DESIGN.md §5.4)
W_EMPTY, W_HEADER, W_TYPED, W_BARS, W_LONG, W_BROKEN, W_GHOST = 0.05, 0.15, 0.35, 0.15, 0.10, 0.20, 0.15

LocaleFn = Callable[[list[str]], Any]
NumberFn = Callable[[str, Any], int | float | None]
DateFn = Callable[[str], date | None]

# --------------------------------------------------------------------------- fallback de tipado
# Mínimo local con las mismas firmas que core/typing_ar.py, usado solo si ese módulo no existe todavía.

_NUM_AR = re.compile(r"^[-+]?\$?\s*(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?$|^[-+]?\$?\s*\d+,\d+$")
_NUM_US = re.compile(r"^[-+]?\$?\s*(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$|^[-+]?\$?\s*\d+\.\d+$")
_PCT = re.compile(r"^[-+]?\d+(?:[.,]\d+)?\s*%$")
_DATE = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2}|\d{4})$")


def _fallback_infer_column_locale(values: list[str]) -> str | None:
    ar = us = candidates = 0
    for raw in values:
        v = raw.strip()
        if not v or _DATE.match(v):
            continue
        if not (_NUM_AR.match(v) or _NUM_US.match(v)):
            continue
        candidates += 1
        if re.search(r",\d{1,2}$", v) or (re.search(r"\.\d{3}(?:[.,]|$)", v) and "," in v):
            ar += 1
        if (re.search(r"\.\d{1,2}$", v) and not re.search(r"\.\d{3}$", v)) or (
            re.search(r",\d{3}(?:[.,]|$)", v) and "." in v
        ):
            us += 1
    if candidates == 0:
        return None
    return "ar" if ar >= us else "us"


def _fallback_parse_number(s: str, locale: Any) -> int | float | None:
    v = s.strip()
    if not v:
        return None
    if _PCT.match(v):
        n = _fallback_parse_number(v.rstrip("% "), locale)
        return None if n is None else n / 100
    compact = v.replace(" ", "")
    if locale == "ar":
        if not _NUM_AR.match(compact):
            return None
        digits = compact.replace("$", "").replace(".", "").replace(",", ".")
    else:
        if not _NUM_US.match(compact):
            return None
        digits = compact.replace("$", "").replace(",", "")
    try:
        f = float(digits)
    except ValueError:
        return None
    return int(f) if f.is_integer() and "." not in digits else f


def _fallback_parse_date(s: str) -> date | None:
    m = _DATE.match(s.strip())
    if not m:
        return None
    d, mo, y = int(m[1]), int(m[2]), int(m[3])
    if y < 100:
        y += 2000 if y < 70 else 1900
    try:
        return date(y, mo, d)
    except ValueError:
        return None


@lru_cache(maxsize=1)
def _typing_backend() -> tuple[LocaleFn, NumberFn, DateFn]:
    """Devuelve (infer_column_locale, parse_number, parse_date): typing_ar si existe, si no el fallback."""
    try:
        from pdf2excel.core.typing_ar import infer_column_locale, parse_date, parse_number
    except ImportError:
        return _fallback_infer_column_locale, _fallback_parse_number, _fallback_parse_date
    return infer_column_locale, parse_number, parse_date


# --------------------------------------------------------------------------- puntaje


def _pad(rows: list[list[str]]) -> tuple[list[list[str]], int]:
    n = max((len(r) for r in rows), default=0)
    return [[str(c) if c is not None else "" for c in r] + [""] * (n - len(r)) for r in rows], n


def _has_content(cell: str) -> bool:
    """Celda "llena": algo que no sea solo barras/puntos/guiones (bordes alucinados por OCR no cuentan)."""
    c = cell.strip()
    return bool(c) and not _BARS_RE.fullmatch(c)


def _typed_score(body: list[list[str]], n: int) -> float:
    """Fracción de columnas tipables ponderada por su tasa de parseo; 0.5 si no hay columnas candidatas."""
    infer_column_locale, parse_number, parse_date = _typing_backend()
    typed = 0.0
    n_typed = 0
    for c in range(n):
        col = [r[c] for r in body if r[c].strip()]
        if not col:
            continue
        if all(parse_date(v) is not None for v in col):
            typed += 1.0
            n_typed += 1
            continue
        loc = infer_column_locale(col)
        if loc:
            ok = sum(1 for v in col if parse_number(v, loc) is not None) / len(col)
            if ok >= 0.5:
                typed += ok
                n_typed += 1
    return typed / n_typed if n_typed else 0.5


def ghost_column_pairs(rows: list[list[str]]) -> int:
    """Pares de columnas adyacentes mutuamente excluyentes (ninguna fila con ambas llenas)."""
    rows, n = _pad(rows)
    ghost = 0
    for c in range(n - 1):
        a = [bool(r[c].strip()) for r in rows]
        b = [bool(r[c + 1].strip()) for r in rows]
        if any(a) and any(b) and not any(x and y for x, y in zip(a, b, strict=True)):
            ghost += 1
    return ghost


def table_quality(rows: list[list[str]]) -> float:
    """Puntaje ∈ [0, 1] sin ground truth sobre el cuerpo (sin encabezado).

    ``q = 0.05·(1−vacías) + 0.15·encabezado_ok + 0.35·tipado + 0.15·(1−barras) + 0.10·(1−largas)
    + 0.20·(1−rotas) − 0.15·fantasma`` (DESIGN.md §5.4). Para `vacías` y `rotas` una celda que es solo
    barras (``|¦!.-``) no cuenta como llena: una fila con un valor y un ``|`` está rota.
    """
    if len(rows) < 2:
        return 0.0
    padded, n = _pad(rows)
    if n == 0:
        return 0.0
    body = padded[1:]
    cells = [c for r in body for c in r]
    if not cells:
        return 0.0

    empty = sum(1 for c in cells if not _has_content(c)) / len(cells)
    bars = sum(1 for c in cells if c.strip() and _BARS_RE.fullmatch(c.strip())) / len(cells)
    header_ok = sum(1 for c in padded[0] if c.strip() and not _NUMERIC_HEADER_RE.match(c)) / n
    typed = _typed_score(body, n)
    long_cells = sum(1 for c in cells if len(c) > _LONG_CELL_CHARS) / len(cells)
    broken = sum(1 for r in body if sum(1 for c in r if _has_content(c)) <= 1) / max(1, len(body))
    ghost = ghost_column_pairs(padded)

    q = (
        W_EMPTY * (1 - empty)
        + W_HEADER * header_ok
        + W_TYPED * typed
        + W_BARS * (1 - bars)
        + W_LONG * (1 - long_cells)
        + W_BROKEN * (1 - broken)
        - W_GHOST * ghost
    )
    return max(0.0, min(1.0, q))


# --------------------------------------------------------------------------- selección


def _filled_cells(tables: list[RawTable]) -> int:
    return sum(1 for t in tables for r in t.rows for c in r if str(c).strip())


def _weighted_quality(tables: list[RawTable]) -> float:
    n_rows = sum(len(t.rows) for t in tables)
    if n_rows == 0:
        return 0.0
    return sum(t.quality * len(t.rows) for t in tables) / n_rows


def select_best(cands: dict[str, list[RawTable]]) -> tuple[str, list[RawTable]]:
    """Elige el candidato con mayor ``0.7·calidad + 0.3·cobertura``.

    calidad = media de `quality` ponderada por filas; cobertura = celdas no vacías / máximo entre candidatos;
    candidato sin tablas = −1; empate → orden de inserción del dict ("hybrid_union" primero).
    Devuelve ``("hybrid", [])`` si ningún candidato tiene tablas.
    """
    if not any(cands.values()):
        return "hybrid", []
    max_cells = max(1, max(_filled_cells(v) for v in cands.values()))
    best_key: str | None = None
    best_score = float("-inf")
    for name, tables in cands.items():
        score = 0.7 * _weighted_quality(tables) + 0.3 * _filled_cells(tables) / max_cells if tables else -1.0
        if score > best_score:  # estricto: a igualdad gana el primero insertado
            best_key, best_score = name, score
    assert best_key is not None
    return best_key, list(cands[best_key])


__all__ = ["ghost_column_pairs", "select_best", "table_quality"]
