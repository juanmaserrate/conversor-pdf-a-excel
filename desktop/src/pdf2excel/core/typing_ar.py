"""Tipado argentino por columna (DESIGN.md §5.6).

Funciones puras, sin pymupdf ni Qt. La decisión de tipo se toma por **columna** (no por celda): primero
identificadores, después fechas, porcentajes, números (con locale AR/US inferido sobre toda la columna) y,
si nada aplica, texto. Lo que no parsea con el tipo de su columna queda como texto y va a ``review_cells``:
nunca se inventa un número.
"""

from __future__ import annotations

import re
from datetime import date

from pdf2excel.models import CellValue, ColumnInfo, ColumnKind, Locale, RawTable, TypedTable

# --------------------------------------------------------------------------- regexes públicas

# Número ya normalizado (sin signo, moneda ni espacios; ver normalize_number_text).
NUM_AR = re.compile(r"^(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?$")
NUM_US = re.compile(r"^(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")
PCT = re.compile(r"^[-+]?\d+(?:[.,]\d+)?\s*%$")
DATE_RE = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2}|\d{4})$")
CUIT_RE = re.compile(r"^\d{2}-\d{8}-\d$")
ID_HEADER_RE = re.compile(
    # Todas las alternativas van ancladas a límites de palabra: "Importe facturado" no es una factura,
    # "Circuito" no es un CUIT y "Turno"/"Vino"/"Destino" no son "N°". DESIGN escribe `n[°ºo]\b`, pero `\b`
    # tras "°" no matchea "N°" a fin de texto: el caso "N°/Nº" se ancla con (?<!\w) y el "no" con \b.
    r"\b(?:comprobante|c[oó]digo|cuit|cuil|cbu|dni|nro|referencia|id|art[ií]culos?|sku|facturas?"
    r"|remitos?|legajo)\b"
    r"|(?<!\w)n[°º]"
    r"|\bno\b",
    re.IGNORECASE,
)
# Celdas "sin dato": guiones, rayas o puntos solos ("-", "—", "–", "..."). Se tratan como vacías.
BLANK_CELL_RE = re.compile(r"^[-–—\.\s]+$")

# Internas
_CURRENCY_RE = re.compile(r"U\$S|USD|ARS|\$", re.IGNORECASE)
_DIGITS_ID_RE = re.compile(r"^\d{5,}$")
_TOTAL_ROW_RE = re.compile(r"^(?:(?:sub)?total(?:es)?|iva|transporte|saldo)\b", re.IGNORECASE)

DATE_FORMAT = "dd/mm/yyyy"
PERCENT_FORMAT = "0.00%"
INT_FORMAT = "#,##0"
FLOAT_FORMAT = "#,##0.00"
CURRENCY_FORMAT = '"$" #,##0.00'


# --------------------------------------------------------------------------- normalización


def normalize_number_text(s: str) -> tuple[str, bool, bool]:
    """Limpia un texto numérico: quita moneda, espacios y signo.

    Devuelve ``(texto_limpio_sin_signo, es_negativo, tiene_moneda)``. El negativo se detecta por ``-``
    inicial, ``-`` final (``"1.234,56-"``, extractos) o paréntesis (``"(1.234,56)"``). Se aceptan
    ``$``, ``ARS``, ``U$S`` y ``USD`` como símbolos de moneda.
    """
    text = s.replace(" ", " ").replace("−", "-").strip()
    has_currency = bool(_CURRENCY_RE.search(text))
    if has_currency:
        text = _CURRENCY_RE.sub("", text)
    text = re.sub(r"\s+", "", text)
    negative = False
    if len(text) >= 2 and text[0] == "(" and text[-1] == ")":
        negative = True
        text = text[1:-1]
    if text.startswith("-"):
        negative = True
        text = text[1:]
    elif text.startswith("+"):
        text = text[1:]
    if text.endswith("-"):
        negative = True
        text = text[:-1]
    return text, negative, has_currency


def _strip_percent(s: str) -> str:
    text = s.strip()
    return text[:-1].rstrip() if text.endswith("%") else text


# --------------------------------------------------------------------------- locale


def infer_column_locale(values: list[str]) -> Locale | None:
    """Infiere ``"ar"`` o ``"us"`` mirando toda la columna.

    Evidencia AR: termina en ``,dd`` o usa ``.`` de miles junto con ``,`` (solo ``NUM_AR`` lo explica).
    Evidencia US: termina en ``.dd`` o usa ``,`` de miles junto con ``.`` (solo ``NUM_US`` lo explica).
    Los enteros sin separador son candidatos pero no votan; los ambiguos (``"1.234"``, ``"1,234"`` solos,
    que parsean en ambos locales) no votan ni cuentan como candidatos. Si no hay candidatos pero sí algún
    ambiguo → ``"ar"`` (sin evidencia → AR, DESIGN §5.6/§10: una tabla de una sola fila con "1.500" se tipa
    igual que una de dos). Sin candidatos ni ambiguos → ``None``; empate → ``"ar"``.
    """
    ar_votes = us_votes = candidates = ambiguous = 0
    for raw in values:
        if not raw or not raw.strip():
            continue
        clean, _neg, _cur = normalize_number_text(_strip_percent(raw))
        if not clean or DATE_RE.match(clean):
            continue
        is_ar = bool(NUM_AR.match(clean))
        is_us = bool(NUM_US.match(clean))
        if not (is_ar or is_us):
            continue
        if is_ar != is_us:
            # Solo un locale lo explica: "1.234,56" / "1.234.567" → AR; "350.00" / "1,234.56" → US.
            candidates += 1
            if is_ar:
                ar_votes += 1
            else:
                us_votes += 1
        elif "." in clean or "," in clean:
            ambiguous += 1  # "1.234" / "1,234" solos: no votan
        else:
            candidates += 1  # entero sin separador: parsea en ambos locales, no vota
    if candidates == 0:
        return "ar" if ambiguous >= 1 else None
    if us_votes > ar_votes:
        return "us"
    return "ar"


# --------------------------------------------------------------------------- parsers


def parse_number(s: str, locale: Locale) -> int | float | None:
    """Parsea un número en el locale dado; ``int`` si no hay parte decimal explícita.

    AR: ``.`` miles y ``,`` decimal. US: al revés. Enteros sin separador parsean en ambos locales.
    ``"1.234"`` en AR → 1234; en US → 1.234. Lo que no cumple el formato → ``None``.
    """
    if s is None:
        return None
    clean, negative, _cur = normalize_number_text(s)
    if not clean:
        return None
    if locale == "ar":
        if not NUM_AR.match(clean):
            return None
        has_decimal = "," in clean
        plain = clean.replace(".", "").replace(",", ".")
    else:
        if not NUM_US.match(clean):
            return None
        has_decimal = "." in clean
        plain = clean.replace(",", "")
    try:
        value: int | float = float(plain) if has_decimal else int(plain)
    except ValueError:
        return None
    return -value if negative else value


def parse_percent(s: str, locale: Locale) -> float | None:
    """``"21%"`` → 0.21; ``"12,5 %"`` → 0.125. Siempre ``float``."""
    if s is None:
        return None
    text = s.strip()
    if not PCT.match(text):
        return None
    number = parse_number(_strip_percent(text), locale)
    if number is None:
        return None
    return number / 100


def parse_date(s: str) -> date | None:
    """``dd/mm/aa(aa)`` con ``/``, ``-`` o ``.``. Año de 2 dígitos: 00–69 → 20xx, 70–99 → 19xx."""
    if s is None:
        return None
    m = DATE_RE.match(s.strip())
    if not m:
        return None
    day, month, year = int(m[1]), int(m[2]), int(m[3])
    if len(m[3]) == 2:
        year += 2000 if year <= 69 else 1900
    try:
        return date(year, month, day)
    except ValueError:
        return None


def is_identifier_header(header: str) -> bool:
    """¿El encabezado nombra un identificador (comprobante, código, CUIT, N°, SKU, …)?"""
    return bool(header) and bool(ID_HEADER_RE.search(header))


# --------------------------------------------------------------------------- clasificación


def is_blank_cell(text: str | None) -> bool:
    """Vacía o "sin dato" (``-``, ``—``, ``–``, ``...``): no cuenta para el tipado ni se escribe."""
    return not text or not text.strip() or bool(BLANK_CELL_RE.match(text.strip()))


def _nonempty(values: list[str]) -> list[str]:
    return [v.strip() for v in values if not is_blank_cell(v)]


def _looks_like_id_column(header: str, nonempty: list[str]) -> bool:
    if is_identifier_header(header):
        return True
    if not nonempty:
        return False
    if all(_DIGITS_ID_RE.match(v) for v in nonempty) and len({len(v) for v in nonempty}) == 1:
        return True
    cuits = sum(1 for v in nonempty if CUIT_RE.match(v))
    return cuits >= 0.8 * len(nonempty)


def classify_column(header: str, values: list[str]) -> ColumnInfo:
    """Decide el tipo de una columna a partir de su encabezado y sus celdas (cuerpo, sin encabezado).

    Orden exacto de DESIGN.md §5.6: id → date (≥ 90 %) → percent (≥ 50 %) → number (≥ 50 % con el locale
    inferido) → text.
    """
    header = (header or "").strip()
    nonempty = _nonempty(values)
    if _looks_like_id_column(header, nonempty):
        return ColumnInfo(kind="id", header=header)
    if not nonempty:
        return ColumnInfo(kind="text", header=header)
    n = len(nonempty)

    dates = sum(1 for v in nonempty if parse_date(v) is not None)
    if dates >= 0.9 * n:
        return ColumnInfo(kind="date", number_format=DATE_FORMAT, header=header)

    pcts = sum(1 for v in nonempty if PCT.match(v))
    if pcts >= 0.5 * n:
        locale: Locale = infer_column_locale(nonempty) or "ar"
        return ColumnInfo(kind="percent", locale=locale, number_format=PERCENT_FORMAT, header=header)

    inferred = infer_column_locale(nonempty)
    if inferred is not None:
        parsed = [parse_number(v, inferred) for v in nonempty]
        ok = [(v, p) for v, p in zip(nonempty, parsed, strict=True) if p is not None]
        if len(ok) >= 0.5 * n:
            with_currency = sum(1 for v, _p in ok if normalize_number_text(v)[2])
            if with_currency >= 0.5 * len(ok):
                fmt = CURRENCY_FORMAT
            elif all(isinstance(p, int) for _v, p in ok):
                fmt = INT_FORMAT
            else:
                fmt = FLOAT_FORMAT
            return ColumnInfo(kind="number", locale=inferred, number_format=fmt, header=header)

    return ColumnInfo(kind="text", header=header)


def _parse_cell(text: str, col: ColumnInfo) -> CellValue:
    """Parsea una celda según el tipo de su columna; ``None`` si no parsea."""
    kind: ColumnKind = col.kind
    if kind == "number":
        return parse_number(text, col.locale or "ar")
    if kind == "date":
        return parse_date(text)
    if kind == "percent":
        return parse_percent(text, col.locale or "ar")
    return text


def _is_total_row(row: list[str]) -> bool:
    for cell in row:
        if cell and cell.strip():
            return bool(_TOTAL_ROW_RE.match(cell.strip()))
    return False


def type_table(raw: RawTable, typed_values: bool = True) -> TypedTable:
    """Convierte una ``RawTable`` (texto) en una ``TypedTable`` con tipos reales por columna.

    ``rows[0]`` es el encabezado y el resto el cuerpo; se igualan las longitudes. Cada celda se parsea con
    el tipo de su columna; vacía (o un guión "sin dato" en una columna tipada) → ``None``; lo que no parsea
    en una columna number/date/percent queda ``str`` y entra en ``review_cells`` (índices sobre ``rows``
    sin encabezado). Con ``typed_values=False`` todo queda texto, las columnas son ``kind="text"`` y no hay
    revisión.
    """
    src_rows = [[("" if c is None else str(c)) for c in r] for r in raw.rows]
    ncols = max((len(r) for r in src_rows), default=0)
    padded = [r + [""] * (ncols - len(r)) for r in src_rows]
    header = [h.strip() for h in padded[0]] if padded else []
    body = padded[1:]

    if typed_values:
        columns = [classify_column(header[c], [r[c] for r in body]) for c in range(ncols)]
    else:
        columns = [ColumnInfo(kind="text", header=header[c]) for c in range(ncols)]

    rows: list[list[CellValue]] = []
    review: set[tuple[int, int]] = set()
    total_rows: set[int] = set()
    for i, src in enumerate(body):
        out: list[CellValue] = []
        for c, text in enumerate(src):
            text = text.strip()
            if not text:
                out.append(None)
                continue
            if not typed_values or columns[c].kind in ("id", "text"):
                out.append(text)
                continue
            if is_blank_cell(text):
                out.append(None)  # "-" / "—" en una columna de importes = sin dato, no para revisar
                continue
            value = _parse_cell(text, columns[c])
            if value is None:
                out.append(text)
                review.add((i, c))
            else:
                out.append(value)
        rows.append(out)
        if _is_total_row(src):
            total_rows.add(i)

    return TypedTable(
        name=raw.title or "",
        header=header,
        rows=rows,
        columns=columns,
        review_cells=review,
        total_rows=total_rows,
        pages=(raw.page, raw.page_end),
        source=raw.source,
        quality=raw.quality,
    )


__all__ = [
    "BLANK_CELL_RE",
    "CUIT_RE",
    "DATE_RE",
    "ID_HEADER_RE",
    "NUM_AR",
    "NUM_US",
    "PCT",
    "classify_column",
    "infer_column_locale",
    "is_blank_cell",
    "is_identifier_header",
    "normalize_number_text",
    "parse_date",
    "parse_number",
    "parse_percent",
    "type_table",
]
