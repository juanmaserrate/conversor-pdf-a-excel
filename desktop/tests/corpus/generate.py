"""Corpus sintético determinista para medir la calidad de extracción (DESIGN.md §9).

Genera todos los PDFs de prueba (digitales, escaneados, cifrado, vacío, mixto, texto basura) con reportlab,
pymupdf y Pillow a partir de una semilla fija: dos corridas producen bytes idénticos (reportlab
``invariant=1``; la única excepción es ``con_password.pdf`` porque el cifrado AES-256 usa sales aleatorias).
Los PDFs **no** se commitean: CI corre ``python -m tests.corpus.generate tests/corpus/_generated``.

El ground truth de cada caso sale de los mismos datos que se dibujan (``ground_truth()``) y vive, ya
tipado, en ``tests/corpus/ground_truth.json`` (``--ground-truth`` lo reescribe).
"""

from __future__ import annotations

import argparse
import io
import json
import random
import sys
import tempfile
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import pymupdf
from PIL import Image, ImageFilter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

DEFAULT_SEED = 7
PASSWORD = "1234"  # contraseña de con_password.pdf
SCAN_SEED = 3  # ruido de make_scan (independiente de la semilla de datos, como en make_scans.py)
_PDF_ENCRYPT_AES_256 = int(getattr(pymupdf, "PDF_ENCRYPT_AES_256", 4))  # los stubs de pymupdf no lo declaran

GtCell = dict[str, Any] | None
GtTable = dict[str, Any]
GroundTruth = dict[str, dict[str, Any]]

_STYLES = getSampleStyleSheet()
_GRID_STYLE = ("FONTSIZE", (0, 0), (-1, -1), 8)


# ------------------------------------------------------------------------------------------- helpers GT


def gt_float(value: float) -> GtCell:
    return {"v": round(float(value), 2), "t": "float"}


def gt_int(value: int) -> GtCell:
    return {"v": int(value), "t": "int"}


def gt_str(value: str) -> GtCell:
    return {"v": str(value), "t": "str"}


def gt_date(value: date | str) -> GtCell:
    """``value`` como ``date`` o texto ``dd/mm/aaaa`` / ``dd/mm/aa``."""
    if isinstance(value, str):
        day, month, year = (int(p) for p in value.split("/"))
        if year < 100:
            year += 2000
        value = date(year, month, day)
    return {"v": value.isoformat(), "t": "date"}


def gt_table(header: Sequence[str], rows: Sequence[Sequence[GtCell]]) -> GtTable:
    return {"header": list(header), "rows": [list(r) for r in rows]}


def gt_entry(
    tables: Sequence[GtTable], kinds: Sequence[str], notes: str = "", **extra: Any
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "tables": list(tables),
        "expected_tables": len(tables),
        "kinds": list(kinds),
        "notes": notes,
    }
    entry.update(extra)
    return entry


def ar(n: float, money: bool = False) -> str:
    """``1234.5`` → ``"1.234,50"`` (``"$ 1.234,50"`` con ``money``)."""
    s = f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("$ " + s) if money else s


def us(n: float, money: bool = False) -> str:
    s = f"{n:,.2f}"
    return ("$ " + s) if money else s


def _fecha(day: int, month: int, year: int = 2026, short: bool = False) -> str:
    return f"{day:02d}/{month:02d}/{year % 100:02d}" if short else f"{day:02d}/{month:02d}/{year}"


# ------------------------------------------------------------------------------------------- reportlab


def _doc(path: Path, pagesize: tuple[float, float] = A4, **margins: float) -> SimpleDocTemplate:
    kwargs: dict[str, Any] = {
        "leftMargin": 15 * mm,
        "rightMargin": 15 * mm,
        "topMargin": 20 * mm,
        "bottomMargin": 20 * mm,
    }
    kwargs.update(margins)
    return SimpleDocTemplate(str(path), pagesize=pagesize, invariant=1, **kwargs)


def _canvas(path: Path | io.BytesIO, pagesize: tuple[float, float] = A4) -> canvas.Canvas:
    target = str(path) if isinstance(path, Path) else path
    return canvas.Canvas(target, pagesize=pagesize, invariant=1)


def _grid_table(data: list[list[Any]], col_widths: list[float], **kw: Any) -> Table:
    table = Table(data, colWidths=col_widths, **kw)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDDDDD")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                _GRID_STYLE,
            ]
        )
    )
    return table


def _page_count(path: Path) -> int:
    with pymupdf.open(path) as doc:
        return int(doc.page_count)


# ------------------------------------------------------------------------------------------- casos


@dataclass
class Case:
    """Un PDF del corpus: cómo se genera y su ground truth (``None`` para los escaneos, que heredan)."""

    name: str
    build: Callable[[Path, random.Random], dict[str, Any] | None]
    depends_on: tuple[str, ...] = ()
    files: list[Path] = field(default_factory=list)


REMITO_ROWS: list[tuple[str, int, float]] = [
    ("Mezcal joven", 10, 350.0),
    ("Tequila reposado", 5, 420.5),
    ("Agua mineral", 48, 12.9),
    ("Hielo bolsa 5kg", 20, 35.0),
    ("Limón kg", 8, 28.75),
    ("Sal de gusano", 3, 90.0),
]
REMITO_HEADER = ["Código", "Producto", "Cantidad", "Precio unitario", "Total"]


def _remito_data() -> tuple[list[list[str]], list[list[GtCell]]]:
    data = [REMITO_HEADER]
    gt: list[list[GtCell]] = []
    for i, (product, qty, price) in enumerate(REMITO_ROWS):
        code = f"A-{100 + i}"
        data.append([code, product, str(qty), f"{price:.2f}", f"{qty * price:.2f}"])
        gt.append([gt_str(code), gt_str(product), gt_int(qty), gt_float(price), gt_float(qty * price)])
    return data, gt


def _remito_story(title: str, grid: bool) -> list[Any]:
    data, _ = _remito_data()
    table = Table(data, colWidths=[60, 140, 70, 100, 80])
    style: list[Any] = [("FONTSIZE", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]
    if grid:
        style.append(("GRID", (0, 0), (-1, -1), 0.8, colors.black))
    table.setStyle(TableStyle(style))
    return [
        Paragraph(title, _STYLES["Title"]),
        Paragraph("Cliente: Bar El Catorce — Fecha: 15/10/2026", _STYLES["Normal"]),
        Spacer(1, 12),
        table,
    ]


def _build_remito(path: Path, grid: bool, title: str = "Remito N° 4521") -> dict[str, Any]:
    SimpleDocTemplate(str(path), pagesize=A4, invariant=1).build(_remito_story(title, grid))
    _, gt = _remito_data()
    return gt_entry(
        [gt_table(REMITO_HEADER, gt)],
        ["digital"],
        "Remito de una página, 7×5, importes en formato US (350.00)"
        + (" con grilla" if grid else " sin bordes"),
    )


def build_remito_grid(path: Path, _rng: random.Random) -> dict[str, Any]:
    return _build_remito(path, grid=True)


def build_remito_nogrid(path: Path, _rng: random.Random) -> dict[str, Any]:
    return _build_remito(path, grid=False)


def build_prueba(path: Path, _rng: random.Random) -> dict[str, Any]:
    data = [["Producto", "Cantidad", "Precio"], ["Mezcal", "10", "350"], ["Tequila", "5", "420"]]
    table = Table(data)
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.8, colors.black)]))
    SimpleDocTemplate(str(path), pagesize=A4, invariant=1).build(
        [Paragraph("Reporte de stock", _STYLES["Title"]), table]
    )
    rows = [
        [gt_str("Mezcal"), gt_int(10), gt_int(350)],
        [gt_str("Tequila"), gt_int(5), gt_int(420)],
    ]
    return gt_entry([gt_table(data[0], rows)], ["digital"], "Tabla mínima 3×3 con enteros")


# --- hard_factura_ar ---------------------------------------------------------------------------

_FACTURA_PRODUCTS = [
    "Mezcal joven 750 ml",
    "Tequila reposado añejo premium edición limitada 1 L",
    "Agua mineral sin gas x 6",
    "Hielo bolsa 5kg",
    "Limón kg",
    "Sal de gusano 100 g",
    "Vermut rosso 950 ml",
    "Gin london dry 700 ml",
    "Tónica x 24 latas",
    "Jugo de naranja 1 L",
]
_FACTURA_PRICES = [350.0, 12.9, 1250.5, 28.75, 90.0, 15999.99, 8.5]
FACTURA_HEADER = ["Código", "Descripción", "Cant.", "Precio unit.", "Importe"]


def _factura_rows(rng: random.Random, n: int = 60) -> list[tuple[str, str, int, float, float]]:
    rows = []
    for i in range(n):
        qty = rng.randint(1, 48)
        unit = rng.choice(_FACTURA_PRICES)
        rows.append(
            (f"A-{100 + i}", _FACTURA_PRODUCTS[i % len(_FACTURA_PRODUCTS)], qty, unit, round(qty * unit, 2))
        )
    return rows


def build_hard_factura_ar(path: Path, rng: random.Random) -> dict[str, Any]:
    rows = _factura_rows(rng)
    subtotal = round(sum(r[4] for r in rows), 2)
    iva = round(subtotal * 0.21, 2)
    total = round(subtotal + iva, 2)

    def on_page(c: canvas.Canvas, d: Any) -> None:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(15 * mm, 285 * mm, "DISTRIBUIDORA DEL SUR S.A.  —  FACTURA A  N° 0003-00012345")
        c.setFont("Helvetica", 9)
        c.drawString(
            15 * mm, 280 * mm, "CUIT 30-71234567-8   Fecha de emisión: 03/11/2026   Vencimiento: 18/11/2026"
        )
        c.drawRightString(195 * mm, 10 * mm, f"Página {d.page}")

    data: list[list[Any]] = [FACTURA_HEADER]
    data += [
        [r[0], Paragraph(r[1], _STYLES["Normal"]), str(r[2]), ar(r[3], True), ar(r[4], True)] for r in rows
    ]
    data += [
        ["", "", "", "Subtotal", ar(subtotal, True)],
        ["", "", "", "IVA 21%", ar(iva, True)],
        ["", "", "", "TOTAL", ar(total, True)],
    ]
    n = len(rows)
    table = Table(data, colWidths=[22 * mm, 78 * mm, 15 * mm, 30 * mm, 35 * mm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, n), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDDDDD")),
                ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                _GRID_STYLE,
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("SPAN", (0, n + 1), (2, n + 1)),
                ("SPAN", (0, n + 2), (2, n + 2)),
                ("SPAN", (0, n + 3), (2, n + 3)),
                ("LINEABOVE", (3, n + 1), (-1, -1), 0.5, colors.black),
                ("FONTNAME", (3, n + 3), (-1, -1), "Helvetica-Bold"),
            ]
        )
    )
    story = [
        Paragraph(
            "Cliente: Bar El Catorce — Av. Corrientes 1234, CABA — Cond. IVA: Responsable Inscripto",
            _STYLES["Normal"],
        ),
        Spacer(1, 4 * mm),
        table,
        Spacer(1, 6 * mm),
        Paragraph(
            "Observaciones: mercadería entregada en depósito. CAE 76123456789012 — Vto. CAE 13/11/2026",
            _STYLES["Normal"],
        ),
    ]
    _doc(path, topMargin=25 * mm).build(story, onFirstPage=on_page, onLaterPages=on_page)

    gt_rows: list[list[GtCell]] = [
        [gt_str(r[0]), gt_str(r[1]), gt_int(r[2]), gt_float(r[3]), gt_float(r[4])] for r in rows
    ]
    gt_rows += [
        [None, None, None, gt_str("Subtotal"), gt_float(subtotal)],
        [None, None, None, gt_str("IVA 21%"), gt_float(iva)],
        [None, None, None, gt_str("TOTAL"), gt_float(total)],
    ]
    return gt_entry(
        [gt_table(FACTURA_HEADER, gt_rows)],
        ["digital", "digital"],
        "Factura A de 2 páginas: repeatRows, descripciones con wrap, importes '$ 1.234,56', "
        "Subtotal/IVA/TOTAL con SPAN y sin grilla, encabezado y pie de página",
        dates=["03/11/2026", "18/11/2026", "13/11/2026"],
    )


# --- hard_extracto_banco -----------------------------------------------------------------------

_CONCEPTOS = [
    "Transferencia recibida MP",
    "Débito automático EDESUR",
    "Compra Débito COTO 0432",
    "Pago servicios ARBA",
    "Acreditación haberes",
    "Comisión mantenimiento cta",
    "Extracción ATM Red Link",
    "Transf. enviada CBU 285...9981",
    "Impuesto Ley 25.413",
    "Cobro cheque 00412",
]
EXTRACTO_HEADER = ["Fecha", "Concepto", "Comprobante", "Débito", "Crédito", "Saldo"]


def build_hard_extracto_banco(path: Path, rng: random.Random) -> dict[str, Any]:
    c = _canvas(path)
    width, height = A4
    rows: list[tuple[str, str, str, float | None, float | None, float]] = []
    saldo = 125430.17
    day = 1
    for i in range(70):
        monto = round(rng.choice([-1, -1, 1]) * rng.uniform(100, 85000), 2)
        saldo = round(saldo + monto, 2)
        fecha = _fecha((day % 28) + 1, 10 if i < 35 else 11, short=True)
        day += rng.randint(0, 2)
        rows.append(
            (
                fecha,
                _CONCEPTOS[i % len(_CONCEPTOS)],
                f"{rng.randint(100000, 999999)}",
                monto if monto < 0 else None,
                monto if monto > 0 else None,
                saldo,
            )
        )
    cols = [
        (20 * mm, "Fecha", "L"),
        (45 * mm, "Concepto", "L"),
        (120 * mm, "Comprobante", "R"),
        (150 * mm, "Débito", "R"),
        (175 * mm, "Crédito", "R"),
        (200 * mm, "Saldo", "R"),
    ]

    def draw(x: float, y: float, text: str, align: str) -> None:
        (c.drawRightString if align == "R" else c.drawString)(x, y, text)

    def head(page: int) -> None:
        c.setFont("Helvetica-Bold", 13)
        c.drawString(20 * mm, height - 20 * mm, "BANCO RÍO DE LA PLATA — Resumen de cuenta corriente")
        c.setFont("Helvetica", 8)
        c.drawString(
            20 * mm,
            height - 26 * mm,
            "Cuenta 0012-345678/9   CBU 0170001200000034567890   Período: 01/10/2026 al 30/11/2026",
        )
        c.setFont("Helvetica-Bold", 8)
        for x, name, align in cols:
            draw(x, height - 36 * mm, name, align)
        c.line(20 * mm, height - 37.5 * mm, 200 * mm, height - 37.5 * mm)
        c.setFont("Helvetica", 7)
        c.drawCentredString(
            width / 2, 10 * mm, f"Hoja {page} — Este resumen no es válido como comprobante fiscal"
        )

    y = height - 42 * mm
    page = 1
    head(page)
    for r in rows:
        if y < 25 * mm:
            c.showPage()
            page += 1
            head(page)
            y = height - 42 * mm
        c.setFont("Helvetica", 8)
        vals = [
            r[0],
            r[1],
            r[2],
            ar(r[3]) if r[3] is not None else "",
            ar(r[4]) if r[4] is not None else "",
            ar(r[5]),
        ]
        for (x, _, align), v in zip(cols, vals, strict=True):
            draw(x, y, v, align)
        y -= 4.6 * mm
    c.setFont("Helvetica-Bold", 8)
    c.drawString(20 * mm, y - 6 * mm, "Saldo final:")
    c.drawRightString(200 * mm, y - 6 * mm, ar(saldo))
    c.save()

    gt_rows: list[list[GtCell]] = [
        [
            gt_date(r[0]),
            gt_str(r[1]),
            gt_str(r[2]),
            gt_float(r[3]) if r[3] is not None else None,
            gt_float(r[4]) if r[4] is not None else None,
            gt_float(r[5]),
        ]
        for r in rows
    ]
    gt_rows.append([gt_str("Saldo final:"), None, None, None, None, gt_float(saldo)])
    return gt_entry(
        [gt_table(EXTRACTO_HEADER, gt_rows)],
        ["digital"] * page,
        "Extracto bancario sin bordes, 70 movimientos en 2 páginas, negativos, Débito/Crédito mutuamente "
        "vacías, fechas dd/mm/aa y 'Saldo final:' suelto al pie (fila de total)",
        saldo_final=saldo,
    )


# --- hard_lista_mixta --------------------------------------------------------------------------

LISTA1_HEADER = ["Código", "Producto", "Presentación", "Precio lista", "Precio mayorista", "Stock"]
LISTA2_HEADER = ["Código", "Producto", "Unidad", "Precio", "Última actualización"]
_LISTA1: list[tuple[str, str, str, float, float | None, int | None]] = [
    ("B-01", "Mezcal joven", "750 ml", 18500.0, 16200.5, 24),
    ("B-02", "Tequila reposado", "1 L", 32999.99, None, 0),
    ("B-03", "Gin london dry", "700 ml", 21450.0, 19800.0, 112),
    ("B-04", "Vermut rosso", "950 ml", 9870.25, 8990.0, None),
]
_LISTA2: list[tuple[str, str, str, float, str]] = [
    ("S-10", "Agua mineral sin gas", "x 6", 4320.0, "28/10/2026"),
    ("S-11", "Tónica", "x 24", 15999.99, "30/10/2026"),
    ("S-12", "Hielo bolsa", "5 kg", 1850.0, "01/11/2026"),
    ("S-13", "Limón", "kg", 1234.56, "01/11/2026"),
    ("S-14", "Jugo de naranja", "1 L", 2890.0, "15/10/2026"),
]


def build_hard_lista_mixta(path: Path, _rng: random.Random) -> dict[str, Any]:
    d1: list[list[Any]] = [LISTA1_HEADER]
    d1 += [
        [
            r[0],
            r[1],
            r[2],
            ar(r[3], True),
            ar(r[4], True) if r[4] is not None else "",
            str(r[5]) if r[5] is not None else "s/d",
        ]
        for r in _LISTA1
    ]
    t1 = Table(d1, colWidths=[18 * mm, 45 * mm, 28 * mm, 30 * mm, 32 * mm, 18 * mm])
    t1.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
                ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
                _GRID_STYLE,
            ]
        )
    )
    d2: list[list[Any]] = [LISTA2_HEADER]
    d2 += [[r[0], r[1], r[2], ar(r[3], True), r[4]] for r in _LISTA2]
    t2 = Table(d2, colWidths=[18 * mm, 55 * mm, 25 * mm, 32 * mm, 40 * mm])
    t2.setStyle(
        TableStyle(
            [
                ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
                ("ALIGN", (3, 1), (3, -1), "RIGHT"),
                _GRID_STYLE,
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]
        )
    )
    story = [
        Paragraph("<b>LISTA DE PRECIOS — Noviembre 2026</b>", _STYLES["Title"]),
        Paragraph(
            "Vigencia desde el 01/11/2026. Precios en pesos argentinos, IVA incluido. "
            "Sujetos a cambio sin previo aviso.",
            _STYLES["Normal"],
        ),
        Spacer(1, 5 * mm),
        Paragraph("<b>1. Bebidas con alcohol</b>", _STYLES["Heading3"]),
        t1,
        Spacer(1, 6 * mm),
        Paragraph(
            "Los precios mayoristas aplican a partir de 12 unidades por renglón. "
            "Consultar disponibilidad de los productos marcados s/d.",
            _STYLES["Normal"],
        ),
        Spacer(1, 6 * mm),
        Paragraph("<b>2. Sin alcohol y varios</b>", _STYLES["Heading3"]),
        t2,
        Spacer(1, 6 * mm),
        Paragraph("Contacto: ventas@distribuidoradelsur.com.ar — Tel. 011 4321-9876", _STYLES["Normal"]),
    ]
    _doc(path).build(story)

    gt1: list[list[GtCell]] = [
        [
            gt_str(r[0]),
            gt_str(r[1]),
            gt_str(r[2]),
            gt_float(r[3]),
            gt_float(r[4]) if r[4] is not None else None,
            gt_int(r[5]) if r[5] is not None else gt_str("s/d"),
        ]
        for r in _LISTA1
    ]
    gt2: list[list[GtCell]] = [
        [gt_str(r[0]), gt_str(r[1]), gt_str(r[2]), gt_float(r[3]), gt_date(r[4])] for r in _LISTA2
    ]
    return gt_entry(
        [gt_table(LISTA1_HEADER, gt1), gt_table(LISTA2_HEADER, gt2)],
        ["digital"],
        "Dos tablas en la misma página (grilla gris + solo línea bajo el encabezado), texto entre medio, "
        "celdas vacías, 's/d' en columna numérica y fechas",
    )


# --- cont_sin_header ---------------------------------------------------------------------------

CONT_HEADER = ["Fecha", "Comprobante", "Detalle", "Cantidad", "Importe"]
_DETALLES = [
    "Entrega sucursal Centro",
    "Entrega sucursal Norte",
    "Reposición depósito",
    "Devolución por rotura",
    "Entrega cliente mayorista",
    "Muestras sin cargo",
]


def build_cont_sin_header(path: Path, rng: random.Random) -> dict[str, Any]:
    rows: list[tuple[str, str, str, int, float]] = []
    day, month = 1, 9
    for i in range(110):
        if i and i % 4 == 0:
            day += 1
            if day > 28:
                day, month = 1, month + 1
        qty = rng.randint(1, 120)
        unit = rng.choice([350.0, 1250.5, 90.0, 28.75, 15999.99])
        rows.append(
            (
                _fecha(day, month),
                f"{rng.randint(10000, 99999)}",
                _DETALLES[i % len(_DETALLES)],
                qty,
                round(qty * unit, 2),
            )
        )
    data: list[list[Any]] = [CONT_HEADER]
    data += [[r[0], r[1], r[2], str(r[3]), ar(r[4])] for r in rows]
    table = _grid_table(data, [22 * mm, 26 * mm, 70 * mm, 20 * mm, 32 * mm])  # sin repeatRows a propósito
    table.setStyle(TableStyle([("ALIGN", (3, 1), (-1, -1), "RIGHT")]))

    def on_page(c: canvas.Canvas, d: Any) -> None:
        c.setFont("Helvetica", 8)
        c.drawRightString(195 * mm, 10 * mm, f"Página {d.page}")

    story = [
        Paragraph("Detalle de entregas — Septiembre/Octubre 2026", _STYLES["Title"]),
        Paragraph("Cliente: Distribuidora del Sur S.A. — Cuenta: 4521", _STYLES["Normal"]),
        Spacer(1, 4 * mm),
        table,
    ]
    _doc(path).build(story, onFirstPage=on_page, onLaterPages=on_page)
    pages = _page_count(path)
    assert pages == 3, f"cont_sin_header debería tener 3 páginas, tiene {pages}"

    gt_rows: list[list[GtCell]] = [
        [gt_date(r[0]), gt_str(r[1]), gt_str(r[2]), gt_int(r[3]), gt_float(r[4])] for r in rows
    ]
    return gt_entry(
        [gt_table(CONT_HEADER, gt_rows)],
        ["digital"] * pages,
        "Tabla de 110 filas que sigue en 3 páginas SIN repetir el encabezado (merge Regla 2)",
    )


# --- header_dos_filas --------------------------------------------------------------------------

HEADER2_HEADER = ["Código", "Producto", "Precio ($) Lista", "Precio ($) Mayorista", "Stock"]


def build_header_dos_filas(path: Path, rng: random.Random) -> dict[str, Any]:
    products = _FACTURA_PRODUCTS
    rows: list[tuple[str, str, float, float, int]] = []
    for i in range(24):
        lista = round(rng.uniform(800, 45000), 2)
        rows.append(
            (
                f"P-{i + 1:02d}",
                products[i % len(products)],
                lista,
                round(lista * 0.88, 2),
                rng.randint(0, 300),
            )
        )
    data: list[list[Any]] = [
        ["Código", "Producto", "Precio ($)", "", "Stock"],
        ["", "", "Lista", "Mayorista", ""],
    ]
    data += [[r[0], r[1], ar(r[2]), ar(r[3]), str(r[4])] for r in rows]
    table = _grid_table(data, [20 * mm, 70 * mm, 30 * mm, 30 * mm, 18 * mm])
    table.setStyle(
        TableStyle(
            [
                ("SPAN", (2, 0), (3, 0)),
                ("SPAN", (0, 0), (0, 1)),
                ("SPAN", (1, 0), (1, 1)),
                ("SPAN", (4, 0), (4, 1)),
                ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#DDDDDD")),
                ("ALIGN", (2, 0), (-1, -1), "CENTER"),
                ("ALIGN", (2, 2), (-1, -1), "RIGHT"),
            ]
        )
    )
    story = [
        Paragraph("Lista de precios con encabezado de dos filas", _STYLES["Title"]),
        Spacer(1, 4 * mm),
        table,
    ]
    _doc(path).build(story)
    gt_rows: list[list[GtCell]] = [
        [gt_str(r[0]), gt_str(r[1]), gt_float(r[2]), gt_float(r[3]), gt_int(r[4])] for r in rows
    ]
    return gt_entry(
        [gt_table(HEADER2_HEADER, gt_rows)],
        ["digital"],
        "Encabezado de 2 filas con celda combinada 'Precio ($)' sobre Lista/Mayorista. "
        "El GT usa el encabezado combinado; hoy la 2.ª fila del encabezado sale como fila de datos",
    )


# --- apaisado ----------------------------------------------------------------------------------

_MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct"]
APAISADO_HEADER = ["Sucursal", "Código", *_MESES]
_SUCURSALES = ["Centro", "Norte", "Sur", "Oeste", "Caballito", "Belgrano"]


def build_apaisado(path: Path, rng: random.Random) -> dict[str, Any]:
    rows: list[list[Any]] = []
    for i in range(24):
        rows.append(
            [_SUCURSALES[i % len(_SUCURSALES)], f"S-{i + 1:02d}", *[rng.randint(80, 9999) for _ in _MESES]]
        )
    data: list[list[Any]] = [APAISADO_HEADER]
    data += [[r[0], r[1], *[f"{v:,}".replace(",", ".") for v in r[2:]]] for r in rows]
    widths = [28 * mm, 18 * mm] + [20 * mm] * len(_MESES)
    table = _grid_table(data, widths)
    table.setStyle(TableStyle([("ALIGN", (2, 1), (-1, -1), "RIGHT"), ("FONTSIZE", (0, 0), (-1, -1), 7)]))
    story = [
        Paragraph("Unidades vendidas por sucursal y mes — 2026", _STYLES["Title"]),
        Spacer(1, 4 * mm),
        table,
    ]
    _doc(path, pagesize=landscape(A4)).build(story)
    gt_rows: list[list[GtCell]] = [[gt_str(r[0]), gt_str(r[1]), *[gt_int(v) for v in r[2:]]] for r in rows]
    return gt_entry(
        [gt_table(APAISADO_HEADER, gt_rows)],
        ["digital"],
        "Página apaisada con tabla ancha de 12 columnas, enteros con punto de miles ('1.234')",
    )


# --- porcentajes -------------------------------------------------------------------------------

PCT_HEADER = ["Producto", "Precio lista", "IVA", "Descuento", "Precio final"]
_IVAS = [("21%", 0.21), ("10,5 %", 0.105), ("27%", 0.27)]
_DESCUENTOS = [("0%", 0.0), ("5 %", 0.05), ("12,5 %", 0.125), ("10%", 0.10)]


def build_porcentajes(path: Path, rng: random.Random) -> dict[str, Any]:
    rows: list[tuple[str, float, tuple[str, float], tuple[str, float], float]] = []
    for i in range(18):
        price = round(rng.uniform(500, 30000), 2)
        iva = _IVAS[i % len(_IVAS)]
        desc = _DESCUENTOS[rng.randrange(len(_DESCUENTOS))]
        final = round(price * (1 + iva[1]) * (1 - desc[1]), 2)
        rows.append((_FACTURA_PRODUCTS[i % len(_FACTURA_PRODUCTS)], price, iva, desc, final))
    data: list[list[Any]] = [PCT_HEADER]
    data += [[r[0], ar(r[1], True), r[2][0], r[3][0], ar(r[4], True)] for r in rows]
    table = _grid_table(data, [70 * mm, 30 * mm, 22 * mm, 24 * mm, 32 * mm])
    table.setStyle(TableStyle([("ALIGN", (1, 1), (-1, -1), "RIGHT")]))
    _doc(path).build([Paragraph("Precios con IVA y descuentos", _STYLES["Title"]), Spacer(1, 4 * mm), table])
    gt_rows: list[list[GtCell]] = [
        [gt_str(r[0]), gt_float(r[1]), gt_float(r[2][1]), gt_float(r[3][1]), gt_float(r[4])] for r in rows
    ]
    return gt_entry(
        [gt_table(PCT_HEADER, gt_rows)],
        ["digital"],
        "Columnas de porcentaje ('21%', '10,5 %') que deben quedar como fracción (0,21)",
    )


# --- numeros_us --------------------------------------------------------------------------------

US_HEADER = ["SKU", "Description", "Qty", "Unit price", "Amount"]
_US_PRICES = [350.0, 1234.56, 12500.0, 89.99, 4075.25, 10.5]
_US_ITEMS = [
    "Steel bracket",
    "Copper wire 10 m",
    "Hex bolt M8",
    "LED panel 60x60",
    "Power supply 24 V",
    "Cable tie x100",
]


def build_numeros_us(path: Path, rng: random.Random) -> dict[str, Any]:
    rows: list[tuple[str, str, int, float, float]] = []
    for i in range(20):
        qty = rng.randint(1, 40)
        unit = _US_PRICES[i % len(_US_PRICES)]
        rows.append((f"US-{100 + i}", _US_ITEMS[i % len(_US_ITEMS)], qty, unit, round(qty * unit, 2)))
    data: list[list[Any]] = [US_HEADER]
    data += [[r[0], r[1], str(r[2]), us(r[3]), us(r[4])] for r in rows]
    table = _grid_table(data, [22 * mm, 70 * mm, 16 * mm, 30 * mm, 34 * mm])
    table.setStyle(TableStyle([("ALIGN", (2, 1), (-1, -1), "RIGHT")]))
    _doc(path).build([Paragraph("Invoice #2026-0412", _STYLES["Title"]), Spacer(1, 4 * mm), table])
    gt_rows: list[list[GtCell]] = [
        [gt_str(r[0]), gt_str(r[1]), gt_int(r[2]), gt_float(r[3]), gt_float(r[4])] for r in rows
    ]
    return gt_entry(
        [gt_table(US_HEADER, gt_rows)],
        ["digital"],
        "Todas las columnas numéricas en formato US ('1,234.56', '350.00')",
    )


# --- negativos_banco ---------------------------------------------------------------------------

NEG_HEADER = ["Fecha", "Concepto", "Débito", "Crédito", "Saldo"]


def _neg_ar(value: float, style: int) -> str:
    """Negativo al estilo de los extractos: ``1.234,56-`` (style 0) o ``(1.234,56)`` (style 1)."""
    body = ar(abs(value))
    return f"{body}-" if style == 0 else f"({body})"


def build_negativos_banco(path: Path, rng: random.Random) -> dict[str, Any]:
    rows: list[tuple[str, str, float | None, float | None, float]] = []
    saldo = 12500.0
    day = 1
    for i in range(32):  # una sola página: el caso mide formatos negativos, no paginación
        monto = round(rng.choice([-1, -1, 1]) * rng.uniform(100, 25000), 2)
        saldo = round(saldo + monto, 2)
        rows.append(
            (
                _fecha((day % 28) + 1, 10),
                _CONCEPTOS[i % len(_CONCEPTOS)],
                monto if monto < 0 else None,
                monto if monto > 0 else None,
                saldo,
            )
        )
        day += rng.randint(0, 1)
    data: list[list[Any]] = [NEG_HEADER]
    for i, r in enumerate(rows):
        debito = _neg_ar(r[2], i % 2) if r[2] is not None else ""
        credito = ar(r[3]) if r[3] is not None else ""
        saldo_txt = ar(r[4]) if r[4] >= 0 else _neg_ar(r[4], 1)
        data.append([r[0], r[1], debito, credito, saldo_txt])
    table = _grid_table(data, [22 * mm, 70 * mm, 28 * mm, 28 * mm, 30 * mm])
    table.setStyle(
        TableStyle([("ALIGN", (2, 1), (-1, -1), "RIGHT"), ("GRID", (0, 0), (-1, -1), 0.4, colors.grey)])
    )
    _doc(path).build(
        [Paragraph("Movimientos de cuenta — Octubre 2026", _STYLES["Title"]), Spacer(1, 4 * mm), table]
    )
    gt_rows: list[list[GtCell]] = [
        [
            gt_date(r[0]),
            gt_str(r[1]),
            gt_float(r[2]) if r[2] is not None else None,
            gt_float(r[3]) if r[3] is not None else None,
            gt_float(r[4]),
        ]
        for r in rows
    ]
    pages = _page_count(path)
    assert pages == 1, f"negativos_banco debería tener 1 página, tiene {pages}"
    return gt_entry(
        [gt_table(NEG_HEADER, gt_rows)],
        ["digital"],
        "Negativos con signo al final ('1.234,56-') y entre paréntesis ('(1.234,56)')",
    )


# --- con_password, sin_tablas, vacio, mixto, texto_basura ---------------------------------------


def build_con_password(path: Path, _rng: random.Random) -> dict[str, Any]:
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, invariant=1).build(
        _remito_story("Remito N° 4522 — confidencial", True)
    )
    with pymupdf.open(stream=buf.getvalue(), filetype="pdf") as doc:
        doc.save(
            str(path),
            encryption=_PDF_ENCRYPT_AES_256,
            user_pw=PASSWORD,
            owner_pw=PASSWORD,
        )
    _, gt = _remito_data()
    return gt_entry(
        [gt_table(REMITO_HEADER, gt)],
        ["digital"],
        f"Remito con grilla cifrado AES-256 (contraseña de usuario '{PASSWORD}'). "
        "No es reproducible byte a byte: el cifrado usa sales aleatorias",
        password=PASSWORD,
    )


_PARRAFOS = [
    "Estimado cliente: le escribimos para informarle que a partir del 1 de noviembre de 2026 nuestros "
    "horarios de atención cambian. El depósito de Avellaneda abrirá de lunes a viernes de 8 a 17 y los "
    "sábados de 9 a 13.",
    "Las entregas programadas se mantienen sin cambios. Si necesita coordinar una entrega fuera de ese "
    "horario, comuníquese con su vendedor asignado o escríbanos a ventas@distribuidoradelsur.com.ar.",
    "Agradecemos su confianza y quedamos a disposición por cualquier consulta. Saludos cordiales, "
    "Distribuidora del Sur S.A.",
]


def build_sin_tablas(path: Path, _rng: random.Random) -> dict[str, Any]:
    story: list[Any] = [Paragraph("Comunicación a clientes", _STYLES["Title"]), Spacer(1, 4 * mm)]
    for text in _PARRAFOS:
        story += [Paragraph(text, _STYLES["Normal"]), Spacer(1, 4 * mm)]
    _doc(path).build(story)
    return gt_entry(
        [], ["digital"], "Carta de 3 párrafos sin ninguna tabla: el Excel lleva solo la hoja 'Texto'"
    )


def build_vacio(path: Path, _rng: random.Random) -> dict[str, Any]:
    c = _canvas(path)
    c.showPage()
    c.save()
    return gt_entry([], ["empty"], "Una página en blanco (sin texto ni imágenes)")


def _scan_jpeg(
    src_pdf: bytes, page_no: int, dpi: int, rng: random.Random, rot_deg: float = 0.0
) -> tuple[bytes, tuple[float, float]]:
    """Render gris de una página + rotación leve + blur + ruido → JPEG q70.

    Devuelve ``(jpeg, (ancho, alto) en pt)``.
    """
    with pymupdf.open(stream=src_pdf, filetype="pdf") as doc:
        page = doc[page_no]
        size = (float(page.rect.width), float(page.rect.height))
        pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
        png = pix.tobytes("png")
    return _degrade(png, rng, rot_deg, blur=0.6, noise=3000, jpeg_q=70), size


def _degrade(
    png: bytes,
    rng: random.Random,
    rot_deg: float,
    blur: float,
    noise: int,
    jpeg_q: int,
    rotate90: bool = False,
) -> bytes:
    im = Image.open(io.BytesIO(png)).convert("L")
    if rot_deg:
        im = im.rotate(rot_deg, resample=Image.Resampling.BICUBIC, fillcolor=255, expand=False)
    if blur:
        im = im.filter(ImageFilter.GaussianBlur(blur))
    if noise:
        px = im.load()
        assert px is not None
        for _ in range(noise):
            x, y = rng.randrange(im.width), rng.randrange(im.height)
            px[x, y] = rng.randint(120, 200)
    if rotate90:
        im = im.rotate(90, expand=True)  # antihorario: el texto queda leyéndose de abajo hacia arriba
    out = io.BytesIO()
    im.save(out, "JPEG", quality=jpeg_q)
    return out.getvalue()


def build_mixto(path: Path, _rng: random.Random) -> dict[str, Any]:
    # Página 2: carta renderizada como imagen (escaneo), sin tablas, así expected_tables no depende del OCR.
    letter = io.BytesIO()
    story: list[Any] = [Paragraph("Condiciones de entrega", _STYLES["Title"]), Spacer(1, 4 * mm)]
    for text in _PARRAFOS:
        story += [Paragraph(text, _STYLES["Normal"]), Spacer(1, 4 * mm)]
    SimpleDocTemplate(letter, pagesize=A4, invariant=1, leftMargin=15 * mm, rightMargin=15 * mm).build(story)
    jpeg, size = _scan_jpeg(letter.getvalue(), 0, 200, random.Random(SCAN_SEED), rot_deg=0.8)

    data, gt = _remito_data()
    table = Table(data, colWidths=[60, 140, 70, 100, 80])
    table.setStyle(
        TableStyle([("FONTSIZE", (0, 0), (-1, -1), 10), ("GRID", (0, 0), (-1, -1), 0.8, colors.black)])
    )
    c = _canvas(path)
    width, height = A4
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2, height - 96, "Remito N° 4523")
    c.setFont("Helvetica", 10)
    c.drawString(72, height - 120, "Cliente: Bar El Catorce — Fecha: 16/10/2026")
    _, th = table.wrapOn(c, width - 144, height)
    table.drawOn(c, 72, height - 140 - th)
    c.showPage()
    c.setPageSize(size)
    c.drawImage(ImageReader(io.BytesIO(jpeg)), 0, 0, size[0], size[1])
    c.showPage()
    c.save()
    return gt_entry(
        [gt_table(REMITO_HEADER, gt)],
        ["digital", "scanned"],
        "Página 1 digital con un remito; página 2 solo imagen (carta escaneada sin tablas)",
    )


_MOJIBAKE = [
    "Ã‰Ã“Ã‘ Â¿Â¡ Ã€Ã‚Ãƒ Ã„Ã…Ã† Ã‡ÃˆÃŠ Ã‹ÃŒÃŽ ÃÃÃ’ Ã”Ã•Ã–Ã˜Ã™",
    "Ã›ÃÃžÃŸ Ã Ã¢Ã£Ã¤Ã¥ Ã¦Ã§Ã¨ÃªÃ« Â§Â¶Â©Â®Â± ÂµÂ¼Â½Â¾ Ã—Ã·Â¢Â£Â¥",
    "Â¦Â¨ÂªÂ«Â¬ Â¯Â²Â³Â´Â¸ Â¹Â»Â¿ ÃŒÃÃŽÃ ÃÃ‘Ã’Ã“Ã” Ã•Ã–Ã—Ã˜Ã™Ãš",
    "ÃžÃŸÃ Ã¡Ã¢ Ã£Ã¤Ã¥Ã¦Ã§ Ã¨Ã©ÃªÃ«Ã¬ Ã­Ã®Ã¯Ã°Ã± Ã²Ã³Ã´ÃµÃ¶ Ã·Ã¸Ã¹ÃºÃ»",
]


def build_texto_basura(path: Path, _rng: random.Random) -> dict[str, Any]:
    """Texto cuya capa extraíble es basura (como una fuente embebida sin ToUnicode)."""
    c = _canvas(path)
    _width, height = A4
    c.setFont("Helvetica-Bold", 14)
    c.drawString(72, height - 80, "Ã‰Ã“Ã‘ Â¿Â¡Ã€Ã‚Ãƒ Ã„Ã…Ã† ÂµÂ¼Â½Â¾")
    c.setFont("Helvetica", 11)
    y = height - 110.0
    for _ in range(4):
        for line in _MOJIBAKE:
            c.drawString(72, y, line)
            y -= 16
        y -= 10
    c.showPage()
    c.save()
    return gt_entry(
        [],
        ["garbage_text"],
        "Texto con > 30 % de caracteres fuera del alfabeto esperable (simula fuente sin ToUnicode): "
        "se trata como escaneo",
    )


# ------------------------------------------------------------------------------------------- escaneos


def make_scan(
    src: Path,
    dst: Path,
    dpi: int = 200,
    rot_deg: float = 0.0,
    blur: float = 0.6,
    noise: int = 3000,
    jpeg_q: int = 70,
    rotate90: bool = False,
    seed: int = SCAN_SEED,
) -> Path:
    """Versión "escaneada" de ``src``: cada página se renderiza a ``dpi`` en gris, se rota ``rot_deg``
    (alternando el signo en cada página), se le aplica blur gaussiano, ``noise`` píxeles de ruido y
    compresión JPEG ``jpeg_q``; con ``rotate90`` la página queda girada 90° (apaisada). Solo imagen, sin
    capa de texto. Determinista (``seed``)."""
    rng = random.Random(seed)
    with pymupdf.open(src) as doc:
        c: canvas.Canvas | None = None
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
            angle = rot_deg if i % 2 == 0 else -rot_deg
            jpeg = _degrade(pix.tobytes("png"), rng, angle, blur, noise, jpeg_q, rotate90=rotate90)
            w, h = float(page.rect.width), float(page.rect.height)
            if rotate90:
                w, h = h, w
            if c is None:
                c = _canvas(dst, pagesize=(w, h))
            else:
                c.setPageSize((w, h))
            c.drawImage(ImageReader(io.BytesIO(jpeg)), 0, 0, w, h)
            c.showPage()
        assert c is not None, f"{src} no tiene páginas"
        c.save()
    return dst


@dataclass(frozen=True)
class ScanSpec:
    name: str  # nombre del archivo de salida (sin .pdf)
    source: str  # caso digital de origen
    dpi: int = 200
    rotate90: bool = False


SCANS: tuple[ScanSpec, ...] = (
    ScanSpec("hard_factura_ar_scan", "hard_factura_ar"),
    ScanSpec("hard_lista_mixta_scan", "hard_lista_mixta"),
    ScanSpec("remito_grid_scan", "remito_grid"),
    ScanSpec("remito_nogrid_scan", "remito_nogrid"),
    ScanSpec("cont_sin_header_scan", "cont_sin_header"),
    ScanSpec("hard_factura_ar_scan90", "hard_factura_ar", rotate90=True),
    ScanSpec("remito_grid_scan150", "remito_grid", dpi=150),
)

# Orden fijo: la semilla alimenta un único RNG, así que cambiar el orden cambia los datos (y el GT).
CASES: tuple[Case, ...] = (
    Case("hard_factura_ar", build_hard_factura_ar),
    Case("hard_extracto_banco", build_hard_extracto_banco),
    Case("hard_lista_mixta", build_hard_lista_mixta),
    Case("remito_grid", build_remito_grid),
    Case("remito_nogrid", build_remito_nogrid),
    Case("prueba", build_prueba),
    Case("cont_sin_header", build_cont_sin_header),
    Case("header_dos_filas", build_header_dos_filas),
    Case("apaisado", build_apaisado),
    Case("porcentajes", build_porcentajes),
    Case("numeros_us", build_numeros_us),
    Case("negativos_banco", build_negativos_banco),
    Case("con_password", build_con_password),
    Case("sin_tablas", build_sin_tablas),
    Case("vacio", build_vacio),
    Case("mixto", build_mixto),
    Case("texto_basura", build_texto_basura),
)

DIGITAL_NAMES: tuple[str, ...] = tuple(c.name for c in CASES)
SCAN_NAMES: tuple[str, ...] = tuple(s.name for s in SCANS)
ALL_NAMES: tuple[str, ...] = DIGITAL_NAMES + SCAN_NAMES
# Archivos que no son reproducibles byte a byte entre corridas (cifrado con sales aleatorias).
NON_DETERMINISTIC: frozenset[str] = frozenset({"con_password.pdf"})


def generate_with_ground_truth(out_dir: Path, seed: int = DEFAULT_SEED) -> tuple[list[Path], GroundTruth]:
    """Genera todos los PDFs en ``out_dir`` y devuelve ``(rutas, ground_truth)``."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)
    paths: list[Path] = []
    gt: GroundTruth = {}
    for case in CASES:
        path = out_dir / f"{case.name}.pdf"
        entry = case.build(path, rng)
        paths.append(path)
        if entry is not None:
            gt[path.name] = entry
    scan_rng = random.Random(seed + 1)
    for spec in SCANS:
        angle = round(scan_rng.uniform(-1.2, 1.2), 2)
        paths.append(
            make_scan(
                out_dir / f"{spec.source}.pdf",
                out_dir / f"{spec.name}.pdf",
                dpi=spec.dpi,
                rot_deg=angle,
                rotate90=spec.rotate90,
            )
        )
    return paths, gt


def generate_all(out_dir: Path, seed: int = DEFAULT_SEED) -> list[Path]:
    """Genera el corpus completo en ``out_dir`` (determinista para ``seed``) y devuelve las rutas."""
    return generate_with_ground_truth(out_dir, seed)[0]


def ground_truth(seed: int = DEFAULT_SEED) -> GroundTruth:
    """Ground truth de todos los casos (genera en un directorio temporal: los datos salen del mismo RNG)."""
    with tempfile.TemporaryDirectory(prefix="pdf2excel-corpus-") as tmp:
        return generate_with_ground_truth(Path(tmp), seed)[1]


def write_ground_truth(target: Path, seed: int = DEFAULT_SEED) -> Path:
    data = ground_truth(seed)
    target.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return target


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m tests.corpus.generate",
        description="Genera el corpus sintético de PDFs de prueba (determinista).",
    )
    parser.add_argument("out_dir", type=Path, help="carpeta de salida (se crea si no existe)")
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED)
    parser.add_argument(
        "--ground-truth",
        type=Path,
        default=None,
        metavar="JSON",
        help="además escribe el ground truth en ese archivo (p. ej. tests/corpus/ground_truth.json)",
    )
    args = parser.parse_args(argv)
    paths, gt = generate_with_ground_truth(args.out_dir, args.seed)
    for p in paths:
        print(f"{p.stat().st_size:>9} B  {p}")
    if args.ground_truth is not None:
        args.ground_truth.write_text(json.dumps(gt, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        print(f"ground truth → {args.ground_truth} ({len(gt)} entradas)")
    print(f"{len(paths)} PDFs en {args.out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
