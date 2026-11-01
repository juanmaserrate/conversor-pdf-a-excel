"""Corpus sintético de FORMULARIOS para medir el modo "igual al PDF" (docs/LAYOUT_MODE.md §7).

Genera PDFs de formularios (hoja de ruta, remito con cajas, factura, etc.) con reportlab y Pillow. Todos los
nombres, empresas, CUIT y patentes son **inventados**. La salida es determinista (reportlab ``invariant=1``,
semillas fijas): dos corridas producen los mismos bytes.

La ``hoja_ruta`` imita lo que hace un navegador al imprimir: bordes finos dibujados como **rectángulos
rellenos**, encabezado/pie de impresión, y dos textos "N/A" dibujados y luego **tapados** por un rectángulo
blanco (el caso del bug de los textos fantasma).

Junto con los PDFs se escribe ``ground_truth_forms.json``: por PDF, las palabras visibles esperadas (sin las
tapadas), las palabras tapadas y las cajas esperadas (cantidad, y cuáles son "etiqueta + valor"). Las
coordenadas de las cajas están en puntos, con origen arriba a la izquierda del marco *visual* de la página.

Uso: ``python -m tests.corpus.generate_forms <out_dir>``.
"""

from __future__ import annotations

import argparse
import json
import random
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pymupdf
from PIL import Image, ImageDraw
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

GROUND_TRUTH_NAME = "ground_truth_forms.json"
W_A4, H_A4 = A4

INK = "#1f2a5a"
GRAY = "#8a8f9c"
LINE = "#c3c8d2"
ORANGE = "#b45309"
HEADER_FILL = "#eef1f6"
SECTION_FILL = "#f4f6fa"


# --------------------------------------------------------------------------------------------- registro


@dataclass
class Recorder:
    """Lo que se dibujó: palabras visibles/tapadas y cajas esperadas (ground truth)."""

    visible: list[str] = field(default_factory=list)
    covered: list[str] = field(default_factory=list)
    boxes: list[dict[str, Any]] = field(default_factory=list)
    tables: list[dict[str, Any]] = field(default_factory=list)
    pages: int = 0

    def add_text(self, text: str, covered: bool = False) -> None:
        (self.covered if covered else self.visible).extend(text.split())

    def add_box(
        self,
        page: int,
        bbox: tuple[float, float, float, float],
        kind: str,
        label: str = "",
        value: str = "",
    ) -> None:
        self.boxes.append(
            {
                "page": page,
                "bbox": [round(v, 1) for v in bbox],
                "kind": kind,
                "label": label,
                "value": value,
            }
        )

    def ground_truth(self, description: str, **extra: Any) -> dict[str, Any]:
        label_value = [b for b in self.boxes if b["kind"] == "label_value"]
        gt: dict[str, Any] = {
            "description": description,
            "pages": self.pages,
            "visible_words": self.visible,
            "covered_words": self.covered,
            "boxes": {
                "count": len(self.boxes),
                "label_value": len(label_value),
                "items": self.boxes,
            },
            "tables": self.tables,
        }
        gt.update(extra)
        return gt


class Pen:
    """Pincel sobre un ``canvas`` con coordenadas arriba-izquierda (``y`` crece hacia abajo)."""

    def __init__(self, c: canvas.Canvas, rec: Recorder, width: float = W_A4, height: float = H_A4) -> None:
        self.c = c
        self.rec = rec
        self.w = width
        self.h = height
        self.page = 1

    # -- texto
    def text(
        self,
        x: float,
        y: float,
        s: str,
        size: float = 8,
        bold: bool = False,
        color: str = "#000000",
        align: str = "l",
        covered: bool = False,
    ) -> None:
        c = self.c
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.setFillColor(colors.HexColor(color))
        y_pdf = self.h - y
        if align == "r":
            c.drawRightString(x, y_pdf, s)
        elif align == "c":
            c.drawCentredString(x, y_pdf, s)
        else:
            c.drawString(x, y_pdf, s)
        self.rec.add_text(s, covered=covered)

    @staticmethod
    def text_width(s: str, size: float, bold: bool = False) -> float:
        return float(pymupdf.get_text_length(s, fontname="hebo" if bold else "helv", fontsize=size))

    # -- figuras
    def fill(self, x: float, y: float, w: float, h: float, color: str) -> None:
        self.c.setFillColor(colors.HexColor(color))
        self.c.rect(x, self.h - y - h, w, h, stroke=0, fill=1)

    def stroke(
        self, x: float, y: float, w: float, h: float, color: str = "#000000", lw: float = 0.6, r: float = 0
    ) -> None:
        self.c.setStrokeColor(colors.HexColor(color))
        self.c.setLineWidth(lw)
        if r:
            self.c.roundRect(x, self.h - y - h, w, h, r, stroke=1, fill=0)
        else:
            self.c.rect(x, self.h - y - h, w, h, stroke=1, fill=0)

    def line(
        self, x0: float, y0: float, x1: float, y1: float, color: str = "#000000", lw: float = 0.6
    ) -> None:
        self.c.setStrokeColor(colors.HexColor(color))
        self.c.setLineWidth(lw)
        self.c.line(x0, self.h - y0, x1, self.h - y1)

    def thin_border(self, x: float, y: float, w: float, h: float, color: str = LINE, t: float = 0.75) -> None:
        """Borde de celda como el de un navegador: cuatro rectángulos rellenos finos."""
        self.fill(x, y, w, t, color)
        self.fill(x, y + h - t, w, t, color)
        self.fill(x, y, t, h, color)
        self.fill(x + w - t, y, t, h, color)

    def image(self, img: Image.Image, x: float, y: float, w: float, h: float) -> None:
        self.c.drawImage(ImageReader(img), x, self.h - y - h, w, h)

    def next_page(self) -> None:
        self.c.showPage()
        self.page += 1
        self.rec.pages = max(self.rec.pages, self.page)

    def label_box(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        label: str,
        value: str = "",
        radius: float = 3,
        rounded: bool = True,
    ) -> None:
        """Caja "ETIQUETA chica gris / Valor en negrita" (el valor puede quedar vacío)."""
        self.stroke(x, y, w, h, LINE, 0.7, radius if rounded else 0)
        self.text(x + 5, y + 9, label, 6, color=GRAY)
        if value:
            self.text(x + 5, y + h - 6, value, 9, bold=True, color=INK)
        self.rec.add_box(self.page, (x, y, x + w, y + h), "label_value" if value else "empty", label, value)


def _new_canvas(path: Path, pagesize: tuple[float, float] = A4) -> canvas.Canvas:
    c = canvas.Canvas(str(path), pagesize=pagesize, invariant=1)
    c.setTitle(path.stem)
    c.setAuthor("corpus sintético")
    return c


def _money(value: float) -> str:
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# --------------------------------------------------------------------------------------------- imágenes


def make_logo(seed: int = 5) -> Image.Image:
    """Logo raster ficticio (formas geométricas, sin tipografía)."""
    rnd = random.Random(seed)
    img = Image.new("RGB", (360, 170), "white")
    d = ImageDraw.Draw(img)
    navy = (31, 42, 90)
    d.ellipse((10, 20, 110, 120), outline=navy, width=14)
    d.rectangle((135, 20, 160, 120), fill=navy)
    d.polygon([(135, 70), (205, 20), (225, 20), (160, 70), (225, 120), (200, 120)], fill=navy)
    d.rectangle((250, 20, 275, 120), fill=navy)
    d.rectangle((250, 20, 330, 42), fill=navy)
    d.ellipse((340, 100, 356, 116), fill=navy)
    for i in range(6):
        x = 12 + i * 18 + rnd.randint(0, 2)
        d.rectangle((x, 140, x + 10, 150), fill=(120, 126, 150))
    return img


def make_map(width: int = 1090, height: int = 676, seed: int = 11) -> Image.Image:
    """Mapa falso: calles, parques, un río y una ruta roja con pines."""
    rnd = random.Random(seed)
    img = Image.new("RGB", (width, height), (238, 235, 228))
    d = ImageDraw.Draw(img)
    for _ in range(7):  # manzanas verdes (parques)
        x, y = rnd.randint(0, width - 120), rnd.randint(0, height - 100)
        d.rectangle((x, y, x + rnd.randint(40, 120), y + rnd.randint(30, 90)), fill=(206, 226, 196))
    d.polygon(
        [
            (0, height * 0.82),
            (width * 0.35, height * 0.7),
            (width * 0.7, height * 0.9),
            (width, height * 0.78),
            (width, height),
            (0, height),
        ],
        fill=(190, 214, 235),
    )
    for i in range(0, width, 54):  # calles
        d.line((i + rnd.randint(-6, 6), 0, i + rnd.randint(-20, 20), height), fill=(255, 255, 255), width=5)
    for j in range(0, height, 48):
        d.line((0, j + rnd.randint(-5, 5), width, j + rnd.randint(-14, 14)), fill=(255, 255, 255), width=5)
    for _ in range(5):  # avenidas
        d.line((rnd.randint(0, width), 0, rnd.randint(0, width), height), fill=(250, 214, 150), width=9)
    route = [(160, 600), (330, 560), (520, 590), (610, 470), (640, 260), (780, 215), (850, 140), (900, 90)]
    d.line(route, fill=(224, 68, 68), width=7, joint="curve")
    for i, (px, py) in enumerate(route[1:], start=1):
        d.ellipse((px - 11, py - 11, px + 11, py + 11), fill=(214, 40, 40), outline="white", width=3)
        if i % 2 == 0:
            d.rectangle((px - 3, py - 3, px + 3, py + 3), fill="white")
    d.ellipse((149, 589, 171, 611), fill=(31, 42, 90), outline="white", width=3)
    return img


# --------------------------------------------------------------------------------------------- hoja_ruta

_STOPS: list[tuple[str, str, str]] = [
    ("ALMACEN LA ESQUINA", "06:14", "06:28"),
    ("KIOSCO EL TREBOL", "06:40", "06:53"),
    ("SUPER DOS HERMANOS", "07:02", "07:11"),
    ("DESPENSA SANTA ROSA", "07:22", "07:36"),
    ("MINIMERCADO ALAMOS", "07:48", "08:01"),
    ("AUTOSERVICIO EL PINO", "08:28", "08:49"),
    ("ALMACEN DON ERNESTO", "08:55", "09:07"),
    ("KIOSCO LA PLAZA", "09:07", "09:34"),
    ("SUPER LOS CEDROS", "09:44", "09:56"),
    ("DESPENSA LUCIA", "10:07", "10:12"),
    ("MINIMERCADO SOL", "10:18", "10:30"),
]
_ENVASES = ["PALLETS", "CAJONES", "BIDONES", "PLÁSTICO", "CANASTO", "DEPÓSITO", "JAULA", "CARRETA"]
_REFS_COLS = ((32, 0), (222, 4), (400, 8))


def _print_header(p: Pen, title: str, page_no: int, total: int) -> None:
    p.text(24, 20, "27/10/26, 19:04", 8)
    p.text(302, 20, title, 8)
    p.text(24, 823, "about:blank", 8)
    p.text(571, 823, f"{page_no}/{total}", 8, align="r")


def _hr_header(p: Pen, logo: Image.Image, title: str, subtitle: tuple[str, ...], with_km: bool) -> None:
    p.image(logo, 27, 26, 62, 29)
    p.text(94, 42, title, 15, bold=True, color=INK)
    for i, line in enumerate(subtitle):
        p.text(94, 53 + i * 9, line, 8, color=GRAY)
    # caja del número de viaje
    p.stroke(289 if with_km else 500, 32, 43 if with_km else 70, 29, INK, 0.9, 4)
    cx = 310.5 if with_km else 535
    p.text(cx, 40, "VIAJE N°", 5.5, bold=True, color=INK, align="c")
    p.text(cx, 56, "7141", 13, bold=True, color=INK, align="c")
    p.rec.add_box(
        p.page, (289 if with_km else 500, 32, 332 if with_km else 570, 61), "label_value", "VIAJE N°", "7141"
    )
    if with_km:
        for x0, x1, label in ((354, 455, "KM SALIDA DEL DEPÓSITO"), (464, 570, "KM LLEGADA AL DEPÓSITO")):
            p.stroke(x0, 32, x1 - x0, 29, INK, 0.9, 4)
            p.text((x0 + x1) / 2, 44, label, 6.5, bold=True, color=INK, align="c")
            p.fill(x0 + 10, 55, x1 - x0 - 20, 0.7, GRAY)
            p.rec.add_box(p.page, (x0, 32, x1, 61), "empty", label, "")
    p.fill(25, 70, 545, 2, INK)


def _hoja_ruta(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path)
    p = Pen(c, rec)
    logo = make_logo()
    title = "Reparto Demo — Hoja de Ruta"
    _print_header(p, title, 1, 2)
    _hr_header(p, logo, "HOJA DE RUTA", ("VIAJE N° 7141 · REPARTO R3 ·", "23/10/2026"), with_km=True)

    # dos filas de cajas etiqueta/valor (algunas vacías)
    row1 = [
        (25, 134, "TRANSPORTISTA", "ROBERTO QUINTANA"),
        (165, 129, "AUXILIAR", "SOFIA MARQUEZ"),
        (299, 133, "VEHÍCULO", ""),
        (438, 132, "PATENTE", ""),
    ]
    row2 = [
        (25, 134, "REPARTO", "R3"),
        (165, 129, "MOTIVO / UNIDAD", "NORTE GRANDE"),
        (299, 133, "TEMPERATURA", ""),
        (438, 132, "REFRIGERADO", "NO"),
    ]
    for y, row in ((80, row1), (108, row2)):
        for x, w, label, value in row:
            p.label_box(x, y, w, 24, label, value)

    # tabla de paradas (bordes finos como rectángulos rellenos)
    cols = [25, 42, 283, 337, 389, 461, 570]
    top = 138
    heads = ["#", "DESTINO", "HORA\nLLEGADA", "HORA\nSALIDA", "MOTIVO", "OBSERVACIONES"]
    for i, head in enumerate(heads):
        x0, x1 = cols[i], cols[i + 1]
        p.fill(x0, top, x1 - x0, 24, HEADER_FILL)
        p.thin_border(x0, top, x1 - x0, 24)
        color = ORANGE if head == "MOTIVO" else (INK if "HORA" in head else GRAY)
        bold = head != "#"
        lines = head.split("\n")
        if len(lines) == 2:
            p.text((x0 + x1) / 2, top + 10, lines[0], 6.5, bold=bold, color=color, align="c")
            p.text((x0 + x1) / 2, top + 19, lines[1], 6.5, bold=bold, color=color, align="c")
        else:
            p.text((x0 + x1) / 2, top + 14, head, 6.5, bold=bold, color=color, align="c")
    y = top + 24
    rh = 20
    rows_info: list[dict[str, Any]] = []

    def plain_row(idx: str, name: str, arr: str, dep: str, motive: str, bold_name: bool = True) -> None:
        nonlocal y
        for i in range(6):
            p.thin_border(cols[i], y, cols[i + 1] - cols[i], rh)
        p.text(cols[0] + 4, y + 13, idx, 7, color=GRAY)
        p.text(cols[1] + 8, y + 13, name, 7.5, bold=bold_name, color="#111111")
        if arr:
            p.text((cols[2] + cols[3]) / 2, y + 13, arr, 7.5, bold=True, color=INK, align="c")
        if dep:
            p.text((cols[3] + cols[4]) / 2, y + 13, dep, 7.5, bold=True, color=INK, align="c")
        if motive:
            p.text((cols[4] + cols[5]) / 2, y + 13, motive, 6.5, bold=True, color=ORANGE, align="c")
        rows_info.append({"y": y, "name": name})
        y += rh

    def section_row(label: str) -> None:
        nonlocal y
        p.fill(cols[0], y, cols[-1] - cols[0], 14, SECTION_FILL)
        p.thin_border(cols[0], y, cols[-1] - cols[0], 14)
        p.text(cols[0] + 5, y + 10, label, 6.5, bold=True, color=INK)
        rows_info.append({"y": y, "name": label})
        y += 14

    # fila 0: salida del depósito (combina # y DESTINO); la hora de llegada trae un "N/A" que se TAPA
    p.thin_border(cols[0], y, cols[2] - cols[0], rh)  # # y DESTINO combinados
    for i in range(2, 6):
        p.thin_border(cols[i], y, cols[i + 1] - cols[i], rh)
    p.text(cols[0] + 8, y + 13, "DEPÓSITO CENTRAL — SALIDA", 7.5, bold=True, color="#111111")
    p.text((cols[2] + cols[3]) / 2, y + 13, "N/A", 7.5, bold=True, color=INK, align="c", covered=True)
    p.fill(cols[2] + 1, y + 1, cols[3] - cols[2] - 2, rh - 2, "#ffffff")
    p.text((cols[3] + cols[4]) / 2, y + 13, "05:22", 7.5, bold=True, color=INK, align="c")
    y += rh
    section_row("R3 · NORTE GRANDE")
    for n, (name, arr, dep) in enumerate(_STOPS, start=1):
        plain_row(str(n), name, arr, dep, "NORTE GRANDE")
    section_row("VUELTA AL DEPÓSITO · NORTE GRANDE")
    # última fila: la hora de salida trae otro "N/A" tapado
    for i in range(6):
        p.thin_border(cols[i], y, cols[i + 1] - cols[i], rh)
    p.text(cols[0] + 4, y + 13, "12", 7, color=GRAY)
    p.text(cols[1] + 8, y + 13, "DEPÓSITO CENTRAL", 7.5, bold=True, color="#111111")
    p.text((cols[2] + cols[3]) / 2, y + 13, "10:38", 7.5, bold=True, color=INK, align="c")
    p.text((cols[3] + cols[4]) / 2, y + 13, "N/A", 7.5, bold=True, color=INK, align="c", covered=True)
    p.fill(cols[3] + 1, y + 1, cols[4] - cols[3] - 2, rh - 2, "#ffffff")
    p.text((cols[4] + cols[5]) / 2, y + 13, "NORTE GRANDE", 6.5, bold=True, color=ORANGE, align="c")
    y += rh
    rec.tables.append({"page": 1, "bbox": [25, top, 570, y], "columns": 6, "rows": len(rows_info) + 2})

    # observaciones (recuadro vacío)
    y += 14
    p.text(25, y + 4, "OBSERVACIONES", 6.5, bold=True, color="#a3a8b4")
    p.stroke(25, y + 9, 545, 33, LINE, 0.7)
    rec.add_box(1, (25, y + 9, 570, y + 42), "empty", "OBSERVACIONES", "")

    # envases
    y += 62
    p.text(25, y, "ENVASES · ENTRADA Y SALIDA DE CAJONES CON SEÑA", 6.5, bold=True, color="#a3a8b4")
    y += 6
    ecols = [25, 123, 168, 213, 258, 303, 348, 393, 438, 483, 570]
    p.fill(25, y, 545, 16, "#f3f4f6")
    for i, head in enumerate(["SECTOR", *_ENVASES, "FIRMA Y ACLARACIÓN"]):
        p.thin_border(ecols[i], y, ecols[i + 1] - ecols[i], 16, "#b9bec9")
        if i in (0, 9):
            p.text(ecols[i] + 4, y + 11, head, 6, color="#a3a8b4")
        else:
            p.text((ecols[i] + ecols[i + 1]) / 2, y + 11, head, 6, color="#a3a8b4", align="c")
    y += 16
    for sector in ("ARMADO (salida)", "CARGA (salida)", "RECEPCIÓN (entrada)"):
        for i in range(len(ecols) - 1):
            p.thin_border(ecols[i], y, ecols[i + 1] - ecols[i], 22, "#b9bec9")
        p.text(ecols[0] + 4, y + 14, sector, 7, bold=True, color="#111111")
        y += 22
    rec.tables.append({"page": 1, "bbox": [25, y - 82, 570, y], "columns": 10, "rows": 4})

    # firmas
    p.fill(25, 773, 258, 0.8, "#111111")
    p.fill(312, 773, 258, 0.8, "#111111")
    p.text(154, 784, "FIRMA Y ACLARACIÓN CHOFER", 7, bold=True, color="#333333", align="c")
    p.text(441, 784, "FIRMA Y ACLARACIÓN AUXILIAR", 7, bold=True, color="#333333", align="c")
    p.next_page()

    # ---- página 2: recorrido
    _print_header(p, title, 2, 2)
    _hr_header(p, logo, "RECORRIDO", ("VIAJE N° 7141 · REPARTO R3 · 23/10/2026",), with_km=False)
    cursor = 25.0
    for bold, s in (
        (True, "22.6"),
        (False, " km de recorrido     "),
        (True, "0h 53m"),
        (False, " de manejo     "),
        (True, "7"),
        (False, " paradas en el mapa"),
    ):
        p.text(cursor, 92, s.strip(), 7, bold=bold, color=INK if bold else "#444444")
        cursor += Pen.text_width(s, 7, bold) + (3 if not s.endswith(" ") else 0)
    p.image(make_map(), 25, 99, 545, 338)
    p.stroke(25, 99, 545, 338, INK, 0.8, 4)
    rec.add_box(2, (25, 99, 570, 437), "image", "mapa", "")
    p.fill(25, 442, 545, 15, "#fff8e1")
    p.thin_border(25, 442, 545, 15, "#e6c200")
    p.text(
        30,
        452,
        "Sin ubicación cargada, no figuran en el mapa: 4 CL.07, 5 CL.11, 9 CL.31, 10 CL.22.",
        6.5,
        color="#92600a",
    )
    rec.add_box(2, (25, 442, 570, 457), "text", "", "")
    p.text(25, 470, "REFERENCIA DE PARADAS", 6.5, bold=True, color="#a3a8b4")
    p.stroke(25, 475, 545, 54, LINE, 0.7)
    rec.add_box(2, (25, 475, 570, 529), "text", "", "")
    names = [n for n, _, _ in _STOPS] + ["DEPÓSITO CENTRAL"]
    for k, name in enumerate(names):
        colx, start = next((cx, s) for cx, s in reversed(_REFS_COLS) if k >= s)
        yy = 488 + (k - start) * 12
        p.text(colx, yy, str(k + 1), 7, bold=True, color="#c0392b")
        p.text(colx + 14, yy, name, 7, color="#222222")
    c.save()
    return rec.ground_truth(
        "Hoja de ruta: encabezado con cajas, tabla de paradas con filas de sección, envases, firmas; "
        "2 textos N/A tapados; página 2 con mapa raster y referencias en 3 columnas.",
        expected_modes={"mode": "layout"},
        covered_spans=2,
    )


# --------------------------------------------------------------------------------------------- remito

_ITEMS = [
    ("A-1001", "Tornillo hexagonal 6mm x 40", 120, "UN", 12.5),
    ("A-1002", "Arandela plana 6mm", 300, "UN", 3.2),
    ("B-2010", "Cable unipolar 2,5mm rollo", 4, "RL", 18450.0),
    ("B-2044", "Interruptor termomagnético 2x20A", 6, "UN", 9850.75),
    ("C-3100", "Cinta aisladora negra", 24, "UN", 1250.0),
    ("C-3105", "Precinto plástico 300mm bolsa", 10, "BL", 2780.4),
    ("D-4001", "Pintura látex interior 20L", 3, "LT", 74500.0),
    ("D-4020", "Rodillo antigota 22cm", 8, "UN", 6320.0),
    ("E-5003", "Guantes de trabajo par", 20, "PR", 3150.0),
    ("E-5010", "Casco de seguridad blanco", 5, "UN", 11900.0),
    ("F-6002", "Bulón anclaje 3/8 x 80", 60, "UN", 480.0),
    ("F-6015", "Disco de corte 115mm", 40, "UN", 1890.0),
]


def _remito_draw(p: Pen) -> None:
    rec = p.rec
    p.stroke(30, 25, 255, 60, "#333333", 0.8)
    p.text(40, 45, "DISTRIBUIDORA LOS ARROYOS S.A.", 10, bold=True)
    p.text(40, 58, "Av. Siempreviva 742 · Rosario", 7.5, color="#444444")
    p.text(40, 69, "CUIT 30-71234567-8 · IVA Responsable Inscripto", 7.5, color="#444444")
    p.stroke(310, 25, 255, 60, "#333333", 0.8)
    p.text(320, 45, "REMITO", 15, bold=True)
    p.text(320, 62, "N° 0001-00004521", 10, bold=True)
    p.text(320, 77, "Documento no válido como factura", 7, color="#666666")
    rec.add_box(p.page, (30, 25, 285, 85), "text", "", "")
    rec.add_box(p.page, (310, 25, 565, 85), "text", "", "")
    row1 = [
        (30, 270, "CLIENTE", "FERRETERÍA EL MARTILLO S.R.L."),
        (305, 95, "CUIT", "30-70123456-7"),
        (405, 75, "FECHA", "12/03/2025"),
        (485, 80, "N° PEDIDO", "PD-00871"),
    ]
    row2 = [
        (30, 270, "DOMICILIO", "Mitre 1520, Funes"),
        (305, 135, "LOCALIDAD", "Funes, Santa Fe"),
        (445, 120, "CONDICIÓN DE VENTA", "Cuenta corriente"),
    ]
    for y, row in ((95, row1), (127, row2)):
        for x, w, label, value in row:
            p.label_box(x, y, w, 27, label, value, rounded=False)
    cols = [30, 90, 300, 350, 410, 490, 565]
    top = 165
    heads = ["CÓDIGO", "DESCRIPCIÓN", "CANT.", "UNIDAD", "PRECIO UNIT.", "IMPORTE"]
    p.fill(30, top, 535, 18, "#e5e7eb")
    rows = len(_ITEMS) + 1
    ys = [top + 18 * i for i in range(rows + 1)]
    p.c.setStrokeColor(colors.HexColor("#555555"))
    p.c.setLineWidth(0.5)
    p.c.grid(cols, [p.h - v for v in ys])
    for i, head in enumerate(heads):
        mid = (cols[i] + cols[i + 1]) / 2
        p.text(mid, top + 12, head, 7.5, bold=True, align="c")
    total = 0.0
    for r, (code, desc, qty, unit, price) in enumerate(_ITEMS):
        yy = ys[r + 1] + 12
        amount = qty * price
        total += amount
        p.text(cols[0] + 4, yy, code, 7.5)
        p.text(cols[1] + 4, yy, desc, 7.5)
        p.text(cols[3] - 5, yy, str(qty), 7.5, align="r")
        p.text((cols[3] + cols[4]) / 2, yy, unit, 7.5, align="c")
        p.text(cols[5] - 5, yy, _money(price), 7.5, align="r")
        p.text(cols[6] - 5, yy, _money(amount), 7.5, align="r")
    rec.tables.append({"page": p.page, "bbox": [30, top, 565, ys[-1]], "columns": 6, "rows": rows})
    y0 = ys[-1] + 12
    p.stroke(30, y0, 340, 60, "#555555", 0.5)
    p.text(36, y0 + 10, "OBSERVACIONES", 6, color=GRAY)
    p.text(36, y0 + 24, "Entregar en horario de 8 a 12 hs.", 7.5)
    rec.add_box(p.page, (30, y0, 370, y0 + 60), "text", "OBSERVACIONES", "Entregar en horario de 8 a 12 hs.")
    iva = total * 0.21
    for i, (label, amount, bold) in enumerate(
        (("Subtotal", total, False), ("IVA 21%", iva, False), ("TOTAL", total + iva, True))
    ):
        yy = y0 + i * 20
        p.stroke(385, yy, 180, 20, "#555555", 0.5)
        p.text(391, yy + 13, label, 8, bold=bold)
        p.text(560, yy + 13, _money(amount), 8, bold=bold, align="r")
    p.fill(30, 770, 200, 0.7, "#000000")
    p.fill(365, 770, 200, 0.7, "#000000")
    p.text(130, 782, "Firma y aclaración — Entrega", 7.5, align="c")
    p.text(465, 782, "Firma y aclaración — Recibe", 7.5, align="c")


def _remito_cajas(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path)
    _remito_draw(Pen(c, rec))
    c.save()
    return rec.ground_truth(
        "Remito con cajas de encabezado (cliente, CUIT, fecha, N°), tabla de ítems con bordes, totales a la "
        "derecha y firmas.",
        expected_modes={"mode": "layout"},
    )


def _escaneado_formulario(path: Path) -> dict[str, Any]:
    """El remito rasterizado a imagen (sin texto) para probar OCR."""
    buf = Path(str(path) + ".tmp.pdf")
    rec = Recorder(pages=1)
    c = _new_canvas(buf)
    _remito_draw(Pen(c, rec))
    c.save()
    doc = pymupdf.open(str(buf))
    pix = doc[0].get_pixmap(dpi=170, colorspace=pymupdf.csGRAY)
    img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    doc.close()
    buf.unlink()
    out = _new_canvas(path)
    out.drawImage(ImageReader(img), 0, 0, W_A4, H_A4)
    out.save()
    return rec.ground_truth(
        "El remito rasterizado a 170 dpi (sin capa de texto): prueba OCR + modo layout.",
        scanned=True,
        expected_modes={"mode": "tables"},
    )


# --------------------------------------------------------------------------------------------- factura


def _factura_bloques(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path)
    p = Pen(c, rec)
    p.stroke(30, 30, 535, 100, "#000000", 0.8)
    p.line(297.5, 80, 297.5, 130, "#000000", 0.6)
    p.stroke(275, 30, 45, 50, "#000000", 0.8)
    p.fill(276, 31, 43, 48, "#ffffff")
    p.text(297.5, 62, "A", 26, bold=True, align="c")
    p.text(297.5, 76, "COD. 01", 6, align="c")
    p.text(40, 52, "COMERCIAL DEL LITORAL S.A.", 11, bold=True)
    for i, s in enumerate(
        (
            "Razón social: Comercial del Litoral S.A.",
            "Domicilio: Bv. Oroño 1234 - Rosario",
            "Condición frente al IVA: Responsable Inscripto",
        )
    ):
        p.text(40, 72 + i * 12, s, 7.5)
    p.text(335, 52, "FACTURA", 16, bold=True)
    for i, s in enumerate(
        (
            "Punto de Venta: 0003    Comp. Nro: 00001234",
            "Fecha de Emisión: 15/04/2025",
            "CUIT: 30-71234560-1",
            "Inicio de Actividades: 01/03/2010",
        )
    ):
        p.text(335, 70 + i * 12, s, 7.5)
    rec.add_box(1, (30, 30, 297.5, 130), "text", "", "")
    rec.add_box(1, (297.5, 30, 565, 130), "text", "", "")
    rec.add_box(1, (275, 30, 320, 80), "text", "A", "COD. 01")  # el recuadro de la letra, que cruza la mitad
    p.stroke(30, 140, 535, 54, "#000000", 0.6)
    left = (
        "CUIT: 20-30456789-5",
        "Apellido y Nombre / Razón Social: Pereyra Hnos. S.R.L.",
        "Domicilio: San Martín 450",
    )
    right = (
        "Condición frente al IVA: Responsable Inscripto",
        "Condición de venta: Contado",
        "Localidad: Rosario",
    )
    for i, s in enumerate(left):
        p.text(38, 156 + i * 14, s, 7.5)
    for i, s in enumerate(right):
        p.text(320, 156 + i * 14, s, 7.5)
    rec.add_box(1, (30, 140, 297.5, 194), "text", "", "")
    rec.add_box(1, (297.5, 140, 565, 194), "text", "", "")
    cols = [30, 80, 285, 335, 400, 445, 510, 565]
    heads = ["Código", "Producto / Servicio", "Cantidad", "Precio Unit.", "% Bonif", "Alíc. IVA", "Subtotal"]
    top = 210
    p.fill(30, top, 535, 18, "#d9d9d9")
    for i, h in enumerate(heads):
        if i in (0, 1):
            p.text(cols[i] + 3, top + 12, h, 7.5, bold=True)
        else:
            p.text(cols[i + 1] - 4, top + 12, h, 7.5, bold=True, align="r")
    p.line(30, top + 18, 565, top + 18, "#000000", 0.8)
    rows = [
        ("S-001", "Servicio de mantenimiento mensual", 1, 45000.0),
        ("S-002", "Visita técnica adicional", 2, 12500.0),
        ("P-010", "Filtro de aire industrial", 4, 8900.0),
        ("P-014", "Correa de transmisión B-52", 3, 5600.0),
        ("P-020", "Lubricante sintético 5L", 6, 14300.0),
        ("S-010", "Capacitación de personal", 1, 30000.0),
        ("P-031", "Juego de juntas", 2, 7400.0),
        ("P-040", "Sensor de temperatura", 5, 9650.0),
    ]
    y = top + 18
    net = 0.0
    for code, desc, qty, price in rows:
        sub = qty * price
        net += sub
        base = y + 12
        p.text(cols[0] + 3, base, code, 7.5)
        p.text(cols[1] + 3, base, desc, 7.5)
        p.text(cols[3] - 4, base, f"{qty:.2f}".replace(".", ","), 7.5, align="r")
        p.text(cols[4] - 4, base, _money(price), 7.5, align="r")
        p.text(cols[5] - 4, base, "0,00", 7.5, align="r")
        p.text(cols[6] - 4, base, "21%", 7.5, align="r")
        p.text(cols[7] - 4, base, _money(sub), 7.5, align="r")
        y += 18
        p.line(30, y, 565, y, "#9a9a9a", 0.4)
    rec.tables.append(
        {"page": 1, "bbox": [30, top, 565, y], "columns": 7, "rows": len(rows) + 1, "rules": "horizontal"}
    )
    y += 16
    p.stroke(330, y, 235, 62, "#000000", 0.6)
    iva = net * 0.21
    for i, (label, value, bold) in enumerate(
        (
            ("Importe Neto Gravado: $", net, False),
            ("IVA 21%: $", iva, False),
            ("Importe Total: $", net + iva, True),
        )
    ):
        yy = y + 16 + i * 17
        p.text(338, yy, label, 8, bold=bold)
        p.text(558, yy, _money(value), 8, bold=bold, align="r")
    rec.add_box(1, (330, y, 565, y + 62), "text", "", "")
    p.text(30, 760, "CAE N°: 71234567890123", 8, bold=True)
    p.text(30, 772, "Fecha de Vto. de CAE: 25/04/2025", 8)
    p.text(565, 760, "Comprobante Autorizado", 8, bold=True, align="r")
    p.text(565, 772, "Pág. 1/1", 8, align="r")
    p.text(
        297.5,
        800,
        "Esta Administración Federal no se responsabiliza por los datos ingresados en el detalle de la operación",
        6.5,
        color="#555555",
        align="c",
    )
    c.save()
    return rec.ground_truth(
        "Factura con bloque emisor/receptor en dos columnas, tabla solo con líneas horizontales, bloque de "
        "impuestos y total, y pie con leyendas.",
        expected_modes={"mode": "layout"},
    )


# --------------------------------------------------------------------------------------------- form 2 columnas


def _checkbox(p: Pen, x: float, y: float, checked: bool, label: str) -> None:
    p.stroke(x, y, 9, 9, "#000000", 0.7)
    if checked:
        p.line(x + 1.5, y + 4.5, x + 3.8, y + 7.5, "#000000", 1)
        p.line(x + 3.8, y + 7.5, x + 8, y + 1.5, "#000000", 1)
    p.text(x + 14, y + 8, label, 8)
    p.rec.add_box(p.page, (x, y, x + 9, y + 9), "checkbox", label, "x" if checked else "")


def _form_dos_columnas(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path)
    p = Pen(c, rec)
    p.text(297.5, 50, "SOLICITUD DE ALTA DE CLIENTE", 14, bold=True, align="c")
    p.fill(40, 70, 515, 16, "#2f4b7c")
    p.text(46, 82, "DATOS PERSONALES", 9, bold=True, color="#ffffff")
    left = [
        ("Apellido y nombres:", "Quiroga, Marcela Inés"),
        ("Documento:", "DNI 30.123.456"),
        ("Fecha de nacimiento:", "14/08/1984"),
        ("Nacionalidad:", "Argentina"),
        ("Estado civil:", "Casada"),
        ("Profesión:", "Contadora"),
    ]
    right = [
        ("Domicilio:", "Belgrano 880, 3° B"),
        ("Localidad:", "Rosario"),
        ("Provincia:", "Santa Fe"),
        ("Código postal:", "S2000"),
        ("Teléfono:", "341-555-0142"),
        ("Correo:", "mquiroga@ejemplo.com"),
    ]
    for i, (label, value) in enumerate(left):
        y = 112 + i * 26
        p.text(40, y, label, 8, bold=True)
        p.line(40 + Pen.text_width(label, 8, True) + 4, y + 2, 285, y + 2, "#000000", 0.5)
        p.text(40 + Pen.text_width(label, 8, True) + 8, y - 1, value, 8, color="#1a3a8a")
    for i, (label, value) in enumerate(right):
        y = 112 + i * 26
        p.text(310, y, label, 8, bold=True)
        p.line(310 + Pen.text_width(label, 8, True) + 4, y + 2, 555, y + 2, "#000000", 0.5)
        p.text(310 + Pen.text_width(label, 8, True) + 8, y - 1, value, 8, color="#1a3a8a")
    p.fill(40, 284, 515, 16, "#fde9a8")
    p.text(46, 296, "TIPO DE CLIENTE", 9, bold=True)
    p.text(40, 322, "Tipo:", 8, bold=True)
    for k, (label, checked) in enumerate(
        (("Persona física", True), ("Empresa", False), ("Monotributista", False))
    ):
        _checkbox(p, 85 + k * 120, 314, checked, label)
    p.text(40, 346, "Medios de contacto:", 8, bold=True)
    for k, (label, checked) in enumerate(
        (("Correo", True), ("Teléfono", True), ("Carta", False), ("Mensaje", False))
    ):
        _checkbox(p, 135 + k * 100, 338, checked, label)
    p.fill(40, 374, 515, 16, "#2f4b7c")
    p.text(46, 386, "AUTORIZACIONES", 9, bold=True, color="#ffffff")
    for i, (label, checked) in enumerate(
        (
            ("Autorizo el tratamiento de mis datos personales.", True),
            ("Acepto recibir novedades comerciales.", False),
            ("Declaro que los datos consignados son verdaderos.", True),
        )
    ):
        _checkbox(p, 40, 402 + i * 20, checked, label)
    p.line(60, 520, 250, 520, "#000000", 0.6)
    p.line(345, 520, 535, 520, "#000000", 0.6)
    p.text(155, 532, "Firma del solicitante", 8, align="c")
    p.text(440, 532, "Aclaración y fecha", 8, align="c")
    c.save()
    return rec.ground_truth(
        "Formulario de campos 'etiqueta: valor' con línea de subrayado en dos columnas, casilleros con y sin "
        "tilde, secciones con fondo de color.",
        expected_modes={"mode": "layout"},
    )


# --------------------------------------------------------------------------------------------- tabla combinada


def _tabla_seccion_combinada(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path)
    p = Pen(c, rec)
    p.text(297.5, 50, "RESULTADO POR ZONA — PRIMER SEMESTRE 2025", 12, bold=True, align="c")
    cols = [40, 170, 240, 305, 370, 430, 495, 555]
    top = 80

    def cell(
        i0: int,
        i1: int,
        y: float,
        h: float,
        s: str,
        fill: str | None = None,
        bold: bool = False,
        align: str = "c",
    ) -> None:
        x0, x1 = cols[i0], cols[i1]
        if fill:
            p.fill(x0, y, x1 - x0, h, fill)
        p.stroke(x0, y, x1 - x0, h, "#444444", 0.6)
        tx = (x0 + x1) / 2 if align == "c" else (x1 - 4 if align == "r" else x0 + 4)
        if s:
            p.text(tx, y + h / 2 + 3, s, 8, bold=bold, align=align)

    cell(0, 1, top, 36, "CONCEPTO", "#d9e2f3", True)
    cell(1, 4, top, 18, "VENTAS", "#d9e2f3", True)
    cell(4, 6, top, 18, "COSTOS", "#d9e2f3", True)
    cell(6, 7, top, 36, "RESULTADO", "#d9e2f3", True)
    for i, h in enumerate(("Marzo", "Abril", "Mayo")):
        cell(1 + i, 2 + i, top + 18, 18, h, "#eaf0fa", True)
    for i, h in enumerate(("Materiales", "Mano de obra")):
        cell(4 + i, 5 + i, top + 18, 18, h, "#eaf0fa", True)
    y = top + 36
    zones = {
        "Zona Norte": ["Rosario Centro", "Rosario Sur", "Funes"],
        "Zona Sur": ["Villa Constitución", "San Nicolás", "Pérez"],
        "Zona Oeste": ["Casilda", "Roldán", "Granadero Baigorria"],
    }
    rnd = random.Random(21)
    grand = [0.0] * 6
    for zone, towns in zones.items():
        sums = [0.0] * 6
        for town in towns:
            vals = [rnd.randint(80, 900) * 100.0 for _ in range(5)]
            vals.append(sum(vals[:3]) - sum(vals[3:5]))
            cell(0, 1, y, 16, town, align="l")
            for k, v in enumerate(vals):
                cell(1 + k, 2 + k, y, 16, _money(v), align="r")
                sums[k] += v
            y += 16
        cell(0, 1, y, 16, f"Subtotal {zone}", "#e2efda", True, "l")
        for k, v in enumerate(sums):
            cell(1 + k, 2 + k, y, 16, _money(v), "#e2efda", True, "r")
            grand[k] += v
        y += 16
    cell(0, 1, y, 18, "TOTAL GENERAL", "#fce4d6", True, "l")
    for k, v in enumerate(grand):
        cell(1 + k, 2 + k, y, 18, _money(v), "#fce4d6", True, "r")
    y += 18
    rec.tables.append(
        {"page": 1, "bbox": [40, top, 555, y], "columns": 7, "rows": 14, "header_rows": 2, "merged": True}
    )
    c.save()
    return rec.ground_truth(
        "Tabla con encabezado de dos niveles (celdas combinadas) y filas de subtotal con fondo.",
        expected_modes={"mode": "layout"},
    )


# --------------------------------------------------------------------------------------------- página rotada


def _pagina_rotada(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path, (H_A4, W_A4))  # reportlab intercambia la MediaBox con /Rotate 90
    c.setPageRotation(90)
    c.transform(0, 1, -1, 0, W_A4, 0)  # dibujamos en un marco apaisado 842 x 595, origen arriba-izquierda
    p = Pen(c, rec, width=H_A4, height=W_A4)
    p.text(30, 40, "PLANILLA DE CONTROL DE ENTREGAS", 14, bold=True, color=INK)
    p.fill(30, 50, 782, 2, INK)
    p.label_box(30, 62, 250, 28, "CHOFER", "DIEGO BARRIOS", rounded=False)
    p.label_box(290, 62, 200, 28, "UNIDAD", "UT-14", rounded=False)
    p.label_box(500, 62, 150, 28, "FECHA", "03/06/2025", rounded=False)
    p.label_box(660, 62, 152, 28, "OBSERVACIONES", "", rounded=False)
    cols = [30, 70, 290, 400, 500, 600, 700, 812]
    heads = ["#", "CLIENTE", "REMITO", "BULTOS", "SALIDA", "LLEGADA", "RECIBIÓ"]
    top = 110
    p.fill(30, top, 782, 20, "#e5e7eb")
    ys = [top + 20 * i for i in range(10)]
    p.c.setStrokeColor(colors.HexColor("#555555"))
    p.c.setLineWidth(0.5)
    p.c.grid(cols, [p.h - v for v in ys])
    for i, head in enumerate(heads):
        p.text((cols[i] + cols[i + 1]) / 2, top + 13, head, 8, bold=True, align="c")
    clients = [
        "Panadería La Espiga",
        "Kiosco Central",
        "Verdulería Don Pepe",
        "Ferretería Rivera",
        "Librería Mitre",
        "Almacén Los Pinos",
        "Farmacia del Centro",
        "Carnicería El Gaucho",
    ]
    rnd = random.Random(3)
    for r, cl in enumerate(clients):
        yy = ys[r + 1] + 13
        p.text(cols[0] + 5, yy, str(r + 1), 8)
        p.text(cols[1] + 5, yy, cl, 8)
        p.text(cols[2] + 5, yy, f"0001-{4500 + r * 7:08d}", 8)
        p.text(cols[4] - 8, yy, str(rnd.randint(1, 30)), 8, align="r")
        p.text((cols[4] + cols[5]) / 2, yy, f"{8 + r}:{rnd.randint(10, 59)}", 8, align="c")
    rec.tables.append({"page": 1, "bbox": [30, top, 812, ys[-1]], "columns": 7, "rows": 9})
    p.fill(30, 400, 300, 0.7, "#000000")
    p.text(180, 412, "Firma del chofer", 8, align="c")
    c.save()
    return rec.ground_truth(
        "Página con /Rotate 90 y contenido tipo formulario (cajas + tabla).",
        rotate=90,
        expected_modes={"mode": "layout"},
    )


# --------------------------------------------------------------------------------------------- sin bordes


def _solo_texto_alineado(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=1)
    c = _new_canvas(path)
    p = Pen(c, rec)
    p.text(40, 50, "LISTADO DE EXISTENCIAS AL 31/03/2025", 12, bold=True)
    p.text(40, 64, "Depósito Central — Valuación a precio de costo", 8, color="#444444")
    xs = (40, 110, 330, 400, 480, 555)
    heads = ("CÓDIGO", "DESCRIPCIÓN", "UNID.", "CANTIDAD", "COSTO UNIT.", "VALOR")
    pos = ("l", "l", "l", "r", "r", "r")
    heads = ("CÓDIGO", "DESCRIPCIÓN", "UNID.", "CANTIDAD", "COSTO", "VALOR")
    xs = (40, 110, 340, 440, 500, 555)
    for x, h, a in zip(xs, heads, pos, strict=True):
        p.text(x, 92, h, 8, bold=True, align=a)
    rnd = random.Random(9)
    names = [
        "Tornillo autorroscante",
        "Tuerca hexagonal",
        "Perfil de aluminio 2m",
        "Chapa lisa 1x2",
        "Caño estructural",
        "Pintura esmalte 4L",
        "Lija al agua 220",
        "Cemento de contacto",
        "Masilla plástica",
        "Bisagra reforzada",
        "Cerradura de embutir",
        "Candado 40mm",
        "Lámpara LED 12W",
        "Cable taller 3x1,5",
        "Tomacorriente doble",
        "Llave térmica 25A",
        "Cinta métrica 5m",
        "Destornillador Phillips",
        "Pinza universal",
        "Martillo 500g",
        "Taladro percutor",
        "Mecha widia 8mm",
        "Disco de corte",
        "Guante de cuero",
        "Anteojo de seguridad",
    ]
    total = 0.0
    for i, name in enumerate(names):
        y = 110 + i * 14
        qty = rnd.randint(5, 400)
        cost = rnd.randint(300, 90000) / 10
        total += qty * cost
        for x, s, a in zip(
            xs,
            (
                f"AR-{1000 + i * 7:04d}",
                name,
                rnd.choice(["UN", "KG", "MT"]),
                str(qty),
                _money(cost),
                _money(qty * cost),
            ),
            pos,
            strict=True,
        ):
            p.text(x, y, s, 8, align=a)
    p.text(440, 110 + len(names) * 14 + 12, "VALOR TOTAL", 8, bold=True, align="r")
    p.text(555, 110 + len(names) * 14 + 12, _money(total), 8, bold=True, align="r")
    c.save()
    return rec.ground_truth(
        "Sin ningún borde dibujado: texto en columnas alineadas, como un listado impreso.",
        expected_modes={"mode": "tables"},
    )


# --------------------------------------------------------------------------------------------- listado


def _multipagina_listado(path: Path) -> dict[str, Any]:
    rec = Recorder(pages=3)
    rnd = random.Random(17)
    heads = ["Fecha", "Comprobante", "Concepto", "Debe", "Haber", "Saldo"]
    rows: list[list[str]] = []
    saldo = 100000.0
    for i in range(118):
        debe = rnd.choice([0.0, rnd.randint(100, 9000) * 10.0])
        haber = 0.0 if debe else rnd.randint(100, 9000) * 10.0
        saldo += haber - debe
        rows.append(
            [
                f"{1 + i // 5:02d}/04/2025",
                f"FA-{5000 + i}",
                rnd.choice(["Venta mostrador", "Cobranza", "Pago proveedor", "Ajuste"]),
                _money(debe) if debe else "",
                _money(haber) if haber else "",
                _money(saldo),
            ]
        )
    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        invariant=1,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40,
        title=path.stem,
    )
    table = Table([heads, *rows], colWidths=[70, 80, 140, 75, 75, 75], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dddddd")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
            ]
        )
    )
    doc.build([table])
    n_pages = len(pymupdf.open(str(path)))
    visible = [w for r in (heads, *rows) for cell in r for w in cell.split()]
    # el encabezado se repite en cada página
    visible += [w for _ in range(n_pages - 1) for h in heads for w in h.split()]
    rec.visible = visible
    rec.pages = n_pages
    rec.tables.append({"page": 1, "columns": 6, "rows": len(rows) + 1, "repeat_header": True})
    return rec.ground_truth(
        "Listado de 3 páginas con encabezado repetido: debe seguir yendo a modo 'solo tablas'.",
        expected_modes={"mode": "tables"},
    )


# --------------------------------------------------------------------------------------------- API

_BUILDERS: dict[str, Callable[[Path], dict[str, Any]]] = {
    "hoja_ruta": _hoja_ruta,
    "remito_cajas": _remito_cajas,
    "factura_bloques": _factura_bloques,
    "form_dos_columnas": _form_dos_columnas,
    "tabla_seccion_combinada": _tabla_seccion_combinada,
    "pagina_rotada": _pagina_rotada,
    "solo_texto_alineado": _solo_texto_alineado,
    "multipagina_listado": _multipagina_listado,
    "escaneado_formulario": _escaneado_formulario,
}


def generate_forms(out_dir: Path) -> dict[str, Path]:
    """Genera todos los PDFs de formularios en ``out_dir`` y ``ground_truth_forms.json``.

    Devuelve ``nombre -> ruta del PDF``. Determinista: dos corridas dan los mismos bytes.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    result: dict[str, Path] = {}
    truth: dict[str, Any] = {}
    for name, builder in _BUILDERS.items():
        pdf = out_dir / f"{name}.pdf"
        truth[f"{name}.pdf"] = builder(pdf)
        result[name] = pdf
    (out_dir / GROUND_TRUTH_NAME).write_text(
        json.dumps(truth, ensure_ascii=False, indent=1, sort_keys=True) + "\n", encoding="utf-8"
    )
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m tests.corpus.generate_forms",
        description="Genera el corpus sintético de formularios (determinista).",
    )
    parser.add_argument("out_dir", type=Path, help="carpeta de salida (se crea si no existe)")
    args = parser.parse_args(argv)
    paths = generate_forms(args.out_dir)
    for name, path in paths.items():
        print(f"{path.stat().st_size:>9} B  {path}")
    print(f"{len(paths)} PDFs en {args.out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
