"""Writer del modo "Igual al PDF" (`excel/layout_writer.py`): páginas armadas a mano, sin extractor."""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from datetime import date
from io import BytesIO
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText
from PIL import Image

from pdf2excel.excel.layout_writer import (
    cell_font_scale,
    col_width_chars,
    row_height_pt,
    write_layout_sheets,
)
from pdf2excel.models import (
    BorderSide,
    ConversionReport,
    LayoutCell,
    LayoutImage,
    LayoutPage,
    Options,
    TextRun,
)

THIN = BorderSide(0.6, "444444")
MED = BorderSide(1.8, "000000")
THICK = BorderSide(3.0, "1F3864")


def _png(w: int = 60, h: int = 30, color: tuple[int, int, int] = (200, 60, 30)) -> bytes:
    buf = BytesIO()
    Image.new("RGB", (w, h), color).save(buf, format="PNG")
    return buf.getvalue()


def _box(
    row: int, col: int, label: str, value: str = "", *, rs: int = 1, cs: int = 1, fill: str | None = None
) -> LayoutCell:
    lines = [[TextRun(label, size_pt=6.5, color="666666")]]
    if value:
        lines.append([TextRun(value, bold=True, size_pt=9)])
    return LayoutCell(
        row=row,
        col=col,
        row_span=rs,
        col_span=cs,
        lines=lines,
        halign="left",
        valign="center",
        fill=fill,
        border_left=THIN,
        border_top=THIN,
        border_right=THIN,
        border_bottom=THIN,
    )


def make_form_page(page: int = 0) -> LayoutPage:
    """Réplica chica (ficticia) de una hoja de ruta: cajas etiqueta+valor, tabla con encabezado de color,
    fila de sección combinada, firma e imagen."""
    cols = [70.0, 70.0, 70.0, 70.0, 70.0, 70.0, 70.0, 70.0]
    rows = [34.0, 24.0, 24.0, 6.0, 18.0, 16.0, 16.0, 16.0, 16.0, 16.0, 40.0, 30.0]
    cells: list[LayoutCell] = [
        LayoutCell(
            row=0,
            col=2,
            col_span=4,
            lines=[[TextRun("HOJA DE RUTA", bold=True, size_pt=16)]],
            halign="center",
        ),
        _box(0, 6, "VIAJE", "000123", cs=1),
        _box(0, 7, "KM", "", cs=1),
        _box(1, 0, "TRANSPORTISTA", "Juan Pérez S.A.", cs=2),
        _box(1, 2, "AUXILIAR", "Ana Gómez", cs=2),
        _box(1, 4, "VEHÍCULO", "", cs=2),
        _box(1, 6, "PATENTE", "AB123CD", cs=2),
        _box(2, 0, "REPARTO", "R-77", cs=2),
        _box(2, 2, "TEMPERATURA", "", cs=2),
        _box(2, 4, "FECHA", "", cs=4),
    ]
    heads = ["N°", "Cliente", "Dirección", "Bultos", "Importe", "Fecha", "Hora", "Estado"]
    for c, h in enumerate(heads):
        cells.append(
            LayoutCell(
                row=4,
                col=c,
                lines=[[TextRun(h, bold=True, size_pt=8, color="FFFFFF")]],
                halign="center",
                fill="1F3864",
                border_left=THIN,
                border_top=THIN,
                border_right=THIN,
                border_bottom=THIN,
            )
        )
    cells.append(
        LayoutCell(
            row=5,
            col=0,
            col_span=8,
            lines=[[TextRun("ZONA NORTE", bold=True, size_pt=8)]],
            fill="F4B183",
            border_left=THIN,
            border_top=THIN,
            border_right=THIN,
            border_bottom=THIN,
        )
    )
    data = [
        (
            "1",
            "Distribuidora Sur",
            "Av. Siempreviva 742",
            12,
            1234.5,
            date(2025, 3, 14),
            "09:30",
            "Entregado",
        ),
        ("2", "Kiosco Uno", "Calle Falsa 123", 3, 99.9, date(2025, 3, 15), "10:15", "Pendiente"),
        ("3", "Mercado Central", "Ruta 5 km 12", 40, 15000.0, date(2025, 3, 16), "11:00", "Pendiente"),
    ]
    for i, rowdata in enumerate(data):
        r = 6 + i
        for c, v in enumerate(rowdata):
            is_num = isinstance(v, (int, float))
            is_date = isinstance(v, date)
            cells.append(
                LayoutCell(
                    row=r,
                    col=c,
                    lines=[[TextRun(str(v), size_pt=8)]],
                    halign="right" if is_num else "left",
                    value=v if (is_num or is_date) else None,
                    number_format=(
                        "#,##0.00" if isinstance(v, float) else "0" if isinstance(v, int) else "dd/mm/yyyy"
                    )
                    if (is_num or is_date)
                    else None,
                    border_left=THIN,
                    border_top=THIN,
                    border_right=THIN,
                    border_bottom=THIN,
                )
            )
    cells.append(_box(10, 0, "OBSERVACIONES", "", rs=1, cs=5))
    cells.append(
        LayoutCell(
            row=11,
            col=5,
            col_span=3,
            lines=[[TextRun("Firma del conductor", size_pt=7, italic=True)]],
            halign="center",
            valign="bottom",
            border_top=MED,
        )
    )
    img = LayoutImage(
        data=_png(),
        format="png",
        bbox=(5.0, 3.0, 65.0, 33.0),
        row=0,
        col=0,
        dx_pt=5,
        dy_pt=3,
        width_pt=60,
        height_pt=30,
    )
    return LayoutPage(
        page=page,
        width_pt=595.0,
        height_pt=842.0,
        col_widths_pt=cols,
        row_heights_pt=rows,
        cells=cells,
        images=[img],
        n_words=80,
    )


def _report(*pages: LayoutPage) -> ConversionReport:
    return ConversionReport(
        source=Path("x.pdf"), pages=[], tables=[], text_items=[], layout_pages=list(pages)
    )


def _build(*pages: LayoutPage, prefix: str = "", opts: Options | None = None) -> tuple[Workbook, list[str]]:
    wb = Workbook()
    default = wb.active
    assert default is not None
    wb.remove(default)
    names = write_layout_sheets(wb, _report(*pages), opts or Options(), set(), prefix)
    return wb, names


def _save(wb: Workbook, path: Path) -> Path:
    wb.save(path)
    return path


def _xml_texts(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return "".join(z.read(n).decode("utf-8", "replace") for n in z.namelist() if n.endswith(".xml"))


def _merged_ok(ws) -> None:
    seen: set[tuple[int, int]] = set()
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                assert (r, c) not in seen, f"solapado en {rng}"
                seen.add((r, c))


# ------------------------------------------------------------------------------------------ tests


def test_basic_sheet_structure(tmp_path: Path) -> None:
    wb, names = _build(make_form_page(0), make_form_page(1))
    assert names == ["Página 1", "Página 2"]
    ws = wb["Página 1"]
    assert ws.sheet_view.showGridLines is False
    assert ws.page_setup.paperSize == 9
    assert ws.page_setup.orientation == "portrait"
    assert ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 0
    assert ws.sheet_properties.pageSetUpPr.fitToPage is True
    assert ws.print_area is not None
    assert ws.freeze_panes is None and ws.auto_filter.ref is None
    assert len(ws._images) == 1

    wb2 = load_workbook(_save(wb, tmp_path / "a.xlsx"))
    ws2 = wb2["Página 1"]
    _merged_ok(ws2)
    assert "A2:B2" in {str(r) for r in ws2.merged_cells.ranges}
    assert ws2["A2"].value == "TRANSPORTISTA\nJuan Pérez S.A."


def test_merged_range_styles_on_every_cell(tmp_path: Path) -> None:
    wb, _ = _build(make_form_page())
    ws = wb["Página 1"]  # openpyxl descarta los estilos de celdas combinadas al RELEER: se mira en memoria
    assert "A6:H6" in {str(r) for r in ws.merged_cells.ranges}
    for c in range(1, 9):
        cell = ws.cell(row=6, column=c)
        assert str(cell.fill.fgColor.rgb).endswith("F4B183")
        assert cell.border.top.style == "thin"
        assert cell.border.bottom.style == "thin"
    assert ws["A6"].border.left.style == "thin"
    assert ws["H6"].border.right.style == "thin"
    assert ws["E11"].border.right.style == "thin" and ws["A11"].border.left.style == "thin"
    # y el archivo guardado trae el estilo de cada celda del rango (atributo s distinto de 0)
    path = _save(wb, tmp_path / "a.xlsx")
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode()
    row6 = re.search(r'<row r="6".*?</row>', xml, re.S)
    assert row6 is not None
    styles = re.findall(r'<c r="[A-H]6" s="(\d+)"', row6.group(0))
    assert len(styles) == 8 and all(int(x) > 0 for x in styles)


def test_typed_values_and_formats(tmp_path: Path) -> None:
    wb, _ = _build(make_form_page())
    ws = load_workbook(_save(wb, tmp_path / "a.xlsx"))["Página 1"]
    assert ws["E7"].value == 1234.5 and ws["E7"].number_format == "#,##0.00"
    assert ws["D7"].value == 12 and ws["D7"].number_format == "0"
    assert ws["F7"].value.date() == date(2025, 3, 14)
    assert ws["F7"].number_format == "dd/mm/yyyy"
    assert ws["A7"].value == "1" and ws["A7"].data_type == "s"  # el N° queda como texto


def test_rich_text_valid_and_single_style(tmp_path: Path) -> None:
    wb, _ = _build(make_form_page())
    path = _save(wb, tmp_path / "a.xlsx")
    ws = load_workbook(path, rich_text=True)["Página 1"]
    rich = ws["A2"].value
    assert isinstance(rich, CellRichText)
    assert "".join(b.text if hasattr(b, "text") else str(b) for b in rich) == "TRANSPORTISTA\nJuan Pérez S.A."
    # un solo estilo: no es rich text
    assert not isinstance(ws["A5"].value, CellRichText)
    assert ws["A5"].font.b is True and ws["A5"].font.color.rgb.endswith("FFFFFF")


def test_no_control_chars_in_xml(tmp_path: Path) -> None:
    page = make_form_page()
    page.cells.append(
        LayoutCell(
            row=9,
            col=6,
            col_span=2,
            lines=[[TextRun("Mal\x00o\x0b", size_pt=8), TextRun("\x01x", bold=True, size_pt=8)]],
        )
    )
    wb, _ = _build(page)
    path = _save(wb, tmp_path / "a.xlsx")
    assert not re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", _xml_texts(path))
    ws = load_workbook(path)["Página 1"]
    assert "Malo" in str(ws["G10"].value)


def test_formula_like_text_stays_text(tmp_path: Path) -> None:
    page = make_form_page()
    page.cells.append(LayoutCell(row=9, col=6, lines=[[TextRun("=SUMA(A1)", size_pt=8)]]))
    wb, _ = _build(page)
    ws = load_workbook(_save(wb, tmp_path / "a.xlsx"))["Página 1"]
    assert ws["G10"].data_type == "s" and ws["G10"].value == "=SUMA(A1)"


def test_empty_page_and_no_cells() -> None:
    empty = LayoutPage(page=0, width_pt=595, height_pt=842, col_widths_pt=[], row_heights_pt=[])
    no_cells = LayoutPage(page=1, width_pt=842, height_pt=595, col_widths_pt=[100.0], row_heights_pt=[50.0])
    wb, names = _build(empty, no_cells)
    assert names == ["Página 1", "Página 2"]
    assert wb["Página 2"].page_setup.orientation == "landscape"


def test_names_prefix_unique_and_short() -> None:
    wb = Workbook()
    used: set[str] = set()
    rep = _report(make_form_page(0), make_form_page(0))
    names = write_layout_sheets(wb, rep, Options(), used, "Remito muy largo de proveedor importante")
    assert len(set(n.casefold() for n in names)) == 2
    assert all(len(n) <= 31 for n in names)
    assert all(n in used for n in names)
    assert all(n.endswith("Página 1") for n in names)
    more = write_layout_sheets(wb, _report(make_form_page(0)), Options(), used)
    assert more[0] not in names and len(more[0]) <= 31


def test_limits_and_dimensions() -> None:
    assert col_width_chars(0.1) == 0.2
    assert col_width_chars(100000) == 255
    assert row_height_pt(5000) == 409
    page = LayoutPage(
        page=0,
        width_pt=595,
        height_pt=842,
        col_widths_pt=[30000.0, 0.0, 40.0],
        row_heights_pt=[9000.0, 20.0],
        cells=[
            LayoutCell(row=0, col=0, col_span=3, row_span=2, lines=[[TextRun("x")]]),
            LayoutCell(row=50, col=99, lines=[[TextRun("fuera")]]),
        ],
    )
    wb, _ = _build(page)
    ws = wb["Página 1"]
    assert ws.column_dimensions["A"].width <= 255
    assert ws.row_dimensions[1].height <= 409
    assert ws.max_row == 2 and ws.max_column == 3


def test_overlapping_cells_do_not_overlap_merges(tmp_path: Path) -> None:
    page = LayoutPage(
        page=0,
        width_pt=300,
        height_pt=300,
        col_widths_pt=[60.0] * 4,
        row_heights_pt=[20.0] * 4,
        cells=[
            LayoutCell(row=0, col=0, col_span=3, lines=[[TextRun("a")]]),
            LayoutCell(row=0, col=1, col_span=3, lines=[[TextRun("b")]]),
            LayoutCell(row=0, col=1, lines=[[TextRun("c")]]),
            LayoutCell(row=1, col=0, row_span=2, col_span=2, lines=[[TextRun("d")]]),
            LayoutCell(row=2, col=1, col_span=2, lines=[[TextRun("e")]]),
        ],
    )
    wb, _ = _build(page)
    ws = load_workbook(_save(wb, tmp_path / "a.xlsx"))["Página 1"]
    _merged_ok(ws)
    assert ws["A1"].value == "a"


def test_images_respect_option_and_not_duplicated(tmp_path: Path) -> None:
    page = make_form_page()
    wb, _ = _build(page, opts=Options(layout_images=False))
    assert len(wb["Página 1"]._images) == 0
    wb, _ = _build(page)
    path = _save(wb, tmp_path / "a.xlsx")
    with zipfile.ZipFile(path) as z:
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        assert len(media) == 1
        assert z.read(media[0]) == page.images[0].data  # no se re-codificó
        drawing = z.read("xl/drawings/drawing1.xml").decode()
    assert "twoCellAnchor" in drawing


def test_columns_and_rows_keep_exact_pdf_proportions() -> None:
    """Sin agrandar la grilla: como la hoja se imprime ajustada al ancho, agrandarla achicaba el texto."""
    page = LayoutPage(
        page=0,
        width_pt=300,
        height_pt=300,
        col_widths_pt=[21.0, 52.5, 105.0],
        row_heights_pt=[30.0, 12.0],
        cells=[LayoutCell(row=0, col=0, col_span=3, lines=[[TextRun("Observación", size_pt=9)]])],
    )
    wb, _ = _build(page)
    ws = wb["Página 1"]
    # 1 unidad de ancho = 5.25 pt (el atributo ya incluye el relleno de 5 px)
    assert [ws.column_dimensions[c].width for c in "ABC"] == [4.0, 10.0, 20.0]
    assert ws.row_dimensions[1].height == 30.0 and ws.row_dimensions[2].height == 12.0


def test_font_shrinks_only_in_the_cell_that_needs_it() -> None:
    long = "Observación muy larga que no entra en una sola línea de la celda " * 3
    tight = LayoutCell(row=0, col=0, col_span=2, lines=[[TextRun(long, size_pt=9)]], wrap=True)
    roomy = LayoutCell(row=1, col=0, lines=[[TextRun("Ok", size_pt=9)]])
    assert cell_font_scale(tight, 300.0, 30.0) < 1.0
    assert cell_font_scale(roomy, 300.0, 30.0) == 1.0
    page = LayoutPage(
        page=0,
        width_pt=300,
        height_pt=300,
        col_widths_pt=[150.0, 150.0],
        row_heights_pt=[30.0, 20.0],
        cells=[tight, roomy],
    )
    wb, _ = _build(page)
    ws = wb["Página 1"]
    assert ws["A1"].alignment.wrap_text is True
    assert ws["A1"].font.sz < 9  # la celda apretada achica su letra...
    assert ws["A2"].font.sz == 9  # ...y las demás conservan el tamaño del PDF


def test_text_that_fits_in_the_pdf_is_not_shrunk() -> None:
    """Un texto que entra con las métricas de Helvetica no se toca (medido, no estimado a ojo)."""
    cell = LayoutCell(row=0, col=0, lines=[[TextRun("HOJA DE RUTA", bold=True, size_pt=16, font="Arial")]])
    import pymupdf

    needed = pymupdf.Font("hebo").text_length("HOJA DE RUTA", fontsize=16)  # lo que mide en el PDF
    assert cell_font_scale(cell, needed + 6.0, 22.0) == 1.0
    assert cell_font_scale(cell, needed * 0.8, 22.0) < 1.0


def test_image_is_anchored_by_two_corners_to_the_grid() -> None:
    page = LayoutPage(
        page=0,
        width_pt=300,
        height_pt=300,
        col_widths_pt=[50.0, 50.0, 50.0, 50.0],
        row_heights_pt=[21.0, 21.0, 21.0, 21.0],  # 28 px exactos: Excel no redondea el alto de fila
        cells=[],
        images=[
            LayoutImage(
                data=_png(),
                bbox=(60.0, 26.0, 160.0, 68.0),
                row=1,
                col=1,
                dx_pt=10.0,
                dy_pt=5.0,
                width_pt=100.0,
                height_pt=42.0,
                format="png",
            )
        ],
    )
    wb, _ = _build(page)
    anchor = wb["Página 1"]._images[0].anchor
    assert (anchor._from.col, anchor._from.row) == (1, 1)
    assert (anchor.to.col, anchor.to.row) == (3, 3)  # termina donde termina la imagen, no donde "alcanza"
    assert anchor._from.colOff == 10 * 12700 and anchor.to.colOff == 10 * 12700
    assert anchor._from.rowOff == 5 * 12700 and anchor.to.rowOff == 5 * 12700


def test_perf_40_pages(tmp_path: Path) -> None:
    pages = [make_form_page(i) for i in range(40)]
    t0 = time.perf_counter()
    wb, names = _build(*pages)
    path = _save(wb, tmp_path / "big.xlsx")
    elapsed = time.perf_counter() - t0
    assert len(names) == 40 and elapsed < 5.0
    assert path.stat().st_size < 3_000_000


# ------------------------------------------------------------------------------- LibreOffice


def _soffice() -> str:
    exe = shutil.which("soffice")
    if exe is None:
        pytest.skip("LibreOffice no está instalado")
    return exe


def render_pdf(xlsx: Path, out_dir: Path) -> tuple[Path, str]:
    """Convierte con LibreOffice (perfil temporal propio). Devuelve el PDF y el stderr."""
    exe = _soffice()
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    try:
        proc = subprocess.run(
            [
                exe,
                f"-env:UserInstallation=file://{profile}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(xlsx),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            env={"HOME": str(profile), "PATH": "/usr/bin:/bin"},
        )
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    pdf = out_dir / (xlsx.stem + ".pdf")
    assert pdf.exists(), proc.stderr
    return pdf, proc.stderr


@pytest.mark.slow
def test_libreoffice_opens_and_fits_width(tmp_path: Path) -> None:
    import pymupdf

    wb, _ = _build(make_form_page(0), make_form_page(1))
    path = _save(wb, tmp_path / "form.xlsx")
    pdf, err = render_pdf(path, tmp_path)
    assert "error" not in err.lower()
    with pymupdf.open(str(pdf)) as doc:
        assert len(doc) == 2  # A4 ajustado al ancho: una página por hoja
        text = doc[0].get_text()
        for word in (
            "HOJA DE RUTA",
            "TRANSPORTISTA",
            "Distribuidora Sur",
            "ZONA NORTE",
            "Firma del conductor",
        ):
            assert word in text
        assert len(doc[0].get_images()) + len(doc[0].get_image_info()) >= 1
        assert abs(doc[0].rect.width - 595) < 2


# ------------------------------------------------------------------ filas justas de alto de letra (R2)


def _tight_page(extra: list[LayoutCell] | None = None, heights: list[float] | None = None) -> LayoutPage:
    """Una fila de 8 pt para una letra de 8 pt entre dos pistas de relleno (como arma el extractor)."""
    text = LayoutCell(row=1, col=0, lines=[[TextRun("Subtotal", size_pt=8)]])
    return LayoutPage(
        page=0,
        width_pt=200,
        height_pt=100,
        col_widths_pt=[80.0, 60.0],
        row_heights_pt=heights or [6.6, 8.0, 5.4],
        cells=[text, *(extra or [])],
    )


def test_tight_text_row_borrows_height_from_empty_neighbours() -> None:
    from pdf2excel.excel.layout_writer import balanced_row_heights

    page = _tight_page()
    heights = balanced_row_heights(page, 3, 2)
    assert abs(sum(heights) - 20.0) < 0.01  # el total no cambia: los bordes exteriores quedan donde estaban
    assert heights[1] > 8.0 + 1.0
    wb, _ = _build(page)
    ws = wb["Página 1"]
    assert ws["A2"].font.sz == 8  # la letra conserva el tamaño del PDF (antes caía a 6 pt)


def test_underlined_value_grows_upwards_and_stays_bottom_aligned() -> None:
    from pdf2excel.excel.layout_writer import balanced_row_heights

    value = LayoutCell(
        row=1, col=1, lines=[[TextRun("Pérez, Ana", size_pt=8)]], border_bottom=BorderSide(0.6, "444444")
    )
    page = _tight_page([value], heights=[12.0, 8.0, 5.0])
    heights = balanced_row_heights(page, 3, 2)
    assert heights[2] == 5.0  # el renglón de abajo (donde corre la línea) no se toca
    assert heights[0] < 12.0 and heights[1] > 8.0
    wb, _ = _build(page)
    assert wb["Página 1"]["B2"].alignment.vertical == "bottom"


def test_row_height_rebalance_never_takes_height_from_a_row_with_text() -> None:
    from pdf2excel.excel.layout_writer import balanced_row_heights

    neighbour = LayoutCell(row=2, col=0, lines=[[TextRun("Otro renglón", size_pt=8)]])
    page = _tight_page([neighbour], heights=[1.0, 8.0, 11.6])
    heights = balanced_row_heights(page, 3, 2)
    assert cell_font_scale(neighbour, 60.0, heights[2]) == 1.0  # la fila vecina conserva el alto que necesita


# ------------------------------------------------------------------ ronda R3: alto en píxeles, derrame, OCR


def test_quantize_row_heights_are_whole_pixels_and_keep_position() -> None:
    from pdf2excel.excel.layout_writer import PX_PT, quantize_row_heights

    heights = [20.0, 14.0, 5.0, 6.02, 2.6, 22.0, 31.2, 8.0, 17.8] * 5
    out = quantize_row_heights(heights)
    assert all(abs(h / PX_PT - round(h / PX_PT)) < 1e-9 and h >= PX_PT for h in out)
    cum_in = cum_out = 0.0
    for a, b in zip(heights, out, strict=True):
        cum_in += a
        cum_out += b
        assert abs(cum_in - cum_out) <= PX_PT  # el error no se acumula: a lo sumo un píxel
    # un mínimo (lo que necesita el texto) se respeta y las filas siguientes lo compensan
    out2 = quantize_row_heights([10.9, 10.9, 10.9, 10.9], [10.95] * 4)
    assert all(h >= 10.95 for h in out2)


def test_written_row_heights_are_multiples_of_a_pixel() -> None:
    wb, _ = _build(make_form_page())
    ws = wb["Página 1"]
    for dim in ws.row_dimensions.values():
        if dim.height is not None:
            assert abs(dim.height / 0.75 - round(dim.height / 0.75)) < 1e-9


def _checkbox_row_page(spacer: float = 14.0) -> LayoutPage:
    """Renglón de 8.8 pt de una casilla con borde arriba y abajo y un texto de 8 pt al lado (como arma el
    extractor un formulario): las pistas vecinas tienen líneas, así que hay que pedir prestado más lejos."""
    box = LayoutCell(
        row=2,
        col=0,
        lines=[[TextRun("✔", size_pt=8)]],
        halign="center",
        border_left=THIN,
        border_top=THIN,
        border_right=THIN,
        border_bottom=THIN,
    )
    label = LayoutCell(row=2, col=2, col_span=2, lines=[[TextRun("Persona física", size_pt=8)]])
    return LayoutPage(
        page=0,
        width_pt=200,
        height_pt=100,
        col_widths_pt=[12.0, 6.0, 40.0, 40.0],
        row_heights_pt=[spacer, 3.0, 8.8, 3.0, spacer],
        cells=[box, label],
    )


def test_row_with_lines_on_both_sides_borrows_height_from_a_distant_empty_row() -> None:
    from pdf2excel.excel.layout_writer import balanced_row_heights

    page = _checkbox_row_page()
    heights = balanced_row_heights(page, 5, 4)
    assert heights[2] >= 11.2  # entra una letra de 8 pt sin achicarla (con el interlineado de Excel)
    assert abs(sum(heights) - sum(page.row_heights_pt)) < 0.01  # el total no cambia
    wb, _ = _build(page)
    assert wb["Página 1"]["C3"].font.sz == 8


def test_far_borrowing_never_takes_height_from_a_row_with_text() -> None:
    from pdf2excel.excel.layout_writer import balanced_row_heights

    page = _checkbox_row_page(spacer=9.0)
    page.cells.append(LayoutCell(row=0, col=2, lines=[[TextRun("Título", size_pt=8)]]))
    page.cells.append(LayoutCell(row=4, col=2, lines=[[TextRun("Pie", size_pt=8)]]))
    heights = balanced_row_heights(page, 5, 4)
    assert heights[0] >= 9.0 and heights[4] >= 9.0  # los renglones con texto nunca prestan alto
    assert abs(sum(heights) - sum(page.row_heights_pt)) < 0.01


def test_single_line_text_spills_over_free_tracks_instead_of_shrinking() -> None:
    from pdf2excel.excel.layout_writer import _occupancy, overflow_layout

    # "1/2" pegado al borde derecho de la página: no entra en su pista pero hay lugar a la izquierda
    page = LayoutPage(
        page=0,
        width_pt=200,
        height_pt=50,
        col_widths_pt=[60.0, 60.0, 12.0],
        row_heights_pt=[12.0],
        cells=[LayoutCell(row=0, col=2, lines=[[TextRun("1/2", size_pt=8)]], halign="left", wrap=False)],
    )
    cell = page.cells[0]
    xs = [0.0, 60.0, 120.0, 132.0]
    room, align = overflow_layout(cell, (0, 2, 0, 2), _occupancy(page, 1, 3), xs)
    assert room == 120.0 and align == "right"
    assert cell_font_scale(cell, 12.0, 12.0) < 1.0
    assert cell_font_scale(cell, 12.0, 12.0, room) == 1.0
    wb, _ = _build(page)
    c = wb["Página 1"]["C1"]
    assert c.font.sz == 8 and c.alignment.horizontal == "right"


def test_text_does_not_spill_over_an_occupied_neighbour_or_its_own_border() -> None:
    from pdf2excel.excel.layout_writer import _occupancy, overflow_layout

    xs = [0.0, 30.0, 60.0, 90.0]
    neighbour = LayoutCell(row=0, col=1, lines=[[TextRun("otro", size_pt=8)]])
    text = LayoutCell(row=0, col=0, lines=[[TextRun("Un texto largo", size_pt=8)]], wrap=False)
    page = LayoutPage(
        page=0,
        width_pt=90,
        height_pt=20,
        col_widths_pt=[30.0] * 3,
        row_heights_pt=[12.0],
        cells=[text, neighbour],
    )
    assert overflow_layout(text, (0, 0, 0, 0), _occupancy(page, 1, 3), xs)[0] == 0.0
    boxed = LayoutCell(
        row=0, col=0, lines=[[TextRun("Un texto largo", size_pt=8)]], border_right=THIN, wrap=False
    )
    page2 = LayoutPage(
        page=0, width_pt=90, height_pt=20, col_widths_pt=[30.0] * 3, row_heights_pt=[12.0], cells=[boxed]
    )
    assert overflow_layout(boxed, (0, 0, 0, 0), _occupancy(page2, 1, 3), xs)[0] == 0.0


def test_doubtful_ocr_text_gets_a_note_only_on_pages_marked_as_scanned() -> None:
    from pdf2excel.core.layout_text import OCR_DOUBT_WARNING

    def page(warnings: list[str]) -> LayoutPage:
        return LayoutPage(
            page=0,
            width_pt=100,
            height_pt=30,
            col_widths_pt=[50.0, 50.0],
            row_heights_pt=[20.0],
            cells=[
                LayoutCell(row=0, col=0, lines=[[TextRun("curT", size_pt=8)]]),
                LayoutCell(row=0, col=1, lines=[[TextRun("CUIT", size_pt=8)]]),
            ],
            warnings=warnings,
        )

    ws = _build(page([OCR_DOUBT_WARNING]))[0]["Página 1"]
    assert ws["A1"].comment is not None and "curT" in ws["A1"].comment.text
    assert ws["B1"].comment is None
    assert _build(page([]))[0]["Página 1"]["A1"].comment is None
