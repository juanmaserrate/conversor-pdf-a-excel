"""UI del modo "Igual al PDF": LayoutPreview, selector de formato, Opciones avanzadas y progreso."""

from __future__ import annotations

import dataclasses
import os
import time
from pathlib import Path
from typing import Any

import pytest
from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QSettings, Qt
from PySide6.QtGui import QColor, QImage
from PySide6.QtTest import QSignalSpy
from PySide6.QtWidgets import QApplication

from pdf2excel.models import (
    BorderSide,
    ColumnInfo,
    ConversionReport,
    LayoutCell,
    LayoutImage,
    LayoutPage,
    Options,
    PageInfo,
    TextRun,
    TypedTable,
)
from pdf2excel.ui import strings_es as S
from pdf2excel.ui.layout_preview import LayoutPreview
from pdf2excel.ui.pages.progress_page import ProgressPage
from pdf2excel.ui.pages.review_page import ReviewPage
from pdf2excel.ui.settings_dialog import SettingsDialog, load_options, save_options

THIN = BorderSide(0.6, "444444")
BOX = dict(border_left=THIN, border_top=THIN, border_right=THIN, border_bottom=THIN)
SCREENSHOT_DIR = os.environ.get("L5_SHOTS")


@pytest.fixture(autouse=True)
def isolated_settings(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path / "cfg"))


def _run(text: str, **kw: Any) -> list[TextRun]:
    return [TextRun(text, **kw)]


def _png(width: int, height: int, color: str) -> bytes:
    image = QImage(width, height, QImage.Format.Format_RGB32)
    image.fill(QColor(color))
    data = QByteArray()
    buf = QBuffer(data)
    buf.open(QIODevice.OpenModeFlag.WriteOnly)
    image.save(buf, "PNG")
    return bytes(data.data())


def _box(row: int, col: int, label: str, value: str, col_span: int = 1, row_span: int = 1) -> LayoutCell:
    lines = [_run(label, size_pt=6.5, color="666666")]
    lines.append(_run(value, bold=True, size_pt=9) if value else _run(""))
    return LayoutCell(row, col, row_span, col_span, lines=lines, valign="top", **BOX)  # type: ignore[arg-type]


def make_route_page(page: int = 0) -> LayoutPage:
    """Hoja de ruta sintética (todo ficticio): logo, cajas etiqueta/valor, tabla de paradas y firmas."""
    cols = [48.0] * 12
    rows = [26.0, 8.0, 30.0, 30.0, 8.0, 18.0] + [16.0] * 8 + [10.0, 40.0, 40.0]
    cells: list[LayoutCell] = [
        LayoutCell(0, 4, 1, 4, lines=[_run("HOJA DE RUTA", bold=True, size_pt=16)], halign="center"),
        _box(0, 9, "VIAJE", "N° 0000123", col_span=3),
        _box(2, 0, "TRANSPORTISTA", "Ana Pérez Demo", col_span=4),
        _box(2, 4, "AUXILIAR", "Luis Gómez Demo", col_span=4),
        _box(2, 8, "REPARTO", "R9 Zona Norte", col_span=4),
        _box(3, 0, "VEHÍCULO", "", col_span=4),
        _box(3, 4, "PATENTE", "AA 000 ZZ", col_span=4),
        _box(3, 8, "TEMPERATURA", "", col_span=4),
    ]
    heads = ["N°", "Cliente", "Dirección", "Bultos"]
    spans = [(0, 1), (1, 5), (6, 4), (10, 2)]
    for (c, w), head in zip(spans, heads, strict=True):
        cells.append(
            LayoutCell(
                5,
                c,
                1,
                w,
                lines=[_run(head, bold=True, size_pt=8, color="FFFFFF")],
                fill="2F5597",
                **BOX,  # type: ignore[arg-type]
            )
        )
    for r in range(6, 14):
        values = [str(r - 5), f"Cliente de prueba {r - 5}", f"Calle Falsa {100 + r}", str(r % 4 + 1)]
        for (c, w), val in zip(spans, values, strict=True):
            fill = "FFF2CC" if r == 9 else None
            cells.append(
                LayoutCell(
                    r,
                    c,
                    1,
                    w,
                    lines=[_run(val, size_pt=8)],
                    fill=fill,
                    halign="right" if c == 10 else "left",
                    **BOX,  # type: ignore[arg-type]
                )
            )
    cells.append(LayoutCell(15, 0, 1, 6, lines=[_run("Firma transportista")], valign="bottom", **BOX))  # type: ignore[arg-type]
    cells.append(LayoutCell(15, 6, 1, 6, lines=[_run("Firma cliente")], valign="bottom", **BOX))  # type: ignore[arg-type]
    cells.append(LayoutCell(14, 0, 1, 12, lines=[]))
    logo = LayoutImage(
        data=_png(120, 60, "#2f9e44"),
        format="png",
        bbox=(10, 4, 130, 24),
        row=0,
        col=0,
        dx_pt=4,
        dy_pt=2,
        width_pt=120,
        height_pt=22,
    )
    return LayoutPage(
        page=page,
        width_pt=576.0,
        height_pt=sum(rows),
        col_widths_pt=cols,
        row_heights_pt=rows,
        cells=cells,
        images=[logo],
        n_words=60,
    )


def make_tables_table() -> TypedTable:
    header = ["N°", "Cliente", "Dirección", "Bultos"]
    rows: list[list[Any]] = [
        [i, f"Cliente de prueba {i}", f"Calle Falsa {100 + i}", i % 4 + 1] for i in range(1, 9)
    ]
    cols = [
        ColumnInfo(kind="number", header="N°", number_format="0"),
        ColumnInfo(kind="text", header="Cliente"),
        ColumnInfo(kind="text", header="Dirección"),
        ColumnInfo(kind="number", header="Bultos", number_format="0"),
    ]
    return TypedTable(name="Paradas", header=header, rows=rows, columns=cols, pages=(0, 0))


def make_report(
    path: Path, mode: str = "layout", with_layout: bool = True, n_pages: int = 2
) -> ConversionReport:
    return ConversionReport(
        source=path,
        pages=[PageInfo(index=i, kind="digital", n_chars=300, image_coverage=0.0) for i in range(n_pages)],
        tables=[make_tables_table()],
        text_items=[],
        layout_pages=[make_route_page(i) for i in range(n_pages)] if with_layout else [],
        suggested_mode="layout" if with_layout else "tables",
        mode=mode,  # type: ignore[arg-type]
        mode_reason=(
            "Parece un formulario (varias cajas y datos sueltos): conviene respetar el diseño del PDF."
            if with_layout
            else "Es un documento muy largo (parece un listado): conviene una tabla limpia."
        ),
    )


# --------------------------------------------------------------------------- LayoutPreview


def test_layout_preview_spans_sizes_and_readonly(qtbot: Any) -> None:
    view = LayoutPreview()
    qtbot.addWidget(view)
    view.resize(700, 500)
    view.show()
    page = make_route_page()
    view.set_page(page)
    assert view.rowCount() == len(page.row_heights_pt)
    assert view.columnCount() == len(page.col_widths_pt)
    assert view.columnSpan(2, 0) == 4
    assert view.columnSpan(0, 4) == 4
    assert view.editTriggers() == view.EditTrigger.NoEditTriggers
    assert not view.item(2, 0).flags() & Qt.ItemFlag.ItemIsEditable
    # Proporcional: la columna ancha 48 pt mide el doble que una de 24 pt a igual zoom.
    assert view.columnWidth(0) == round(48 * view.zoom)
    assert view.rowHeight(0) == round(26 * view.zoom)
    assert view.item(2, 0).toolTip().startswith("TRANSPORTISTA")
    view.set_page(None)
    assert view.rowCount() == 0


def test_layout_preview_renders_cells_and_image(qtbot: Any) -> None:
    view = LayoutPreview()
    qtbot.addWidget(view)
    view.resize(700, 520)
    view.show()
    view.set_page(make_route_page())
    img = view.viewport().grab().toImage()
    # El logo verde está pintado en su posición y la cabecera azul se ve.
    logo_x = view.columnViewportPosition(0) + round(10 * view.zoom)
    logo_y = view.rowViewportPosition(0) + round(8 * view.zoom)
    assert img.pixelColor(logo_x, logo_y).green() > 120 and img.pixelColor(logo_x, logo_y).red() < 100
    head_x = view.columnViewportPosition(1) + round(40 * view.zoom)
    head_y = view.rowViewportPosition(5) + round(3 * view.zoom)
    assert img.pixelColor(head_x, head_y).blue() > 120 and img.pixelColor(head_x, head_y).red() < 90


def test_layout_preview_handles_overlap_and_out_of_grid(qtbot: Any) -> None:
    view = LayoutPreview()
    qtbot.addWidget(view)
    page = LayoutPage(
        page=0,
        width_pt=100,
        height_pt=100,
        col_widths_pt=[50, 50],
        row_heights_pt=[20, 20],
        cells=[
            LayoutCell(0, 0, 1, 2, lines=_lines("a")),
            LayoutCell(0, 1, 1, 1, lines=_lines("solapada")),
            LayoutCell(5, 5, lines=_lines("fuera")),
            LayoutCell(1, 1, 3, 3, lines=_lines("pasada")),
        ],
    )
    view.set_page(page)
    assert view.columnSpan(0, 0) == 2
    assert view.rowSpan(1, 1) == 1


def _lines(text: str) -> list[list[TextRun]]:
    return [_run(text)]


def test_layout_preview_is_fast_with_big_grid(qtbot: Any) -> None:
    cols, rows = 80, 250
    cells = [
        LayoutCell(r, c, lines=_lines(f"{r},{c}"), **BOX)  # type: ignore[arg-type]
        for r in range(0, rows, 2)
        for c in range(0, cols, 2)
    ]
    page = LayoutPage(
        page=0,
        width_pt=800,
        height_pt=1200,
        col_widths_pt=[10.0] * cols,
        row_heights_pt=[5.0] * rows,
        cells=cells,
    )
    view = LayoutPreview()
    qtbot.addWidget(view)
    view.resize(700, 500)
    view.show()
    start = time.perf_counter()
    view.set_page(page)
    view.viewport().grab()
    assert time.perf_counter() - start < 3.0


# --------------------------------------------------------------------------- ReviewPage


def _page_with(qtbot: Any, report: ConversionReport, name: str = "demo.pdf") -> tuple[ReviewPage, str]:
    page = ReviewPage()
    qtbot.addWidget(page)
    page.resize(760, 520)
    page.show()
    path = f"/tmp/{name}"
    page.add_report(path, report)
    return page, path


def test_selector_preselects_report_mode_and_shows_reason(qtbot: Any, tmp_path: Path) -> None:
    report = make_report(tmp_path / "a.pdf", mode="layout")
    page, _ = _page_with(qtbot, report)
    assert page.layout_radio.isChecked() and page.layout_radio.isEnabled()
    assert page.current_mode() == "layout"
    assert page.preview_stack.currentWidget() is page.layout_preview
    assert page.mode_reason_label.text() == report.mode_reason
    assert page.table_combo.count() == 2 and page.table_combo.itemText(1) == "Página 2 de 2"
    assert page.layout_preview.page() is report.layout_pages[0]
    # Controles intactos
    assert page.typed_check.isChecked() and page.save_button.isEnabled() and page.back_button.isEnabled()


def test_switching_mode_updates_report_preview_and_signal(qtbot: Any, tmp_path: Path) -> None:
    report = make_report(tmp_path / "a.pdf", mode="layout")
    page, path = _page_with(qtbot, report)
    spy = QSignalSpy(page.mode_changed)
    page.tables_radio.click()
    assert report.mode == "tables"
    assert page.preview_stack.currentWidget() is page.table_view
    assert page.preview_model.rowCount() == 8
    assert spy.count() == 1 and spy.at(0) == [path, "tables"]
    page.layout_radio.click()
    assert report.mode == "layout"
    assert page.preview_stack.currentWidget() is page.layout_preview
    page.table_combo.setCurrentIndex(1)
    assert page.layout_preview.page() is report.layout_pages[1]


def test_mode_is_per_file(qtbot: Any, tmp_path: Path) -> None:
    first = make_report(tmp_path / "a.pdf", mode="layout")
    second = make_report(tmp_path / "b.pdf", mode="layout")
    page, p1 = _page_with(qtbot, first, "a.pdf")
    page.add_report("/tmp/b.pdf", second)
    page.show_report("/tmp/b.pdf")
    page.tables_radio.click()
    assert second.mode == "tables" and first.mode == "layout"
    page.show_report(p1)
    assert page.layout_radio.isChecked()
    page.show_report("/tmp/b.pdf")
    assert page.tables_radio.isChecked()


def test_layout_disabled_without_layout_pages(qtbot: Any, tmp_path: Path) -> None:
    report = make_report(tmp_path / "big.pdf", mode="layout", with_layout=False)
    page, _ = _page_with(qtbot, report)
    assert not page.layout_radio.isEnabled()
    assert page.layout_radio.toolTip() == S.FORMAT_LAYOUT_DISABLED_TIP
    assert page.tables_radio.isChecked()
    assert report.mode == "tables"
    assert page.preview_stack.currentWidget() is page.table_view
    page.set_mode("layout")
    assert report.mode == "tables"


def test_keyboard_navigation_between_radios(qtbot: Any, tmp_path: Path) -> None:
    report = make_report(tmp_path / "a.pdf", mode="layout")
    page, _ = _page_with(qtbot, report)
    assert page.layout_radio.focusPolicy() != Qt.FocusPolicy.NoFocus
    assert page.layout_radio.accessibleName() and page.tables_radio.accessibleName()
    page.tables_radio.setFocus()
    qtbot.keyClick(page.tables_radio, Qt.Key.Key_Space)
    assert page.tables_radio.isChecked() and report.mode == "tables"
    qtbot.keyClick(page.layout_radio, Qt.Key.Key_Space)
    assert report.mode == "layout"


def test_typed_checkbox_still_works_in_tables_mode(qtbot: Any, tmp_path: Path) -> None:
    report = make_report(tmp_path / "a.pdf", mode="tables")
    page, _ = _page_with(qtbot, report)
    assert page.preview_stack.currentWidget() is page.table_view
    page.typed_check.setChecked(False)
    assert page.preview_model.rowCount() == 8


def test_error_file_hides_selector(qtbot: Any, tmp_path: Path) -> None:
    from pdf2excel.errors import UserError

    page = ReviewPage()
    qtbot.addWidget(page)
    page.show()
    page.add_error("/tmp/mal.pdf", UserError("corrupt"))
    assert not page.format_box.isVisible()


# --------------------------------------------------------------------------- Opciones avanzadas y progreso


def test_settings_layout_choice_roundtrip(qtbot: Any) -> None:
    dialog = SettingsDialog(Options())
    qtbot.addWidget(dialog)
    assert [dialog.layout_combo.itemData(i) for i in range(3)] == ["auto", "layout", "tables"]
    assert dialog.layout_combo.currentData() == "auto"
    dialog.layout_combo.setCurrentIndex(1)
    opts = dialog.options()
    assert opts.layout == "layout"
    save_options(opts)
    assert load_options().layout == "layout"
    dialog2 = SettingsDialog(load_options())
    qtbot.addWidget(dialog2)
    assert dialog2.layout_combo.currentData() == "layout"
    save_options(dataclasses.replace(opts, layout="tables"))
    assert load_options().layout == "tables"


def test_settings_garbage_layout_falls_back_to_auto(tmp_path: Path) -> None:
    QSettings("ConversorPDF", "ConversorPDFaExcel").setValue("layout", "cualquiera")
    assert load_options().layout == "auto"


def test_progress_layout_stage_text(qtbot: Any) -> None:
    page = ProgressPage()
    qtbot.addWidget(page)
    page.set_progress("/tmp/a.pdf", 3, 10, "layout", 1, 1)
    assert page.stage_label.text() == "Armando el diseño de la hoja…"
    assert "3" in page.counter_label.text() and "10" in page.counter_label.text()


# --------------------------------------------------------------------------- capturas (manual)


@pytest.mark.skipif(not SCREENSHOT_DIR, reason="definí L5_SHOTS para guardar capturas")
def test_screenshots(qtbot: Any, tmp_path: Path) -> None:
    out = Path(str(SCREENSHOT_DIR))
    out.mkdir(parents=True, exist_ok=True)
    theme = Path(__file__).resolve().parents[1] / "src" / "pdf2excel" / "ui" / "theme.qss"
    app = QApplication.instance()
    if app is not None and theme.is_file() and isinstance(app, QApplication):
        app.setStyleSheet(theme.read_text(encoding="utf-8"))
    for size in ((900, 620), (640, 440)):
        for mode in ("layout", "tables"):
            report = make_report(tmp_path / "hoja_ruta.pdf", mode=mode)
            page = ReviewPage()
            qtbot.addWidget(page)
            page.resize(*size)
            page.show()
            page.add_file("/tmp/otro_archivo_demo.pdf")
            page.add_report("/tmp/hoja_ruta_demo.pdf", report)
            page.file_list.select_file("/tmp/hoja_ruta_demo.pdf")
            page.show_report("/tmp/hoja_ruta_demo.pdf")
            QApplication.processEvents()
            page.grab().save(str(out / f"review_{mode}_{size[0]}x{size[1]}.png"))
    big = make_report(tmp_path / "x.pdf", mode="layout", with_layout=False)
    page = ReviewPage()
    qtbot.addWidget(page)
    page.resize(760, 500)
    page.show()
    page.add_report("/tmp/listado_largo.pdf", big)
    QApplication.processEvents()
    page.grab().save(str(out / "review_layout_disabled.png"))


# --------------------------------------------------------------------------- MainWindow


def test_main_window_wires_modes(qtbot: Any, tmp_path: Path) -> None:
    from pdf2excel.ui.main_window import MainWindow

    window = MainWindow()
    qtbot.addWidget(window)
    report = make_report(tmp_path / "a.pdf", mode="layout")
    broken = make_report(tmp_path / "b.pdf", mode="layout", with_layout=False)
    broken.mode = "layout"  # inconsistente a propósito: sin diseño no se puede exportar como layout
    for name, rep in (("a.pdf", report), ("b.pdf", broken)):
        path = str(tmp_path / name)
        window.order.append(path)
        window.reports[path] = rep
        window.review_page.add_report(path, rep)
    path_a = str(tmp_path / "a.pdf")
    window.review_page.show_report(path_a)
    window.review_page.tables_radio.click()
    assert window._mode_overrides == {path_a: "tables"} and report.mode == "tables"
    window._sanitize_modes(list(window.reports.values()))
    assert broken.mode == "tables"
    # Opciones avanzadas "Siempre igual al PDF" manda sobre la elección previa y sobre reportes sin diseño.
    window.opts.layout = "layout"
    window._apply_layout_preference()
    assert report.mode == "layout" and broken.mode == "tables" and not window._mode_overrides
    window.opts.layout = "auto"
    window._apply_layout_preference()
    assert report.mode == report.suggested_mode
