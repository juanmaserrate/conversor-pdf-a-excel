# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller onedir de "Conversor PDF a Excel".

Build (desde desktop/, con el paquete instalado: pip install -e .):

    pyinstaller --noconfirm --clean packaging/conversor.spec
    python packaging/prune_qt.py dist/ConversorPDFaExcel
    dist/ConversorPDFaExcel/ConversorPDFaExcel --selftest tests/corpus/_generated/remito_grid_scan.pdf

Decisiones (ver docs/DESIGN.md §8):
- onedir (onefile tarda 3-3,6 s en arrancar y molesta a los antivirus).
- SE INCLUYEN pymupdf-layout + onnxruntime (el modelo de layout es lo que detecta tablas sin bordes y
  sobre OCR) y tesseract vendoreado en Windows (resources/tesseract, lo baja packaging/fetch_tesseract.py).
- Qt: solo Widgets; todo lo demás va a `excludes` y lo que PyInstaller arrastre igual lo saca prune_qt.py.
"""

import os
import sys
from pathlib import Path

from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
)

ROOT = Path(SPECPATH).parent  # desktop/  (SPECPATH = desktop/packaging)
APP_NAME = "ConversorPDFaExcel"
BUNDLE_ID = "ar.com.conversorpdf.app"
RESOURCES = ROOT / "resources"
ICONS = RESOURCES / "icons"

# Entry point real: src/pdf2excel/__main__.py (expone main() y --selftest <pdf>).
# PDF2EXCEL_ENTRY sirve solo para probar el empaquetado con un script alternativo (CI no lo usa).
ENTRY = Path(os.environ.get("PDF2EXCEL_ENTRY") or (ROOT / "src" / "pdf2excel" / "__main__.py"))

sys.path.insert(0, str(ROOT / "src"))
from pdf2excel import __version__  # noqa: E402  (única fuente de versión)

# --------------------------------------------------------------------------- datas
# OBLIGATORIO: collect_data_files("pymupdf") copia pymupdf/layout/resources/onnx/*.onnx y *.yaml.
# Sin esto el bundle arranca pero find_tables() falla con
#   FileNotFoundError: .../pymupdf/layout/resources/onnx/layout_rf2.4.1+imf1.yaml
# (pymupdf 2.0 autoimporta pymupdf.layout; la dist "pymupdf_layout" solo trae metadata y los datos
# viven dentro del paquete pymupdf, por eso collect_data_files("pymupdf_layout") devuelve [] y es inocuo).
datas = collect_data_files("pymupdf") + collect_data_files("pymupdf_layout")
datas += [(str(ICONS), "icons")]
# El tema: app.py lo lee como Path(__file__).parent / "ui" / "theme.qss"; congelado, __file__ apunta a
# _internal/pdf2excel/app.pyc, así que el .qss tiene que estar en _internal/pdf2excel/ui/ (es el único
# archivo no-.py de src/ y PyInstaller no lo arrastra solo). Sin esto la app arranca sin tema (WARNING).
THEME_QSS = ROOT / "src" / "pdf2excel" / "ui" / "theme.qss"
if THEME_QSS.is_file():
    datas.append((str(THEME_QSS), "pdf2excel/ui"))
else:
    raise SystemExit(f"conversor.spec: falta {THEME_QSS} (la app empaquetada quedaría sin tema)")
if (RESOURCES / "tessdata").is_dir():
    datas.append((str(RESOURCES / "tessdata"), "tessdata"))  # spa/eng/osd + pdf.ttf + configs/pdf
else:
    print("conversor.spec: AVISO resources/tessdata no existe (corré packaging/fetch_tesseract.py resources)")
if (RESOURCES / "tesseract").is_dir():
    datas.append((str(RESOURCES / "tesseract"), "tesseract"))  # tesseract.exe + DLL (solo Windows)

# --------------------------------------------------------------------------- binaries / hiddenimports
binaries = collect_dynamic_libs("onnxruntime")

hiddenimports = (
    collect_submodules("pymupdf_layout")
    + collect_submodules("pymupdf.layout")
    + collect_submodules("pdf2excel")
    + [
        "onnxruntime",
        "openpyxl.cell._writer",
        "cv2",
        "platformdirs",
        "numpy",
    ]
)

# --------------------------------------------------------------------------- excludes
QT_UNUSED = (
    "QtWebEngineCore", "QtWebEngineWidgets", "QtWebEngineQuick",
    "QtQuick", "QtQuickWidgets", "QtQuickControls2", "QtQuick3D", "QtQml",
    "Qt3DCore", "Qt3DRender", "Qt3DInput", "Qt3DLogic", "Qt3DAnimation", "Qt3DExtras",
    "QtMultimedia", "QtMultimediaWidgets", "QtPdf", "QtPdfWidgets",
    "QtCharts", "QtDataVisualization", "QtLocation", "QtPositioning", "QtBluetooth", "QtNfc",
    "QtSensors", "QtSerialPort", "QtSerialBus", "QtRemoteObjects", "QtScxml", "QtStateMachine",
    "QtTest", "QtSql", "QtNetworkAuth", "QtWebSockets", "QtWebChannel", "QtDesigner", "QtHelp",
    "QtUiTools", "QtOpenGLWidgets", "QtTextToSpeech", "QtSpatialAudio", "QtGraphs", "QtHttpServer",
    "QtWebView", "QtXml", "QtSvgWidgets", "QtConcurrent", "QtAsyncio",
)
excludes = [f"PySide6.{m}" for m in QT_UNUSED] + [
    "pandas", "camelot", "pypdfium2", "pdfplumber", "pdfminer", "lxml", "matplotlib", "scipy",
    "IPython", "tkinter", "rapidocr_onnxruntime", "pytest", "reportlab", "pymupdf4llm", "PIL.ImageQt",
    # pymupdf importa PIL y fontTools solo dentro de funciones que no usamos (pil_save, subset_fonts):
    # sacarlos ahorra ~27 MB. Si algún día hace falta PIL, quitarlo de acá y de la lista de abajo.
    "PIL", "fontTools", "uharfbuzz",
    "pyinstaller", "setuptools", "pkg_resources", "wheel", "pip",
]

# En el venv de desarrollo puede estar instalado opencv-python (completo, con Qt5+GTK+ffmpeg, +116 MB)
# además de opencv-python-headless: PyInstaller empaqueta el cv2 que resuelva el import. CI instala solo
# headless; acá avisamos para que la medición de tamaño no confunda.
try:
    from importlib.metadata import distribution

    distribution("opencv-python")
    print("conversor.spec: AVISO opencv-python (NO headless) está instalado: el bundle crece ~116 MB. "
          "Desinstalalo (pip uninstall opencv-python) para medir el tamaño real.")
except Exception:  # PackageNotFoundError u otro: no está, perfecto
    pass

# --------------------------------------------------------------------------- build
a = Analysis(
    [str(ENTRY)],
    pathex=[str(ROOT / "src")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

if sys.platform == "win32":
    icon = str(ICONS / "app.ico")
    version_file = str(ROOT / "packaging" / "file_version_info.txt")
elif sys.platform == "darwin":
    icon = str(ICONS / "app.icns") if (ICONS / "app.icns").is_file() else str(ICONS / "app.png")
    version_file = None
else:
    icon = None
    version_file = None

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # sin ventana negra; --selftest igual escribe a stdout si está redirigido
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
    version=version_file,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=f"{APP_NAME}.app",
        icon=icon,
        bundle_identifier=BUNDLE_ID,
        info_plist={
            "NSHighResolutionCapable": True,
            "CFBundleName": "Conversor PDF a Excel",
            "CFBundleDisplayName": "Conversor PDF a Excel",
            "CFBundleShortVersionString": __version__,
            "CFBundleVersion": __version__,
            "LSMinimumSystemVersion": "12.0",
            "NSRequiresAquaSystemAppearance": False,
            "CFBundleDocumentTypes": [
                {
                    "CFBundleTypeName": "PDF",
                    "CFBundleTypeRole": "Viewer",
                    "LSItemContentTypes": ["com.adobe.pdf"],
                    "LSHandlerRank": "Alternate",
                }
            ],
        },
    )
