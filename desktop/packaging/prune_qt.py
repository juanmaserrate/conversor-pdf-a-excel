"""Post-build de PyInstaller: borra del bundle lo que Qt arrastra por dependencias binarias y no usamos.

PyInstaller copia módulos Qt enteros (Qml/Quick/VirtualKeyboard/Pdf/Wayland/EglFS…) aunque solo importemos
QtWidgets, más plugins y traducciones para 40 idiomas. Acá se eliminan de forma conservadora:

- Bibliotecas Qt que no son Core/Gui/Widgets/Network/Svg/OpenGL/DBus/XcbQpa (lista NEVER_DROP) y los
  módulos Python de PySide6 correspondientes.
- Directorios de plugins que no se usan (PLUGIN_DIRS_DROP). ``platforms/`` nunca se toca.
- Traducciones que no sean español/inglés.
- Windows: ``opengl32sw.dll`` (renderer software, 20 MB), ``d3dcompiler_47.dll`` y, si apareciera,
  ``opencv_videoio_ffmpeg*.dll`` (opencv-headless no debería traerlo).

Uso::

    python packaging/prune_qt.py dist/ConversorPDFaExcel            # onedir (Windows/Linux)
    python packaging/prune_qt.py dist/ConversorPDFaExcel.app        # bundle de macOS

Imprime los MB liberados. ``prune(dist) -> int`` devuelve los bytes liberados.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
from collections.abc import Iterable
from pathlib import Path

# Bibliotecas Qt / módulos PySide6 a borrar (nombres de módulo sin el prefijo Qt6 / PySide6.Qt).
_DROP_MODULES = (
    "Qml",
    "Quick",
    "QuickWidgets",
    "Quick3D",
    "QuickControls2",
    "QuickTemplates2",
    "QuickLayouts",
    "QuickTest",
    "QuickTimeline",
    "QuickShapes",
    "QuickParticles",
    "QuickEffects",
    "QuickDialogs2",
    "QuickVectorImage",
    "VirtualKeyboard",
    "Pdf",
    "PdfWidgets",
    "PdfQuick",
    "EglFS",
    "EglFs",
    "Wayland",
    "WlShell",
    "Labs",
    "ShaderTools",
    "WebEngine",
    "WebEngineCore",
    "WebEngineWidgets",
    "WebEngineQuick",
    "WebView",
    "WebChannel",
    "WebSockets",
    "Multimedia",
    "MultimediaWidgets",
    "MultimediaQuick",
    "SpatialAudio",
    "3DCore",
    "3DRender",
    "3DInput",
    "3DLogic",
    "3DAnimation",
    "3DExtras",
    "3DQuick",
    "Charts",
    "ChartsQml",
    "DataVisualization",
    "Graphs",
    "Location",
    "Positioning",
    "PositioningQuick",
    "Bluetooth",
    "Nfc",
    "Sensors",
    "SensorsQuick",
    "SerialPort",
    "SerialBus",
    "RemoteObjects",
    "Scxml",
    "StateMachine",
    "Test",
    "Sql",
    "NetworkAuth",
    "Designer",
    "DesignerComponents",
    "Help",
    "UiTools",
    "OpenGLWidgets",
    "TextToSpeech",
    "HttpServer",
    "Xml",
    "SvgWidgets",
    "Concurrent",
)
# Regex de nombres de archivo: libQt6Qml.so.6, Qt6Quick.dll, QtQml.abi3.so, QtQuick.pyd, Qt6Pdf.framework…
_DROP_ALT = "|".join(sorted(set(_DROP_MODULES), key=len, reverse=True))
LIB_DROP = re.compile(
    rf"^(?:lib)?(?:Qt6|Qt)(?:{_DROP_ALT})(?:\w*)\.(?:dll|so(?:\.[0-9.]+)?|dylib|pyd|abi3\.so|framework)$",
    re.I,
)

#: Nunca se borran (sin importar las reglas de arriba).
NEVER_DROP = re.compile(r"^(?:lib)?(?:Qt6|Qt)(?:Core|Gui|Widgets|Network|Svg|OpenGL|DBus|XcbQpa)\b", re.I)

#: Directorios de plugins Qt que no usamos. ``platforms/``, ``imageformats/``, ``iconengines/``,
#: ``styles/`` y ``platformthemes/`` se conservan siempre.
PLUGIN_DIRS_DROP = frozenset(
    {
        "platforminputcontexts",
        "wayland-shell-integration",
        "wayland-graphics-integration-client",
        "wayland-graphics-integration-server",
        "wayland-decoration-client",
        "egldeviceintegrations",
        "xcbglintegrations",
        "qmltooling",
        "qmllint",
        "scenegraph",
        "networkinformation",
        "tls",
        "sqldrivers",
        "multimedia",
        "position",
        "sensors",
        "canbus",
        "texttospeech",
        "virtualkeyboard",
        "webview",
        "assetimporters",
        "sceneparsers",
        "renderers",
        "renderplugins",
        "geometryloaders",
        "designer",
        "help",
        "generic",
        "printsupport",
    }
)

#: Directorios enteros de PySide6 que son solo para desarrollo (qml, herramientas, typesystems…).
DEV_DIRS_DROP = frozenset(
    {"qml", "libexec", "typesystems", "glue", "include", "scripts", "examples", "resources"}
)
#: Ejecutables de herramientas de Qt que PySide6 trae en su raíz.
TOOL_EXES = re.compile(
    r"^(?:assistant|designer|linguist|lupdate|lrelease|lconvert|qmlls|qmllint|qmlformat|qmlimportscanner|"
    r"qmltyperegistrar|qsb|balsam|balsamui|svgtoqml|uic|rcc|pyside6-\w+|shiboken6|QtWebEngineProcess|"
    r"qtdiag|qtpaths|windeployqt|macdeployqt|androiddeployqt|qmlcachegen|qmlplugindump|qmlprofiler|"
    r"qmlpreview|qmlscene|qmltestrunner|qml|qdbus|qdbuscpp2xml|qdbusxml2cpp|qdistancefieldgenerator|"
    r"qmlsc|qmleasing|qtattributionsscanner)(?:\.exe)?$",
    re.I,
)

#: Traducciones que se conservan: qtbase_es.qm, qt_es.qm, qtbase_en.qm… (y variantes como es_AR).
KEEP_TRANSLATIONS = re.compile(r"^(?:qt|qtbase)_(?:es|en)(?:_[A-Za-z]+)?\.qm$", re.I)

#: Archivos sueltos a borrar en Windows (el renderer software de OpenGL pesa ~20 MB y nunca se usa en
#: una app de widgets sobre Windows 10+).
WINDOWS_FILES_DROP = re.compile(
    r"^(?:opengl32sw\.dll|d3dcompiler_47\.dll|opencv_videoio_ffmpeg\w*\.dll)$", re.I
)

_QT_MARKERS = ("PySide6", "shiboken6", "Qt6", "Qt/")


def _size(path: Path) -> int:
    if path.is_symlink():
        return 0
    if path.is_dir():
        return sum(f.stat().st_size for f in path.rglob("*") if f.is_file() and not f.is_symlink())
    return path.stat().st_size if path.is_file() else 0


def _remove(path: Path) -> int:
    freed = _size(path)
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path, ignore_errors=True)
    else:
        path.unlink(missing_ok=True)
    return freed


def _internal_dir(dist: Path) -> Path:
    """Carpeta donde viven los binarios: ``_internal`` (onedir), ``Contents/Frameworks`` (.app) o ``dist``."""
    if dist.suffix == ".app" and (dist / "Contents" / "Frameworks").is_dir():
        return dist / "Contents" / "Frameworks"
    if (dist / "_internal").is_dir():
        return dist / "_internal"
    return dist


def _in_qt_tree(rel: str) -> bool:
    return any(marker in rel for marker in _QT_MARKERS)


def _should_drop(path: Path, rel: str) -> bool:
    name = path.name
    parts = rel.split("/")

    if path.is_dir():
        # Plugins no usados: .../PySide6/Qt/plugins/<dir>  o  .../PySide6/plugins/<dir>
        if "plugins" in parts[:-1] and name in PLUGIN_DIRS_DROP and _in_qt_tree(rel):
            return True
        # Directorios de desarrollo de PySide6 (qml/, libexec/, typesystems/…)
        if name in DEV_DIRS_DROP and parts[0] == "PySide6" and len(parts) <= 3:
            return True
        # Frameworks de macOS (Qt6Quick.framework es un directorio)
        return bool(name.endswith(".framework") and LIB_DROP.match(name) and not NEVER_DROP.match(name))

    if not path.is_file() and not path.is_symlink():
        return False
    if NEVER_DROP.match(name):
        return False
    if _in_qt_tree(rel) and LIB_DROP.match(name):
        return True
    if "translations" in parts[:-1] and name.endswith(".qm") and not KEEP_TRANSLATIONS.match(name):
        return True
    if parts[0] == "PySide6" and len(parts) == 2 and TOOL_EXES.match(name):
        return True
    return bool(sys.platform == "win32" and WINDOWS_FILES_DROP.match(name))


def _iter_candidates(internal: Path) -> Iterable[Path]:
    # topdown para poder saltar subárboles ya borrados
    for root, dirs, files in os.walk(internal):
        base = Path(root)
        for d in list(dirs):
            yield base / d
        for f in files:
            yield base / f


def _remove_dangling_symlinks(root: Path) -> int:
    removed = 0
    for p in root.rglob("*"):
        if p.is_symlink() and not p.exists():
            p.unlink(missing_ok=True)
            removed += 1
    return removed


def prune(dist: Path) -> int:
    """Borra lo prescindible dentro de ``dist`` y devuelve los bytes liberados."""
    dist = Path(dist)
    if not dist.exists():
        raise FileNotFoundError(f"No existe el bundle: {dist}")
    internal = _internal_dir(dist)
    freed = 0
    for path in _iter_candidates(internal):
        if not path.exists() and not path.is_symlink():
            continue  # ya se fue con un directorio padre
        rel = path.relative_to(internal).as_posix()
        if _should_drop(path, rel):
            freed += _remove(path)
    if dist.suffix == ".app":
        # En el .app PyInstaller enlaza Contents/Resources -> Contents/Frameworks; limpiar enlaces rotos.
        _remove_dangling_symlinks(dist / "Contents" / "Resources")
    return freed


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print("Uso: python packaging/prune_qt.py dist/ConversorPDFaExcel", file=sys.stderr)
        return 2
    try:
        freed = prune(Path(argv[1]))
    except FileNotFoundError as exc:
        print(f"prune_qt: {exc}", file=sys.stderr)
        return 1
    print(f"prune_qt: liberados {freed / 1048576:.1f} MB")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
