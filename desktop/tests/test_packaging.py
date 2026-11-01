"""Tests de empaquetado: update_check, scripts de packaging/, conversor.spec, installer.iss y el workflow.

No construyen nada con PyInstaller: validan los scripts sobre árboles sintéticos y el contenido de los
archivos.
"""

from __future__ import annotations

import http.server
import importlib.util
import json
import re
import shutil
import sys
import threading
from collections.abc import Iterator
from pathlib import Path
from types import ModuleType

import pytest

DESKTOP = Path(__file__).resolve().parents[1]
PACKAGING = DESKTOP / "packaging"
WORKFLOW = DESKTOP.parent / ".github" / "workflows" / "build-desktop.yml"


def _load(name: str, path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


stamp_version = _load("wi10_stamp_version", PACKAGING / "stamp_version.py")
prune_qt = _load("wi10_prune_qt", PACKAGING / "prune_qt.py")
fetch_tesseract = _load("wi10_fetch_tesseract", PACKAGING / "fetch_tesseract.py")

# --------------------------------------------------------------------------- update_check


class _Handler(http.server.BaseHTTPRequestHandler):
    payload: bytes = b"{}"
    status: int = 200

    def do_GET(self) -> None:
        self.send_response(self.status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(self.payload)))
        self.end_headers()
        self.wfile.write(self.payload)

    def log_message(self, *_args: object) -> None:
        pass


@pytest.fixture
def latest_server() -> Iterator[http.server.HTTPServer]:
    server = http.server.HTTPServer(("127.0.0.1", 0), _Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield server
    finally:
        server.shutdown()
        server.server_close()


def _serve(
    monkeypatch: pytest.MonkeyPatch, server: http.server.HTTPServer, data: object, status: int = 200
) -> None:
    _Handler.payload = json.dumps(data).encode() if not isinstance(data, bytes) else data
    _Handler.status = status
    monkeypatch.setenv("PDF2EXCEL_UPDATE_URL", f"http://127.0.0.1:{server.server_port}/latest.json")


def test_update_check_module_has_no_qt() -> None:
    from pdf2excel import update_check

    source = Path(update_check.__file__).read_text(encoding="utf-8")
    assert not re.search(r"^\s*(import|from)\s+(PySide6|shiboken6|PyQt\d)", source, re.M)


def test_update_check_offline_returns_none(monkeypatch: pytest.MonkeyPatch) -> None:
    from pdf2excel.update_check import check_for_update

    monkeypatch.setenv("PDF2EXCEL_UPDATE_URL", "http://127.0.0.1:9/latest.json")  # puerto cerrado
    assert check_for_update("0.0.1", timeout=1) is None


def test_update_check_default_url_without_network() -> None:
    from pdf2excel.update_check import LATEST_URL, check_for_update, latest_url

    assert LATEST_URL.endswith("/releases/latest/download/latest.json")
    assert latest_url() == LATEST_URL
    assert check_for_update("0.0.1", timeout=1) is None  # sin red (o con red y sin release): nunca lanza


def test_update_check_newer_version(
    monkeypatch: pytest.MonkeyPatch, latest_server: http.server.HTTPServer
) -> None:
    from pdf2excel.update_check import UpdateInfo, check_for_update

    _serve(
        monkeypatch,
        latest_server,
        {"version": "0.2.0", "url": "https://x/setup.exe", "notes_url": "https://x/notes"},
    )
    info = check_for_update("0.1.0", timeout=3)
    assert info == UpdateInfo(version="0.2.0", url="https://x/setup.exe", notes_url="https://x/notes")


def test_update_check_same_or_older_version(
    monkeypatch: pytest.MonkeyPatch, latest_server: http.server.HTTPServer
) -> None:
    from pdf2excel.update_check import check_for_update

    _serve(monkeypatch, latest_server, {"version": "0.1.0", "url": "https://x/setup.exe"})
    assert check_for_update("0.1.0", timeout=3) is None
    assert check_for_update("0.1.1", timeout=3) is None
    assert check_for_update("1.0", timeout=3) is None


def test_update_check_notes_url_defaults_to_url(
    monkeypatch: pytest.MonkeyPatch, latest_server: http.server.HTTPServer
) -> None:
    from pdf2excel.update_check import check_for_update

    _serve(monkeypatch, latest_server, {"version": "v9.0.0", "url": "https://x/setup.exe"})
    info = check_for_update("0.1.0", timeout=3)
    assert info is not None and info.version == "9.0.0" and info.notes_url == "https://x/setup.exe"


@pytest.mark.parametrize(
    "payload",
    [
        b"not json",
        json.dumps([1, 2]).encode(),
        json.dumps({"version": 3, "url": "https://x"}).encode(),
        json.dumps({"version": "9.9.9", "url": "file:///etc/passwd"}).encode(),
        json.dumps({"version": "9.9.9"}).encode(),
    ],
)
def test_update_check_bad_payloads(
    monkeypatch: pytest.MonkeyPatch, latest_server: http.server.HTTPServer, payload: bytes
) -> None:
    from pdf2excel.update_check import check_for_update

    _serve(monkeypatch, latest_server, payload)
    assert check_for_update("0.1.0", timeout=3) is None


def test_update_check_http_error(
    monkeypatch: pytest.MonkeyPatch, latest_server: http.server.HTTPServer
) -> None:
    from pdf2excel.update_check import check_for_update

    _serve(monkeypatch, latest_server, {"version": "9.0.0", "url": "https://x"}, status=404)
    assert check_for_update("0.1.0", timeout=3) is None


def test_update_check_rejects_non_http_scheme(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    from pdf2excel.update_check import check_for_update

    f = tmp_path / "latest.json"
    f.write_text(json.dumps({"version": "9.0.0", "url": "https://x"}))
    monkeypatch.setenv("PDF2EXCEL_UPDATE_URL", f.as_uri())
    assert check_for_update("0.1.0", timeout=1) is None


def test_update_check_async(monkeypatch: pytest.MonkeyPatch, latest_server: http.server.HTTPServer) -> None:
    from pdf2excel.update_check import UpdateInfo, check_for_update_async

    _serve(
        monkeypatch, latest_server, {"version": "0.3.0", "url": "https://x/s.exe", "notes_url": "https://x/n"}
    )
    results: list[UpdateInfo | None] = []
    thread = check_for_update_async(results.append, current="0.1.0", timeout=3)
    assert thread.daemon
    thread.join(10)
    assert not thread.is_alive()
    assert results == [UpdateInfo("0.3.0", "https://x/s.exe", "https://x/n")]


def test_update_check_async_callback_exception_is_swallowed(monkeypatch: pytest.MonkeyPatch) -> None:
    from pdf2excel.update_check import check_for_update_async

    monkeypatch.setenv("PDF2EXCEL_UPDATE_URL", "http://127.0.0.1:9/latest.json")

    def boom(_: object) -> None:
        raise RuntimeError("callback roto")

    thread = check_for_update_async(boom, timeout=1)
    thread.join(10)
    assert not thread.is_alive()


@pytest.mark.parametrize(
    ("a", "b", "newer"),
    [
        ("0.2.0", "0.1.0", True),
        ("v1.0.0", "0.9.9", True),
        ("1.0", "1.0.0", False),
        ("1.2.3-rc1", "1.2.3", False),
        ("1.10.0", "1.9.0", True),
        ("garbage", "0.1.0", False),
        ("0.1.0", "garbage", True),
    ],
)
def test_version_comparison(a: str, b: str, newer: bool) -> None:
    from pdf2excel.update_check import is_newer

    assert is_newer(a, b) is newer


# --------------------------------------------------------------------------- stamp_version


@pytest.fixture
def stamp_tree(tmp_path: Path) -> Path:
    (tmp_path / "src" / "pdf2excel").mkdir(parents=True)
    (tmp_path / "packaging").mkdir()
    shutil.copy(DESKTOP / "src" / "pdf2excel" / "__init__.py", tmp_path / "src" / "pdf2excel" / "__init__.py")
    for name in ("installer.iss", "file_version_info.txt"):
        shutil.copy(PACKAGING / name, tmp_path / "packaging" / name)
    return tmp_path


@pytest.mark.parametrize(
    ("tag", "expected"),
    [
        ("v1.2.3", "1.2.3"),
        ("1.2.3", "1.2.3"),
        ("refs/tags/v0.2.0", "0.2.0"),
        ("V2.0", "2.0"),
        ("v1.2.3-rc1", "1.2.3-rc1"),
    ],
)
def test_normalize(tag: str, expected: str) -> None:
    assert stamp_version.normalize(tag) == expected


@pytest.mark.parametrize("bad", ["main", "", "v", "1.2.3.4.5", "release-1", "v1..2"])
def test_normalize_rejects(bad: str) -> None:
    with pytest.raises(ValueError):
        stamp_version.normalize(bad)


def test_numeric_tuple() -> None:
    assert stamp_version.numeric_tuple("1.2.3") == (1, 2, 3, 0)
    assert stamp_version.numeric_tuple("1.2.3-rc1") == (1, 2, 3, 0)
    assert stamp_version.numeric_tuple("2") == (2, 0, 0, 0)
    assert stamp_version.numeric_tuple("1.2.3.4") == (1, 2, 3, 4)


def test_stamp_version_main_rewrites_three_files(
    stamp_tree: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert stamp_version.main("v1.2.3", root=stamp_tree) == "1.2.3"
    assert capsys.readouterr().out.strip() == "1.2.3"
    init = (stamp_tree / "src" / "pdf2excel" / "__init__.py").read_text()
    assert '__version__ = "1.2.3"' in init
    assert 'APP_NAME = "Conversor PDF a Excel"' in init  # el resto queda intacto
    iss = (stamp_tree / "packaging" / "installer.iss").read_text()
    assert '#define AppVersion "1.2.3"' in iss
    assert iss.count("#define AppVersion") == 1
    fvi = (stamp_tree / "packaging" / "file_version_info.txt").read_text()
    assert "filevers=(1, 2, 3, 0)" in fvi and "prodvers=(1, 2, 3, 0)" in fvi
    assert "StringStruct('FileVersion', '1.2.3')" in fvi and "StringStruct('ProductVersion', '1.2.3')" in fvi
    # idempotente / re-estampable
    stamp_version.main("v1.2.4", root=stamp_tree)
    assert "filevers=(1, 2, 4, 0)" in (stamp_tree / "packaging" / "file_version_info.txt").read_text()


def test_stamp_version_cli(stamp_tree: Path) -> None:
    import subprocess

    shutil.copy(PACKAGING / "stamp_version.py", stamp_tree / "packaging" / "stamp_version.py")
    out = subprocess.run(
        [sys.executable, str(stamp_tree / "packaging" / "stamp_version.py"), "v3.4.5"],
        capture_output=True,
        text=True,
        check=True,
    )
    assert out.stdout.strip() == "3.4.5"
    assert '__version__ = "3.4.5"' in (stamp_tree / "src" / "pdf2excel" / "__init__.py").read_text()
    bad = subprocess.run(
        [sys.executable, str(stamp_tree / "packaging" / "stamp_version.py"), "main"],
        capture_output=True,
        text=True,
    )
    assert bad.returncode != 0 and "no es una versión válida" in bad.stderr


def test_file_version_info_is_valid_and_spanish() -> None:
    text = (PACKAGING / "file_version_info.txt").read_text(encoding="utf-8")
    assert "StringTable('040A04B0'" in text
    assert "VarStruct('Translation', [1034, 1200])" in text
    assert "filevers=(0, 1, 0, 0)" in text
    for key in ("CompanyName", "ProductName", "FileDescription", "OriginalFilename"):
        assert f"StringStruct('{key}'" in text

    # PyInstaller lo evalúa con eval() sobre sus clases: tiene que ser una expresión Python válida
    class _Node:
        def __init__(self, *args: object, **kwargs: object) -> None:
            self.args, self.kwargs = args, kwargs

    names = {
        n: _Node
        for n in (
            "VSVersionInfo",
            "FixedFileInfo",
            "StringFileInfo",
            "StringTable",
            "StringStruct",
            "VarFileInfo",
            "VarStruct",
        )
    }
    code = "\n".join(line for line in text.splitlines() if not line.lstrip().startswith("#"))
    node = eval(code, {"__builtins__": {}}, names)  # archivo propio del repo
    assert node.kwargs["ffi"].kwargs["filevers"] == (0, 1, 0, 0)
    try:
        from PyInstaller.utils.win32 import versioninfo  # necesita pefile + win32api: solo en Windows
    except ImportError:
        pytest.skip("PyInstaller.utils.win32.versioninfo solo se importa en Windows")
    info = versioninfo.load_version_info_from_text_file(str(PACKAGING / "file_version_info.txt"))
    assert info.ffi.fileVersionMS == (0 << 16) | 1


# --------------------------------------------------------------------------- prune_qt


def _touch(path: Path, size: int = 1024) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"\0" * size)


@pytest.fixture
def fake_dist(tmp_path: Path) -> Path:
    dist = tmp_path / "ConversorPDFaExcel"
    internal = dist / "_internal"
    _touch(dist / "ConversorPDFaExcel", 10)
    qt = internal / "PySide6" / "Qt"
    for lib in (
        "libQt6Core.so.6",
        "libQt6Gui.so.6",
        "libQt6Widgets.so.6",
        "libQt6Network.so.6",
        "libQt6Svg.so.6",
        "libQt6OpenGL.so.6",
        "libQt6DBus.so.6",
        "libQt6XcbQpa.so.6",
    ):
        _touch(qt / "lib" / lib, 1 << 20)
    for lib in (
        "libQt6Qml.so.6",
        "libQt6Quick.so.6",
        "libQt6VirtualKeyboard.so.6",
        "libQt6Pdf.so.6",
        "libQt6WaylandClient.so.6",
        "libQt6EglFSDeviceIntegration.so.6",
        "libQt6LabsPlatform.so.6",
        "libQt6ShaderTools.so.6",
        "libQt6Sql.so.6",
        "libQt6Xml.so.6",
        "libQt6WebEngineCore.so.6",
    ):
        _touch(qt / "lib" / lib, 3 << 20)
    for mod in ("QtCore", "QtGui", "QtWidgets", "QtNetwork", "QtSvg"):
        _touch(internal / "PySide6" / f"{mod}.abi3.so", 1 << 20)
    for mod in ("QtQml", "QtQuick", "QtPdf", "QtSql", "QtXml", "QtMultimedia"):
        _touch(internal / "PySide6" / f"{mod}.abi3.so", 1 << 20)
    for keep in (
        "platforms/libqxcb.so",
        "platforms/libqoffscreen.so",
        "imageformats/libqsvg.so",
        "iconengines/libqsvgicon.so",
        "styles/libqwindowsvistastyle.dll",
        "platformthemes/libqgtk3.so",
    ):
        _touch(qt / "plugins" / keep, 1 << 20)
    for drop in (
        "platforminputcontexts/libcomposeplatforminputcontextplugin.so",
        "tls/libqopensslbackend.so",
        "sqldrivers/libqsqlite.so",
        "qmltooling/libqmldbg_debugger.so",
        "wayland-shell-integration/x.so",
        "egldeviceintegrations/x.so",
        "xcbglintegrations/libqxcb-glx-integration.so",
        "networkinformation/x.so",
    ):
        _touch(qt / "plugins" / drop, 1 << 20)
    for tr in ("qt_es.qm", "qtbase_es.qm", "qtbase_en.qm", "qt_de.qm", "qtbase_fr.qm", "qtmultimedia_ja.qm"):
        _touch(qt / "translations" / tr, 100 << 10)
    _touch(qt / "qml" / "QtQuick" / "libqtquick2plugin.so", 5 << 20)
    _touch(qt / "libexec" / "QtWebEngineProcess", 2 << 20)
    _touch(internal / "PySide6" / "designer", 2 << 20)
    _touch(internal / "PySide6" / "lupdate", 1 << 20)
    _touch(internal / "numpy" / "core.so", 1 << 20)
    _touch(internal / "pymupdf" / "layout" / "resources" / "onnx" / "model.onnx", 1 << 20)
    _touch(internal / "tessdata" / "spa.traineddata", 1 << 20)
    _touch(internal / "opengl32sw.dll", 20 << 20)
    _touch(internal / "d3dcompiler_47.dll", 4 << 20)
    _touch(internal / "opencv_videoio_ffmpeg500_64.dll", 20 << 20)
    return dist


def test_prune_drops_unused_and_keeps_essentials(fake_dist: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(prune_qt.sys, "platform", "win32")
    freed = prune_qt.prune(fake_dist)
    assert freed > 20 << 20
    internal = fake_dist / "_internal"
    qt = internal / "PySide6" / "Qt"
    for lib in (
        "libQt6Core.so.6",
        "libQt6Gui.so.6",
        "libQt6Widgets.so.6",
        "libQt6Network.so.6",
        "libQt6Svg.so.6",
        "libQt6OpenGL.so.6",
        "libQt6DBus.so.6",
        "libQt6XcbQpa.so.6",
    ):
        assert (qt / "lib" / lib).exists(), lib
    for lib in (
        "libQt6Qml.so.6",
        "libQt6Quick.so.6",
        "libQt6VirtualKeyboard.so.6",
        "libQt6Pdf.so.6",
        "libQt6WaylandClient.so.6",
        "libQt6EglFSDeviceIntegration.so.6",
        "libQt6LabsPlatform.so.6",
        "libQt6ShaderTools.so.6",
        "libQt6Sql.so.6",
        "libQt6Xml.so.6",
        "libQt6WebEngineCore.so.6",
    ):
        assert not (qt / "lib" / lib).exists(), lib
    for mod in ("QtCore", "QtGui", "QtWidgets", "QtNetwork", "QtSvg"):
        assert (internal / "PySide6" / f"{mod}.abi3.so").exists()
    for mod in ("QtQml", "QtQuick", "QtPdf", "QtSql", "QtXml", "QtMultimedia"):
        assert not (internal / "PySide6" / f"{mod}.abi3.so").exists()
    assert (qt / "plugins" / "platforms" / "libqxcb.so").exists()
    assert (qt / "plugins" / "platforms" / "libqoffscreen.so").exists()
    assert (qt / "plugins" / "imageformats" / "libqsvg.so").exists()
    assert (qt / "plugins" / "iconengines").is_dir() and (qt / "plugins" / "styles").is_dir()
    assert (qt / "plugins" / "platformthemes").is_dir()
    for gone in (
        "platforminputcontexts",
        "tls",
        "sqldrivers",
        "qmltooling",
        "wayland-shell-integration",
        "egldeviceintegrations",
        "xcbglintegrations",
        "networkinformation",
    ):
        assert not (qt / "plugins" / gone).exists(), gone
    assert sorted(p.name for p in (qt / "translations").iterdir()) == [
        "qt_es.qm",
        "qtbase_en.qm",
        "qtbase_es.qm",
    ]
    assert not (qt / "qml").exists() and not (qt / "libexec").exists()
    assert (
        not (internal / "PySide6" / "designer").exists() and not (internal / "PySide6" / "lupdate").exists()
    )
    assert (internal / "numpy" / "core.so").exists()
    assert (internal / "pymupdf" / "layout" / "resources" / "onnx" / "model.onnx").exists()
    assert (internal / "tessdata" / "spa.traineddata").exists()
    assert (fake_dist / "ConversorPDFaExcel").exists()
    for dll in ("opengl32sw.dll", "d3dcompiler_47.dll", "opencv_videoio_ffmpeg500_64.dll"):
        assert not (internal / dll).exists(), dll


def test_prune_windows_dlls_only_on_windows(fake_dist: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(prune_qt.sys, "platform", "linux")
    prune_qt.prune(fake_dist)
    assert (fake_dist / "_internal" / "opengl32sw.dll").exists()


def test_prune_is_idempotent_and_reports_bytes(fake_dist: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(prune_qt.sys, "platform", "win32")
    first = prune_qt.prune(fake_dist)
    assert first > 0
    assert prune_qt.prune(fake_dist) == 0


def test_prune_mac_app_bundle(tmp_path: Path) -> None:
    app = tmp_path / "ConversorPDFaExcel.app"
    fw = app / "Contents" / "Frameworks"
    res = app / "Contents" / "Resources"
    _touch(fw / "PySide6" / "Qt" / "lib" / "QtQuick.framework" / "Versions" / "A" / "QtQuick", 5 << 20)
    _touch(fw / "PySide6" / "Qt" / "lib" / "QtCore.framework" / "Versions" / "A" / "QtCore", 5 << 20)
    _touch(fw / "PySide6" / "QtQml.abi3.so", 1 << 20)
    _touch(fw / "PySide6" / "QtWidgets.abi3.so", 1 << 20)
    res.mkdir(parents=True)
    (res / "PySide6").symlink_to(fw / "PySide6")  # PyInstaller enlaza Resources -> Frameworks
    (res / "roto").symlink_to(fw / "PySide6" / "QtQml.abi3.so")
    freed = prune_qt.prune(app)
    assert freed >= 6 << 20
    assert not (fw / "PySide6" / "Qt" / "lib" / "QtQuick.framework").exists()
    assert (fw / "PySide6" / "Qt" / "lib" / "QtCore.framework").exists()
    assert not (fw / "PySide6" / "QtQml.abi3.so").exists()
    assert (fw / "PySide6" / "QtWidgets.abi3.so").exists()
    assert not (res / "roto").exists() and not (res / "roto").is_symlink()


def test_prune_main_cli(fake_dist: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert prune_qt.main(["prune_qt.py", str(fake_dist)]) == 0
    assert "prune_qt: liberados" in capsys.readouterr().out
    assert prune_qt.main(["prune_qt.py", str(fake_dist / "nope")]) == 1
    assert prune_qt.main(["prune_qt.py"]) == 2


def test_prune_never_drop_regex() -> None:
    for name in (
        "libQt6Core.so.6",
        "Qt6Gui.dll",
        "QtWidgets.abi3.so",
        "Qt6Svg.dll",
        "libQt6XcbQpa.so.6",
        "Qt6Network.dll",
        "Qt6OpenGL.dll",
        "libQt6DBus.so.6",
    ):
        assert prune_qt.NEVER_DROP.match(name), name
    for name in (
        "Qt6Qml.dll",
        "Qt6Quick.dll",
        "Qt6Pdf.dll",
        "QtPdf.pyd",
        "libQt6WaylandClient.so.6",
        "Qt6Svg.dll",
    ):
        assert prune_qt.LIB_DROP.match(name) or prune_qt.NEVER_DROP.match(name), name
    assert not prune_qt.LIB_DROP.match("libicuuc.so.73")
    assert not prune_qt.LIB_DROP.match("python313.dll")


# --------------------------------------------------------------------------- fetch_tesseract


def test_fetch_dry_run_lists_urls_and_hashes(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    assert fetch_tesseract.main(tmp_path / "res", platform="linux", dry_run=True) == 0
    out = capsys.readouterr().out.strip().splitlines()
    assert len(out) == 5
    for line in out:
        sha, url, _arrow, rel = line.split()
        assert (
            re.fullmatch(r"[0-9a-f]{64}", sha) and url.startswith("https://") and rel.startswith("tessdata/")
        )
    assert not (tmp_path / "res").exists()  # no descargó nada
    assert any("spa.traineddata" in line for line in out)
    assert any("osd.traineddata" in line for line in out)
    assert any("eng.traineddata" in line for line in out)
    assert any("pdf.ttf" in line for line in out) and any("configs/pdf" in line for line in out)


def test_fetch_dry_run_win32_adds_installer(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    assert fetch_tesseract.main(tmp_path, platform="win32", dry_run=True) == 0
    lines = capsys.readouterr().out.strip().splitlines()
    assert len(lines) == 6 and "tesseract-ocr-w64-setup-5." in lines[-1] and "->  tesseract/" in lines[-1]
    sha, _url, _arrow, _rel = lines[-1].split()
    assert re.fullmatch(r"[0-9a-f]{64}", sha), "el sha256 del instalador de Windows tiene que estar fijado"


def test_fetch_constants_are_pinned() -> None:
    for item in fetch_tesseract.TESSDATA_DOWNLOADS:
        assert item.verified, item
        assert "/tessdata_fast/4." in item.url or "/tesseract/5." in item.url
        assert "master" not in item.url and "main" not in item.url
    assert re.fullmatch(r"5\.\d+\.\d+\.\d{8}", fetch_tesseract.TESSERACT_WIN_VERSION)
    assert fetch_tesseract.TESSERACT_WIN_VERSION in fetch_tesseract.TESSERACT_WIN_URL
    assert fetch_tesseract.TESSERACT_WIN_URL.startswith("https://github.com/UB-Mannheim/tesseract/releases/")


def test_fetch_windows_installer_hash_is_pinned(monkeypatch: pytest.MonkeyPatch) -> None:
    """Sin sha256 real el job build-windows falla en cada corrida (fetch_one exige un hash válido)."""
    monkeypatch.delenv("PDF2EXCEL_TESSERACT_WIN_URL", raising=False)
    monkeypatch.delenv("PDF2EXCEL_TESSERACT_WIN_SHA256", raising=False)
    item = fetch_tesseract.windows_download()
    assert item.verified, item.sha256
    assert item.sha256 == fetch_tesseract.TESSERACT_WIN_SHA256 != "PENDIENTE"
    assert all(d.verified for d in fetch_tesseract.downloads_for("win32"))


def test_fetch_one_skips_when_hash_matches_and_fails_on_mismatch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    data = b"hola tessdata"
    sha = __import__("hashlib").sha256(data).hexdigest()
    item = fetch_tesseract.Download("https://example.invalid/x.traineddata", sha, "tessdata/x.traineddata")
    calls: list[str] = []

    def fake_download(url: str, target: Path, timeout: float = 0) -> None:
        calls.append(url)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)

    monkeypatch.setattr(fetch_tesseract, "download_to", fake_download)
    assert fetch_tesseract.fetch_one(item, tmp_path).read_bytes() == data
    assert calls == [item.url]
    assert fetch_tesseract.fetch_one(item, tmp_path) == tmp_path / "tessdata" / "x.traineddata"
    assert calls == [item.url]  # idempotente: no vuelve a bajar

    bad = fetch_tesseract.Download(
        "https://example.invalid/y.traineddata", "0" * 64, "tessdata/y.traineddata"
    )
    with pytest.raises(RuntimeError, match="NO coincide"):
        fetch_tesseract.fetch_one(bad, tmp_path)
    assert not (tmp_path / "tessdata" / "y.traineddata").exists()  # borra el archivo malo

    pending = fetch_tesseract.Download("https://example.invalid/z.exe", "PENDIENTE", "tesseract/z.exe")
    with pytest.raises(RuntimeError, match="sha256 no definido"):
        fetch_tesseract.fetch_one(pending, tmp_path)


def test_fetch_windows_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(
        "PDF2EXCEL_TESSERACT_WIN_URL", "https://mirror.local/tesseract-ocr-w64-setup-5.3.4.exe"
    )
    monkeypatch.setenv("PDF2EXCEL_TESSERACT_WIN_SHA256", "A" * 64)
    item = fetch_tesseract.windows_download()
    assert item.url.startswith("https://mirror.local/") and item.sha256 == "a" * 64 and item.verified
    assert item.relpath == "tesseract/tesseract-ocr-w64-setup-5.3.4.exe"


def test_fetch_extract_windows_installer_filters_files(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def fake_run(cmd: list[str], **_kw: object) -> None:
        out = Path(cmd[3][2:])
        _touch(out / "tesseract.exe")
        _touch(out / "libtesseract-5.dll")
        _touch(out / "uninstall.exe")
        _touch(out / "$PLUGINSDIR" / "nsis.dll")
        _touch(out / "tessdata" / "eng.traineddata")
        _touch(out / "doc" / "README.html")

    monkeypatch.setattr(fetch_tesseract, "find_7z", lambda: "7z")
    monkeypatch.setattr(fetch_tesseract.subprocess, "run", fake_run)
    n = fetch_tesseract.extract_windows_installer(tmp_path / "setup.exe", tmp_path / "tesseract")
    assert n == 2
    assert sorted(p.name for p in (tmp_path / "tesseract").iterdir()) == [
        "libtesseract-5.dll",
        "tesseract.exe",
    ]


def test_fetch_requires_7z_message(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(fetch_tesseract, "find_7z", lambda: None)
    with pytest.raises(RuntimeError, match="7z"):
        fetch_tesseract.extract_windows_installer(tmp_path / "setup.exe", tmp_path / "t")


# -------------------------------------------------------------------------- spec / iss / workflow / gitignore


def test_spec_contents() -> None:
    spec = (PACKAGING / "conversor.spec").read_text(encoding="utf-8")
    assert 'collect_data_files("pymupdf")' in spec and "OBLIGATORIO" in spec
    assert "theme.qss" in spec and '"pdf2excel/ui"' in spec  # app.py lo busca en Path(__file__).parent/ui
    assert 'collect_dynamic_libs("onnxruntime")' in spec
    assert "name=APP_NAME" in spec and 'APP_NAME = "ConversorPDFaExcel"' in spec
    assert "console=False" in spec and "optimize=1" in spec
    assert "bundle_identifier=BUNDLE_ID" in spec and 'BUNDLE_ID = "ar.com.conversorpdf.app"' in spec
    assert '"NSHighResolutionCapable": True' in spec
    for mod in (
        "QtWebEngineCore",
        "QtQuick",
        "QtQml",
        "Qt3DCore",
        "QtMultimedia",
        "QtPdf",
        "QtCharts",
        "QtSql",
        "QtTest",
        "QtXml",
        "QtSvgWidgets",
        "QtConcurrent",
    ):
        assert f'"{mod}"' in spec, mod
    for pkg in (
        "pandas",
        "camelot",
        "pypdfium2",
        "pdfplumber",
        "pdfminer",
        "lxml",
        "matplotlib",
        "scipy",
        "IPython",
        "tkinter",
        "rapidocr_onnxruntime",
        "pytest",
        "reportlab",
        "pymupdf4llm",
        "PIL.ImageQt",
    ):
        assert f'"{pkg}"' in spec, pkg
    for hidden in ("onnxruntime", "openpyxl.cell._writer", "cv2", "platformdirs"):
        assert f'"{hidden}"' in spec, hidden
    compile(spec, "conversor.spec", "exec")  # sintaxis Python válida


def test_installer_iss_contents() -> None:
    iss = (PACKAGING / "installer.iss").read_text(encoding="utf-8")
    assert '#define AppName "Conversor PDF a Excel"' in iss
    assert '#define AppExe "ConversorPDFaExcel.exe"' in iss
    guid = re.search(r'#define AppId "\{\{([0-9A-F-]{36})\}"', iss)
    assert guid, "AppId debe ser un GUID con doble llave"
    assert re.fullmatch(r"[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}", guid.group(1))
    for line in (
        "PrivilegesRequired=lowest",
        "PrivilegesRequiredOverridesAllowed=dialog",
        "CloseApplications=yes",
        "Compression=lzma2/ultra64",
        "OutputBaseFilename=ConversorPDFaExcel-{#AppVersion}-setup-win64",
        "AppPublisher=",
        "UninstallDisplayIcon={app}\\{#AppExe}",
        'MessagesFile: "compiler:Languages\\Spanish.isl"',
        'Name: "desktopicon"',
        'Name: "sendto"',
        "{usersendto}",
        'Type: filesandordirs; Name: "{app}\\_internal"',
    ):
        assert line in iss, line
    # cada sección referenciada existe
    for section in ("[Setup]", "[Languages]", "[Tasks]", "[Files]", "[Icons]", "[Run]", "[InstallDelete]"):
        assert section in iss
    # las directivas de [Setup] son clave=valor sin espacios raros
    setup = iss.split("[Setup]")[1].split("[Languages]")[0]
    for raw in setup.splitlines():
        line = raw.strip()
        if line and not line.startswith(";"):
            assert re.fullmatch(r"[A-Za-z0-9]+=.*", line), line


def test_workflow_pins_actions_and_has_required_jobs() -> None:
    import yaml

    text = WORKFLOW.read_text(encoding="utf-8")
    wf = yaml.safe_load(text)
    assert set(wf["jobs"]) >= {"test", "build-windows", "build-macos", "release", "pip-audit"}
    assert wf["defaults"]["run"]["working-directory"] == "desktop"
    on = wf[True] if True in wf else wf["on"]  # PyYAML parsea "on" como True
    assert (
        "desktop/**" in on["push"]["paths"]
        and "v*" in on["push"]["tags"]
        and "main" in on["push"]["branches"]
    )
    assert "workflow_dispatch" in on and "pull_request" in on and "schedule" in on
    uses = re.findall(r"uses:\s*(\S+)", text)
    assert uses
    for action in uses:
        assert re.fullmatch(r"[\w.-]+/[\w.-]+@v\d+\.\d+\.\d+", action), f"acción sin versión fija: {action}"
    release_if = wf["jobs"]["release"]["if"]
    assert "always()" in release_if and "needs.build-windows.result == 'success'" in release_if
    assert "startsWith(github.ref, 'refs/tags/v')" in release_if and "inputs.release_tag" in release_if
    assert on["workflow_dispatch"]["inputs"]["release_tag"]["required"] is False
    assert "RELEASE_TAG" in wf["env"]
    assert wf["jobs"]["release"]["needs"] == ["build-windows", "build-macos"]
    assert wf["jobs"]["build-windows"]["runs-on"] == "windows-2025"
    assert wf["jobs"]["build-macos"]["runs-on"] == "macos-14"
    assert wf["jobs"]["test"]["runs-on"] == "ubuntu-24.04"
    for snippet in (
        "packaging/fetch_tesseract.py resources",
        'packaging/stamp_version.py "$RELEASE_TAG"',
        "pyinstaller --noconfirm --clean packaging/conversor.spec",
        "packaging/prune_qt.py",
        "--selftest",
        "selftest ok",
        "ISCC.exe",
        "7z a -mx=7",
        "signtool",
        "timestamp.digicert.com",
        "hdiutil create",
        "notarytool",
        "latest.json",
        "softprops/action-gh-release",
        "pip-audit",
        "ruff check src tests",
        "mypy src",
        "--runslow --corpus-report corpus_report.md",
        "GITHUB_STEP_SUMMARY",
        "tesseract-ocr-spa",
        "xvfb",
        "libegl1",
        "PDF2EXCEL_TESSERACT_WIN_URL: ${{ vars.TESSERACT_WIN_URL }}",
        "PDF2EXCEL_TESSERACT_WIN_SHA256: ${{ vars.TESSERACT_WIN_SHA256 }}",
    ):
        assert snippet in text, snippet


def test_gitignore_entries() -> None:
    gi = (DESKTOP / ".gitignore").read_text(encoding="utf-8").splitlines()
    for entry in (
        "__pycache__/",
        "*.egg-info/",
        ".pytest_cache/",
        ".ruff_cache/",
        ".mypy_cache/",
        "build/",
        "dist/",
        "resources/tessdata/",
        "resources/tesseract/",
        "tests/corpus/*.pdf",
        "tests/corpus_real/",
        "tests/corpus/_generated/",
        "*.spec.bak",
    ):
        assert entry in gi, entry


def test_packaging_readme_covers_required_topics() -> None:
    readme = (PACKAGING / "README.md").read_text(encoding="utf-8").lower()
    for topic in (
        "windows",
        "firmar",
        "git tag v0.1.0",
        "git push --tags",
        "smartscreen",
        "tamaño",
        "agpl",
        "fetch_tesseract",
        "sha256",
        "stamp_version",
        "prune_qt",
        "latest.json",
    ):
        assert topic in readme, topic
