"""Descarga, con sha256 fijado, lo que la app vendorea para el OCR y que NO se commitea.

- En todas las plataformas: ``tessdata_fast`` ``spa``/``eng``/``osd`` + ``pdf.ttf`` + ``configs/pdf`` del
  repo de tesseract (hacen falta para ``tesseract … -c textonly_pdf=1 pdf``) → ``<dest>/tessdata/``.
- Solo Windows: el instalador NSIS de Tesseract 5.3.x (build UB-Mannheim) se extrae con ``7z x`` y se
  copian ``tesseract.exe`` + las DLL → ``<dest>/tesseract/``.

Uso (desde ``desktop/``)::

    python packaging/fetch_tesseract.py resources                 # descarga lo que falte
    python packaging/fetch_tesseract.py resources --dry-run       # solo lista URLs y sha256
    python packaging/fetch_tesseract.py resources --platform win32
    python packaging/fetch_tesseract.py resources --print-hashes  # descarga a tmp y muestra los sha256 reales

Es idempotente: si un archivo ya está y su sha256 coincide, se salta. Si el sha256 de una descarga NO
coincide con la constante, falla (exit 1) y borra el archivo. Solo stdlib (+ ``7z`` en Windows).

Cómo actualizar las constantes: ver ``packaging/README.md`` ("Actualizar Tesseract / tessdata").
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path

# --------------------------------------------------------------------------- constantes fijas

#: Tag de tesseract-ocr/tessdata_fast del que se bajan los .traineddata (los archivos no cambiaron desde 4.0).
TESSDATA_FAST_TAG = "4.1.0"
TESSDATA_FAST_BASE = f"https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/{TESSDATA_FAST_TAG}/"
#: Tag de tesseract-ocr/tesseract del que se bajan ``pdf.ttf`` y ``configs/pdf``.
TESSERACT_SRC_TAG = "5.3.4"
TESSERACT_SRC_BASE = (
    f"https://raw.githubusercontent.com/tesseract-ocr/tesseract/{TESSERACT_SRC_TAG}/tessdata/"
)

#: Build de Tesseract para Windows (UB-Mannheim, instalador NSIS de 64 bits).
#: La URL/sha256 se pueden pisar con ``PDF2EXCEL_TESSERACT_WIN_URL`` / ``PDF2EXCEL_TESSERACT_WIN_SHA256``
#: (el workflow las toma de las variables del repo ``TESSERACT_WIN_URL`` / ``TESSERACT_WIN_SHA256``).
#:
#: Se usa 5.4.0.20240606 y no 5.3.x porque UB-Mannheim publicó en GitHub Releases el instalador de 5.4.0
#: pero NO los de 5.3.3/5.3.4 (esas URLs devuelven 404; los viejos viven solo en digi.bib.uni-mannheim.de).
#: Desvío documentado de DESIGN §0/§8 ("5.3.x"): mismo build, misma CLI (--psm, textonly_pdf).
TESSERACT_WIN_VERSION = "5.4.0.20240606"
TESSERACT_WIN_URL = (
    "https://github.com/UB-Mannheim/tesseract/releases/download/"
    f"v{TESSERACT_WIN_VERSION}/tesseract-ocr-w64-setup-{TESSERACT_WIN_VERSION}.exe"
)
#: sha256 del instalador (50 175 248 bytes), calculado sobre dos descargas independientes del asset.
TESSERACT_WIN_SHA256 = "c885fff6998e0608ba4bb8ab51436e1c6775c2bafc2559a19b423e18678b60c9"


@dataclass(frozen=True)
class Download:
    """Un archivo a bajar: URL, sha256 esperado (hex minúsculas) y ruta relativa dentro de ``dest``."""

    url: str
    sha256: str
    relpath: str  # relativo a <dest_resources>

    @property
    def verified(self) -> bool:
        return len(self.sha256) == 64 and all(c in "0123456789abcdef" for c in self.sha256)


#: Archivos de OCR comunes a todas las plataformas (sha256 verificados contra tessdata_fast 4.1.0 y
#: tesseract 5.3.4; coinciden con los paquetes tesseract-ocr-spa/eng de Debian/Ubuntu).
TESSDATA_DOWNLOADS: tuple[Download, ...] = (
    Download(
        TESSDATA_FAST_BASE + "spa.traineddata",
        "6f2e04d02774a18f01bed44b1111f2cd7f3ba7ac9dc4373cd3f898a40ea6b464",
        "tessdata/spa.traineddata",
    ),
    Download(
        TESSDATA_FAST_BASE + "eng.traineddata",
        "7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2",
        "tessdata/eng.traineddata",
    ),
    Download(
        TESSDATA_FAST_BASE + "osd.traineddata",
        "9cf5d576fcc47564f11265841e5ca839001e7e6f38ff7f7aacf46d15a96b00ff",
        "tessdata/osd.traineddata",
    ),
    Download(
        TESSERACT_SRC_BASE + "pdf.ttf",
        "c7845420925a23d88ed830a63957b8af85a66a8daf8d9fc90e843673b2ef1a59",
        "tessdata/pdf.ttf",
    ),
    Download(
        TESSERACT_SRC_BASE + "configs/pdf",
        "54d56e81dfefe289b0d2e4c7bc9bbe662711f5b8255c128b73bcd66efb6bba1a",
        "tessdata/configs/pdf",
    ),
)

#: Qué se copia del instalador extraído a ``resources/tesseract/`` (lo demás: docs, uninstaller, tessdata).
_WIN_KEEP_SUFFIXES = (".exe", ".dll")
_WIN_SKIP_NAMES = {"uninstall.exe", "uninstall-tesseract.exe"}
_WIN_SKIP_DIRS = {"$pluginsdir", "doc", "tessdata", "training"}


def windows_download() -> Download:
    """Instalador de Windows (URL y sha256 de las constantes o del entorno)."""
    url = os.environ.get("PDF2EXCEL_TESSERACT_WIN_URL", "").strip() or TESSERACT_WIN_URL
    sha = os.environ.get("PDF2EXCEL_TESSERACT_WIN_SHA256", "").strip().lower() or TESSERACT_WIN_SHA256
    return Download(url, sha, "tesseract/" + url.rsplit("/", 1)[-1])


def downloads_for(platform: str) -> list[Download]:
    items = list(TESSDATA_DOWNLOADS)
    if platform.startswith("win"):
        items.append(windows_download())
    return items


# --------------------------------------------------------------------------- utilidades


def sha256_of(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _log(msg: str) -> None:
    print(f"fetch_tesseract: {msg}", flush=True)


def download_to(url: str, target: Path, timeout: float = 120.0) -> None:
    """Baja ``url`` a ``target`` pasando por un archivo temporal (no deja parciales)."""
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(target.suffix + ".part")
    request = urllib.request.Request(url, headers={"User-Agent": "pdf2excel-fetch-tesseract"})
    with urllib.request.urlopen(request, timeout=timeout) as response, tmp.open("wb") as out:
        shutil.copyfileobj(response, out, 1 << 20)
    tmp.replace(target)


def fetch_one(item: Download, dest: Path) -> Path:
    """Descarga (si hace falta) y verifica un archivo. Lanza ``RuntimeError`` si el sha256 no coincide."""
    if not item.verified:
        raise RuntimeError(
            f"sha256 no definido para {item.url} ({item.sha256!r}). Completá la constante en "
            "packaging/fetch_tesseract.py (ver README: 'Actualizar Tesseract / tessdata')."
        )
    target = dest / item.relpath
    if target.is_file() and sha256_of(target) == item.sha256:
        _log(f"ya está: {item.relpath}")
        return target
    _log(f"bajando {item.url}")
    download_to(item.url, target)
    actual = sha256_of(target)
    if actual != item.sha256:
        target.unlink(missing_ok=True)
        raise RuntimeError(
            f"sha256 NO coincide para {item.url}\n  esperado {item.sha256}\n  obtenido {actual}"
        )
    _log(f"ok: {item.relpath} ({target.stat().st_size / 1048576:.1f} MB)")
    return target


def find_7z() -> str | None:
    for name in ("7z", "7za", "7zr"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (r"C:\Program Files\7-Zip\7z.exe", r"C:\Program Files (x86)\7-Zip\7z.exe"):
        if Path(candidate).is_file():
            return candidate
    return None


def extract_windows_installer(installer: Path, tesseract_dir: Path) -> int:
    """Extrae el NSIS con ``7z x`` y copia ``tesseract.exe`` + DLL a ``tesseract_dir``; devuelve cuántos."""
    seven_zip = find_7z()
    if seven_zip is None:
        raise RuntimeError("No se encontró 7z (hace falta para extraer el instalador NSIS de Tesseract).")
    with tempfile.TemporaryDirectory(prefix="tesseract-nsis-") as tmpdir:
        subprocess.run(
            [seven_zip, "x", "-y", f"-o{tmpdir}", str(installer)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
        )
        copied = 0
        for src in Path(tmpdir).rglob("*"):
            if not src.is_file() or src.suffix.lower() not in _WIN_KEEP_SUFFIXES:
                continue
            rel_parts = {p.lower() for p in src.relative_to(tmpdir).parts[:-1]}
            if src.name.lower() in _WIN_SKIP_NAMES or rel_parts & _WIN_SKIP_DIRS:
                continue
            tesseract_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, tesseract_dir / src.name)
            copied += 1
    if not (tesseract_dir / "tesseract.exe").is_file():
        raise RuntimeError("El instalador se extrajo pero no apareció tesseract.exe")
    return copied


def have_windows_binary(dest: Path) -> bool:
    return (dest / "tesseract" / "tesseract.exe").is_file()


# --------------------------------------------------------------------------- entrada


def main(dest_resources: Path, platform: str = sys.platform, dry_run: bool = False) -> int:
    """Descarga/verifica todo para ``platform`` dentro de ``dest_resources``. Devuelve 0 si está todo."""
    dest = Path(dest_resources)
    items = downloads_for(platform)
    if dry_run:
        for item in items:
            print(f"{item.sha256}  {item.url}  ->  {item.relpath}")
        return 0

    for item in TESSDATA_DOWNLOADS:
        fetch_one(item, dest)

    if platform.startswith("win"):
        if have_windows_binary(dest):
            _log("ya está: tesseract/tesseract.exe")
        else:
            installer = fetch_one(windows_download(), dest)
            n = extract_windows_installer(installer, dest / "tesseract")
            installer.unlink(missing_ok=True)
            _log(f"ok: tesseract/ ({n} archivos)")
    return 0


def print_hashes(platform: str) -> int:
    """Baja cada archivo a un temporal y muestra su sha256 real (para actualizar las constantes)."""
    with tempfile.TemporaryDirectory(prefix="fetch-tesseract-hashes-") as tmpdir:
        for item in downloads_for(platform):
            target = Path(tmpdir) / Path(item.relpath).name
            _log(f"bajando {item.url}")
            download_to(item.url, target)
            print(f"{sha256_of(target)}  {item.url}")
    return 0


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Descarga tessdata (y tesseract.exe en Windows) con sha256 fijo."
    )
    parser.add_argument("dest", type=Path, help="carpeta resources/ de destino (crea tessdata/ y tesseract/)")
    parser.add_argument(
        "--platform", default=sys.platform, help="win32 | darwin | linux (default: el actual)"
    )
    parser.add_argument("--dry-run", action="store_true", help="solo listar URLs y sha256, sin descargar")
    parser.add_argument("--print-hashes", action="store_true", help="descargar a tmp y mostrar sha256 reales")
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = _parse_args(sys.argv[1:])
    try:
        if args.print_hashes:
            sys.exit(print_hashes(args.platform))
        sys.exit(main(args.dest, args.platform, args.dry_run))
    except (RuntimeError, OSError, subprocess.CalledProcessError) as exc:
        print(f"fetch_tesseract: ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
