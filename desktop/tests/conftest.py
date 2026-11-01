"""Configuración compartida de pytest: corpus sintético, ground truth, OCR y reporte de CI.

- ``--runslow``: corre también los tests marcados ``slow`` (OCR, corpus escaneado). En CI (``CI=true``)
  corren siempre.
- ``--corpus-report PATH``: al terminar escribe la tabla markdown archivo × métricas del corpus.
- ``corpus_dir`` usa ``tests/corpus/_generated/`` si está completo (CI lo genera); si no, genera el corpus una
  vez por sesión en ``tmp_path_factory``.
"""

from __future__ import annotations

import json
import os
import sys
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

TESTS_DIR = Path(__file__).resolve().parent
DESKTOP_DIR = TESTS_DIR.parent
CORPUS_DIR = TESTS_DIR / "corpus"
GENERATED_DIR = CORPUS_DIR / "_generated"
GROUND_TRUTH_PATH = CORPUS_DIR / "ground_truth.json"
CORPUS_DIR_ENV = "PDF2EXCEL_CORPUS_DIR"  # override explícito de la carpeta del corpus (ya generado)

# `tests.scorer` y `tests.corpus.generate` se importan como paquetes namespace desde la raíz de desktop/.
if str(DESKTOP_DIR) not in sys.path:
    sys.path.insert(0, str(DESKTOP_DIR))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

CORPUS_RESULTS_KEY = pytest.StashKey[dict[str, Any]]()  # nombre de PDF → Score
CORPUS_META_KEY = pytest.StashKey[
    dict[str, dict[str, Any]]
]()  # nombre de PDF → {"ocr": bool, "seconds": float}


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--runslow",
        action="store_true",
        default=False,
        help="corre también los tests lentos (OCR, corpus escaneado); en CI (CI=true) corren siempre",
    )
    parser.addoption(
        "--corpus-report",
        action="store",
        default=None,
        metavar="PATH",
        help="escribe la tabla markdown del corpus (celdas %, filas %, espurias %, tipos %, tablas, OCR, s)",
    )


def _in_ci() -> bool:
    return os.environ.get("CI", "").strip().lower() in {"true", "1", "yes"}


def slow_enabled(config: pytest.Config) -> bool:
    return bool(config.getoption("--runslow")) or _in_ci()


def pytest_configure(config: pytest.Config) -> None:
    config.stash[CORPUS_RESULTS_KEY] = {}
    config.stash[CORPUS_META_KEY] = {}


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if slow_enabled(config):
        return
    skip = pytest.mark.skip(reason="test lento: corré pytest con --runslow (en CI corre siempre)")
    for item in items:
        if "slow" in item.keywords:
            item.add_marker(skip)


def pytest_sessionfinish(session: pytest.Session, exitstatus: int) -> None:
    target = session.config.getoption("--corpus-report")
    if not target:
        return
    from tests.scorer import format_table

    results = session.config.stash.get(CORPUS_RESULTS_KEY, {})
    meta = session.config.stash.get(CORPUS_META_KEY, {})
    path = Path(target)
    path.parent.mkdir(parents=True, exist_ok=True)
    note = "" if results else "_Sin resultados: no corrió test_corpus_accuracy._\n\n"
    path.write_text(note + format_table(results, meta), encoding="utf-8")


# --------------------------------------------------------------------------- fixtures


def _corpus_complete(directory: Path) -> bool:
    from tests.corpus.generate import ALL_NAMES

    return directory.is_dir() and all((directory / f"{name}.pdf").is_file() for name in ALL_NAMES)


@pytest.fixture(scope="session")
def corpus_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Carpeta con el corpus sintético completo (se genera una vez por sesión si hace falta)."""
    override = os.environ.get(CORPUS_DIR_ENV)
    if override:
        candidate = Path(override)
        if _corpus_complete(candidate):
            return candidate
        pytest.fail(
            f"{CORPUS_DIR_ENV}={override} no tiene el corpus completo; generalo con tests.corpus.generate"
        )
    if _corpus_complete(GENERATED_DIR):
        return GENERATED_DIR
    from tests.corpus.generate import generate_all

    out = tmp_path_factory.mktemp("corpus")
    generate_all(out)
    return out


@pytest.fixture(scope="session")
def ground_truth() -> dict[str, Any]:
    with GROUND_TRUTH_PATH.open(encoding="utf-8") as fh:
        data: dict[str, Any] = json.load(fh)
    return data


@pytest.fixture(scope="session")
def has_tesseract() -> bool:
    from pdf2excel.core.ocr import ocr_available

    return ocr_available()


@pytest.fixture(scope="session")
def ocr_cache_dir(tmp_path_factory: pytest.TempPathFactory) -> Iterator[Path]:
    """Caché de OCR aislada para la sesión: los tests nunca tocan la caché del usuario.

    Se usa la variable ``PDF2EXCEL_CACHE_DIR`` (la misma que honra ``--selftest``) en vez de reemplazar
    ``default_cache_dir``: así los tests de esa función siguen probando el código real.
    """
    import pdf2excel.core.ocr as ocr_mod

    cache = tmp_path_factory.mktemp("ocr-cache")
    patch = pytest.MonkeyPatch()
    patch.setenv(ocr_mod.CACHE_DIR_ENV, str(cache))
    try:
        yield cache
    finally:
        patch.undo()


@pytest.fixture(scope="session")
def corpus_results(request: pytest.FixtureRequest) -> dict[str, Any]:
    """Acumulador nombre → Score que alimenta ``--corpus-report``."""
    return request.config.stash[CORPUS_RESULTS_KEY]


@pytest.fixture(scope="session")
def corpus_meta(request: pytest.FixtureRequest) -> dict[str, dict[str, Any]]:
    """Acumulador nombre → {"ocr": bool, "seconds": float} para el reporte."""
    return request.config.stash[CORPUS_META_KEY]
