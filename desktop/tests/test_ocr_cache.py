"""Caché de OCR: ubicación (``PDF2EXCEL_CACHE_DIR``) y poda (``prune_cache``). Puros, sin tesseract."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from pdf2excel.core import ocr


def test_default_cache_dir_env_override(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv(ocr.CACHE_DIR_ENV, str(tmp_path / "mi_cache"))
    assert ocr.default_cache_dir() == tmp_path / "mi_cache" / "ocr"
    monkeypatch.setenv(ocr.CACHE_DIR_ENV, "   ")  # en blanco = sin override
    assert "mi_cache" not in str(ocr.default_cache_dir())
    monkeypatch.delenv(ocr.CACHE_DIR_ENV)
    assert ocr.default_cache_dir().name == "ocr" and "mi_cache" not in str(ocr.default_cache_dir())


def test_cache_path_for_honours_env_override(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    src = tmp_path / "a.pdf"
    src.write_bytes(b"%PDF-1.4 stub")
    monkeypatch.setenv(ocr.CACHE_DIR_ENV, str(tmp_path / "c"))
    assert ocr.cache_path_for(src, [0], "spa").parent == tmp_path / "c" / "ocr"


def test_prune_cache_by_age_and_size(tmp_path: Path) -> None:
    cache = tmp_path / "ocr"
    cache.mkdir()
    now = 1_700_000_000.0
    day = 86400
    files = {
        "viejo.pdf": (now - 40 * day, 10),
        "a.pdf": (now - 5 * day, 100),
        "b.pdf": (now - 3 * day, 100),
        "c.pdf": (now - 1 * day, 100),
        "nuevo.pdf": (now, 100),
        "otro.part": (now - 90 * day, 100),  # no es .pdf: no se toca
    }
    for name, (mtime, size) in files.items():
        path = cache / name
        path.write_bytes(b"x" * size)
        os.utime(path, (mtime, mtime))
    removed = ocr.prune_cache(cache, max_bytes=250, max_age_days=30, now=now)
    assert removed == 3  # viejo.pdf por edad; a.pdf y b.pdf por tamaño (los menos recientes)
    assert sorted(p.name for p in cache.iterdir()) == ["c.pdf", "nuevo.pdf", "otro.part"]
    assert ocr.prune_cache(cache, max_bytes=250, max_age_days=30, now=now) == 0
    assert ocr.prune_cache(tmp_path / "no_existe") == 0  # nunca lanza


def test_prune_cache_defaults_are_bounded() -> None:
    assert ocr.CACHE_MAX_BYTES == 500 * 1024 * 1024
    assert ocr.CACHE_MAX_AGE_DAYS == 30
