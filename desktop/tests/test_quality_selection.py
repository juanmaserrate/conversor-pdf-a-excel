"""El puntaje sin ground truth (``select_best``) elige el candidato correcto (DESIGN.md §5.4 y §9).

Para cada página de los PDFs con GT: ``page_candidates`` → ``select_best``; cada candidato se tipa con
``type_table`` y se puntúa con el scorer. El elegido debe tener ``cells_pct`` ≥ máximo − 0,5: el puntaje
sin GT correlaciona con el real. Los escaneos (``slow``, requieren tesseract) se evalúan sobre el PDF
OCR'd, que es lo que ve el pipeline.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from tests.corpus import generate as gen
from tests.scorer import gt_for, score

from pdf2excel.core.candidates import page_candidates
from pdf2excel.core.document import classify_document, needs_ocr, open_document
from pdf2excel.core.ocr import ocr_document
from pdf2excel.core.quality import select_best
from pdf2excel.core.typing_ar import type_table
from pdf2excel.models import RawTable

TOLERANCE = 0.5
_GT_PATH = Path(__file__).resolve().parent / "corpus" / "ground_truth.json"
_GT: dict[str, Any] = json.loads(_GT_PATH.read_text(encoding="utf-8"))


def _digital_pages() -> list[tuple[str, int]]:
    """(nombre, página) de los PDFs digitales con tablas que page_candidates puede abrir sin contraseña."""
    out: list[tuple[str, int]] = []
    for name in gen.DIGITAL_NAMES:
        entry = _GT[f"{name}.pdf"]
        if not entry["tables"] or "password" in entry:
            continue
        out.extend((name, i) for i, kind in enumerate(entry["kinds"]) if kind == "digital")
    return out


def _scan_pages() -> list[tuple[str, int]]:
    out: list[tuple[str, int]] = []
    for name in gen.SCAN_NAMES:
        base = gt_for(name, _GT)
        out.extend((name, i) for i in range(len(base["kinds"])))
    return out


def _cells_pct(tables: list[RawTable], gt: dict[str, Any]) -> float:
    return score([type_table(t) for t in tables], gt).cells_pct


def _check_selection(path: Path, page_no: int, gt: dict[str, Any]) -> None:
    cands = page_candidates(path, page_no)
    chosen, chosen_tables = select_best(cands)
    if not any(cands.values()):
        pytest.skip(f"{path.name} p.{page_no + 1}: ningún candidato encontró tablas")
    scores = {name: _cells_pct(tables, gt) for name, tables in cands.items()}
    best = max(scores.values())
    shapes = {k: [(len(t.rows), len(t.rows[0]) if t.rows else 0) for t in v] for k, v in cands.items()}
    detail = f"{path.name} p.{page_no + 1}: elegido {chosen!r} ({scores[chosen]:.2f} %) vs {scores}; " + (
        f"formas={shapes}"
    )
    assert chosen in cands and chosen_tables == cands[chosen], detail
    assert scores[chosen] >= best - TOLERANCE, detail


@pytest.mark.parametrize(("name", "page_no"), _digital_pages())
def test_selected_candidate_is_best_digital(
    name: str, page_no: int, corpus_dir: Path, ground_truth: dict[str, Any]
) -> None:
    _check_selection(corpus_dir / f"{name}.pdf", page_no, gt_for(name, ground_truth))


@pytest.fixture(scope="module")
def ocr_pdf(corpus_dir: Path, has_tesseract: bool, ocr_cache_dir: Path) -> Any:
    """Devuelve la ruta del PDF OCR'd de un escaneo del corpus (cacheado en la sesión)."""
    if not has_tesseract:
        pytest.skip("tesseract no disponible")

    def _get(name: str) -> Path:
        src = corpus_dir / f"{name}.pdf"
        with open_document(src) as doc:
            pages = [p.index for p in classify_document(doc) if needs_ocr(p)]
        assert pages, f"{name} no necesita OCR: no es un escaneo"
        return ocr_document(src, pages, cache_dir=ocr_cache_dir)

    return _get


@pytest.mark.slow
@pytest.mark.parametrize(("name", "page_no"), _scan_pages())
def test_selected_candidate_is_best_scanned(
    name: str, page_no: int, ocr_pdf: Any, ground_truth: dict[str, Any]
) -> None:
    _check_selection(ocr_pdf(name), page_no, gt_for(name, ground_truth))


def test_select_best_prefers_candidate_with_tables() -> None:
    raw = RawTable(
        page=0, page_end=0, bbox=(0, 0, 1, 1), rows=[["a", "b"], ["1", "2"]], source="hybrid", quality=0.9
    )
    assert select_best({"hybrid_union": [], "hybrid": [raw]}) == ("hybrid", [raw])
    assert select_best({"hybrid_union": [], "hybrid": []}) == ("hybrid", [])
