"""Fixtures pytest del corpus de FORMULARIOS (``tests/corpus/generate_forms.py``).

Se importan en los módulos de test que las usan::

    from tests.forms_fixtures import forms_corpus, forms_ground_truth  # noqa: F401
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from tests.corpus.generate_forms import GROUND_TRUTH_NAME, generate_forms


@pytest.fixture(scope="session")
def forms_corpus(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    """``nombre -> PDF`` de todos los formularios sintéticos (se generan una vez por sesión)."""
    return generate_forms(tmp_path_factory.mktemp("forms_corpus"))


@pytest.fixture(scope="session")
def forms_ground_truth(forms_corpus: dict[str, Path]) -> dict[str, Any]:
    """Ground truth de formularios: ``"<nombre>.pdf" -> {visible_words, covered_words, boxes, ...}``."""
    path = next(iter(forms_corpus.values())).parent / GROUND_TRUTH_NAME
    data: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    return data
