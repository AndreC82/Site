import json

import pytest

from pdf_takeoff.legend import LegendEntry, load_legend, match_code


def test_load_legend(tmp_path):
    data = {
        "P-01": {"category": "wall_paint", "description": "Branco"},
        "GB-FL-DL": {"category": "drywall", "description": "Fireline dupla", "layers": 2},
    }
    path = tmp_path / "legend.json"
    path.write_text(json.dumps(data), encoding="utf-8")

    legend = load_legend(str(path))
    assert legend["P-01"].category == "wall_paint"
    assert legend["GB-FL-DL"].layers == 2


def test_invalid_category_raises():
    with pytest.raises(ValueError):
        LegendEntry(code="X", category="nao_existe", description="")


def test_match_code_exact_and_prefix():
    legend = {"P-01": LegendEntry(code="P-01", category="wall_paint", description="Branco")}
    assert match_code("P-01", legend) is not None
    assert match_code("p-01", legend) is not None
    assert match_code("P-01(TIPICO)", legend) is not None
    assert match_code("P-99", legend) is None


def test_match_code_no_false_positive_on_unrelated_text():
    legend = {"P-01": LegendEntry(code="P-01", category="wall_paint", description="Branco")}
    assert match_code("SALA", legend) is None
    assert match_code("3.50", legend) is None
