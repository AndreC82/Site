from pdf_takeoff.gib_spec_extract import (
    BoardSpec,
    extract_ceiling_legend,
    extract_stud_height_m,
    extract_wall_default,
)

SAMPLE_TEXT = """Wall Linings
10mm Gibboard internal linings (Aqualine to wet areas)
Refer to sheet A8/A10 for fire control requirements
Ceiling Linings 1
13mm Standard Gibboard fixed to Rondo steel ceiling
battens @ 600crs. Refer to sheets A8/A10 for fire control
requirements
Ceiling Linings 2
13mm Aqualine Gibboard fixed to Rondo steel ceiling
battens @ 600crs. Refer to sheets A8/A10 for fire control
requirements
Ceiling Linings 3
1/16mm Fyreline Gibboard installed as per Gibboard detail
GBUC30a and the fire report fixed to Rondo steel ceiling
battens @ 600crs. Refer to sheets A8/A10 for fire control
requirements
Insulation
2720mm Stud Height (Unless noted otherwise)
FFL: 36.700
"""


def test_extract_ceiling_legend_reads_thickness_product_and_layers():
    legend = extract_ceiling_legend(SAMPLE_TEXT)
    assert legend["C1"] == BoardSpec(thickness_mm=13, product="Standard", layers=1)
    assert legend["C2"] == BoardSpec(thickness_mm=13, product="Aqualine", layers=1)
    assert legend["C3"] == BoardSpec(thickness_mm=16, product="Fireline", layers=1)


def test_extract_wall_default_reads_standard_and_aqualine_override():
    default_spec, wet_spec = extract_wall_default(SAMPLE_TEXT)
    assert default_spec == BoardSpec(thickness_mm=10, product="Standard", layers=1)
    assert wet_spec == BoardSpec(thickness_mm=10, product="Aqualine", layers=1)


def test_extract_wall_default_returns_none_when_absent():
    default_spec, wet_spec = extract_wall_default("no relevant text here")
    assert default_spec is None
    assert wet_spec is None


def test_extract_stud_height_m():
    assert extract_stud_height_m(SAMPLE_TEXT) == 2.72


def test_extract_stud_height_m_returns_none_when_absent():
    assert extract_stud_height_m("no height info") is None


def test_board_spec_label_and_known_check():
    spec = BoardSpec(thickness_mm=10, product="Standard")
    assert spec.board_type_label == "10mm Standard"
    assert spec.is_known_board_type is True

    weird = BoardSpec(thickness_mm=7, product="Exotic")
    assert weird.is_known_board_type is False
