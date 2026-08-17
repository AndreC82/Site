from pdf_takeoff.extract import TextWord
from pdf_takeoff.calibration import auto_detect_scale


def test_auto_detect_scale_with_consistent_dimensions():
    segments = [
        ((0, 0), (200, 0)),
        ((0, 0), (0, 150)),
        ((200, 0), (200, 150)),
        ((0, 150), (200, 150)),
    ]
    words = [
        TextWord(text="4.00", bbox=(90, -10, 110, 0)),  # perto do segmento de topo (200 unid.)
        TextWord(text="3.00", bbox=(-10, 65, 10, 85)),  # perto do segmento da esquerda (150 unid.)
    ]
    result = auto_detect_scale(segments, words)
    assert abs(result.scale_m_per_unit - 0.02) < 1e-6


def test_no_candidates_returns_unconfident_fallback():
    result = auto_detect_scale([], [])
    assert result.confident is False
    assert result.candidates == []


def test_mm_values_are_converted_to_meters():
    segments = [((0, 0), (1000, 0))]
    words = [TextWord(text="2000", bbox=(490, -10, 510, 0))]  # 2000mm perto de um segmento de 1000 unid.
    result = auto_detect_scale(segments, words)
    assert abs(result.scale_m_per_unit - 0.002) < 1e-9
