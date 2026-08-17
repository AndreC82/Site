import pytest

from pdf_takeoff.geometry import area_m2, build_room_polygons, perimeter_m


def _rect_segments(x0, y0, x1, y1):
    return [
        ((x0, y0), (x1, y0)),
        ((x1, y0), (x1, y1)),
        ((x1, y1), (x0, y1)),
        ((x0, y1), (x0, y0)),
    ]


def test_single_rectangle_area_and_perimeter():
    segments = _rect_segments(0, 0, 200, 150)
    polygons = build_room_polygons(segments, min_area=1)
    assert len(polygons) == 1
    poly = polygons[0]
    scale = 0.02  # m por unidade
    # Tolerância folgada: o snap de vértices (proposital, para tolerar imprecisões
    # reais de traçado em PDFs de CAD) desloca levemente os cantos.
    assert area_m2(poly, scale) == pytest.approx(200 * 150 * scale**2, abs=0.05)
    assert perimeter_m(poly, scale) == pytest.approx(2 * (200 + 150) * scale, abs=0.02)


def test_two_adjacent_rooms_share_wall():
    room_a = _rect_segments(0, 0, 200, 150)
    room_b = _rect_segments(200, 0, 450, 150)
    polygons = build_room_polygons(room_a + room_b, min_area=1)
    assert len(polygons) == 2
    areas = sorted(p.area for p in polygons)
    assert areas[0] == pytest.approx(200 * 150, abs=50)
    assert areas[1] == pytest.approx(250 * 150, abs=50)


def test_small_noise_polygons_filtered_by_min_area():
    room = _rect_segments(0, 0, 200, 150)
    noise = _rect_segments(500, 500, 502, 502)  # 2x2 - hachura/ruído
    polygons = build_room_polygons(room + noise, min_area=25)
    assert len(polygons) == 1
