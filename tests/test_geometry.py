import pytest

from pdf_takeoff.extract import TextWord
from pdf_takeoff.geometry import area_m2, build_room_polygons, filter_title_block_noise, perimeter_m


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


def test_filter_title_block_noise_removes_full_span_strip():
    # Sala real: pequena, um único rótulo dentro.
    room = _rect_segments(50, 50, 150, 150)
    # "Bloco de notas": ocupa quase a largura inteira da página (título/legenda).
    notes_strip = _rect_segments(0, 800, 990, 950)
    polygons = build_room_polygons(room + notes_strip, min_area=25)
    assert len(polygons) == 2

    words = [TextWord(text="SALA", bbox=(90, 90, 110, 100))]
    kept = filter_title_block_noise(
        polygons, page_width=1000, page_height=1000, scale_m_per_unit=1.0, words=words
    )
    assert len(kept) == 1
    assert kept[0].bounds == pytest.approx((50, 50, 150, 150), abs=1)


def test_filter_title_block_noise_removes_dense_text_block():
    room = _rect_segments(50, 50, 150, 150)
    notes_block = _rect_segments(300, 300, 340, 340)  # pequeno, mas cheio de texto
    polygons = build_room_polygons(room + notes_block, min_area=25)
    assert len(polygons) == 2

    words = [TextWord(text="SALA", bbox=(90, 90, 110, 100))]
    # 30 palavras espremidas num bloco pequeno -> densidade alta em m² reais
    for i in range(30):
        words.append(TextWord(text=f"nota{i}", bbox=(305 + i % 10, 305 + i // 10, 308 + i % 10, 310 + i // 10)))

    # escala realista (~3cm por unidade de PDF, como num desenho CAD típico)
    kept = filter_title_block_noise(
        polygons, page_width=1000, page_height=1000, scale_m_per_unit=0.03, words=words
    )
    assert len(kept) == 1
    assert kept[0].bounds == pytest.approx((50, 50, 150, 150), abs=1)
