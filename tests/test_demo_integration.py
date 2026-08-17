from pdf_takeoff.demo import run_demo


def test_full_pipeline_end_to_end(tmp_path):
    result = run_demo(str(tmp_path))

    assert result["confident"] is True
    assert abs(result["scale"] - 0.02) < 1e-6

    rooms = {r.room_id: r for r in result["rooms"]}
    assert len(rooms) == 2

    room_a = rooms["P1-01"]
    room_b = rooms["P1-02"]

    # Tolerância generosa: o snap de nós na reconstrução do polígono
    # desloca levemente os cantos em relação ao retângulo "ideal".
    assert abs(room_a.ceiling_area_m2 - 12.0) < 0.1
    assert abs(room_b.ceiling_area_m2 - 15.0) < 0.1
    assert room_a.wall_paint_codes == ["P-01"]
    assert room_b.wall_paint_codes == ["P-02"]
    assert room_a.ceiling_paint_codes == ["PT-01"]
    assert room_b.ceiling_paint_codes == ["PT-01"]
    assert room_b.drywall_codes == ["GB-AQ-DL"]
    assert room_a.drywall_codes == []

    import os

    assert os.path.exists(result["xlsx"])
    assert os.path.exists(result["review_pdf"])
