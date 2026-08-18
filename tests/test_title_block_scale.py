from pdf_takeoff.calibration import extract_scale_from_title_block


def test_extracts_scale_from_common_title_block_format():
    text = "some text\n1:100 @ A2\nRev.\nRef No\nScale:\n"
    scale = extract_scale_from_title_block(text)
    assert scale == 100 * 25.4 / 1000 / 72


def test_returns_none_when_not_present():
    assert extract_scale_from_title_block("no scale info here") is None


def test_extracts_without_paper_size_suffix():
    text = "Scale 1:50\n"
    scale = extract_scale_from_title_block(text)
    assert scale == 50 * 25.4 / 1000 / 72
