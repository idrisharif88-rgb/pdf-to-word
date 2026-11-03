import pytest

from pdf_to_word.config import Settings
from pdf_to_word.text import (TextUnit, batches, clean_text_smart,
                              detect_placement, detect_rtl, heading_level,
                              is_footnote, is_ocr_noise, parse_page_range,
                              parse_page_spec, running_head_texts,
                              unwrap_lines)


def test_clean_text_arabic_punctuation_and_tatweel():
    assert clean_text_smart("كتـــاب , جميل ?") == "كتاب، جميل؟"


def test_clean_text_keeps_line_breaks_and_strips_bidi():
    assert clean_text_smart("‏سطر  أول\n\n\n\nسطر ثان") == "سطر أول\n\nسطر ثان"


def test_detect_rtl():
    assert detect_rtl("مرحبا بالعالم")
    assert not detect_rtl("E-mail: info@example.com")


def test_unwrap_lines_joins_single_breaks_only():
    assert unwrap_lines("a\nb\n\nc") == "a b\n\nc"


def test_noise():
    assert is_ocr_noise(" . ")
    assert not is_ocr_noise("7")


@pytest.mark.parametrize("spec,total,expected", [
    ("all", 3, [1, 2, 3]),
    ("", 2, [1, 2]),
    ("1-3, 5", 10, [1, 2, 3, 5]),
    ("2,2,9", 5, [2]),
])
def test_parse_page_range(spec, total, expected):
    assert parse_page_range(spec, total) == expected


@pytest.mark.parametrize("bad", ["a", "1-", "3-x"])
def test_parse_page_range_rejects_garbage(bad):
    with pytest.raises(ValueError):
        parse_page_range(bad, 10)


def test_parse_page_spec_ignores_garbage():
    assert parse_page_spec("1-3, x, 7") == {1, 2, 3, 7}


def test_heading_level():
    assert heading_level(0.05, 0.03, "عنوان", 1, 0.3, 1.5, 1.25) == 1
    assert heading_level(0.04, 0.03, "عنوان", 1, 0.3, 1.5, 1.25) == 2
    assert heading_level(0.05, 0.03, "عنوان", 1, 0.9, 1.5, 1.25) == 0  # full width
    assert heading_level(0.05, 0.03, "x" * 100, 1, 0.3, 1.5, 1.25) == 0


def test_footnote():
    assert is_footnote(0.85, 0.02, 0.03, 0.78, 0.88)
    assert not is_footnote(0.5, 0.02, 0.03, 0.78, 0.88)


def test_placement():
    assert detect_placement(0.1, 0.9) == "full"
    assert detect_placement(0.35, 0.65) == "center"
    assert detect_placement(0.05, 0.4) == "left"
    assert detect_placement(0.6, 0.95) == "right"


def test_running_heads_need_most_pages():
    head = TextUnit("اسم الكتاب 12", 0.02, 0.02, 0.04, 0.4, 0.6)
    body = TextUnit("نص", 0.02, 0.5, 0.52, 0.1, 0.9)
    pages = [[head, body]] * 4
    assert running_head_texts(pages, 0.12) == {"اسم الكتاب #"}
    assert running_head_texts(pages[:2], 0.12) == set()


def test_batches():
    assert batches([1, 2, 3, 4, 5], 2) == [[1, 2], [3, 4], [5]]


def test_settings_from_env_types():
    s = Settings.from_env({
        "ARABIC_SIZE_PT": "16", "EXACT_LAYOUT": "true", "BATCH_SIZE": "5",
        "EXTRACT_TABLES": "0", "GCP_FORM_PROCESSOR_ID": "abc",
        "ARABIC_FONT": "Amiri", "IMAGE_PAGES": "",
    })
    assert s.arabic_size_pt == 16.0 and s.exact_layout and s.batch_size == 5
    assert s.arabic_font == "Amiri" and s.image_pages == ""
    assert s.form_parser_available and not s.use_form_parser


def test_settings_defaults():
    s = Settings.from_env({})
    assert not s.extract_tables and not s.form_parser_available
    assert not s.use_form_parser
