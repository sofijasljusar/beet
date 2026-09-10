"""
Unit Tests for ESAIC 2025 Congress Abstract Extractor
------------------------------------------------------
Uses pytest for testing business logic, text normalization,
consortium filtering, underline detection, and regex patterns.

Run with:
    pytest -v
"""

import pytest
import pymupdf

from extract_abstracts import (
    dehyphenate,
    clean_title,
    clean_text,
    is_consortium,
    is_author_underlined,
    PRESENTATION_ID_PATTERN,
    STANDALONE_FIG_PATTERN,
    BACK_MATTER_PATTERN,
)


# =====================================================================
# 1. Text Normalization & De-hyphenation Tests
# =====================================================================

def test_dehyphenate_line_wraps():
    """Line-wrap hyphens like 'preopera-\\n tive' should be reconnected to 'preoperative'."""
    text = "Patients with preopera-\n tive compli-\n cations were evaluated."
    expected = "Patients with preoperative complications were evaluated."
    assert dehyphenate(text) == expected


def test_dehyphenate_preserves_legitimate_hyphens():
    """Intentional hyphens like 'AI-driven' or 'GLP-1' must not be stripped."""
    text = "AI-driven predictions and GLP-1 receptor agonists in a population-based study."
    assert dehyphenate(text) == text


@pytest.mark.parametrize("empty_val", ["", None])
def test_dehyphenate_empty_and_none(empty_val):
    """Gracefully handles empty strings and None."""
    assert dehyphenate(empty_val) == ""


def test_clean_title_collapses_whitespace_and_newlines():
    """Multiline titles and irregular spacing must collapse into a clean single line."""
    multiline_title = (
        "Epidemiological trends of acute respiratory\n"
        "distress syndrome in the 21st century:   a nationwide,\n"
        "population-based   study"
    )
    expected = (
        "Epidemiological trends of acute respiratory distress syndrome in the 21st century: "
        "a nationwide, population-based study"
    )
    assert clean_title(multiline_title) == expected


def test_clean_text_normalizes_paragraphs():
    """Abstract body text is stripped of trailing spaces and line-wrap hyphens."""
    raw_body = " Background and Goal: postopera-\n tive pain was assessed. \n"
    assert clean_text(raw_body) == "Background and Goal: postoperative pain was assessed."


# =====================================================================
# 2. Consortium / Study Group Filtering Tests
# =====================================================================

@pytest.mark.parametrize("study_group", [
    "PEEP LAP Study Group",
    "MET-REPAIR Investigators",
    "ESAIC Clinical Trials Network",
    "International Sepsis Consortium",
    "COVID-19 Advisory Committee",
    "National Airway Taskforce",
    "Collaborators Group",
])
def test_is_consortium_matches_study_groups(study_group):
    """Consortium and study groups must return True so they are excluded from author records."""
    assert is_consortium(study_group) is True


@pytest.mark.parametrize("real_name", [
    "S. Saxena",
    "I. Balenović",
    "B. Silbert",
    "E. Gómez Pesquera",
    "F. Verdina",
    "K. Ando",
    "C. Peng",
    "T. Ildikó",
    "B. Morel",
])
def test_is_consortium_preserves_human_authors(real_name):
    """Real human names must return False so they are included as authors."""
    assert is_consortium(real_name) is False


# =====================================================================
# 3. Presenter Vector Underline Detection Tests
# =====================================================================

def test_author_underlined_positive():
    """Returns True when vector underline overlaps horizontally and sits directly beneath text."""
    # Author text span bbox: x0=50, y0=100, x1=120, y1=110
    author_spans = [("S. Saxena", [50.0, 100.0, 120.0, 110.0])]
    # Vector underline directly beneath at y0=111, spanning x0=50 to x1=120
    underlines = [pymupdf.Rect(50.0, 111.0, 120.0, 112.0)]
    assert is_author_underlined(author_spans, underlines) is True


def test_author_underlined_negative_no_horizontal_overlap():
    """Returns False when underline is located in a different column."""
    author_spans = [("B. Morel", [50.0, 100.0, 120.0, 110.0])]
    underlines = [pymupdf.Rect(300.0, 111.0, 380.0, 112.0)]
    assert is_author_underlined(author_spans, underlines) is False


def test_author_underlined_negative_vertical_distance():
    """Returns False when line is too far below text (e.g. section separator line)."""
    author_spans = [("B. Morel", [50.0, 100.0, 120.0, 110.0])]
    underlines = [pymupdf.Rect(50.0, 130.0, 120.0, 131.0)]
    assert is_author_underlined(author_spans, underlines) is False


# =====================================================================
# 4. Regex Pattern Tests (Presentation IDs, Visuals, Back-Matter)
# =====================================================================

@pytest.mark.parametrize("line, expected_id", [
    ("BAPC-01 Epidemiological trends", "BAPC-01"),
    ("BAPC-06 Lung aeration following", "BAPC-06"),
    ("10AP01-2 A Late-Onset Case", "10AP01-2"),
    ("32AP04-12 Cardiovascular Study", "32AP04-12"),
    ("60AP01-6 Influence of Decision Fatigue", "60AP01-6"),
])
def test_presentation_id_pattern_valid(line, expected_id):
    """Matches all valid presentation ID formats across congress sections."""
    m = PRESENTATION_ID_PATTERN.match(line)
    assert m is not None
    assert m.group(1) == expected_id


@pytest.mark.parametrize("invalid_line", [
    "Background and Goal of Study",
    "ESAIC 2025 Annual Congress",
    "Perioperative Care",
    "Heart: Cardiovascular Anaesthesiology and Haemodynamics",
    "Fig 1. Consort Diagram",
])
def test_presentation_id_pattern_invalid(invalid_line):
    """Rejects lines that are headings or body text, not presentation IDs."""
    assert PRESENTATION_ID_PATTERN.match(invalid_line) is None


@pytest.mark.parametrize("tag", ["Fig 1.", "Fig. 2", "Figure 1.", "Table 1.", "Table 2:"])
def test_standalone_figure_pattern_matches_isolated_tags(tag):
    """Matches standalone figure/table labels to discard them."""
    assert STANDALONE_FIG_PATTERN.match(tag) is not None


@pytest.mark.parametrize("caption_or_mention", [
    "(Fig 1)",
    "Figure 1: Forest plot showing odds ratios",
    "as depicted in Table 1.",
    "Table 1 summarizes demographic characteristics",
])
def test_standalone_figure_pattern_preserves_captions_and_mentions(caption_or_mention):
    """Preserves full captions and in-text references like '(Fig 1)'."""
    assert STANDALONE_FIG_PATTERN.match(caption_or_mention) is None


@pytest.mark.parametrize("header", [
    "Author Index",
    "AUTHOR INDEX",
    "Subject Index",
    "SUBJECT INDEX",
    "Author Index continued",
])
def test_back_matter_pattern_matches_indexes(header):
    """Detects both Author Index and Subject Index headers."""
    assert BACK_MATTER_PATTERN.search(header) is not None


@pytest.mark.parametrize("body_line", [
    "Background: Authors conducted a multi-center trial",
    "Subject: 40 patients randomized to intervention",
    "Perioperative Care",
])
def test_back_matter_pattern_ignores_regular_text(body_line):
    """Does not trigger on regular abstract body text."""
    assert BACK_MATTER_PATTERN.search(body_line) is None
