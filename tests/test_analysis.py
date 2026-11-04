"""
Step 3 tests — verify measured analysis checks against known fixtures.

For each fixture the test asserts:
  1. The expected check fires at least once.
  2. No unexpected checks fire (clean deck never produces any issue).
"""

from pathlib import Path

import pytest

from app.analysis import analyze_pptx

FIXTURES = Path(__file__).parent / "fixtures"


def fp(name: str) -> Path:
    return FIXTURES / name


def slide_issues(result: dict, slide: int = 1) -> list[dict]:
    return result["slides"][slide - 1]["issues"]


def check_names(result: dict, slide: int = 1) -> set[str]:
    return {i["check"] for i in slide_issues(result, slide)}


def deck_check_names(result: dict) -> set[str]:
    return {i["check"] for i in result["deck_issues"]}


# ── Top-level structure ───────────────────────────────────────────────────────

def test_result_has_expected_keys():
    r = analyze_pptx(fp("clean.pptx"))
    for key in ("file_name", "slide_width_inches", "slide_height_inches",
                "slide_count", "fonts_used", "deck_issues", "slides"):
        assert key in r, f"Missing key: {key}"


def test_result_slide_dimensions():
    r = analyze_pptx(fp("clean.pptx"))
    assert r["slide_width_inches"] == pytest.approx(10.0, abs=0.01)
    assert r["slide_height_inches"] == pytest.approx(7.5, abs=0.01)


def test_result_shape_has_expected_keys():
    r = analyze_pptx(fp("clean.pptx"))
    shape = r["slides"][0]["shapes"][0]
    for key in ("xml_index", "name", "shape_type_name", "left_inches",
                "top_inches", "width_inches", "height_inches",
                "right_inches", "bottom_inches", "is_fully_on_slide",
                "has_text_frame", "text", "font_sizes_pt", "font_names",
                "alt_text", "is_picture"):
        assert key in shape, f"Missing shape key: {key}"


def test_issues_carry_measured_type():
    r = analyze_pptx(fp("overlapping_objects.pptx"))
    for issue in slide_issues(r):
        assert issue["type"] == "Measured"


# ── clean.pptx — zero issues ──────────────────────────────────────────────────

def test_clean_produces_no_slide_issues():
    r = analyze_pptx(fp("clean.pptx"))
    assert slide_issues(r) == [], (
        f"clean.pptx should produce no issues; got: {slide_issues(r)}"
    )


def test_clean_produces_no_deck_issues():
    r = analyze_pptx(fp("clean.pptx"))
    assert r["deck_issues"] == [], (
        f"clean.pptx should produce no deck issues; got: {r['deck_issues']}"
    )


def test_clean_title_extracted():
    r = analyze_pptx(fp("clean.pptx"))
    assert "Quarterly Results" in r["slides"][0]["title"]


# ── overlapping_objects.pptx ──────────────────────────────────────────────────

def test_overlapping_objects_check_fires():
    r = analyze_pptx(fp("overlapping_objects.pptx"))
    assert "overlapping_objects" in check_names(r)


def test_overlapping_objects_names_both_shapes():
    r = analyze_pptx(fp("overlapping_objects.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "overlapping_objects")
    assert "Red Box" in issue["shapes"]
    assert "Blue Box" in issue["shapes"]


def test_overlapping_objects_no_other_checks():
    r = analyze_pptx(fp("overlapping_objects.pptx"))
    unexpected = check_names(r) - {"overlapping_objects"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── object_off_slide.pptx ─────────────────────────────────────────────────────

def test_object_off_slide_check_fires():
    r = analyze_pptx(fp("object_off_slide.pptx"))
    assert "object_off_slide" in check_names(r)


def test_object_off_slide_names_correct_shape():
    r = analyze_pptx(fp("object_off_slide.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "object_off_slide")
    assert issue["shape"] == "Off-Slide Shape"


def test_object_off_slide_no_other_checks():
    r = analyze_pptx(fp("object_off_slide.pptx"))
    unexpected = check_names(r) - {"object_off_slide"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── text_overflow.pptx ───────────────────────────────────────────────────────

def test_text_overflow_check_fires():
    r = analyze_pptx(fp("text_overflow.pptx"))
    assert "text_overflow" in check_names(r)


def test_text_overflow_confidence_is_low():
    r = analyze_pptx(fp("text_overflow.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "text_overflow")
    assert issue.get("confidence") == "low"


def test_text_overflow_no_other_checks():
    r = analyze_pptx(fp("text_overflow.pptx"))
    unexpected = check_names(r) - {"text_overflow"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── small_text.pptx ──────────────────────────────────────────────────────────

def test_small_text_check_fires():
    r = analyze_pptx(fp("small_text.pptx"))
    assert "small_text" in check_names(r)


def test_small_text_reports_correct_size():
    r = analyze_pptx(fp("small_text.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "small_text")
    assert issue["font_size_pt"] <= 10.0


def test_small_text_no_other_checks():
    r = analyze_pptx(fp("small_text.pptx"))
    unexpected = check_names(r) - {"small_text"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── low_contrast.pptx ────────────────────────────────────────────────────────

def test_low_contrast_check_fires():
    r = analyze_pptx(fp("low_contrast.pptx"))
    assert "low_contrast" in check_names(r)


def test_low_contrast_ratio_below_aa():
    r = analyze_pptx(fp("low_contrast.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "low_contrast")
    assert issue["contrast_ratio"] < 4.5


def test_low_contrast_reports_colors():
    r = analyze_pptx(fp("low_contrast.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "low_contrast")
    assert issue["text_color"].startswith("#")
    assert issue["bg_color"].startswith("#")


def test_low_contrast_no_other_checks():
    r = analyze_pptx(fp("low_contrast.pptx"))
    unexpected = check_names(r) - {"low_contrast"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── missing_alt_text.pptx ────────────────────────────────────────────────────

def test_missing_alt_text_check_fires():
    r = analyze_pptx(fp("missing_alt_text.pptx"))
    assert "missing_alt_text" in check_names(r)


def test_missing_alt_text_names_correct_shape():
    r = analyze_pptx(fp("missing_alt_text.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "missing_alt_text")
    assert issue["shape"] == "Chart Image"


def test_missing_alt_text_no_other_checks():
    r = analyze_pptx(fp("missing_alt_text.pptx"))
    unexpected = check_names(r) - {"missing_alt_text"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── no_title.pptx ────────────────────────────────────────────────────────────

def test_no_title_check_fires():
    r = analyze_pptx(fp("no_title.pptx"))
    assert "missing_slide_title" in check_names(r)


def test_no_title_slide_has_empty_title_field():
    r = analyze_pptx(fp("no_title.pptx"))
    assert r["slides"][0]["title"] == ""


def test_no_title_no_other_checks():
    r = analyze_pptx(fp("no_title.pptx"))
    unexpected = check_names(r) - {"missing_slide_title"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── reading_order.pptx ───────────────────────────────────────────────────────

def test_reading_order_check_fires():
    r = analyze_pptx(fp("reading_order.pptx"))
    assert "reading_order" in check_names(r)


def test_reading_order_reports_inversion_count():
    r = analyze_pptx(fp("reading_order.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "reading_order")
    assert issue["inversion_count"] >= 1


def test_reading_order_no_other_checks():
    r = analyze_pptx(fp("reading_order.pptx"))
    unexpected = check_names(r) - {"reading_order"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── crowded.pptx ─────────────────────────────────────────────────────────────

def test_crowded_check_fires():
    r = analyze_pptx(fp("crowded.pptx"))
    assert "crowded_slide" in check_names(r)


def test_crowded_reports_shape_count():
    r = analyze_pptx(fp("crowded.pptx"))
    issue = next(i for i in slide_issues(r) if i["check"] == "crowded_slide")
    assert issue["shape_count"] >= 10


def test_crowded_no_other_checks():
    r = analyze_pptx(fp("crowded.pptx"))
    unexpected = check_names(r) - {"crowded_slide"}
    assert unexpected == set(), f"Unexpected checks fired: {unexpected}"


# ── Deck-level checks ─────────────────────────────────────────────────────────

def test_duplicate_title_fires_on_two_slide_deck(tmp_path):
    """Build a small two-slide deck with the same title and verify detection."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    blank = next(l for l in prs.slide_layouts if l.name.lower() == "blank")
    for _ in range(2):
        slide = prs.slides.add_slide(blank)
        txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.9))
        txb.text_frame.text = "Same Title"
        txb.text_frame.paragraphs[0].runs[0].font.size = Pt(28)

    path = tmp_path / "dup_title.pptx"
    prs.save(str(path))

    r = analyze_pptx(path)
    assert "duplicate_title" in deck_check_names(r)


def test_many_fonts_fires_on_four_font_deck(tmp_path):
    """Build a deck with four explicit fonts and verify detection."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank = next(l for l in prs.slide_layouts if l.name.lower() == "blank")
    slide = prs.slides.add_slide(blank)

    fonts = ["Arial", "Times New Roman", "Courier New", "Georgia"]
    for i, font in enumerate(fonts):
        txb = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5 + i * 1.5), Inches(9), Inches(1)
        )
        run = txb.text_frame.paragraphs[0].add_run()
        run.text = f"Text in {font}"
        run.font.name = font
        run.font.size = Pt(20)

    path = tmp_path / "many_fonts.pptx"
    prs.save(str(path))

    r = analyze_pptx(path)
    assert "many_fonts" in deck_check_names(r)
