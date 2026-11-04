"""
Step 2 tests — verify that every fixture file exists, is a valid PPTX,
and has the structural properties its manifest entry claims.
"""

import json
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

FIXTURES_DIR = Path(__file__).parent / "fixtures"
ALL_FIXTURES = [
    "clean.pptx",
    "overlapping_objects.pptx",
    "object_off_slide.pptx",
    "text_overflow.pptx",
    "small_text.pptx",
    "low_contrast.pptx",
    "missing_alt_text.pptx",
    "no_title.pptx",
    "reading_order.pptx",
    "crowded.pptx",
]


@pytest.fixture(scope="module")
def manifest() -> dict:
    path = FIXTURES_DIR / "manifest.json"
    assert path.exists(), "manifest.json missing — run: python tests/create_fixtures.py"
    return json.loads(path.read_text(encoding="utf-8"))


def fp(name: str) -> Path:
    return FIXTURES_DIR / name


# ── Existence and validity ────────────────────────────────────────────────────

@pytest.mark.parametrize("name", ALL_FIXTURES)
def test_fixture_file_exists(name, manifest):
    assert fp(name).exists(), f"{name} not found in tests/fixtures/"
    assert name in manifest, f"{name} not in manifest.json"


@pytest.mark.parametrize("name", ALL_FIXTURES)
def test_fixture_is_valid_pptx(name):
    prs = Presentation(str(fp(name)))
    assert prs.slides is not None


@pytest.mark.parametrize("name", ALL_FIXTURES)
def test_fixture_slide_count(name, manifest):
    prs = Presentation(str(fp(name)))
    expected = manifest[name]["slide_count"]
    assert len(prs.slides) == expected


# ── clean.pptx — zero known problems ─────────────────────────────────────────

def test_clean_has_no_known_problems(manifest):
    assert manifest["clean.pptx"]["known_problems"] == []


def test_clean_has_title_text():
    prs = Presentation(str(fp("clean.pptx")))
    slide = prs.slides[0]
    all_text = " ".join(
        s.text_frame.text for s in slide.shapes if s.has_text_frame
    )
    assert "Quarterly Results" in all_text


def test_clean_image_has_alt_text():
    prs = Presentation(str(fp("clean.pptx")))
    slide = prs.slides[0]
    pics = [s for s in slide.shapes if s.shape_type == 13]  # 13 = PICTURE
    assert len(pics) >= 1
    for pic in pics:
        descr = pic._element.nvPicPr.cNvPr.get("descr", "")
        assert descr != "", f"Picture '{pic.name}' in clean.pptx should have alt text"


# ── overlapping_objects.pptx ──────────────────────────────────────────────────

def test_overlapping_objects_shapes_exist():
    prs = Presentation(str(fp("overlapping_objects.pptx")))
    names = {s.name for s in prs.slides[0].shapes}
    assert "Red Box" in names
    assert "Blue Box" in names


def test_overlapping_objects_do_overlap():
    prs = Presentation(str(fp("overlapping_objects.pptx")))
    slide = prs.slides[0]
    red = next(s for s in slide.shapes if s.name == "Red Box")
    blue = next(s for s in slide.shapes if s.name == "Blue Box")
    # Horizontal overlap check
    assert red.left < blue.left + blue.width
    assert blue.left < red.left + red.width
    # Vertical overlap check
    assert red.top < blue.top + blue.height
    assert blue.top < red.top + red.height


# ── object_off_slide.pptx ─────────────────────────────────────────────────────

def test_off_slide_shape_is_outside_bounds():
    prs = Presentation(str(fp("object_off_slide.pptx")))
    slide = prs.slides[0]
    off = next(s for s in slide.shapes if s.name == "Off-Slide Shape")
    assert off.left >= prs.slide_width, (
        f"Off-Slide Shape left ({off.left}) should be ≥ slide width ({prs.slide_width})"
    )


def test_on_slide_shape_is_within_bounds():
    prs = Presentation(str(fp("object_off_slide.pptx")))
    slide = prs.slides[0]
    on = next(s for s in slide.shapes if s.name == "On-Slide Shape")
    assert on.left >= 0
    assert on.left + on.width <= prs.slide_width


# ── text_overflow.pptx ───────────────────────────────────────────────────────

def test_text_overflow_box_has_no_auto_size():
    from pptx.enum.text import MSO_AUTO_SIZE
    prs = Presentation(str(fp("text_overflow.pptx")))
    slide = prs.slides[0]
    # Find the small box (not the title)
    small_boxes = [
        s for s in slide.shapes
        if s.has_text_frame and s.width <= Inches(3.5) and s.top > Inches(1)
    ]
    assert len(small_boxes) >= 1
    box = small_boxes[0]
    assert box.text_frame.auto_size == MSO_AUTO_SIZE.NONE


def test_text_overflow_text_is_long():
    prs = Presentation(str(fp("text_overflow.pptx")))
    slide = prs.slides[0]
    small_boxes = [
        s for s in slide.shapes
        if s.has_text_frame and s.width <= Inches(3.5) and s.top > Inches(1)
    ]
    text = small_boxes[0].text_frame.text
    assert len(text) > 100, "Expected many characters in the overflow text box"


# ── small_text.pptx ──────────────────────────────────────────────────────────

def test_small_text_has_tiny_font():
    prs = Presentation(str(fp("small_text.pptx")))
    slide = prs.slides[0]
    tiny_found = any(
        run.font.size is not None and run.font.size <= Pt(10)
        for shape in slide.shapes
        if shape.has_text_frame
        for para in shape.text_frame.paragraphs
        for run in para.runs
    )
    assert tiny_found, "Expected at least one run with font size ≤ 10pt"


# ── low_contrast.pptx ────────────────────────────────────────────────────────

def test_low_contrast_text_color():
    prs = Presentation(str(fp("low_contrast.pptx")))
    slide = prs.slides[0]
    colors_found = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                color = run.font.color.rgb
                if color is not None:
                    colors_found.append(str(color))
    assert "BBBBBB" in colors_found, "Expected #BBBBBB text color in low_contrast.pptx"


# ── missing_alt_text.pptx ────────────────────────────────────────────────────

def test_missing_alt_text_image_exists():
    prs = Presentation(str(fp("missing_alt_text.pptx")))
    slide = prs.slides[0]
    pics = [s for s in slide.shapes if s.shape_type == 13]
    assert len(pics) >= 1, "Expected at least one picture on the slide"


def test_missing_alt_text_has_no_description():
    prs = Presentation(str(fp("missing_alt_text.pptx")))
    slide = prs.slides[0]
    chart_img = next(s for s in slide.shapes if s.name == "Chart Image")
    descr = chart_img._element.nvPicPr.cNvPr.get("descr", "")
    assert descr == "", f"'Chart Image' should have no alt text, found: {descr!r}"


# ── no_title.pptx ────────────────────────────────────────────────────────────

def test_no_title_slide_has_one_shape():
    prs = Presentation(str(fp("no_title.pptx")))
    slide = prs.slides[0]
    assert len(slide.shapes) == 1, (
        f"no_title.pptx should have exactly 1 shape, found {len(slide.shapes)}"
    )


def test_no_title_shape_is_in_body_region():
    """The only shape should be below the top 15% of the slide (not a title)."""
    prs = Presentation(str(fp("no_title.pptx")))
    slide = prs.slides[0]
    shape = slide.shapes[0]
    title_zone_bottom = prs.slide_height * 0.20
    assert shape.top >= title_zone_bottom, (
        "The only shape in no_title.pptx should be below the title zone"
    )


# ── reading_order.pptx ───────────────────────────────────────────────────────

def test_reading_order_has_three_step_shapes():
    prs = Presentation(str(fp("reading_order.pptx")))
    slide = prs.slides[0]
    step_shapes = [s for s in slide.shapes if s.name.startswith("Step")]
    assert len(step_shapes) == 3


def test_reading_order_xml_first_is_not_visually_first():
    prs = Presentation(str(fp("reading_order.pptx")))
    slide = prs.slides[0]
    step_shapes = [s for s in slide.shapes if s.name.startswith("Step")]
    xml_first = step_shapes[0]
    visually_first = min(step_shapes, key=lambda s: s.top)
    assert xml_first.name != visually_first.name, (
        "Expected the first shape in XML order to differ from the visually topmost shape"
    )


# ── crowded.pptx ─────────────────────────────────────────────────────────────

def test_crowded_slide_has_many_shapes():
    prs = Presentation(str(fp("crowded.pptx")))
    slide = prs.slides[0]
    assert len(slide.shapes) >= 10, (
        f"crowded.pptx should have ≥ 10 shapes, found {len(slide.shapes)}"
    )


def test_crowded_slide_shape_count_matches_manifest(manifest):
    prs = Presentation(str(fp("crowded.pptx")))
    slide = prs.slides[0]
    expected = manifest["crowded.pptx"]["known_problems"][0]["shape_count"]
    assert len(slide.shapes) == expected
