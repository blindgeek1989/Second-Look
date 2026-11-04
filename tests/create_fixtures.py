#!/usr/bin/env python3
"""
Generate test fixture PPTX files for Second Look.

Run:  python tests/create_fixtures.py

Each file documents the specific accessibility or visual problems it contains.
The companion manifest.json is consumed by Step 3 tests.
"""

import io
import json
import struct
import zlib
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE

FIXTURES_DIR = Path(__file__).parent / "fixtures"
SLIDE_W = Inches(10)
SLIDE_H = Inches(7.5)
RECT = MSO_AUTO_SHAPE_TYPE.RECTANGLE


# ── Helpers ──────────────────────────────────────────────────────────────────

def make_png(width: int = 80, height: int = 60, color: tuple = (180, 120, 100)) -> bytes:
    """Return bytes of a minimal valid RGB PNG — no Pillow needed."""
    r, g, b = color

    def chunk(name: bytes, data: bytes) -> bytes:
        payload = name + data
        return (
            struct.pack(">I", len(data))
            + payload
            + struct.pack(">I", zlib.crc32(payload) & 0xFFFFFFFF)
        )

    ihdr = chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    rows = b"".join(bytes([0]) + bytes([r, g, b] * width) for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + ihdr
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs: Presentation):
    """Add a slide using the blank layout (no placeholders)."""
    for layout in prs.slide_layouts:
        if layout.name.lower() == "blank":
            return prs.slides.add_slide(layout)
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_title_box(slide, text: str):
    """Add a text box styled as a slide title in the top strip."""
    txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.9))
    tf = txb.text_frame
    tf.text = text
    run = tf.paragraphs[0].runs[0]
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    return txb


def set_pic_alt(pic, text: str) -> None:
    pic._element.nvPicPr.cNvPr.set("descr", text)


# ── 1. Clean ─────────────────────────────────────────────────────────────────

def build_clean() -> Presentation:
    """Single slide, no problems — used to catch false positives."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Quarterly Results — Q3 2024")

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(8), Inches(3.5))
    tf = body.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = (
        "Revenue increased 12 percent year over year. "
        "Operating margin improved to 24 percent. "
        "The team delivered all three product milestones on schedule."
    )
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    img = io.BytesIO(make_png(80, 60, (100, 149, 237)))
    pic = slide.shapes.add_picture(img, Inches(7.2), Inches(2.0), Inches(2.0), Inches(1.5))
    set_pic_alt(pic, "Blue placeholder representing a bar chart")

    return prs


# ── 2. Overlapping objects ────────────────────────────────────────────────────

def build_overlapping_objects() -> Presentation:
    """Red Box and Blue Box overlap. Red: x=1.5–5.0", Blue: x=3.5–7.0"."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Overlapping Objects")

    s1 = slide.shapes.add_shape(RECT, Inches(1.5), Inches(2.0), Inches(3.5), Inches(2.5))
    s1.name = "Red Box"
    s1.fill.solid()
    s1.fill.fore_color.rgb = RGBColor(0xCC, 0x22, 0x22)
    s1.line.color.rgb = RGBColor(0x99, 0x00, 0x00)

    # Starts at x=3.5" — inside Red Box (which ends at x=5.0")
    s2 = slide.shapes.add_shape(RECT, Inches(3.5), Inches(2.5), Inches(3.5), Inches(2.5))
    s2.name = "Blue Box"
    s2.fill.solid()
    s2.fill.fore_color.rgb = RGBColor(0x22, 0x22, 0xCC)
    s2.line.color.rgb = RGBColor(0x00, 0x00, 0x99)

    return prs


# ── 3. Object off slide ──────────────────────────────────────────────────────

def build_object_off_slide() -> Presentation:
    """Off-Slide Shape starts at x=11", outside the 10"-wide slide."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Object Off the Slide")

    on = slide.shapes.add_shape(RECT, Inches(1.0), Inches(2.0), Inches(3.0), Inches(1.5))
    on.name = "On-Slide Shape"
    on.fill.solid()
    on.fill.fore_color.rgb = RGBColor(0x22, 0xAA, 0x44)
    on.text_frame.text = "This shape is on the slide."
    on.text_frame.paragraphs[0].runs[0].font.size = Pt(16)

    off = slide.shapes.add_shape(RECT, Inches(11.0), Inches(2.0), Inches(3.0), Inches(1.5))
    off.name = "Off-Slide Shape"
    off.fill.solid()
    off.fill.fore_color.rgb = RGBColor(0xCC, 0x44, 0x44)
    off.text_frame.text = "This shape is off the slide."
    off.text_frame.paragraphs[0].runs[0].font.size = Pt(16)

    return prs


# ── 4. Text overflow ─────────────────────────────────────────────────────────

def build_text_overflow() -> Presentation:
    """Fixed-size text box (auto_size=NONE) containing far more text than fits."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Possible Text Overflow")

    txb = slide.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(3.0), Inches(1.2))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE  # box will NOT expand to fit

    run = tf.paragraphs[0].add_run()
    run.text = (
        "This text box is intentionally too small to hold all of this content. "
        "The box is only 3 inches wide and 1.2 inches tall, but this paragraph "
        "contains far more text than will fit in those dimensions. "
        "The overflowing words will not be visible when the slide is displayed, "
        "which means important content could be hidden from sighted viewers."
    )
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    return prs


# ── 5. Small text ────────────────────────────────────────────────────────────

def build_small_text() -> Presentation:
    """A run set to 8pt — below any readable threshold."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Small Text")

    txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(2.0))
    tf = txb.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = (
        "Important footnote: All figures are unaudited estimates for the fiscal year. "
        "See appendix for full methodology and data sources."
    )
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    return prs


# ── 6. Low contrast ──────────────────────────────────────────────────────────

def build_low_contrast() -> Presentation:
    """#BBBBBB text on a white (#FFFFFF) background — contrast ≈ 1.86:1 (WCAG AA needs 4.5:1)."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Low Contrast Text")

    txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(2.0))
    tf = txb.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = (
        "This text is light gray on a white background. "
        "The contrast ratio is approximately 1.86:1, well below "
        "the WCAG 2.2 AA minimum of 4.5:1 for normal text."
    )
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)  # #BBBBBB on white

    return prs


# ── 7. Missing alt text ──────────────────────────────────────────────────────

def build_missing_alt_text() -> Presentation:
    """Image with no alt text (descr attribute absent on cNvPr)."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Image Missing Alt Text")

    img = io.BytesIO(make_png(120, 90, (200, 150, 80)))
    pic = slide.shapes.add_picture(img, Inches(2.0), Inches(2.0), Inches(3.0), Inches(2.25))
    pic.name = "Chart Image"
    # python-pptx auto-sets descr="image.png" for BytesIO streams.
    # Remove it entirely so this fixture truly has no alt text.
    cNvPr = pic._element.nvPicPr.cNvPr
    cNvPr.attrib.pop("descr", None)

    return prs


# ── 8. No title ──────────────────────────────────────────────────────────────

def build_no_title() -> Presentation:
    """Slide with only body text — no title textbox in the top region."""
    prs = new_prs()
    slide = blank_slide(prs)
    # Deliberately no call to add_title_box — this is the problem.

    txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(4.0))
    tf = txb.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = (
        "This slide has no title. A screen reader user navigating by headings "
        "will find nothing here. A person scanning the deck will also lose context "
        "about what this slide is meant to convey."
    )
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    return prs


# ── 9. Reading order ─────────────────────────────────────────────────────────

def build_reading_order() -> Presentation:
    """
    Visual layout (top → bottom):   Step 1 y=1.5"  |  Step 2 y=2.9"  |  Step 3 y=4.0"
    XML / tab order (added order):   Step 3 first   |  Step 1 second  |  Step 2 third

    A screen reader following XML order reads Step 3 → Step 1 → Step 2.
    """
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Slide with Reading Order Problems")

    def labelled_box(label: str, top_in: float):
        s = slide.shapes.add_shape(RECT, Inches(1.5), Inches(top_in), Inches(7.0), Inches(1.1))
        s.name = label
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor(0xE8, 0xF0, 0xFE)
        s.line.color.rgb = RGBColor(0x44, 0x44, 0xAA)
        s.text_frame.text = label
        s.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    # Added in XML order that DOES NOT match visual top-to-bottom order.
    labelled_box("Step 3 — Visually at bottom, first in XML", 4.0)
    labelled_box("Step 1 — Visually at top, second in XML", 1.5)
    labelled_box("Step 2 — Visually in middle, third in XML", 2.9)

    return prs


# ── 10. Crowded ──────────────────────────────────────────────────────────────

def build_crowded() -> Presentation:
    """Thirteen shapes (title + 12 data items) packed onto one slide."""
    prs = new_prs()
    slide = blank_slide(prs)
    add_title_box(slide, "Crowded Slide")

    items = [
        "Q1 Revenue: $4.2M", "Q2 Revenue: $5.1M", "Q3 Revenue: $6.0M",
        "Q4 Forecast: $6.8M", "Headcount: 142", "New hires: 18",
        "Attrition: 4%", "NPS Score: 68", "Churn: 2.1%",
        "Burn rate: $1.1M/mo", "Runway: 22 months", "Open reqs: 9",
    ]

    cols = 3
    for i, item in enumerate(items):
        col = i % cols
        row = i // cols
        txb = slide.shapes.add_textbox(
            Inches(0.3 + col * 3.2),
            Inches(1.3 + row * 1.4),
            Inches(3.0),
            Inches(1.1),
        )
        tf = txb.text_frame
        tf.word_wrap = True
        run = tf.paragraphs[0].add_run()
        run.text = item
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    return prs


# ── Manifest ─────────────────────────────────────────────────────────────────

MANIFEST: dict = {
    "clean.pptx": {
        "description": (
            "A well-formed single-slide deck. Title box, body text in high contrast, "
            "one image with alt text. Used to verify checks do not produce false positives."
        ),
        "slide_count": 1,
        "known_problems": [],
    },
    "overlapping_objects.pptx": {
        "description": "Two rectangles (Red Box, Blue Box) overlap between x=3.5\" and x=5.0\".",
        "slide_count": 1,
        "known_problems": [
            {"check": "overlapping_objects", "slide": 1, "shapes": ["Red Box", "Blue Box"]},
        ],
    },
    "object_off_slide.pptx": {
        "description": "Off-Slide Shape starts at x=11\", outside the 10\"-wide slide boundary.",
        "slide_count": 1,
        "known_problems": [
            {"check": "object_off_slide", "slide": 1, "shape": "Off-Slide Shape"},
        ],
    },
    "text_overflow.pptx": {
        "description": (
            "Text box with auto_size=NONE and dimensions 3\"×1.2\" contains "
            "~250 words — far more than the box can display."
        ),
        "slide_count": 1,
        "known_problems": [
            {"check": "text_overflow", "slide": 1},
        ],
    },
    "small_text.pptx": {
        "description": "Body run set to 8pt — below any recommended reading size.",
        "slide_count": 1,
        "known_problems": [
            {"check": "small_text", "slide": 1, "font_size_pt": 8},
        ],
    },
    "low_contrast.pptx": {
        "description": (
            "Body text color #BBBBBB on white (#FFFFFF) slide background. "
            "Measured contrast ratio ≈ 1.86:1; WCAG 2.2 AA requires 4.5:1."
        ),
        "slide_count": 1,
        "known_problems": [
            {
                "check": "low_contrast",
                "slide": 1,
                "text_color": "#BBBBBB",
                "bg_color": "#FFFFFF",
                "contrast_ratio": 1.86,
            },
        ],
    },
    "missing_alt_text.pptx": {
        "description": "Picture shape 'Chart Image' has no descr attribute on its cNvPr element.",
        "slide_count": 1,
        "known_problems": [
            {"check": "missing_alt_text", "slide": 1, "shape": "Chart Image"},
        ],
    },
    "no_title.pptx": {
        "description": "Single slide with only a body text box — no title element anywhere.",
        "slide_count": 1,
        "known_problems": [
            {"check": "missing_slide_title", "slide": 1},
        ],
    },
    "reading_order.pptx": {
        "description": (
            "Three labeled boxes added to XML in the order Step3→Step1→Step2, "
            "but their visual positions are top-to-bottom Step1→Step2→Step3."
        ),
        "slide_count": 1,
        "known_problems": [
            {
                "check": "reading_order",
                "slide": 1,
                "xml_order": ["Step 3", "Step 1", "Step 2"],
                "visual_order": ["Step 1", "Step 2", "Step 3"],
            },
        ],
    },
    "crowded.pptx": {
        "description": "Title box plus 12 data-item text boxes — 13 shapes on one slide.",
        "slide_count": 1,
        "known_problems": [
            {"check": "crowded_slide", "slide": 1, "shape_count": 13},
        ],
    },
}


# ── Builders map ─────────────────────────────────────────────────────────────

BUILDERS = {
    "clean.pptx":               build_clean,
    "overlapping_objects.pptx": build_overlapping_objects,
    "object_off_slide.pptx":    build_object_off_slide,
    "text_overflow.pptx":       build_text_overflow,
    "small_text.pptx":          build_small_text,
    "low_contrast.pptx":        build_low_contrast,
    "missing_alt_text.pptx":    build_missing_alt_text,
    "no_title.pptx":            build_no_title,
    "reading_order.pptx":       build_reading_order,
    "crowded.pptx":             build_crowded,
}


def main() -> None:
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    for filename, builder in BUILDERS.items():
        prs = builder()
        path = FIXTURES_DIR / filename
        prs.save(str(path))
        problems = len(MANIFEST[filename]["known_problems"])
        flag = "clean" if problems == 0 else f"{problems} known problem(s)"
        print(f"  {filename:35s}  [{flag}]")

    manifest_path = FIXTURES_DIR / "manifest.json"
    manifest_path.write_text(
        json.dumps(MANIFEST, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(f"\n  manifest.json written")
    print(f"\nAll {len(BUILDERS)} fixtures ready in {FIXTURES_DIR}")


if __name__ == "__main__":
    main()
