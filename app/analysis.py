"""
app/analysis.py — Measured analysis of PPTX files for Second Look.

Public API
----------
    analyze_pptx(path: str | Path) -> dict

All findings carry  type="Measured".  No AI opinions are made here.
When a color or size cannot be determined reliably the check is skipped
and that fact is not reported as an issue.
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

# ── Tunable thresholds ────────────────────────────────────────────────────────

SMALL_TEXT_PT: float = 12.0        # runs at or below this size are flagged
CROWDED_THRESHOLD: int = 8         # slides with more shapes than this are flagged
MANY_FONTS_THRESHOLD: int = 3      # more than this many distinct fonts in a deck
ALMOST_ALIGNED_MIN: int = 12_700   # 1 pt — below this edges are "close enough"
ALMOST_ALIGNED_MAX: int = 127_000  # 10 pt — above this misalignment is intentional
NEARBY_MAX: int = 1_828_800        # 2 in — only test alignment for nearby shape pairs
TITLE_REGION: float = 0.15         # top 15 % of slide height is the title zone
WCAG_NORMAL: float = 4.5           # WCAG 2.2 AA contrast for normal text
WCAG_LARGE: float = 3.0            # WCAG 2.2 AA contrast for large text
LARGE_TEXT_PT: float = 18.0        # large text threshold (regular weight)
LARGE_BOLD_PT: float = 14.0        # large text threshold (bold)
OVERFLOW_CHAR_WIDTH: float = 0.6   # fraction of font size for average character width
OVERFLOW_LINE_HEIGHT: float = 1.2  # line height as a multiple of font size
OVERFLOW_BUFFER: float = 1.5       # safety factor before flagging overflow


# ── Low-level helpers ─────────────────────────────────────────────────────────

def _emu_to_in(emu: int) -> float:
    return round(emu / 914_400, 4)


def _srgb_linear(c8: int) -> float:
    """Convert an 8-bit sRGB channel to linear light."""
    c = c8 / 255.0
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _luminance(r: int, g: int, b: int) -> float:
    """WCAG relative luminance for an RGB colour."""
    return (
        0.2126 * _srgb_linear(r)
        + 0.7152 * _srgb_linear(g)
        + 0.0722 * _srgb_linear(b)
    )


def _contrast(rgb1: tuple[int, int, int], rgb2: tuple[int, int, int]) -> float:
    """WCAG contrast ratio between two RGB colours."""
    l1, l2 = _luminance(*rgb1), _luminance(*rgb2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def _alt_text(shape) -> str | None:
    """Alt text from cNvPr.descr for any shape type, or None if absent/empty."""
    for attr in ("nvPicPr", "nvSpPr", "nvCxnSpPr"):
        try:
            cNvPr = getattr(shape._element, attr).cNvPr
            descr = (cNvPr.get("descr") or "").strip()
            return descr or None
        except AttributeError:
            continue
    return None


def _bg_color(shape) -> tuple[int, int, int] | None:
    """
    Effective background RGB for the shape.
    Returns (255, 255, 255) for no-fill or BACKGROUND fill (transparent → white slide).
    Returns None when the fill cannot be determined (theme colour, gradient, …).
    """
    try:
        fill = shape.fill
        fill_type = fill.type
        if fill_type is None:
            return (255, 255, 255)
        try:
            rgb = fill.fore_color.rgb      # raises for non-solid fills
            return (rgb[0], rgb[1], rgb[2])
        except Exception:
            # BACKGROUND (5) = transparent → white; gradient/picture → unknown
            if getattr(fill_type, "name", "") == "BACKGROUND":
                return (255, 255, 255)
            return None
    except Exception:
        return None


def _text_color(run) -> tuple[int, int, int] | None:
    """Explicit RGB text colour, or None if theme-inherited."""
    try:
        rgb = run.font.color.rgb
        return (rgb[0], rgb[1], rgb[2])
    except Exception:
        return None


# ── Shape extraction ──────────────────────────────────────────────────────────

_TYPE_NAMES: dict[int, str] = {
    1: "SHAPE", 3: "FREEFORM", 6: "LINE", 7: "MEDIA",
    11: "OLE_OBJECT", 13: "PICTURE", 14: "PLACEHOLDER", 17: "TEXT_BOX",
}


def _extract_shape(shape, xml_index: int, slide_w: int, slide_h: int) -> dict[str, Any]:
    """Return a plain-dict description of one shape."""
    l = shape.left or 0
    t = shape.top or 0
    w = shape.width or 0
    h = shape.height or 0

    font_sizes: list[float] = []
    font_names: list[str] = []
    full_text = ""
    auto_size_name: str | None = None

    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    size_pt = round(run.font.size / 12_700, 1)
                    if size_pt not in font_sizes:
                        font_sizes.append(size_pt)
                if run.font.name and run.font.name not in font_names:
                    font_names.append(run.font.name)
        full_text = shape.text_frame.text
        try:
            asv = shape.text_frame.auto_size
            auto_size_name = asv.name if asv is not None else None
        except Exception:
            pass

    is_pic = shape.shape_type == 13

    return {
        "xml_index": xml_index,
        "name": shape.name,
        "shape_type_id": int(shape.shape_type),
        "shape_type_name": _TYPE_NAMES.get(shape.shape_type, "OTHER"),
        "left_emu": l,
        "top_emu": t,
        "width_emu": w,
        "height_emu": h,
        "left_inches": _emu_to_in(l),
        "top_inches": _emu_to_in(t),
        "width_inches": _emu_to_in(w),
        "height_inches": _emu_to_in(h),
        "right_inches": _emu_to_in(l + w),
        "bottom_inches": _emu_to_in(t + h),
        "is_fully_on_slide": (
            l >= 0 and t >= 0
            and l + w <= slide_w
            and t + h <= slide_h
        ),
        "has_text_frame": shape.has_text_frame,
        "text": full_text,
        "font_sizes_pt": font_sizes,
        "font_names": font_names,
        "auto_size": auto_size_name,
        "alt_text": _alt_text(shape) if is_pic else None,
        "is_picture": is_pic,
    }


# ── Issue constructor ─────────────────────────────────────────────────────────

def _issue(check: str, severity: str, detail: str, **extra) -> dict[str, Any]:
    return {"type": "Measured", "check": check, "severity": severity,
            "detail": detail, **extra}


# ── Per-slide checks ──────────────────────────────────────────────────────────

def check_objects_off_slide(
    shapes: list[dict], slide_w: int, slide_h: int
) -> list[dict]:
    issues = []
    for s in shapes:
        l, t, w, h = s["left_emu"], s["top_emu"], s["width_emu"], s["height_emu"]
        if l < 0 or t < 0 or l + w > slide_w or t + h > slide_h:
            issues.append(_issue(
                "object_off_slide", "high",
                f"'{s['name']}' is outside the slide boundary "
                f"(left={s['left_inches']}\", top={s['top_inches']}\", "
                f"right={s['right_inches']}\", bottom={s['bottom_inches']}\").",
                shape=s["name"],
            ))
    return issues


def check_overlapping(shapes: list[dict]) -> list[dict]:
    issues = []
    seen: set[tuple] = set()
    for i, a in enumerate(shapes):
        for b in shapes[i + 1 :]:
            pair = (a["name"], b["name"])
            if pair in seen:
                continue
            if (
                a["left_emu"] < b["left_emu"] + b["width_emu"]
                and b["left_emu"] < a["left_emu"] + a["width_emu"]
                and a["top_emu"] < b["top_emu"] + b["height_emu"]
                and b["top_emu"] < a["top_emu"] + a["height_emu"]
            ):
                seen.add(pair)
                issues.append(_issue(
                    "overlapping_objects", "medium",
                    f"'{a['name']}' and '{b['name']}' overlap.",
                    shapes=[a["name"], b["name"]],
                ))
    return issues


def check_missing_title(
    shapes: list[dict], pptx_slide, slide_h: int
) -> list[dict]:
    threshold = int(slide_h * TITLE_REGION)
    # Check real title placeholders first
    for ph in pptx_slide.placeholders:
        if ph.placeholder_format.idx == 0 and ph.has_text_frame and ph.text_frame.text.strip():
            return []
    # Fall back to any text box in the title zone
    for s in shapes:
        if s["has_text_frame"] and s["text"].strip() and s["top_emu"] < threshold:
            return []
    return [_issue(
        "missing_slide_title", "high",
        f"No title found in the top {int(TITLE_REGION * 100)}% of this slide.",
    )]


def check_missing_alt_text(shapes: list[dict]) -> list[dict]:
    return [
        _issue(
            "missing_alt_text", "high",
            f"Image '{s['name']}' has no alt text.",
            shape=s["name"],
        )
        for s in shapes
        if s["is_picture"] and not s["alt_text"]
    ]


def check_small_text(shapes: list[dict]) -> list[dict]:
    issues = []
    for s in shapes:
        for size in s["font_sizes_pt"]:
            if size <= SMALL_TEXT_PT:
                issues.append(_issue(
                    "small_text", "medium",
                    f"'{s['name']}' has text at {size}pt "
                    f"(minimum recommended: {SMALL_TEXT_PT}pt).",
                    shape=s["name"],
                    font_size_pt=size,
                ))
                break  # one issue per shape
    return issues


def check_empty_placeholders(pptx_slide) -> list[dict]:
    issues = []
    for ph in pptx_slide.placeholders:
        if not ph.has_text_frame:
            continue
        if not ph.text_frame.text.strip():
            issues.append(_issue(
                "empty_placeholder", "low",
                f"Placeholder '{ph.name}' (index {ph.placeholder_format.idx}) is empty.",
                shape=ph.name,
            ))
    return issues


def check_text_overflow(shapes: list[dict], pptx_slide) -> list[dict]:
    """
    Flag text boxes where auto_size=NONE and estimated text volume exceeds
    the box capacity.  Confidence: low (heuristic only).
    """
    issues = []
    pptx_map = {s.name: s for s in pptx_slide.shapes}

    for sd in shapes:
        if not sd["has_text_frame"] or sd["auto_size"] != "NONE":
            continue
        ps = pptx_map.get(sd["name"])
        if ps is None:
            continue

        flagged = False
        for para in ps.text_frame.paragraphs:
            if flagged:
                break
            for run in para.runs:
                if not run.font.size:
                    continue
                fs_pt = run.font.size / 12_700
                box_w_pt = sd["width_emu"] / 12_700
                box_h_pt = sd["height_emu"] / 12_700
                text_len = len(ps.text_frame.text)
                chars_per_line = box_w_pt / (fs_pt * OVERFLOW_CHAR_WIDTH)
                if chars_per_line <= 0:
                    continue
                lines_needed = math.ceil(text_len / chars_per_line)
                max_lines = box_h_pt / (fs_pt * OVERFLOW_LINE_HEIGHT)
                if lines_needed > max_lines * OVERFLOW_BUFFER:
                    issues.append(_issue(
                        "text_overflow", "medium",
                        f"'{sd['name']}' may overflow: estimated {lines_needed} "
                        f"lines needed, box fits approximately {max_lines:.0f}. "
                        f"(Heuristic — confidence: low)",
                        shape=sd["name"],
                        confidence="low",
                    ))
                    flagged = True
                    break

    return issues


def check_reading_order(shapes: list[dict]) -> list[dict]:
    """
    Flag slides where XML order differs from top-to-bottom, left-to-right
    visual order.  Reports when any two shapes are 'swapped'.
    """
    if len(shapes) < 2:
        return []

    xml_names = [s["name"] for s in shapes]
    visual_names = [
        s["name"]
        for s in sorted(shapes, key=lambda s: (s["top_emu"], s["left_emu"]))
    ]

    xml_pos = {name: i for i, name in enumerate(xml_names)}
    vis_pos = {name: i for i, name in enumerate(visual_names)}

    inversions = [
        (a, b)
        for i, a in enumerate(xml_names)
        for b in xml_names[i + 1 :]
        if vis_pos[a] > vis_pos[b]
    ]

    if not inversions:
        return []

    ex_a, ex_b = inversions[0]
    return [_issue(
        "reading_order", "medium",
        f"Reading order may not match visual layout. "
        f"'{ex_a}' comes before '{ex_b}' in the file, "
        f"but '{ex_b}' appears higher on the slide. "
        f"{len(inversions)} pair(s) affected.",
        inversion_count=len(inversions),
    )]


def check_crowded(shapes: list[dict]) -> list[dict]:
    if len(shapes) > CROWDED_THRESHOLD:
        return [_issue(
            "crowded_slide", "low",
            f"Slide has {len(shapes)} shapes (threshold: {CROWDED_THRESHOLD}). "
            f"Crowded slides can be hard to follow.",
            shape_count=len(shapes),
        )]
    return []


def check_almost_aligned(shapes: list[dict]) -> list[dict]:
    """
    Flag nearby shape pairs whose edges are almost — but not exactly — aligned.
    """
    issues = []
    seen: set[frozenset] = set()

    for i, a in enumerate(shapes):
        for b in shapes[i + 1 :]:
            pair: frozenset = frozenset({a["name"], b["name"]})
            if pair in seen:
                continue

            # Only check pairs that are "nearby"
            gap_h = max(
                0,
                max(a["left_emu"], b["left_emu"])
                - min(a["left_emu"] + a["width_emu"], b["left_emu"] + b["width_emu"]),
            )
            gap_v = max(
                0,
                max(a["top_emu"], b["top_emu"])
                - min(a["top_emu"] + a["height_emu"], b["top_emu"] + b["height_emu"]),
            )
            if gap_h > NEARBY_MAX and gap_v > NEARBY_MAX:
                continue

            edges = {
                "left edges":   abs(a["left_emu"] - b["left_emu"]),
                "right edges":  abs(
                    (a["left_emu"] + a["width_emu"])
                    - (b["left_emu"] + b["width_emu"])
                ),
                "top edges":    abs(a["top_emu"] - b["top_emu"]),
                "bottom edges": abs(
                    (a["top_emu"] + a["height_emu"])
                    - (b["top_emu"] + b["height_emu"])
                ),
            }

            for edge_name, diff in edges.items():
                if ALMOST_ALIGNED_MIN <= diff <= ALMOST_ALIGNED_MAX:
                    seen.add(pair)
                    gap_pt = round(diff / 12_700, 1)
                    issues.append(_issue(
                        "almost_aligned", "low",
                        f"'{a['name']}' and '{b['name']}' have {edge_name} "
                        f"that are {gap_pt}pt apart — almost but not exactly aligned.",
                        shapes=[a["name"], b["name"]],
                        edge=edge_name,
                        gap_pt=gap_pt,
                    ))
                    break  # one issue per pair

    return issues


def check_low_contrast(shapes: list[dict], pptx_slide) -> list[dict]:
    """
    Check text contrast against background.  Only reports when both colours
    can be read as explicit RGB values.
    """
    issues = []
    pptx_map = {s.name: s for s in pptx_slide.shapes}

    for sd in shapes:
        if not sd["has_text_frame"]:
            continue
        ps = pptx_map.get(sd["name"])
        if ps is None:
            continue

        bg = _bg_color(ps)
        if bg is None:
            continue

        flagged = False
        for para in ps.text_frame.paragraphs:
            if flagged:
                break
            for run in para.runs:
                if not run.text.strip():
                    continue
                fg = _text_color(run)
                if fg is None:
                    continue

                ratio = _contrast(fg, bg)
                fs_emu = run.font.size
                fs_pt = (fs_emu / 12_700) if fs_emu else 12.0
                is_bold = run.font.bold or False

                if fs_pt >= LARGE_TEXT_PT or (is_bold and fs_pt >= LARGE_BOLD_PT):
                    threshold, size_label = WCAG_LARGE, "large"
                else:
                    threshold, size_label = WCAG_NORMAL, "normal"

                if ratio < threshold:
                    fg_hex = "#{:02X}{:02X}{:02X}".format(*fg)
                    bg_hex = "#{:02X}{:02X}{:02X}".format(*bg)
                    issues.append(_issue(
                        "low_contrast", "high",
                        f"'{sd['name']}' has {size_label} text {fg_hex} on "
                        f"{bg_hex} background. "
                        f"Contrast: {ratio:.2f}:1 "
                        f"(WCAG AA requires {threshold}:1).",
                        shape=sd["name"],
                        text_color=fg_hex,
                        bg_color=bg_hex,
                        contrast_ratio=round(ratio, 2),
                        threshold=threshold,
                    ))
                    flagged = True
                    break

    return issues


# ── Deck-level checks ─────────────────────────────────────────────────────────

def check_duplicate_titles(slides: list[dict]) -> list[dict]:
    seen: dict[str, list[int]] = {}
    for slide in slides:
        t = slide.get("title", "").strip()
        if t:
            seen.setdefault(t, []).append(slide["slide_number"])
    return [
        _issue(
            "duplicate_title", "medium",
            f"The title \"{t}\" appears on slides {nums}.",
            title=t,
            slides=nums,
        )
        for t, nums in seen.items()
        if len(nums) > 1
    ]


def check_many_fonts(deck_fonts: list[str]) -> list[dict]:
    if len(deck_fonts) > MANY_FONTS_THRESHOLD:
        return [_issue(
            "many_fonts", "low",
            f"Deck uses {len(deck_fonts)} different fonts: {deck_fonts}. "
            f"More than {MANY_FONTS_THRESHOLD} can look inconsistent.",
            fonts=deck_fonts,
            font_count=len(deck_fonts),
        )]
    return []


# ── Title extraction helper ───────────────────────────────────────────────────

def _slide_title(shapes: list[dict], pptx_slide, slide_h: int) -> str:
    # Priority 1: title placeholder (idx=0)
    for ph in pptx_slide.placeholders:
        if ph.placeholder_format.idx == 0 and ph.has_text_frame:
            text = ph.text_frame.text.strip()
            if text:
                return text
    # Priority 2: text box in the title region
    threshold = int(slide_h * TITLE_REGION)
    for s in shapes:
        if s["has_text_frame"] and s["text"].strip() and s["top_emu"] < threshold:
            return s["text"].strip()
    return ""


# ── Public API ────────────────────────────────────────────────────────────────

def analyze_pptx(path: str | Path) -> dict[str, Any]:
    """
    Open a PPTX file and return structured measured data with issues.

    Return shape:
        file_name         str
        slide_width_inches  float
        slide_height_inches float
        slide_count       int
        fonts_used        list[str]
        deck_issues       list[dict]   — issues spanning the whole deck
        slides            list[dict]   — per-slide data and issues
    """
    path = Path(path)
    prs = Presentation(str(path))
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    all_slides: list[dict] = []
    deck_fonts: set[str] = set()

    for slide_num, pptx_slide in enumerate(prs.slides, start=1):
        shapes = [
            _extract_shape(s, i, slide_w, slide_h)
            for i, s in enumerate(pptx_slide.shapes)
        ]
        for s in shapes:
            deck_fonts.update(s["font_names"])

        title = _slide_title(shapes, pptx_slide, slide_h)

        issues: list[dict] = []
        issues += check_objects_off_slide(shapes, slide_w, slide_h)
        issues += check_overlapping(shapes)
        issues += check_missing_title(shapes, pptx_slide, slide_h)
        issues += check_missing_alt_text(shapes)
        issues += check_small_text(shapes)
        issues += check_empty_placeholders(pptx_slide)
        issues += check_text_overflow(shapes, pptx_slide)
        issues += check_reading_order(shapes)
        issues += check_crowded(shapes)
        issues += check_almost_aligned(shapes)
        issues += check_low_contrast(shapes, pptx_slide)

        all_slides.append({
            "slide_number": slide_num,
            "title": title,
            "shape_count": len(shapes),
            "shapes": shapes,
            "issues": issues,
        })

    sorted_fonts = sorted(deck_fonts)
    deck_issues: list[dict] = []
    deck_issues += check_duplicate_titles(all_slides)
    deck_issues += check_many_fonts(sorted_fonts)

    return {
        "file_name": path.name,
        "slide_width_inches": _emu_to_in(slide_w),
        "slide_height_inches": _emu_to_in(slide_h),
        "slide_count": len(all_slides),
        "fonts_used": sorted_fonts,
        "deck_issues": deck_issues,
        "slides": all_slides,
    }
