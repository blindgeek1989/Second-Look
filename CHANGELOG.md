# Changelog

All changes are listed in plain language, newest first.

---

## Step 3 — Measured analysis (2026-11-04)

- Added `app/analysis.py` with `analyze_pptx(path) -> dict`.
- Returns structured JSON for every slide: dimensions, shape positions, text, fonts, alt text.
- Implements 13 measured checks across two levels:

  **Per-slide checks** (11):
  - `object_off_slide` — shape bounding box extends outside slide boundary
  - `overlapping_objects` — two shapes' bounding boxes intersect
  - `missing_slide_title` — no text in the top 15% of the slide
  - `missing_alt_text` — picture shape has no alt text attribute
  - `small_text` — font run at or below 12pt
  - `empty_placeholder` — template placeholder with no content
  - `text_overflow` — fixed-size text box estimated to exceed its capacity (confidence: low)
  - `reading_order` — XML order does not match top-to-bottom visual order
  - `crowded_slide` — more than 8 shapes on a slide
  - `almost_aligned` — two nearby shapes whose edges are within 1–10pt of aligning
  - `low_contrast` — text/background contrast below WCAG 2.2 AA (4.5:1 normal, 3:1 large)

  **Deck-level checks** (2):
  - `duplicate_title` — the same title text appears on more than one slide
  - `many_fonts` — more than 3 distinct fonts in the deck

- Added `tests/test_analysis.py` — 37 tests; each fixture triggers exactly one check,
  `clean.pptx` triggers none.
- Fixed `tests/create_fixtures.py` — narrowed clean body box to prevent false overlap with image.
- Fixed `_bg_color()` — text boxes have `MSO_FILL_TYPE.BACKGROUND` (not `None`); both are now
  correctly treated as transparent (white background).

---

## Step 2 — Test fixtures (2026-11-04)

- Added `tests/create_fixtures.py` — script that generates all 10 fixture PPTX files.
- Added `tests/fixtures/manifest.json` — machine-readable list of every known problem in each fixture.
- Added `tests/test_fixtures.py` — 49 tests verifying file existence, valid PPTX format, and structural properties.
- Added `tests/conftest.py` — auto-runs the fixture generator if files are missing.
- Fixed a python-pptx quirk: `add_picture()` auto-sets `descr="image.png"` on BytesIO streams; the missing-alt-text fixture now explicitly removes that attribute.

Fixtures and their documented problems:

| File | Problem |
|------|---------|
| `clean.pptx` | None — false-positive baseline |
| `overlapping_objects.pptx` | Red Box and Blue Box overlap |
| `object_off_slide.pptx` | Off-Slide Shape starts at x=11" |
| `text_overflow.pptx` | Fixed-size box with far too much text |
| `small_text.pptx` | Body run at 8pt |
| `low_contrast.pptx` | #BBBBBB text on white (contrast ≈ 1.86:1) |
| `missing_alt_text.pptx` | Image with no alt text attribute |
| `no_title.pptx` | No title element of any kind |
| `reading_order.pptx` | XML order is Step3→Step1→Step2; visual order is Step1→Step2→Step3 |
| `crowded.pptx` | 13 shapes on one slide |

---

## Step 1 — Project setup (2026-11-04)

- Created project folder structure (app, tests, static, templates).
- Added requirements.txt with all Step 1 dependencies.
- Added .gitignore, .env.example, and README.md.
- Created basic FastAPI app with a home page containing one level 1 heading.
- Added pytest smoke test confirming the home page returns HTTP 200.
- Noted that LibreOffice is not yet installed (needed for Step 4).
