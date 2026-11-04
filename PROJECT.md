# Second Look — Project Brief

You are helping Aaron build Second Look, an accessible web app that gives blind and
low-vision professionals a visual second opinion on their work files. Second Look is a
working name.

The first version supports PowerPoint files only. It must explain how a slide deck looks,
find visual and accessibility problems, answer questions, and eventually make safe fixes.

Aaron is a blind accessibility professional who uses JAWS, NVDA, and VoiceOver. He knows
HTML, CSS, JavaScript, and ARIA. Explain your work in plain language. Use words such as
"select" and "navigate" instead of "click." Do not assume he can review anything visually.

---

## HOW TO WORK

1. Save this project brief as PROJECT.md before doing anything else.
2. Work on one step at a time.
3. Before each step, explain your plan in no more than five sentences.
4. After each step, list what you built, how to test it, and any problems.
5. Stop after each step and wait for Aaron to approve the next step.
6. Do not skip tests or mark a step complete if a test fails.
7. Keep a plain-language CHANGELOG.md.
8. Commit to Git after each approved step.
9. Do not add features outside the current step.
10. If a decision would change the product, ask Aaron first.

---

## PRODUCT RULES

1. Accessibility test rule: If a person cannot finish a task from start to end using only
   a keyboard and a screen reader, the feature is not done.
2. Meet WCAG 2.2 Level AA.
3. Lead with the answer. Keep responses short by default, and let the user ask for more detail.
4. Separate measured facts from AI judgment. Label each finding as "Measured" or "AI opinion."
5. Include a confidence level for AI opinions. Never present a guess as a fact.
6. When using the LibreOffice adapter, tell users that slide images come from LibreOffice
   and may differ slightly from PowerPoint. Update this disclaimer if a different rendering
   adapter is used in future.
7. Never overwrite the original file. Always create a new copy.
8. Do not describe a slide as visually perfect or guarantee that a file is accessible.
9. Treat every uploaded file as confidential.
10. Delete uploaded files and generated images after a session ends.
11. Do not put file contents, personal data, or API keys in logs.
12. Keep the API key in an environment variable named ANTHROPIC_API_KEY.
13. Keep the Claude model name in one setting. Check Anthropic's current documentation
    before choosing it.

---

## TECHNICAL DIRECTION

- Python, FastAPI, Jinja templates, semantic HTML, CSS
- Only the JavaScript needed for progressive enhancement
- python-pptx to read PowerPoint files
- PyMuPDF to convert PDF pages to images
- Anthropic API for visual review
- pytest for automated tests
- axe-core for automated accessibility checks
- No user accounts, payments, mobile app, or database yet

### Slide rendering interface

Slide rendering must sit behind a defined interface so the rendering method can be
swapped without touching the analysis module, the AI review, or the web interface.

The interface accepts a PPTX file path and a session output directory. It returns a
list of image file paths, one per slide, in order. That is the only contract the rest
of the application depends on.

**Pilot adapter — LibreOffice**
LibreOffice converts the PPTX to PDF; PyMuPDF converts each PDF page to a PNG.
This is the only adapter that will be built before the pilot.

**Future adapters (do not build yet)**
1. PowerPoint Office Add-in — renders slides in the user's locally installed PowerPoint.
   This avoids server-side Office automation entirely.
2. Microsoft Graph API — converts files stored in OneDrive or SharePoint without
   requiring Office to be installed on the server.
3. Google Slides API — renders native Google Slides files. If this adapter is ever
   built, request only the `drive.file` permission and require the user to choose the
   file themselves. Do not request broad Drive access.

**Constraints**
- Do not install Microsoft Office on a server.
- Do not automate Microsoft Office on a server in any way.
- Do not build the PowerPoint Add-in or the Google adapter yet.
- Product rule 6 (tell users images come from LibreOffice) applies to the LibreOffice
  adapter only. Update the disclaimer text if a different adapter is used in future.

---

## STEPS

### STEP 1: SET UP THE PROJECT
Create the project structure, virtual environment, requirements file, .gitignore,
README.md, and .env.example. Check whether LibreOffice is installed. Create a basic
home page with one level 1 heading.

Done when: app runs locally, tests pass, README explains how to start the app,
home page works by keyboard and screen reader.

### STEP 2: CREATE TEST FILES
Create a small set of PowerPoint test files in tests/fixtures with known problems.

### STEP 3: BUILD THE MEASURED ANALYSIS
Module that accepts a PPTX and returns structured JSON for each slide with measured checks.

### STEP 4: RENDER THE SLIDES
Convert PPTX to PDF with LibreOffice, then create one image per slide.

### STEP 5: ADD THE AI VISUAL REVIEW
Send each slide image and measured data to the Anthropic API for structured JSON review.

### STEP 6: BUILD THE ACCESSIBLE WEB INTERFACE
Upload, processing status, deck summary, slide detail, and issue list screens.

### STEP 7: ADD QUESTIONS
Question box on deck and slide pages with starter questions and "More detail" option.

### STEP 8: PILOT CHECKPOINT
Pilot guide for five blind or low-vision professionals. Stop new features until pilot done.

Include this question in the pilot feedback form:
"Which tool do you mainly use to create or edit presentations?
  a) PowerPoint desktop app
  b) PowerPoint for the web (Office 365 / Microsoft 365 in a browser)
  c) Google Slides
  d) Other (please describe)"

Use the answers to decide whether to build the Microsoft Graph or Google Slides adapter
before the paid beta, and to prioritise which adapter to build first.

### STEP 9: ADD SAFE FIXES (after pilot approval)
Let users request specific changes with confirmation before applying.

### STEP 10: PREPARE FOR A PAID BETA
Security review, privacy policy draft, terms of service draft, support process,
deployment plan. No payments until Aaron approves pricing and business setup.

---

## SCOPE DECISIONS

**Word support (decided 2026-11-04):** Word file review is a confirmed feature, but will
be built as a dedicated phase after all PowerPoint steps are complete. It is not deferred
to pilot demand — it is simply sequenced after the PowerPoint version is working.
