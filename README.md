# Second Look

A visual second opinion on PowerPoint decks — built for blind and low-vision professionals.

Second Look measures what is on each slide, describes how it looks using AI, lists problems
by severity, and answers questions about your deck.

---

## Requirements

- Python 3.11 or later
- LibreOffice 7.6 or later (see installation steps below — required for Step 4)
- An Anthropic API key (required for Step 5)

---

## Install LibreOffice (Windows)

LibreOffice is not currently installed. You need it to convert PowerPoint slides to images
in Step 4. You do not need it for Steps 1–3.

1. Open a browser and navigate to https://www.libreoffice.org/download/libreoffice-still/
2. Select the latest "Still" release for Windows (64-bit).
3. Run the installer and accept the defaults.
4. After installation, copy the full path to `soffice.exe` — typically
   `C:\Program Files\LibreOffice\program\soffice.exe`.
5. Add that path to your `.env` file as `LIBREOFFICE_PATH`.

---

## Set Up

```
# 1. Create and activate the virtual environment
python -m venv .venv
.venv\Scripts\activate        # Windows PowerShell

# 2. Install dependencies
pip install -r requirements.txt

# 3. Copy the example environment file and fill in your values
copy .env.example .env
# Then open .env and add your ANTHROPIC_API_KEY
```

---

## Start the App

```
uvicorn app.main:app --reload
```

Then open a browser and navigate to http://127.0.0.1:8000

---

## Run the Tests

```
pytest
```

---

## Project Status

| Step | Description | Status |
|------|-------------|--------|
| 1 | Project setup | Done |
| 2 | Test fixtures | Pending |
| 3 | Measured analysis | Pending |
| 4 | Slide rendering | Pending |
| 5 | AI visual review | Pending |
| 6 | Accessible web interface | Pending |
| 7 | Questions | Pending |
| 8 | Pilot checkpoint | Pending |
| 9 | Safe fixes | Pending |
| 10 | Paid beta preparation | Pending |

---

## Privacy

Uploaded files are treated as confidential. Temporary files and slide images are deleted
when processing ends. No file contents or personal data are written to logs.
