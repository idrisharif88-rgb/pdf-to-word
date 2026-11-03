# PDF to Word — Arabic OCR

A desktop app that turns scanned Arabic books (PDF or image) into editable Word
documents, using Google Document AI for OCR.

It is built for Arabic text: right-to-left paragraphs, Arabic fonts and
punctuation, and headings detected from the scan's type sizes. Running heads
and printed page numbers are removed, and footnotes are set apart.

## Features

- **Arabic OCR:** Google Document AI OCR with Arabic and English language hints
- **Word output:**
  - Real Word headings, so the Navigation pane and table of contents work
  - Page-number footers and mirrored book margins
- **Born-digital PDFs:** PDFs that already contain text are converted
  directly, with no OCR cost
- **Large books:** pages are sent in batches of 15, in parallel
  - Each request uploads only that batch's pages
  - Requests that time out are retried
- **Resume:** if a run stops, finished batches are kept and skipped next time
- **Optional:**
  - Tables, using a Form Parser processor (an extra paid request)
  - Exact layout that matches the scan
  - "Keep as picture" pages
  - Original page markers (صفحة N)
- **Desktop app:** native GNOME file picker on Linux, completion sound, and
  the finished file shown in your file manager

## Requirements

- Python 3.10 or newer, with Tk (`sudo apt install python3-tk` on Ubuntu)
- A Google Cloud project with Document AI enabled and a **Document OCR**
  processor (optionally a **Form Parser** processor for tables)
- Linux, Windows or macOS. On Linux, `zenity` gives the native file picker.

## Setup

```sh
git clone https://github.com/idrisharif88-rgb/pdf-to-word.git
cd pdf-to-word
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` and fill in your own values:

```ini
GCP_PROJECT_ID=your-gcp-project-id
GCP_LOCATION=eu
GCP_PROCESSOR_ID=your-ocr-processor-id
# Optional, for tables:
GCP_FORM_PROCESSOR_ID=your-form-parser-id
```

Then set up authentication. Either:

- run `gcloud auth application-default login`, or
- set `GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json` in `.env`.

`.env` and key files are git-ignored. Never commit them.

## Run

```sh
.venv/bin/python -m pdf_to_word
```

On Linux, `./install-launcher.sh` adds a "PDF to Word" launcher to the app
menu and the Desktop.

1. Choose the input file. The output path is filled in next to it.
2. Enter the pages, for example `all`, `1-50` or `3,7,12-15`.
3. Switch options on or off, then press **Run OCR**.

The finished file is saved as `<name>_merged.docx`.

## Settings

Every setting can be changed in `.env`. Each one is the upper-case name of a
field in [`pdf_to_word/config.py`](pdf_to_word/config.py), for example:

| Setting | Default | What it does |
|---|---|---|
| `ARABIC_FONT` | Aref Ruqaa | Font for Arabic text |
| `ARABIC_SIZE_PT` | 18 | Body text size |
| `BATCH_SIZE` | 15 | Pages per OCR request |
| `EXTRACT_TABLES` | 0 | Use the Form Parser for tables |
| `NATIVE_CONVERT` | 1 | Convert text PDFs directly, without OCR |
| `MARK_SOURCE_PAGES` | 0 | Add "صفحة N" at the top of each page |
| `REQUEST_TIMEOUT` | 300 | Seconds per Document AI call |

## Cost

Each batch is one Document OCR request. With tables switched on, each batch is
also one Form Parser request. Check Google's Document AI pricing for current
rates. Converting a PDF that already has text costs nothing.

## Project layout

```
pdf_to_word/
  config.py       settings from .env
  ocr_client.py   the only code that calls Google Document AI
  pipeline.py     batching, threads, resume, cancel, merge
  builder.py      OCR result -> Word document for one batch
  text.py         Arabic text cleanup and layout heuristics
  docx_format.py  right-to-left Word formatting
  tables.py, merge.py, pdf_tools.py, desktop.py
  gui/            Tkinter app and round widgets
tests/            pytest suite with a fake Document AI client
```

## Development

```sh
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest
```

The tests use a fake Document AI client, so they need no Google account and
cost nothing.
