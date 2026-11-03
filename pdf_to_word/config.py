"""Settings, read once from the environment (and a .env file if present).

Everything the conversion needs is carried in one immutable ``Settings``
object and passed explicitly, so a job's options can never change under a
worker thread mid-run. The GUI makes a per-job copy with
``dataclasses.replace``.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, fields
from pathlib import Path
from typing import Mapping

PROJECT_DIR = Path(__file__).resolve().parent.parent


def load_dotenv(path: Path = PROJECT_DIR / ".env") -> None:
    """Minimal .env loader; real environment variables win over the file."""
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip("'\""))


def _parse_bool(value: str) -> bool:
    return value.strip().lower() in ("1", "true", "yes", "on")


@dataclass(frozen=True)
class Settings:
    # --- Google Cloud ---
    gcp_project_id: str = "your-gcp-project-id"
    gcp_location: str = "eu"
    gcp_processor_id: str = "your-ocr-processor-id"
    # Optional Form Parser processor. A Document OCR processor returns no
    # tables at all, so table extraction needs a Form Parser to point at.
    gcp_form_processor_id: str = ""
    # Optional service-account JSON key; empty means Application Default
    # Credentials (gcloud auth application-default login).
    google_application_credentials: str = ""
    # Seconds to wait for one Document AI call, and total retry budget.
    request_timeout: float = 300.0
    retry_deadline: float = 900.0

    # --- Batching ---
    batch_size: int = 15
    max_workers: int = 0  # 0 = choose from CPU count

    # --- Arabic typography ---
    # Aref Ruqaa is a calligraphic Ruqaa face; it sits taller than Arial, so
    # it needs looser leading to stay readable.
    arabic_font: str = "Aref Ruqaa"
    # Latin lines read better in a normal text face than in Ruqaa's Latin.
    latin_font: str = "Times New Roman"
    arabic_size_pt: float = 18.0
    line_spacing: float = 1.6
    page_margin_cm: float = 2.5
    mirror_margins: bool = True
    arabic_bold: bool = True
    add_toc: bool = True

    # --- Tables ---
    extract_tables: bool = False  # Form Parser is an extra paid call; opt in
    # A detected grid this empty is almost always a misread layout.
    table_min_fill: float = 0.45
    table_max_cols: int = 10
    table_size_scale: float = 0.75

    # --- Page furniture ---
    # Running heads, folios and watermarks repeat at the same spot page after
    # page; they belong to the scan, not to the text.
    drop_running_heads: bool = True
    margin_band: float = 0.12

    # --- Born-digital PDFs ---
    # Such a PDF already holds its own text; OCR would throw that away.
    native_convert: bool = True
    native_min_chars: int = 150
    # Share of pages that must carry text before the file counts as digital.
    native_min_page_ratio: float = 0.95

    # --- Layout ---
    # Exact layout puts every block in a text box at its measured position.
    # The page then matches the scan, but the text no longer reflows.
    exact_layout: bool = False
    extract_images: bool = True
    # An image this large IS the page scan, not a figure printed on it.
    image_max_page_ratio: float = 0.7
    heading_ratio_1: float = 1.5
    heading_ratio_2: float = 1.25
    # Footnotes sit low on the page and are set smaller than the body.
    footnote_band: float = 0.78
    footnote_max_ratio: float = 0.88
    footnote_size_scale: float = 0.7
    # Pages to carry over as a picture instead of text, e.g. "1-6,140-146".
    image_pages: str = ""
    # A page that yields almost no text is a plate; keep the picture.
    image_fallback_chars: int = 40
    page_image_dpi: int = 150
    # Poetry needs the OCR line breaks; prose does not.
    keep_line_breaks: bool = False

    # --- OCR options chosen per job ---
    enable_quality_scores: bool = True
    cleanup: bool = True

    @property
    def form_parser_available(self) -> bool:
        return bool(self.gcp_form_processor_id)

    @property
    def use_form_parser(self) -> bool:
        return self.extract_tables and self.form_parser_available

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> "Settings":
        """Build settings from environment variables named like the fields
        in upper case (e.g. ``ARABIC_FONT``)."""
        if env is None:
            load_dotenv()
            env = os.environ
        values = {}
        for field in fields(cls):
            raw = env.get(field.name.upper())
            if raw is None or raw == "":
                continue
            kind = type(field.default)
            if kind is bool:
                values[field.name] = _parse_bool(raw)
            elif kind is int:
                values[field.name] = int(raw)
            elif kind is float:
                values[field.name] = float(raw)
            else:
                values[field.name] = raw
        return cls(**values)
