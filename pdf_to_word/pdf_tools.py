"""Local PDF work: text-layer check, page extraction, rendering, figures.

pymupdf also opens JPG/PNG input as a one-page document.

pymupdf documents are not shared between threads, so each helper opens the
file itself.
"""
from __future__ import annotations

import logging
from io import BytesIO
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from pypdf import PdfReader, PdfWriter

from .config import Settings

log = logging.getLogger(__name__)

PDF_MIME = "application/pdf"
IMAGE_MIMES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".png": "image/png", ".tif": "image/tiff",
               ".tiff": "image/tiff"}
SUPPORTED_SUFFIXES = (".pdf", *IMAGE_MIMES)


def mime_type_for(path: str) -> str:
    suffix = Path(path).suffix.lower()
    if suffix == ".pdf":
        return PDF_MIME
    try:
        return IMAGE_MIMES[suffix]
    except KeyError:
        raise ValueError(f"Unsupported file type: {suffix or path}") from None


def is_pdf(path: str) -> bool:
    return Path(path).suffix.lower() == ".pdf"


def page_count(path: str) -> int:
    return len(PdfReader(path).pages) if is_pdf(path) else 1


def has_text_layer(path: str, settings: Settings) -> bool:
    """True only when essentially every page carries real text.

    An average is the wrong test: a deck where most pages are text and the
    rest are screenshots would convert to a mix of text and pictures. A mixed
    file goes through OCR so the whole document comes out the same way. A
    repeated footer is text on every page, so the smallest page count is
    treated as boilerplate.
    """
    if not is_pdf(path):
        return False
    try:
        counts = [len((pg.extract_text() or "").strip())
                  for pg in PdfReader(path).pages]
        if not counts:
            return False
        boilerplate = min(counts)
        with_text = sum(1 for c in counts
                        if c - boilerplate >= settings.native_min_chars)
        return with_text / len(counts) >= settings.native_min_page_ratio
    except Exception:
        log.warning("Text-layer check failed", exc_info=True)
        return False


def convert_native_pdf(path: str, output_path: str,
                       pages: Sequence[int]) -> None:
    """Convert a born-digital PDF straight to Word, keeping its layout."""
    from pdf2docx import Converter  # only this path needs it

    converter = Converter(path)
    try:
        converter.convert(output_path, pages=[n - 1 for n in pages])
    finally:
        converter.close()


def extract_pages(path: str, pages: Sequence[int]) -> bytes:
    """A new PDF holding only ``pages`` (1-based), in that order.

    Uploading just the batch instead of the whole file keeps each request
    small, which is what stops long books from hitting deadline errors.
    """
    reader = PdfReader(path)
    writer = PdfWriter()
    for number in pages:
        writer.add_page(reader.pages[number - 1])
    buffer = BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def render_page_image(path: str, page_number: int, dpi: int) -> Optional[bytes]:
    """The whole page as PNG bytes, for layouts not worth rebuilding."""
    try:
        import pymupdf
    except ImportError:
        log.warning("pymupdf missing; cannot render page images.")
        return None
    try:
        with pymupdf.open(path) as pdf:
            if page_number - 1 >= len(pdf):
                return None
            return pdf[page_number - 1].get_pixmap(dpi=dpi).tobytes("png")
    except Exception:
        log.warning("Could not render page %s", page_number, exc_info=True)
        return None


def page_figures(path: str, page_number: int,
                 max_page_ratio: float) -> List[Tuple[bytes, object]]:
    """Figures printed on the page, excluding the page scan itself."""
    try:
        import pymupdf
    except ImportError:
        return []
    figures = []
    try:
        with pymupdf.open(path) as pdf:
            if page_number - 1 >= len(pdf):
                return []
            page = pdf[page_number - 1]
            page_area = page.rect.width * page.rect.height
            for entry in page.get_images(full=True):
                xref = entry[0]
                for rect in page.get_image_rects(xref):
                    if page_area and rect.width * rect.height / page_area > max_page_ratio:
                        continue
                    if rect.width < 20 or rect.height < 20:
                        continue
                    figures.append((pdf.extract_image(xref)["image"], rect))
    except Exception:
        log.warning("Figure extraction failed on page %s", page_number,
                    exc_info=True)
    return figures


def pdf_page_size(path: str, page_number: int) -> Optional[Tuple[float, float]]:
    """Page size in points, so the Word page can match the original."""
    try:
        import pymupdf
        with pymupdf.open(path) as pdf:
            if page_number - 1 >= len(pdf):
                return None
            rect = pdf[page_number - 1].rect
            return rect.width, rect.height
    except Exception:
        log.warning("Page size lookup failed", exc_info=True)
        return None
