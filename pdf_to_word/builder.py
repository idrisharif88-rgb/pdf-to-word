"""Turn one batch of Document AI output into a Word document."""
from __future__ import annotations

import logging
import statistics
from io import BytesIO
from typing import List, Optional, Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from . import pdf_tools
from .config import Settings
from .docx_format import (add_footnote_rule, add_positioned_textbox,
                          add_text_with_breaks, format_paragraph,
                          setup_document)
from .tables import add_tables, anchor_text, table_regions
from .text import (TextUnit, clean_text_smart, detect_placement, detect_rtl,
                   heading_level, in_spans, is_footnote, is_ocr_noise,
                   parse_page_spec, repeat_key, running_head_texts,
                   unwrap_lines)

log = logging.getLogger(__name__)


def page_text_units(document, page) -> List[TextUnit]:
    """One TextUnit per OCR paragraph (paragraphs track print far better
    than blocks, which fuse columns)."""
    units = []
    for unit in list(page.paragraphs) or list(page.blocks):
        layout = unit.layout
        text = anchor_text(document, layout)
        verts = layout.bounding_poly.normalized_vertices
        top = min((v.y for v in verts), default=0.0)
        bottom = max((v.y for v in verts), default=0.0)
        left = min((v.x for v in verts), default=0.0)
        right = max((v.x for v in verts), default=1.0)
        height = bottom - top if verts else 0.0
        # The box grows with the number of printed lines; compare one line.
        line_height = height / max(1, text.count("\n") + 1)
        units.append(TextUnit(text, line_height, top, bottom, left, right))
    return units


def add_page_picture(doc, path: str, page_number: int, dpi: int) -> bool:
    data = pdf_tools.render_page_image(path, page_number, dpi)
    if not data:
        return False
    try:
        section = doc.sections[0]
        usable = section.page_width - section.left_margin - section.right_margin
        doc.add_picture(BytesIO(data), width=usable)
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return True
    except Exception:
        log.warning("Could not place page %s image", page_number, exc_info=True)
        return False


class BatchBuilder:
    def __init__(self, settings: Settings, path: str, pages: Sequence[int],
                 ocr_document, form_document=None):
        self.s = settings
        self.path = path
        self.pages = list(pages)
        self.ocr = ocr_document
        self.form = form_document
        self.doc = Document()
        setup_document(self.doc, settings)

        self.units = [page_text_units(ocr_document, pg)
                      for pg in ocr_document.pages]
        self.heads = (running_head_texts(self.units, settings.margin_band)
                      if settings.drop_running_heads else set())
        self.image_pages = parse_page_spec(settings.image_pages)
        # One median for the whole batch: a page that is all title would
        # otherwise report no heading, since every line on it is tall.
        heights = [u.line_height for units in self.units for u in units
                   if u.line_height > 0]
        self.median_h = statistics.median(heights) if heights else 0.0

    def build(self):
        if self.s.exact_layout:
            self._match_page_size()
        total = len(self.ocr.pages)
        for i, page in enumerate(self.ocr.pages):
            page_number = self.pages[i] if i < len(self.pages) else i + 1
            self._add_page(i, page, page_number)
            if i < total - 1:
                self.doc.add_page_break()
        return self.doc

    # ----------------------------------------------------------------
    def _match_page_size(self) -> None:
        size = (pdf_tools.pdf_page_size(self.path, self.pages[0])
                if self.pages else None)
        if not size:
            return
        for section in self.doc.sections:
            section.page_width, section.page_height = Pt(size[0]), Pt(size[1])
            for edge in ("left_margin", "right_margin",
                         "top_margin", "bottom_margin"):
                setattr(section, edge, Pt(0))

    def _table_source(self, i: int, page):
        """(page, document) to read tables from, or (None, None)."""
        if not self.s.extract_tables:
            return None, None
        if self.form is not None:
            if i < len(self.form.pages):
                return self.form.pages[i], self.form
            return None, None
        return page, self.ocr

    def _add_page(self, i: int, page, page_number: int) -> None:
        s, doc = self.s, self.doc
        units = self.units[i]

        # Chosen by the user, or a page the OCR found almost nothing on.
        chars = sum(len(u.text.strip()) for u in units)
        if page_number in self.image_pages or chars < s.image_fallback_chars:
            if add_page_picture(doc, self.path, page_number, s.page_image_dpi):
                return

        table_page, table_doc = self._table_source(i, page)
        spans = (table_regions(table_page, table_doc, s)
                 if table_page is not None else [])
        page_size = (pdf_tools.pdf_page_size(self.path, page_number)
                     if s.exact_layout else None)
        rule_added = False
        z = 1

        for u in units:
            if in_spans(u.top, u.bottom, spans):
                continue  # emitted by add_tables
            if repeat_key(u.text) in self.heads:
                continue
            text = clean_text_smart(u.text) if s.cleanup else u.text
            if not s.keep_line_breaks:
                text = unwrap_lines(text)
            if not text.strip() or is_ocr_noise(text):
                continue

            if s.exact_layout and page_size:
                w, h = page_size
                add_positioned_textbox(
                    doc, s, text, u.left * w, u.top * h,
                    u.width * w, (u.bottom - u.top) * h, detect_rtl(text), z)
                z += 1
                continue

            footnote = is_footnote(u.top, u.line_height, self.median_h,
                                   s.footnote_band, s.footnote_max_ratio)
            level = 0 if footnote else heading_level(
                u.line_height, self.median_h, text, u.line_count, u.width,
                s.heading_ratio_1, s.heading_ratio_2)
            if footnote and not rule_added:
                add_footnote_rule(doc)
                rule_added = True

            paragraph = doc.add_paragraph(
                style=f"Heading {level}" if level else None)
            add_text_with_breaks(paragraph, text)
            format_paragraph(
                paragraph, s, heading=bool(level),
                placement=detect_placement(u.left, u.right),
                size_scale=s.footnote_size_scale if footnote else 1.0,
                level=level)

        if s.extract_images and not s.exact_layout:
            for data, rect in pdf_tools.page_figures(
                    self.path, page_number, s.image_max_page_ratio):
                try:
                    doc.add_picture(BytesIO(data),
                                    width=Pt(min(rect.width, 400)))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception:
                    log.warning("Could not place figure", exc_info=True)

        if table_page is not None:
            add_tables(doc, table_doc, table_page, s)


def build_batch_document(settings: Settings, path: str, pages: Sequence[int],
                         ocr_document, form_document: Optional[object] = None):
    return BatchBuilder(settings, path, pages, ocr_document,
                        form_document).build()
