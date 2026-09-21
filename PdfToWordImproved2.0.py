# === OCR GUI v4.0 — Arabic Typography Edition ===
# Part 1 / 10  –  Imports, dark theme, helpers, GUI skeleton
from google.oauth2 import service_account
import os
import re
import gc
import shutil
import logging
import platform
import threading
import multiprocessing
import statistics
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader
from concurrent.futures import ThreadPoolExecutor, as_completed
from google.api_core.client_options import ClientOptions
from google.cloud import documentai  # type: ignore

IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    import winsound  # type: ignore

# --------------------------------------------------------------------
# Configuration: read from the environment, with a .env file if present.
def _load_dotenv() -> None:
    """Minimal .env loader so no extra dependency is required."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(env_path):
        return
    with open(env_path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip("'\""))


_load_dotenv()

GCP_PROJECT_ID = os.environ.get("GCP_PROJECT_ID", "your-gcp-project-id")
GCP_LOCATION = os.environ.get("GCP_LOCATION", "eu")
GCP_PROCESSOR_ID = os.environ.get("GCP_PROCESSOR_ID", "your-ocr-processor-id")
# Optional Form Parser processor. A Document OCR processor returns no
# tables at all, so table extraction needs a Form Parser to point at.
GCP_FORM_PROCESSOR_ID = os.environ.get("GCP_FORM_PROCESSOR_ID", "")
# Optional: path to a service-account JSON key. When unset, Application
# Default Credentials are used (gcloud auth application-default login).
GCP_CREDENTIALS_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")

# --------------------------------------------------------------------
# Logging
logging.basicConfig(
    filename=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "ocr_gui.log"
    ),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# --------------------------------------------------------------------
# Cross-platform notification / file-reveal helpers
def notify_sound() -> None:
    """Play a short completion sound, silently doing nothing if unavailable."""
    try:
        if IS_WINDOWS:
            winsound.MessageBeep()
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        elif shutil.which("canberra-gtk-play"):
            subprocess.Popen(["canberra-gtk-play", "-i", "complete"])
        elif shutil.which("paplay"):
            subprocess.Popen(
                ["paplay", "/usr/share/sounds/freedesktop/stereo/complete.oga"]
            )
        else:
            print("\a", end="", flush=True)
    except Exception as e:
        logging.warning(f"Could not play notification sound: {e}")


def reveal_in_file_manager(path: str) -> None:
    """Open the folder containing `path` in the desktop file manager."""
    if IS_WINDOWS:
        subprocess.Popen(f'explorer /select,"{path}"')
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(path)])

# --------------------------------------------------------------------
# Arabic typography settings (override in .env if you like)
# Aref Ruqaa is a calligraphic Ruqaa face; it sits taller than Arial, so it
# needs looser leading to stay readable.
ARABIC_FONT = os.environ.get("ARABIC_FONT", "Aref Ruqaa")
# Ruqaa ships Latin glyphs too, but they are calligraphic; Latin lines read
# better in a normal text face.
LATIN_FONT = os.environ.get("LATIN_FONT", "Times New Roman")
ARABIC_SIZE_PT = float(os.environ.get("ARABIC_SIZE_PT", "18"))
LINE_SPACING = float(os.environ.get("LINE_SPACING", "1.6"))
PAGE_MARGIN_CM = float(os.environ.get("PAGE_MARGIN_CM", "2.5"))
MIRROR_MARGINS = os.environ.get("MIRROR_MARGINS", "1") in ("1", "true", "True")
EXTRACT_TABLES = os.environ.get("EXTRACT_TABLES", "1") in ("1", "true", "True")
ADD_TOC = os.environ.get("ADD_TOC", "1") in ("1", "true", "True")
# A detected grid this empty is almost always a misread layout, not a table.
TABLE_MIN_FILL = float(os.environ.get("TABLE_MIN_FILL", "0.45"))
TABLE_MAX_COLS = int(os.environ.get("TABLE_MAX_COLS", "10"))
TABLE_SIZE_SCALE = float(os.environ.get("TABLE_SIZE_SCALE", "0.75"))
# Running heads, folios and download watermarks repeat at the same spot on
# page after page; they belong to the scan, not to the text.
DROP_RUNNING_HEADS = os.environ.get("DROP_RUNNING_HEADS", "1") in ("1", "true", "True")
MARGIN_BAND = float(os.environ.get("MARGIN_BAND", "0.12"))
# A born-digital PDF already holds its own text and layout; OCR would throw
# that away and guess it back. Convert those directly instead.
NATIVE_CONVERT = os.environ.get("NATIVE_CONVERT", "1") in ("1", "true", "True")
NATIVE_MIN_CHARS = int(os.environ.get("NATIVE_MIN_CHARS", "150"))
ARABIC_BOLD = os.environ.get("ARABIC_BOLD", "1") not in ("0", "false", "False")
# Poetry needs the OCR line breaks; prose does not. Off by default.
KEEP_LINE_BREAKS = os.environ.get("KEEP_LINE_BREAKS", "0") in ("1", "true", "True")

TATWEEL = "\u0640"
BIDI_MARKS = "\u200e\u200f\u202a-\u202e\u2066-\u2069"
ARABIC_RANGE = "\u0600-\u06FF"

# Order of the children Word expects after w:szCs inside w:rPr.
_RPR_AFTER_SZCS = ("w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
                   "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
                   "w:eastAsianLayout", "w:specVanish", "w:oMath")
# Order of the children Word expects after w:bidi inside w:pPr.
_PPR_AFTER_BIDI = ("w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                   "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
                   "w:jc", "w:textDirection", "w:textAlignment",
                   "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
                   "w:rPr", "w:sectPr", "w:pPrChange")


# --------------------------------------------------------------------
# Arabic-aware text cleaner (regex based, offline)
def has_text_layer(file_path: str, sample: int = 6) -> bool:
    """True when the PDF carries real text rather than scanned images.

    Watermark-only files score a handful of characters per page, so the test is
    the average over a sample rather than the presence of any text at all.
    """
    if not file_path.lower().endswith(".pdf"):
        return False
    try:
        reader = PdfReader(file_path)
        pages = reader.pages[:sample]
        if not pages:
            return False
        total = sum(len((pg.extract_text() or "").strip()) for pg in pages)
        return total / len(pages) >= NATIVE_MIN_CHARS
    except Exception as e:
        logging.warning(f"Text-layer check failed: {e}")
        return False


def convert_native_pdf(file_path: str, output_path: str, pages) -> None:
    """Convert a born-digital PDF straight to Word, keeping its own layout."""
    from pdf2docx import Converter  # imported lazily; only this path needs it

    converter = Converter(file_path)
    try:
        # pdf2docx counts pages from zero.
        converter.convert(output_path, pages=[n - 1 for n in pages])
    finally:
        converter.close()


def clean_text_smart(text: str) -> str:
    """Tidy OCR output without flattening the line structure."""
    # Control characters and bidi overrides, but keep newlines and tabs.
    text = re.sub(r"[\u0000-\u0008\u000b-\u001f]", "", text)
    text = re.sub(f"[{BIDI_MARKS}]", "", text)
    # Tatweel is decorative stretching the OCR sprinkles mid-word.
    text = text.replace(TATWEEL, "")
    # Soft hyphens and scan artefacts.
    text = text.replace("\u00ad", "").replace("\u00ac", "").replace("|", "")
    # Collapse runs of spaces, never collapse a line break away.
    text = re.sub(r"[^\S\n]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    # No space before punctuation.
    text = re.sub(r"[^\S\n]+([,.:;!?\u061f\u060c\u061b])", r"\1", text)
    # Latin punctuation sitting in Arabic text becomes its Arabic form.
    text = re.sub(f"(?<=[{ARABIC_RANGE}]),", "\u060c", text)
    text = re.sub(f"(?<=[{ARABIC_RANGE}]);", "\u061b", text)
    text = re.sub(f"(?<=[{ARABIC_RANGE}])\\?", "\u061f", text)
    return text.strip()


# --------------------------------------------------------------------
# Complex-script font plumbing
def detect_rtl(text: str) -> bool:
    """Decide a paragraph's direction from its own letters.

    A copyright page mixes Arabic blocks with Latin ones; forcing RTL onto the
    Latin lines is what turns "E-mail:" into ":E-mail" and reverses phone
    numbers, so each paragraph is judged separately.
    """
    arabic = len(re.findall(f"[{ARABIC_RANGE}]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    return arabic >= latin


def apply_complex_script(rPr, font_name: str, size_pt: float, bold: bool,
                         rtl: bool = True) -> None:
    """Set the complex-script run properties.

    python-docx's `font.name` only writes w:ascii and w:hAnsi, which Word uses
    for Latin text. Arabic is complex script and is rendered from w:cs, w:szCs
    and w:bCs, so without these the font, size and weight never reach it.
    """
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)

    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.insert_element_before(szCs, *_RPR_AFTER_SZCS)
    szCs.set(qn("w:val"), str(int(round(size_pt * 2))))  # half-points

    rPr.get_or_add_bCs().set(qn("w:val"), "1" if bold else "0")
    rPr.get_or_add_rtl().set(qn("w:val"), "1" if rtl else "0")


# --------------------------------------------------------------------
# Arabic paragraph styling
def apply_arabic_paragraph_format(paragraph, heading: bool = False,
                                  placement: str = "full",
                                  size_scale: float = 1.0):
    try:
        text = paragraph.text
        rtl = detect_rtl(text)

        pPr = paragraph._element.get_or_add_pPr()
        bidi = pPr.find(qn("w:bidi"))
        if bidi is None:
            bidi = OxmlElement("w:bidi")
            pPr.insert_element_before(bidi, *_PPR_AFTER_BIDI)
        bidi.set(qn("w:val"), "1" if rtl else "0")

        size = (ARABIC_SIZE_PT * 1.15 if heading else ARABIC_SIZE_PT) * size_scale
        bold = True if heading else ARABIC_BOLD
        # Justify only real running text; justifying a short line stretches the
        # words across the page.
        if placement == "center":
            side = "center"
        elif placement in ("left", "right"):
            side = placement
        elif len(text) > 120:
            side = "justify"          # running text keeps a clean measure
        else:
            side = "start"            # short line: natural side for its script
        set_physical_alignment(paragraph, side, rtl)
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(12 if heading else 0)
        fmt.space_after = Pt(6)
        fmt.line_spacing = LINE_SPACING
        fmt.keep_with_next = heading

        font_name = ARABIC_FONT if rtl else LATIN_FONT
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.bold = bold
            apply_complex_script(run._element.get_or_add_rPr(),
                                 font_name, size, bold, rtl=rtl)
    except Exception as e:
        logging.warning(f"Could not apply Arabic paragraph format: {e}")


# --------------------------------------------------------------------
# Document-wide setup: RTL section, Arabic Normal style, page-number footer
def set_section_rtl(section) -> None:
    sectPr = section._sectPr
    bidi = sectPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        sectPr.append(bidi)
    bidi.set(qn("w:val"), "1")


def add_page_number_footer(section) -> None:
    """A real Word footer field, so page numbers are not typed into the body."""
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instr, end):
        run._element.append(node)
    run.font.name = ARABIC_FONT
    run.font.size = Pt(11)
    run.bold = False


def style_heading_levels(doc) -> None:
    """Make Word's built-in Heading styles Arabic-aware.

    Using the real styles (rather than ad-hoc bold text) is what gives you the
    Navigation pane, collapsible sections and a working table of contents.
    """
    for level, scale in ((1, 1.6), (2, 1.35), (3, 1.15)):
        try:
            style = doc.styles[f"Heading {level}"]
        except KeyError:
            continue
        size = ARABIC_SIZE_PT * scale
        style.font.name = ARABIC_FONT
        style.font.size = Pt(size)
        style.font.bold = True
        # Word's Heading styles are blue by default; a scanned book is black.
        style.font.color.rgb = RGBColor(0, 0, 0)
        apply_complex_script(style.element.get_or_add_rPr(),
                             ARABIC_FONT, size, True)
        fmt = style.paragraph_format
        fmt.space_before = Pt(18)
        fmt.space_after = Pt(8)
        fmt.keep_with_next = True


def setup_page(section) -> None:
    """Book-style page geometry using Word's own page setup."""
    margin = Cm(PAGE_MARGIN_CM)
    section.top_margin = margin
    section.bottom_margin = margin
    section.left_margin = margin
    section.right_margin = margin
    if MIRROR_MARGINS:
        # Mirrored margins keep the inner (binding) edge consistent when the
        # book is printed double-sided.
        sectPr = section._sectPr
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.set(qn("w:gutter"), str(int(Cm(0.6).twips)))
        settings = section.part.document.settings.element
        for tag in ("w:mirrorMargins", "w:evenAndOddHeaders"):
            if settings.find(qn(tag)) is None:
                settings.append(OxmlElement(tag))


def add_table_of_contents(doc) -> None:
    """Insert a Word TOC field; Word fills it from the Heading styles."""
    paragraph = doc.add_paragraph()
    apply_arabic_paragraph_format(paragraph, heading=True)
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-3" \h \z \u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "اضغط F9 لتحديث الفهرس"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instr, separate, placeholder, end):
        run._element.append(node)
    doc.add_page_break()


def set_table_rtl(table) -> None:
    """Right-to-left column order, so column 1 is the rightmost."""
    tblPr = table._tbl.tblPr
    if tblPr.find(qn("w:bidiVisual")) is None:
        tblPr.append(OxmlElement("w:bidiVisual"))


def cell_text(document, cell) -> str:
    return "".join(
        document.text[int(seg.start_index):int(seg.end_index)]
        for seg in cell.layout.text_anchor.text_segments
    )


def fetch_tables(client, project_id, location, file_content, mime_type, pages):
    """Run the Form Parser over the same pages, purely to recover tables."""
    if not GCP_FORM_PROCESSOR_ID:
        return None
    try:
        name = client.processor_path(project_id, location, GCP_FORM_PROCESSOR_ID)
        request = documentai.ProcessRequest(
            name=name,
            raw_document=documentai.RawDocument(
                content=file_content, mime_type=mime_type
            ),
            process_options=documentai.ProcessOptions(
                individual_page_selector=(
                    documentai.ProcessOptions.IndividualPageSelector(pages=pages)
                )
            ),
        )
        return client.process_document(request=request).document
    except Exception as e:
        logging.warning(f"Form Parser call failed, tables skipped: {e}")
        return None


def is_real_table(document, table) -> bool:
    """Reject grids the parser invented.

    A list laid out in two columns is often returned as a wide, mostly empty
    grid; rendering that as a table is worse than leaving the text as prose.
    """
    rows = list(table.header_rows) + list(table.body_rows)
    if not rows:
        return False
    width = max(len(r.cells) for r in rows)
    if width > TABLE_MAX_COLS:
        return False
    total = sum(len(r.cells) for r in rows)
    filled = sum(
        1 for r in rows for c in r.cells if cell_text(document, c).strip()
    )
    return total > 0 and filled / total >= TABLE_MIN_FILL


def add_tables(doc, document, page, cleanup_enabled: bool) -> None:
    """Rebuild Document AI tables as real Word tables."""
    for table in page.tables:
        if not is_real_table(document, table):
            continue
        rows = list(table.header_rows) + list(table.body_rows)
        if not rows:
            continue
        width = max(len(r.cells) for r in rows)
        word_table = doc.add_table(rows=len(rows), cols=width)
        word_table.style = "Table Grid"
        word_table.autofit = True
        set_table_rtl(word_table)
        for r, row in enumerate(rows):
            for c, cell in enumerate(row.cells):
                if c >= width:
                    continue
                text = cell_text(document, cell)
                if cleanup_enabled:
                    text = clean_text_smart(text)
                text = unwrap_lines(text)
                target = word_table.cell(r, c).paragraphs[0]
                if text.strip():
                    target.add_run(text.strip())
                apply_arabic_paragraph_format(
                    target, heading=(r < len(table.header_rows)),
                    size_scale=TABLE_SIZE_SCALE,
                )
        doc.add_paragraph()


def table_regions(page, document=None):
    """Vertical spans covered by tables, so their text is not emitted twice.

    Only tables that survive the quality gate suppress text; a rejected grid
    must leave its text to flow as normal paragraphs.
    """
    spans = []
    for table in page.tables:
        if document is not None and not is_real_table(document, table):
            continue
        verts = table.layout.bounding_poly.normalized_vertices
        if verts:
            spans.append((min(v.y for v in verts), max(v.y for v in verts)))
    return spans


def setup_document(doc) -> None:
    try:
        style = doc.styles["Normal"]
        style.font.name = ARABIC_FONT
        style.font.size = Pt(ARABIC_SIZE_PT)
        style.font.bold = ARABIC_BOLD
        apply_complex_script(style.element.get_or_add_rPr(),
                             ARABIC_FONT, ARABIC_SIZE_PT, ARABIC_BOLD)
    except Exception as e:
        logging.warning(f"Normal style not set: {e}")
    try:
        doc.styles["Normal"].paragraph_format.line_spacing = LINE_SPACING
    except Exception as e:
        logging.warning(f"Normal spacing not set: {e}")
    style_heading_levels(doc)
    for section in doc.sections:
        try:
            set_section_rtl(section)
            setup_page(section)
            add_page_number_footer(section)
        except Exception as e:
            logging.warning(f"Section setup failed: {e}")


# --------------------------------------------------------------------
# Turning Document AI output into paragraphs
def page_text_units(document, page):
    """Yield (text, relative_height) per paragraph.

    Document AI paragraphs track the printed paragraphs far more closely than
    blocks, which tend to fuse columns together.
    """
    source = list(page.paragraphs) or list(page.blocks)
    units = []
    for unit in source:
        layout = unit.layout
        text = "".join(
            document.text[int(seg.start_index):int(seg.end_index)]
            for seg in layout.text_anchor.text_segments
        )
        verts = layout.bounding_poly.normalized_vertices
        height = (max(v.y for v in verts) - min(v.y for v in verts)) if verts else 0.0
        top = min((v.y for v in verts), default=0.0)
        bottom = max((v.y for v in verts), default=0.0)
        left = min((v.x for v in verts), default=0.0)
        right = max((v.x for v in verts), default=1.0)
        units.append((text, height, top, bottom, left, right))
    return units


def unwrap_lines(text: str) -> str:
    """Join OCR line-wraps inside one paragraph.

    Document AI breaks a paragraph at every printed line. Keeping those as hard
    breaks makes Word justify each fragment separately, which is what spreads
    the words across the page.
    """
    text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    return re.sub(r"[^\S\n]{2,}", " ", text)


def is_ocr_noise(text: str) -> bool:
    """A lone stray character is a scan artefact, not content."""
    stripped = text.strip()
    return len(stripped) <= 1 and not stripped.isdigit()


def running_head_texts(pages_units) -> set:
    """Text that repeats in the top or bottom band across most pages."""
    if len(pages_units) < 3:
        return set()
    seen = {}
    for units in pages_units:
        page_hits = set()
        for text, _h, top, bottom, _l, _r in units:
            in_band = top <= MARGIN_BAND or bottom >= 1.0 - MARGIN_BAND
            key = re.sub(r"\d+", "#", " ".join(text.split()))[:60]
            if in_band and key:
                page_hits.add(key)
        for key in page_hits:
            seen[key] = seen.get(key, 0) + 1
    threshold = max(2, int(len(pages_units) * 0.6))
    return {k for k, n in seen.items() if n >= threshold}


def detect_placement(left: float, right: float) -> str:
    """Recover how the line sat on the printed page.

    Margins are measured on both sides of the block: roughly equal gaps mean it
    was centred, a clearly smaller gap on one side means it was pushed to that
    side, and a block spanning the measure is running text.
    """
    left_gap = max(left, 0.0)
    right_gap = max(1.0 - right, 0.0)
    width = right - left
    if width >= 0.72:
        return "full"
    if left_gap > 0.12 and right_gap > 0.12 and abs(left_gap - right_gap) < 0.09:
        return "center"
    if left_gap + 0.10 < right_gap:
        return "left"
    if right_gap + 0.10 < left_gap:
        return "right"
    return "full"


def set_physical_alignment(paragraph, side: str, rtl: bool) -> None:
    """Align to a physical side of the page.

    w:jc is start/end relative, so inside a bidi paragraph "left" lands on the
    physical right. This maps the side you can see to the value Word wants.
    """
    if side == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif side == "justify":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif side == "start":
        paragraph.alignment = None
    elif rtl:
        paragraph.alignment = (WD_ALIGN_PARAGRAPH.LEFT if side == "right"
                               else WD_ALIGN_PARAGRAPH.RIGHT)
    else:
        paragraph.alignment = (WD_ALIGN_PARAGRAPH.LEFT if side == "left"
                               else WD_ALIGN_PARAGRAPH.RIGHT)


def add_text_with_breaks(paragraph, text: str) -> None:
    """Keep OCR line breaks as real line breaks inside a single paragraph."""
    lines = text.split("\n")
    run = paragraph.add_run(lines[0])
    for line in lines[1:]:
        run.add_break()
        run.add_text(line)

# --------------------------------------------------------------------
# Page-range parser
def parse_page_range(page_input: str, total_pages: int) -> List[int]:
    page_input = page_input.strip().lower()
    if page_input == "all":
        return list(range(1, total_pages + 1))
    pages = set()
    for part in page_input.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-")
            pages.update(range(int(start), int(end) + 1))
        else:
            pages.add(int(part))
    return sorted(p for p in pages if 1 <= p <= total_pages)

# --------------------------------------------------------------------
# Jet-black + neon-green theme
def apply_dark_theme(root):
    root.configure(bg="#000000")
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "TButton",
        background="#00FF00",
        foreground="#000000",
        font=("Consolas", 10, "bold"),
        padding=5,
    )
    style.map("TButton",
              foreground=[("active", "#000000")],
              background=[("active", "#00DD00")])
    style.configure("TProgressbar", troughcolor="#111111", background="#00FF00")

# --------------------------------------------------------------------
# Main GUI class header
class OCRApp:
    def __init__(self, root):
        apply_dark_theme(root)
        self.root = root
        self.root.title("Enterprise OCR v4.0 — Arabic Typography Edition")
        self.root.geometry("820x640")
        self.root.resizable(False, False)

        if IS_WINDOWS:
            try:
                import ctypes
                ctypes.windll.user32.ShowWindow(
                    ctypes.windll.kernel32.GetConsoleWindow(), 0
                )
            except Exception:
                pass

        # Cloud configuration (see .env / .env.example)
        self.project_id = GCP_PROJECT_ID
        self.location = GCP_LOCATION
        self.processor_id = GCP_PROCESSOR_ID

        self.cancel_requested = False

        # --- GUI layout ---
        tk.Label(root, text="Input File:", fg="#00FF00", bg="#000000",
                 font=("Consolas", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        self.file_entry = tk.Entry(root, width=80, bg="#111111", fg="#00FF00",
                                   insertbackground="#00FF00")
        self.file_entry.pack(padx=10, pady=2)
        ttk.Button(root, text="Browse", command=self.browse_file).pack(pady=2)

        tk.Label(root, text="Output DOCX File:", fg="#00FF00", bg="#000000",
                 font=("Consolas", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        self.output_entry = tk.Entry(root, width=80, bg="#111111", fg="#00FF00",
                                     insertbackground="#00FF00")
        self.output_entry.pack(padx=10, pady=2)

        tk.Label(root, text="Page Range (e.g., all, 1-5, 2,5,7):",
                 fg="#00FF00", bg="#000000", font=("Consolas", 10, "bold"))\
            .pack(anchor="w", padx=10, pady=5)
        self.page_entry = tk.Entry(root, width=40, bg="#111111",
                                   fg="#00FF00", insertbackground="#00FF00")
        self.page_entry.insert(0, "all")
        self.page_entry.pack(padx=10, pady=2)

        self.quality_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Enable Quality Scoring",
                       variable=self.quality_var, bg="#000000",
                       fg="#00FF00", selectcolor="#111111",
                       font=("Consolas", 10)).pack(anchor="w", padx=20, pady=5)

        self.cleanup_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Enable Smart Text Cleanup",
                       variable=self.cleanup_var, bg="#000000",
                       fg="#00FF00", selectcolor="#111111",
                       font=("Consolas", 10)).pack(anchor="w", padx=20, pady=2)

        # Buttons
        btn_frame = tk.Frame(root, bg="#000000")
        btn_frame.pack(pady=10)
        self.run_btn = ttk.Button(btn_frame, text="Run OCR", command=self.run_ocr)
        self.run_btn.grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel_ocr).grid(row=0, column=1, padx=5)

        # Status + progress bar + log console
        self.status_label = tk.Label(root, text="Ready.", anchor="w",
                                     fg="#00FF00", bg="#000000",
                                     font=("Consolas", 10, "bold"))
        self.status_label.pack(fill="x", padx=10, pady=8)
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(root, variable=self.progress_var, maximum=100)
        self.progress.pack(fill="x", padx=15, pady=10)
        self.log_box = tk.Text(root, height=8, bg="#000000", fg="#00FF00",
                               font=("Consolas", 9), state="disabled")
        self.log_box.pack(fill="x", padx=10, pady=10)

    # Simple logger for the GUI
    def gui_log(self, msg):
        self.root.after(0, lambda: self._append_log(msg))

    def _append_log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        total_lines = int(self.log_box.index('end-1c').split('.')[0])
        if total_lines > 200:
            self.log_box.delete("1.0", f"{total_lines - 180}.0")
        self.log_box.see("end")
        self.log_box.config(state="disabled")


    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF or Image", "*.pdf *.jpg *.jpeg")])
        if filename:
            # 1. Update the input field with the chosen file
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)
            
            # 2. Extract the file's directory and its name
            input_dir = os.path.dirname(filename)  # Gets the folder containing the file
            base_name = os.path.splitext(os.path.basename(filename))[0]
            
            # 3. Build the path using the same folder as the input
            output_path = os.path.join(input_dir, f"{base_name}.docx")
            
            # 4. Update the GUI output field
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_path)
            
            # Log the destination for the user
            self.gui_log(f"Output path set to: {input_dir}")
    def run_ocr(self):
        """Start OCR process when 'Run OCR' is clicked."""
        if self.cancel_requested:
            self.cancel_requested = False  # reset cancel flag

        file_path = self.file_entry.get().strip()
        output_path = self.output_entry.get().strip()
        page_input = self.page_entry.get().strip()
        enable_quality = self.quality_var.get()
        cleanup = self.cleanup_var.get()

        if not file_path or not output_path:
            messagebox.showerror("Error", "Please select input and output files.")
            return

        try:
            # Determine MIME type
            mime_type = "application/pdf" if file_path.lower().endswith(".pdf") else "image/jpeg"
            total_pages = len(PdfReader(file_path).pages) if mime_type == "application/pdf" else 1
            selected_pages = parse_page_range(page_input, total_pages)

            if not selected_pages:
                messagebox.showwarning("Page Range", "No valid pages selected.")
                return

            self.gui_log(f"Starting OCR for pages: {selected_pages}")
            self.safe_status_update("Starting OCR process...")
            self.safe_progress_update(0)
            self.run_btn.config(state="disabled")

            # Start threaded OCR
            t = threading.Thread(
                target=self.threaded_ocr,
                args=(
                    self.project_id,
                    self.location,
                    self.processor_id,
                    file_path,
                    mime_type,
                    output_path,
                    selected_pages,
                    enable_quality,
                    cleanup,
                ),
                daemon=True,
            )
            t.start()

        except Exception as e:
            logging.error(f"Error during OCR start: {e}")
            messagebox.showerror("Error", str(e))
            
    
    # --- Cancel OCR ---
    def cancel_ocr(self):
        self.cancel_requested = True
        self.safe_status_update("Canceling process...")
        self.gui_log("User requested cancel.")
    
    def safe_status_update(self, msg):
        self.root.after(0, lambda: self.status_label.config(text=msg))
        
    def safe_progress_update(self, value):
        self.root.after(0, lambda: self.progress_var.set(value))

    # Tk is not thread-safe: dialogs raised from worker threads must be
    # marshalled back onto the main loop.
    def safe_info(self, title, msg):
        self.root.after(0, lambda: messagebox.showinfo(title, msg))

    def safe_warning(self, title, msg):
        self.root.after(0, lambda: messagebox.showwarning(title, msg))

    def safe_error(self, title, msg):
        self.root.after(0, lambda: messagebox.showerror(title, msg))
        
    
    def threaded_ocr(self, project_id, location, processor_id,
                     file_path, mime_type, output_path, selected_pages,
                     enable_quality, cleanup):
        """Run OCR in background threads, update progress bar, and handle resume."""
        try:
            if file_path.lower().endswith(".pdf"):
                total_pages = len(PdfReader(file_path).pages)
            else:
                total_pages = 1

            
            output_dir = os.path.dirname(os.path.abspath(output_path))
            os.makedirs(output_dir, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(output_path))[0]

            if NATIVE_CONVERT and has_text_layer(file_path):
                self.gui_log("PDF has a text layer - converting directly, "
                             "no OCR needed.")
                self.safe_status_update("Converting native PDF...")
                convert_native_pdf(file_path, output_path, selected_pages)
                self.safe_progress_update(100)
                self.safe_status_update("Done.")
                notify_sound()
                self.safe_info("Success",
                               f"Converted without OCR:\n{output_path}")
                try:
                    reveal_in_file_manager(output_path)
                except Exception as e:
                    self.gui_log(f"Auto-open failed: {e}")
                self.root.after(3000, lambda: self.safe_status_update("Ready."))
                return
            batch_size = 15
            total_parts = (len(selected_pages) + batch_size - 1) // batch_size
            completed = 0

            # Smart resume detection
            existing_parts = [f for f in os.listdir(output_dir)
                              if f.startswith(base_name + "_part_") and f.endswith(".docx")]
            start_index = len(existing_parts) * batch_size
            if existing_parts:
                self.gui_log(f"Resuming from part {len(existing_parts)+1}")
            else:
                self.gui_log("Starting new OCR job...")

            # Auto thread scaling
            max_workers = max(2, min(6, multiprocessing.cpu_count() - 1))
            self.gui_log(f"Using {max_workers} threads")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = []
                for i in range(start_index, len(selected_pages), batch_size):
                    if self.cancel_requested:
                        break
                    part_num = (i // batch_size) + 1
                    batch_pages = selected_pages[i:i + batch_size]
                    part_output = os.path.join(output_dir, f"{base_name}_part_{part_num}.docx")
                    futures.append(executor.submit(
                        self.process_part,
                        project_id, location, processor_id,
                        file_path, mime_type, batch_pages,
                        enable_quality, part_output, part_num, cleanup
                    ))

                for future in as_completed(futures):
                    if self.cancel_requested:
                        break
                    future.result()
                    completed += 1
                    progress = (completed / total_parts) * 100
                    self.safe_progress_update(progress)
                    self.safe_status_update(f"Processed {completed}/{total_parts} parts...")
                    self.gui_log(f"Completed part {completed}/{total_parts}")

            if not self.cancel_requested:
                self.merge_and_finish(output_dir, base_name)
            else:
                self.safe_status_update("OCR canceled by user.")
                self.gui_log("Processing canceled.")
        except Exception as e:
            logging.error(f"OCR error: {e}")
            self.safe_error("Error", str(e))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            
     
    def process_part(self, project_id, location, processor_id,
                     file_path, mime_type, pages, enable_quality,
                     part_output, part_num, cleanup):
        """Handle one OCR batch (part)."""
        if self.cancel_requested:
            return
        self.safe_status_update(f"Processing Part {part_num} (Pages {pages[0]}–{pages[-1]})...")
        self.gui_log(f"Starting part {part_num} for pages {pages}")
        doc = extract_batch(project_id, location, processor_id, file_path,
                            mime_type, pages, enable_quality, cleanup)
        doc.save(part_output)
        self.gui_log(f"Saved {part_output}")
        del doc
        gc.collect()      
        
        
        
    def merge_and_finish(self, output_dir, base_name):
        """Merge all temporary part files, delete them, and open the merged document."""
        try:
            merged_path = os.path.join(output_dir, f"{base_name}_merged.docx")
            merged_doc = Document()
            setup_document(merged_doc)
            if ADD_TOC:
                add_table_of_contents(merged_doc)
            merged_body = merged_doc.element.body
            trailing_sectPr = merged_body.find(qn("w:sectPr"))
            part_files = []

            # Collect parts and order them numerically. Lexical sorting would
            # give part_1, part_10, part_2, ... and scramble long documents.
            part_pattern = re.compile(
                rf"^{re.escape(base_name)}_part_(\d+)\.docx$"
            )
            numbered_parts = []
            for filename in os.listdir(output_dir):
                match = part_pattern.match(filename)
                if match:
                    numbered_parts.append((int(match.group(1)), filename))
            numbered_parts.sort()

            for _, filename in numbered_parts:
                part_path = os.path.join(output_dir, filename)
                part_files.append(part_path)
                self.gui_log(f"Merging {filename}")
                try:
                    part_doc = Document(part_path)
                    for element in part_doc.element.body:
                        # Each part carries its own sectPr; copying them would
                        # add a section break between every part.
                        if element.tag == qn("w:sectPr"):
                            continue
                        if trailing_sectPr is not None:
                            trailing_sectPr.addprevious(element)
                        else:
                            merged_body.append(element)
                except Exception as e:
                    self.gui_log(f"Error reading {filename}: {e}")

            if not part_files:
                self.safe_warning("Merge", "No part files found to merge.")
                return

            merged_doc.save(merged_path)
            self.gui_log(f"Merged file saved: {merged_path}")

            # Delete temp part files
            for f in part_files:
                try:
                    os.remove(f)
                    self.gui_log(f"Deleted temporary file: {os.path.basename(f)}")
                except Exception as e:
                    self.gui_log(f"Could not delete {os.path.basename(f)}: {e}")

            self.safe_progress_update(100)
            self.safe_status_update("All parts merged successfully!")
            notify_sound()
            self.safe_info("Success", f"OCR completed!\nFinal file:\n{merged_path}")
            self.root.after(3000, lambda: self.safe_status_update("Ready."))
            self.root.after(3000, lambda: self.run_btn.config(state="normal"))

            # Reveal the finished file in the desktop file manager
            try:
                reveal_in_file_manager(merged_path)
                self.gui_log("Opened final file in the file manager.")
            except Exception as e:
                self.gui_log(f"Auto-open failed: {e}")

        except Exception as e:
            logging.error(f"Merge error: {e}")
            self.safe_error("Error", f"Error during merging: {e}")  
    
    # === Part 2 / 10 — OCR extraction logic ===

def extract_batch(project_id, location, processor_id,
                  file_path, mime_type, pages,
                  enable_quality, cleanup_enabled=True) -> Document:
    """Run Google Document AI OCR on a batch of pages and return a Word Document."""
    opts = ClientOptions(api_endpoint=f"{location}-documentai.googleapis.com")

    # Prefer an explicit service-account key when one is configured, otherwise
    # fall back to Application Default Credentials.
    if GCP_CREDENTIALS_FILE:
        if not os.path.isfile(GCP_CREDENTIALS_FILE):
            raise FileNotFoundError(
                f"Service-account key not found: {GCP_CREDENTIALS_FILE}\n"
                "Fix GOOGLE_APPLICATION_CREDENTIALS in .env, or unset it to use "
                "'gcloud auth application-default login'."
            )
        creds = service_account.Credentials.from_service_account_file(
            GCP_CREDENTIALS_FILE
        )
        client = documentai.DocumentProcessorServiceClient(
            credentials=creds, client_options=opts
        )
    else:
        client = documentai.DocumentProcessorServiceClient(client_options=opts)

    name = client.processor_path(project_id, location, processor_id)


    with open(file_path, "rb") as f:
        file_content = f.read()

    # Configure OCR
    ocr_config = documentai.OcrConfig(
        enable_native_pdf_parsing=False,
        enable_image_quality_scores=enable_quality,
        enable_symbol=True,
        hints=documentai.OcrConfig.Hints(language_hints=["ar", "en"]),
    )
    process_options = documentai.ProcessOptions(
        ocr_config=ocr_config,
        individual_page_selector=documentai.ProcessOptions.IndividualPageSelector(pages=pages),
    )
    request = documentai.ProcessRequest(
        name=name,
        raw_document=documentai.RawDocument(content=file_content, mime_type=mime_type),
        process_options=process_options,
    )

    result = client.process_document(request=request)
    document = result.document
    form_document = (
        fetch_tables(client, project_id, location, file_content, mime_type, pages)
        if EXTRACT_TABLES else None
    )
    doc = Document()
    setup_document(doc)

    all_units = [page_text_units(document, pg) for pg in document.pages]
    heads = running_head_texts(all_units) if DROP_RUNNING_HEADS else set()

    for i, page in enumerate(document.pages, start=1):
        units = all_units[i - 1]
        heights = [u[1] for u in units if u[1] > 0]
        median_h = statistics.median(heights) if heights else 0.0
        table_page = None
        if EXTRACT_TABLES and form_document is not None:
            if i - 1 < len(form_document.pages):
                table_page = form_document.pages[i - 1]
        elif EXTRACT_TABLES:
            table_page = page
        if table_page is not None:
            table_source = form_document if form_document is not None else document
            spans = table_regions(table_page, table_source)
        else:
            spans = []

        for raw_text, height, top, bottom, left, right in units:
            # Table cell text is emitted by add_tables; skip it here.
            if any(top >= s - 0.01 and bottom <= e + 0.01 for s, e in spans):
                continue
            key = re.sub(r"\d+", "#", " ".join(raw_text.split()))[:60]
            if key in heads:
                continue
            refined = clean_text_smart(raw_text) if cleanup_enabled else raw_text
            if not KEEP_LINE_BREAKS:
                refined = unwrap_lines(refined)
            if not refined.strip() or is_ocr_noise(refined):
                continue
            # A short, unusually tall line is almost always a heading.
            is_heading = (
                median_h > 0
                and height > median_h * 1.25
                and len(refined) <= 90
                and "\n" not in refined
            )
            placement = detect_placement(left, right)
            style = "Heading 2" if is_heading else None
            paragraph = doc.add_paragraph(style=style)
            add_text_with_breaks(paragraph, refined)
            apply_arabic_paragraph_format(paragraph, heading=is_heading,
                                          placement=placement)

        if table_page is not None:
            add_tables(doc, table_source, table_page, cleanup_enabled)

        if i < len(document.pages):
            doc.add_page_break()

    return doc


# === Part 3 / 10 — File browsing, cancel logic, and small helpers ===

def safe_thread(func):
    """Decorator to run a function in a background thread."""
    def wrapper(self, *args, **kwargs):
        t = threading.Thread(target=func, args=(self, *args), kwargs=kwargs, daemon=True)
        t.start()
    return wrapper



    

    
    # --- Safe GUI update helpers ---
    

    
        # === Part 4 / 10 — Threaded OCR controller and batch handling ===

    


    


    # --- GUI log wrapper (safe thread) ---
    
        
        
    # === Part 5 / 10 — Merge .docx parts and auto-open final file ===

    



    # === Part 6 / 10 — Run OCR button logic ===

    








# === Part 7 / 10 — Application entry point ===
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = OCRApp(root)
        # bring window to front
        root.after(200, lambda: root.focus_force())
        root.mainloop()
    except Exception as e:
        logging.error(f"Fatal error: {e}")
        messagebox.showerror("Startup Error", f"Fatal error: {e}")


