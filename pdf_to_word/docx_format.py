"""Arabic-aware Word formatting on top of python-docx."""
from __future__ import annotations

import logging
from xml.sax.saxutils import escape, quoteattr

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .config import Settings
from .text import detect_rtl

log = logging.getLogger(__name__)

HEADING_SCALES = {1: 1.6, 2: 1.35, 3: 1.15}

# Order of the children Word expects after w:szCs inside w:rPr.
_RPR_AFTER_SZCS = ("w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
                   "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em",
                   "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath")
# Order of the children Word expects after w:bidi inside w:pPr.
_PPR_AFTER_BIDI = ("w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                   "w:contextualSpacing", "w:mirrorIndents",
                   "w:suppressOverlap", "w:jc", "w:textDirection",
                   "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                   "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                   "w:pPrChange")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V_NS = "urn:schemas-microsoft-com:vml"


def apply_complex_script(rPr, font_name: str, size_pt: float, bold: bool,
                         rtl: bool = True) -> None:
    """Set complex-script run properties.

    python-docx's ``font.name`` only writes w:ascii and w:hAnsi. Arabic is
    rendered from w:cs, w:szCs and w:bCs, so without these the font, size
    and weight never reach it.
    """
    rFonts = rPr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), font_name)
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.insert_element_before(szCs, *_RPR_AFTER_SZCS)
    szCs.set(qn("w:val"), str(int(round(size_pt * 2))))  # half-points
    rPr.get_or_add_bCs().set(qn("w:val"), "1" if bold else "0")
    rPr.get_or_add_rtl().set(qn("w:val"), "1" if rtl else "0")


def set_physical_alignment(paragraph, side: str, rtl: bool) -> None:
    """Align to a physical side of the page.

    w:jc is start/end relative, so inside a bidi paragraph "left" lands on
    the physical right. This maps the visible side to what Word wants.
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


def format_paragraph(paragraph, settings: Settings, heading: bool = False,
                     placement: str = "full", size_scale: float = 1.0,
                     level: int = 0) -> None:
    """Direction, alignment, spacing and fonts for one paragraph."""
    try:
        text = paragraph.text
        rtl = detect_rtl(text)

        pPr = paragraph._element.get_or_add_pPr()
        bidi = pPr.find(qn("w:bidi"))
        if bidi is None:
            bidi = OxmlElement("w:bidi")
            pPr.insert_element_before(bidi, *_PPR_AFTER_BIDI)
        bidi.set(qn("w:val"), "1" if rtl else "0")

        # Direct run formatting overrides the style, so a heading carries its
        # own level's size or every level would render identically.
        scale = HEADING_SCALES.get(level, 1.15) if heading else 1.0
        size = settings.arabic_size_pt * scale * size_scale
        bold = True if heading else settings.arabic_bold
        # Justify only real running text; a short line would be stretched.
        if placement in ("center", "left", "right"):
            side = placement
        elif len(text) > 120:
            side = "justify"
        else:
            side = "start"
        set_physical_alignment(paragraph, side, rtl)
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(12 if heading else 0)
        fmt.space_after = Pt(6)
        fmt.line_spacing = settings.line_spacing
        fmt.keep_with_next = heading

        font_name = settings.arabic_font if rtl else settings.latin_font
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.bold = bold
            apply_complex_script(run._element.get_or_add_rPr(),
                                 font_name, size, bold, rtl=rtl)
    except Exception:
        log.warning("Could not apply Arabic paragraph format", exc_info=True)


def add_text_with_breaks(paragraph, text: str) -> None:
    """Keep OCR line breaks as real line breaks inside one paragraph."""
    lines = text.split("\n")
    run = paragraph.add_run(lines[0])
    for line in lines[1:]:
        run.add_break()
        run.add_text(line)


def _field_run(paragraph, instruction: str, placeholder: str = "",
               dirty: bool = False):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    if dirty:
        begin.set(qn("w:dirty"), "true")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    nodes = [begin, instr]
    if placeholder:
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        text = OxmlElement("w:t")
        text.text = placeholder
        nodes += [separate, text]
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in nodes + [end]:
        run._element.append(node)
    return run


def set_section_rtl(section) -> None:
    sectPr = section._sectPr
    bidi = sectPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        sectPr.append(bidi)
    bidi.set(qn("w:val"), "1")


def add_page_number_footer(section, settings: Settings) -> None:
    """A real Word PAGE field, so numbers are not typed into the body."""
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraph = (footer.paragraphs[0] if footer.paragraphs
                 else footer.add_paragraph())
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = _field_run(paragraph, " PAGE ")
    run.font.name = settings.arabic_font
    run.font.size = Pt(11)
    run.bold = False


def style_heading_levels(doc, settings: Settings) -> None:
    """Make Word's built-in Heading styles Arabic-aware (navigation, TOC)."""
    for level, scale in sorted(HEADING_SCALES.items()):
        try:
            style = doc.styles[f"Heading {level}"]
        except KeyError:
            continue
        size = settings.arabic_size_pt * scale
        style.font.name = settings.arabic_font
        style.font.size = Pt(size)
        style.font.bold = True
        # Word's Heading styles are blue by default; a scanned book is black.
        style.font.color.rgb = RGBColor(0, 0, 0)
        apply_complex_script(style.element.get_or_add_rPr(),
                             settings.arabic_font, size, True)
        fmt = style.paragraph_format
        fmt.space_before = Pt(18)
        fmt.space_after = Pt(8)
        fmt.keep_with_next = True


def setup_page(section, settings: Settings) -> None:
    """Book-style page geometry."""
    margin = Cm(settings.page_margin_cm)
    section.top_margin = section.bottom_margin = margin
    section.left_margin = section.right_margin = margin
    if settings.mirror_margins:
        # Mirrored margins keep the binding edge consistent double-sided.
        pgMar = section._sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.set(qn("w:gutter"), str(int(Cm(0.6).twips)))
        doc_settings = section.part.document.settings.element
        for tag in ("w:mirrorMargins", "w:evenAndOddHeaders"):
            if doc_settings.find(qn(tag)) is None:
                doc_settings.append(OxmlElement(tag))


def setup_document(doc, settings: Settings) -> None:
    """Normal/Heading styles, RTL sections, margins and page numbers."""
    try:
        style = doc.styles["Normal"]
        style.font.name = settings.arabic_font
        style.font.size = Pt(settings.arabic_size_pt)
        style.font.bold = settings.arabic_bold
        apply_complex_script(style.element.get_or_add_rPr(),
                             settings.arabic_font, settings.arabic_size_pt,
                             settings.arabic_bold)
        style.paragraph_format.line_spacing = settings.line_spacing
    except Exception:
        log.warning("Normal style not set", exc_info=True)
    style_heading_levels(doc, settings)
    for section in doc.sections:
        try:
            set_section_rtl(section)
            if not settings.exact_layout:
                setup_page(section, settings)
                add_page_number_footer(section, settings)
        except Exception:
            log.warning("Section setup failed", exc_info=True)


def add_table_of_contents(doc, settings: Settings) -> None:
    """Insert a TOC field; Word fills it from the Heading styles."""
    paragraph = doc.add_paragraph()
    format_paragraph(paragraph, settings, heading=True)
    _field_run(paragraph, r' TOC \o "1-3" \h \z \u ',
               placeholder="اضغط F9 لتحديث الفهرس", dirty=True)
    doc.add_page_break()


def add_footnote_rule(doc) -> None:
    """The short line print uses to separate footnotes from the body."""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(10)
    pPr = paragraph._element.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    for key, value in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1")):
        top.set(qn(key), value)
    borders.append(top)
    pPr.append(borders)


def set_table_rtl(table) -> None:
    """Right-to-left column order, so column 1 is the rightmost."""
    tblPr = table._tbl.tblPr
    if tblPr.find(qn("w:bidiVisual")) is None:
        tblPr.append(OxmlElement("w:bidiVisual"))


def _textbox_paragraph_xml(text: str, rtl: bool, size_pt: float,
                           font_name: str) -> str:
    half = str(int(round(size_pt * 2)))
    font = quoteattr(font_name)
    rtl_tag = '<w:rtl w:val="1"/>' if rtl else ""
    rpr = (f'<w:rPr><w:rFonts w:ascii={font} w:hAnsi={font} w:cs={font}/>'
           f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>{rtl_tag}</w:rPr>')
    return "".join(
        f'<w:p><w:pPr><w:bidi w:val="{"1" if rtl else "0"}"/>'
        f'<w:spacing w:after="0"/></w:pPr>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>'
        for line in text.split("\n")
    )


def add_positioned_textbox(doc, settings: Settings, text: str,
                           left_pt: float, top_pt: float, width_pt: float,
                           height_pt: float, rtl: bool, z: int) -> None:
    """Place text at an absolute page position, as a borderless box."""
    rows = text.split("\n")
    lines = max(1, len(rows))
    longest = max((len(r) for r in rows), default=1) or 1
    # Fit the type to the box on both axes; Ruqaa is wider than the scan.
    by_height = (height_pt / lines) * 0.72
    by_width = (width_pt / longest) / 0.46
    size_pt = max(5.0, min(60.0, by_height, by_width))
    font_name = settings.arabic_font if rtl else settings.latin_font
    # A little slack keeps a glyph from touching the edge of the box.
    box_w = max(width_pt * 1.06, 8)
    box_h = max(height_pt * 1.15, 8)
    style = (f"position:absolute;"
             f"margin-left:{max(left_pt - width_pt * 0.03, 0):.1f}pt;"
             f"margin-top:{top_pt:.1f}pt;width:{box_w:.1f}pt;"
             f"height:{box_h:.1f}pt;z-index:{z};"
             "mso-position-horizontal-relative:page;"
             "mso-position-vertical-relative:page")
    xml = (
        f'<w:p xmlns:w="{W_NS}" xmlns:v="{V_NS}">'
        '<w:pPr><w:spacing w:after="0" w:line="20" w:lineRule="exact"/></w:pPr>'
        '<w:r><w:pict>'
        f'<v:shape style="{style}" filled="f" stroked="f">'
        '<v:textbox inset="0,0,0,0"><w:txbxContent>'
        f'{_textbox_paragraph_xml(text, rtl, size_pt, font_name)}'
        '</w:txbxContent></v:textbox></v:shape>'
        '</w:pict></w:r></w:p>'
    )
    doc.element.body.insert_element_before(parse_xml(xml), "w:sectPr")
