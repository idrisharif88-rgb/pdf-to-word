import sys
from pathlib import Path

import pytest
from google.cloud import documentai

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def make_ocr_document(pages):
    """Build a documentai.Document.

    ``pages`` is a list of pages; each page a list of
    (text, top, bottom, left, right) paragraphs in normalised coordinates.
    """
    full_text = ""
    doc_pages = []
    for paragraphs in pages:
        paras = []
        for text, top, bottom, left, right in paragraphs:
            start = len(full_text)
            full_text += text + "\n"
            verts = [documentai.NormalizedVertex(x=x, y=y) for x, y in
                     ((left, top), (right, top), (right, bottom), (left, bottom))]
            paras.append(documentai.Document.Page.Paragraph(
                layout=documentai.Document.Page.Layout(
                    text_anchor=documentai.Document.TextAnchor(text_segments=[
                        documentai.Document.TextAnchor.TextSegment(
                            start_index=start, end_index=start + len(text))]),
                    bounding_poly=documentai.BoundingPoly(
                        normalized_vertices=verts))))
        doc_pages.append(documentai.Document.Page(paragraphs=paras))
    return documentai.Document(text=full_text, pages=doc_pages)


BODY = ("هذا نص عربي طويل يمثل فقرة كاملة من كتاب ممسوح ضوئيا ويحتوي على "
        "كلمات كثيرة حتى يتجاوز طوله مئة وعشرين حرفا ليكون نصا متصلا")


def standard_page(n):
    """Running head, a title, two body paragraphs."""
    return [
        ("اسم الكتاب", 0.03, 0.05, 0.4, 0.6),               # running head
        (f"الفصل {n}", 0.15, 0.20, 0.35, 0.65),             # tall → heading
        (f"{BODY} {n}أ", 0.30, 0.33, 0.1, 0.9),
        (f"{BODY} {n}ب", 0.40, 0.43, 0.1, 0.9),
    ]


@pytest.fixture
def sample_pdf(tmp_path):
    """A 12-page PDF without a real text layer (one stray char per page)."""
    import pymupdf
    path = tmp_path / "book.pdf"
    pdf = pymupdf.open()
    for _ in range(12):
        page = pdf.new_page()
        page.draw_rect(pymupdf.Rect(50, 50, 300, 300), color=(0, 0, 0))
    pdf.save(str(path))
    return path
