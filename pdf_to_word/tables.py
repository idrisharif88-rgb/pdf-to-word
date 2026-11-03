"""Rebuild Document AI tables as Word tables."""
from __future__ import annotations

from typing import List, Tuple

from .config import Settings
from .docx_format import format_paragraph, set_table_rtl
from .text import clean_text_smart, unwrap_lines


def anchor_text(document, layout) -> str:
    return "".join(
        document.text[int(seg.start_index):int(seg.end_index)]
        for seg in layout.text_anchor.text_segments
    )


def _rows(table):
    return list(table.header_rows) + list(table.body_rows)


def is_real_table(document, table, settings: Settings) -> bool:
    """Reject grids the parser invented (e.g. a two-column list returned as
    a wide, mostly empty grid)."""
    rows = _rows(table)
    if not rows:
        return False
    if max(len(r.cells) for r in rows) > settings.table_max_cols:
        return False
    total = sum(len(r.cells) for r in rows)
    filled = sum(1 for r in rows for c in r.cells
                 if anchor_text(document, c.layout).strip())
    return total > 0 and filled / total >= settings.table_min_fill


def table_regions(page, document, settings: Settings) -> List[Tuple[float, float]]:
    """Vertical spans covered by accepted tables, so their text is not also
    emitted as paragraphs. A rejected grid leaves its text to flow."""
    spans = []
    for table in page.tables:
        if not is_real_table(document, table, settings):
            continue
        verts = table.layout.bounding_poly.normalized_vertices
        if verts:
            spans.append((min(v.y for v in verts), max(v.y for v in verts)))
    return spans


def add_tables(doc, document, page, settings: Settings) -> None:
    for table in page.tables:
        if not is_real_table(document, table, settings):
            continue
        rows = _rows(table)
        width = max(len(r.cells) for r in rows)
        word_table = doc.add_table(rows=len(rows), cols=width)
        word_table.style = "Table Grid"
        word_table.autofit = True
        set_table_rtl(word_table)
        for r, row in enumerate(rows):
            for c, cell in enumerate(row.cells[:width]):
                text = anchor_text(document, cell.layout)
                if settings.cleanup:
                    text = clean_text_smart(text)
                text = unwrap_lines(text).strip()
                target = word_table.cell(r, c).paragraphs[0]
                if text:
                    target.add_run(text)
                format_paragraph(target, settings,
                                 heading=r < len(table.header_rows),
                                 size_scale=settings.table_size_scale)
        doc.add_paragraph()
