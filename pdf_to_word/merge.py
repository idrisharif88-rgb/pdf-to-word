"""Batch part files: naming, resume detection and merging."""
from __future__ import annotations

import logging
import os
import re
from io import BytesIO
from pathlib import Path
from typing import Callable, Dict, List

from docx import Document
from docx.oxml.ns import qn

from .config import Settings
from .docx_format import add_table_of_contents, setup_document

log = logging.getLogger(__name__)

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def part_path(output_dir: Path, base_name: str, part_num: int) -> Path:
    return output_dir / f"{base_name}_part_{part_num}.docx"


def existing_parts(output_dir: Path, base_name: str) -> Dict[int, Path]:
    """Finished part files keyed by part number (numeric, not lexical)."""
    pattern = re.compile(rf"^{re.escape(base_name)}_part_(\d+)\.docx$")
    found = {}
    if output_dir.is_dir():
        for entry in output_dir.iterdir():
            match = pattern.match(entry.name)
            if match:
                found[int(match.group(1))] = entry
    return dict(sorted(found.items()))


def save_atomically(doc, target: Path) -> None:
    """Write to a temp name first, so a crash never leaves a half part that
    a later resume would trust."""
    tmp = target.with_name(target.name + ".tmp")
    doc.save(str(tmp))
    os.replace(tmp, target)


def carry_relationships(element, source_part, target_part) -> None:
    """Re-point a copied element's relationship ids at the merged document.

    A picture lives in its own package part; moving the XML alone would leave
    an id pointing at nothing, which Word draws as an empty frame.
    """
    for attr in ("embed", "link", "id", "pict"):
        name = f"{{{R_NS}}}{attr}"
        for node in element.iter():
            rid = node.get(name)
            if not rid:
                continue
            try:
                rel = source_part.rels[rid]
            except KeyError:
                continue
            try:
                if rel.is_external:
                    new_rid = target_part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True)
                else:
                    source = rel.target_part
                    if source.partname.startswith("/word/media/"):
                        # Each part numbers its images from one; adding the
                        # bytes through the merged package renames them and
                        # drops duplicates.
                        source = target_part.package.image_parts \
                            .get_or_add_image_part(BytesIO(source.blob))
                    new_rid = target_part.relate_to(source, rel.reltype)
                node.set(name, new_rid)
            except Exception:
                log.warning("Could not carry relationship %s", rid,
                            exc_info=True)


def merge_parts(parts: List[Path], merged_path: Path, settings: Settings,
                on_log: Callable[[str], None] = lambda m: None) -> Path:
    """Concatenate part files into ``merged_path``; returns the path."""
    merged = Document()
    setup_document(merged, settings)
    if settings.add_toc:
        add_table_of_contents(merged, settings)
    body = merged.element.body
    trailing_sectPr = body.find(qn("w:sectPr"))

    for path in parts:
        on_log(f"Merging {path.name}")
        part_doc = Document(str(path))
        for element in list(part_doc.element.body):
            # Each part has its own sectPr; copying them would add a section
            # break between every part.
            if element.tag == qn("w:sectPr"):
                continue
            carry_relationships(element, part_doc.part, merged.part)
            if trailing_sectPr is not None:
                trailing_sectPr.addprevious(element)
            else:
                body.append(element)

    save_atomically(merged, merged_path)
    return merged_path
