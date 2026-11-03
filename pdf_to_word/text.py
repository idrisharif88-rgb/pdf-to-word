"""Pure text and geometry helpers: no Word, PDF or network access."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, List, Sequence, Set

TATWEEL = "ـ"
BIDI_MARKS = "‎‏‪-‮⁦-⁩"
ARABIC_RANGE = "؀-ۿ"

_ARABIC_RE = re.compile(f"[{ARABIC_RANGE}]")
_LATIN_RE = re.compile(r"[A-Za-z]")


@dataclass(frozen=True)
class TextUnit:
    """One OCR paragraph with its normalised (0..1) bounding box."""
    text: str
    line_height: float
    top: float
    bottom: float
    left: float
    right: float

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def line_count(self) -> int:
        return self.text.count("\n") + 1


def clean_text_smart(text: str) -> str:
    """Tidy OCR output without flattening the line structure."""
    # Control characters and bidi overrides, but keep newlines and tabs.
    text = re.sub(r"[\u0000-\u0008\u000b-\u001f]", "", text)
    text = re.sub(f"[{BIDI_MARKS}]", "", text)
    # Tatweel is decorative stretching the OCR sprinkles mid-word.
    text = text.replace(TATWEEL, "")
    # Soft hyphens and scan artefacts.
    text = text.replace("­", "").replace("¬", "").replace("|", "")
    # Collapse runs of spaces, never collapse a line break away.
    text = re.sub(r"[^\S\n]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    # No space before punctuation.
    text = re.sub(r"[^\S\n]+([,.:;!?؟،؛])", r"\1", text)
    # Latin punctuation sitting in Arabic text becomes its Arabic form.
    text = re.sub(f"(?<=[{ARABIC_RANGE}]),", "،", text)
    text = re.sub(f"(?<=[{ARABIC_RANGE}]);", "؛", text)
    text = re.sub(f"(?<=[{ARABIC_RANGE}])\\?", "؟", text)
    return text.strip()


def detect_rtl(text: str) -> bool:
    """Decide a paragraph's direction from its own letters.

    A copyright page mixes Arabic blocks with Latin ones; forcing RTL onto the
    Latin lines turns "E-mail:" into ":E-mail" and reverses phone numbers, so
    each paragraph is judged separately.
    """
    return len(_ARABIC_RE.findall(text)) >= len(_LATIN_RE.findall(text))


def unwrap_lines(text: str) -> str:
    """Join OCR line-wraps inside one paragraph.

    Document AI breaks a paragraph at every printed line. Keeping those as
    hard breaks makes Word justify each fragment separately.
    """
    text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    return re.sub(r"[^\S\n]{2,}", " ", text)


def is_ocr_noise(text: str) -> bool:
    """A lone stray character is a scan artefact, not content."""
    stripped = text.strip()
    return len(stripped) <= 1 and not stripped.isdigit()


def repeat_key(text: str) -> str:
    """Normalise text so a running head matches across pages (digits vary)."""
    return re.sub(r"\d+", "#", " ".join(text.split()))[:60]


def running_head_texts(pages_units: Sequence[Sequence[TextUnit]],
                       margin_band: float) -> Set[str]:
    """Keys of text that repeats in the top or bottom band on most pages."""
    if len(pages_units) < 3:
        return set()
    seen: dict = {}
    for units in pages_units:
        page_hits = set()
        for u in units:
            in_band = u.top <= margin_band or u.bottom >= 1.0 - margin_band
            key = repeat_key(u.text)
            if in_band and key:
                page_hits.add(key)
        for key in page_hits:
            seen[key] = seen.get(key, 0) + 1
    threshold = max(2, int(len(pages_units) * 0.6))
    return {k for k, n in seen.items() if n >= threshold}


def parse_page_spec(spec: str) -> Set[int]:
    """Page numbers from a spec like "1-6,140-146"; bad parts are ignored."""
    pages: Set[int] = set()
    for part in (spec or "").replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            start, _, end = part.partition("-")
            if start.isdigit() and end.isdigit():
                pages.update(range(int(start), int(end) + 1))
        elif part.isdigit():
            pages.add(int(part))
    return pages


def parse_page_range(page_input: str, total_pages: int) -> List[int]:
    """Pages selected by "all", "1-5" or "2,5,7", clipped to the document.

    Raises ValueError on malformed input so the GUI can say what is wrong.
    """
    page_input = page_input.strip().lower()
    if page_input in ("", "all"):
        return list(range(1, total_pages + 1))
    pages: Set[int] = set()
    for part in page_input.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            start, _, end = part.partition("-")
            if not (start.isdigit() and end.isdigit()):
                raise ValueError(f"Invalid page range: {part!r}")
            pages.update(range(int(start), int(end) + 1))
        elif part.isdigit():
            pages.add(int(part))
        else:
            raise ValueError(f"Invalid page number: {part!r}")
    return sorted(p for p in pages if 1 <= p <= total_pages)


def heading_level(unit_height: float, median_h: float, text: str,
                  lines: int, width: float,
                  ratio_1: float, ratio_2: float) -> int:
    """0 for body text, otherwise the heading level to use.

    Measured against the median line height of the whole batch, so a page
    that is entirely title still gets its heading.
    """
    if median_h <= 0 or len(text) > 90 or lines > 3:
        return 0
    # A short line that runs the full width is a paragraph's last line.
    if width > 0.8:
        return 0
    ratio = unit_height / median_h
    if ratio >= ratio_1:
        return 1
    if ratio >= ratio_2:
        return 2
    return 0


def is_footnote(top: float, height: float, median_h: float,
                band: float, max_ratio: float) -> bool:
    """Small type low on the page is the footnote apparatus."""
    return median_h > 0 and top >= band and height / median_h <= max_ratio


def detect_placement(left: float, right: float) -> str:
    """Recover how a block sat on the printed page: full/center/left/right."""
    left_gap = max(left, 0.0)
    right_gap = max(1.0 - right, 0.0)
    if right - left >= 0.72:
        return "full"
    if left_gap > 0.12 and right_gap > 0.12 and abs(left_gap - right_gap) < 0.09:
        return "center"
    if left_gap + 0.10 < right_gap:
        return "left"
    if right_gap + 0.10 < left_gap:
        return "right"
    return "full"


def in_spans(top: float, bottom: float,
             spans: Iterable[tuple], slack: float = 0.01) -> bool:
    """True when the block lies inside one of the vertical spans."""
    return any(top >= s - slack and bottom <= e + slack for s, e in spans)


def batches(items: Sequence[int], size: int) -> List[List[int]]:
    return [list(items[i:i + size]) for i in range(0, len(items), size)]
