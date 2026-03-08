from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

import fitz

from pdf_toolkit.core import parse_page_spec
from pdf_toolkit.errors import ValidationError


@dataclass(slots=True)
class RedactionBox:
    page_number: int
    x1: float
    y1: float
    x2: float
    y2: float

    @property
    def rect(self) -> fitz.Rect:
        return fitz.Rect(self.x1, self.y1, self.x2, self.y2)


def parse_redaction_box(spec: str) -> RedactionBox:
    parts = [part.strip() for part in spec.split(",")]
    if len(parts) != 5:
        raise ValidationError(f"Invalid redaction box '{spec}'. Expected page,x1,y1,x2,y2.")
    page_number = int(parts[0])
    x1, y1, x2, y2 = (float(value) for value in parts[1:])
    if page_number < 1:
        raise ValidationError("Redaction box page numbers must be 1-based and positive.")
    if x1 >= x2 or y1 >= y2:
        raise ValidationError(f"Invalid redaction box '{spec}'. Coordinates must define a positive rectangle.")
    return RedactionBox(page_number=page_number, x1=x1, y1=y1, x2=x2, y2=y2)


def _page_selection(page_count: int, page_spec: str | None) -> set[int] | None:
    if not page_spec:
        return None
    return set(parse_page_spec(page_spec, page_count, allow_duplicates=False))


def _regex_matches(page: fitz.Page, pattern: str, case_sensitive: bool) -> list[fitz.Rect]:
    flags = 0 if case_sensitive else re.IGNORECASE
    compiled = re.compile(pattern, flags)
    rects: list[fitz.Rect] = []
    for word in page.get_text("words"):
        text = str(word[4])
        if compiled.search(text):
            rects.append(fitz.Rect(word[:4]))
    return rects


def run_redaction(
    input_path: Path,
    *,
    output_path: Path | None,
    patterns: list[str],
    regex: bool,
    case_sensitive: bool,
    page_spec: str | None,
    box_specs: list[str],
    label: str | None,
    dry_run: bool,
) -> dict[str, object]:
    if not patterns and not box_specs:
        raise ValidationError("Provide at least one --pattern or --box redaction target.")
    if not dry_run and output_path is None:
        raise ValidationError("Provide --output when applying redactions, or use --dry-run for report-only mode.")
    document = fitz.open(input_path)
    selected = _page_selection(document.page_count, page_spec)
    manual_boxes = [parse_redaction_box(spec) for spec in box_specs]
    matches: list[dict[str, object]] = []
    pages_touched: set[int] = set()

    try:
        for page_index in range(document.page_count):
            if selected is not None and page_index not in selected:
                continue
            page = document[page_index]
            rects: list[tuple[fitz.Rect, str]] = []

            for pattern in patterns:
                found = _regex_matches(page, pattern, case_sensitive) if regex else page.search_for(pattern)
                for rect in found:
                    rects.append((rect, pattern))

            for box in manual_boxes:
                if box.page_number - 1 == page_index:
                    rects.append((box.rect, "manual-box"))

            for rect, source in rects:
                pages_touched.add(page_index + 1)
                matches.append(
                    {
                        "page": page_index + 1,
                        "source": source,
                        "bbox": [rect.x0, rect.y0, rect.x1, rect.y1],
                    }
                )
                if not dry_run:
                    page.add_redact_annot(rect, text=label or "")

        if not dry_run and output_path is not None:
            for page_number in sorted(pages_touched):
                document[page_number - 1].apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            document.save(output_path, garbage=4, deflate=True, clean=True)
    finally:
        document.close()

    return {
        "outputs": [output_path] if output_path else [],
        "details": {
            "matches": matches,
            "pages_touched": sorted(pages_touched),
            "match_count": len(matches),
            "dry_run": dry_run,
            "regex": regex,
        },
    }
