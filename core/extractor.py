"""
extractor.py
------------
Core PDF data-extraction logic.

Designed for good-quality, text-based PDFs (not scanned images).
For each PDF it pulls out:
  1. Document metadata (filename, page count, dates, etc.)
  2. Every table found on every page (via pdfplumber's table detector)
  3. "Key: Value" style fields found in the free text (e.g. "Invoice No: 4521")

The results are handed back as plain Python data structures so the
Excel writer (excel_writer.py) can stay completely independent of
the PDF library being used.
"""

from __future__ import annotations

import hashlib
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import pdfplumber

logger = logging.getLogger("pdf_extractor")

# ---------------------------------------------------------------------------
# Regex used to spot "Label: Value" or "Label : Value" pairs in free text.
# Tuned to avoid matching inside sentences (label must be short, start of line).
# ---------------------------------------------------------------------------
KEY_VALUE_PATTERN = re.compile(
    r"^\s*([A-Za-z][A-Za-z0-9 /_.\-]{1,40}?)\s*[:#]\s+(.+?)\s*$"
)

# Common numeric / date patterns, extracted separately so they land in their
# own columns even when they are not part of a "Label: Value" line.
DATE_PATTERN = re.compile(
    r"\b(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})\b"
)
AMOUNT_PATTERN = re.compile(
    r"\b(?:[$€£]|USD|EUR|GBP)\s?[\d,]+\.\d{2}\b|\b[\d,]+\.\d{2}\s?(?:[$€£]|USD|EUR|GBP)\b"
)


@dataclass
class TableResult:
    page_number: int
    table_index: int
    headers: list[str]
    rows: list[list[str]]


@dataclass
class PdfResult:
    file_path: Path
    file_name: str
    page_count: int = 0
    extracted_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))
    key_values: dict[str, str] = field(default_factory=dict)
    tables: list[TableResult] = field(default_factory=list)
    dates_found: list[str] = field(default_factory=list)
    amounts_found: list[str] = field(default_factory=list)
    full_text: str = ""
    error: str | None = None


def _clean_cell(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", " ").strip()


def extract_pdf(file_path: Path) -> PdfResult:
    """Extract tables, key-value fields, dates and amounts from one PDF."""
    result = PdfResult(file_path=file_path, file_name=file_path.name)

    try:
        with pdfplumber.open(str(file_path)) as pdf:
            result.page_count = len(pdf.pages)
            all_text_lines: list[str] = []

            for page_number, page in enumerate(pdf.pages, start=1):
                # ---- Tables -----------------------------------------------------
                page_tables = page.extract_tables()
                for t_idx, table in enumerate(page_tables, start=1):
                    if not table or len(table) < 1:
                        continue
                    headers = [_clean_cell(c) or f"col_{i+1}" for i, c in enumerate(table[0])]
                    rows = [[_clean_cell(c) for c in row] for row in table[1:]]
                    result.tables.append(
                        TableResult(page_number=page_number, table_index=t_idx, headers=headers, rows=rows)
                    )

                # ---- Free text ----------------------------------------------------
                text = page.extract_text() or ""
                if text:
                    all_text_lines.extend(text.splitlines())

            result.full_text = "\n".join(all_text_lines)

            # ---- Key-value fields --------------------------------------------
            for line in all_text_lines:
                m = KEY_VALUE_PATTERN.match(line)
                if m:
                    key, value = m.group(1).strip(), m.group(2).strip()
                    if key and value and len(key) < 40:
                        # keep the first occurrence of each key
                        result.key_values.setdefault(key, value)

            # ---- Dates & amounts (deduplicated, order preserved) --------------
            seen_dates: set[str] = set()
            for d in DATE_PATTERN.findall(result.full_text):
                if d not in seen_dates:
                    seen_dates.add(d)
                    result.dates_found.append(d)

            seen_amounts: set[str] = set()
            for a in AMOUNT_PATTERN.findall(result.full_text):
                if a not in seen_amounts:
                    seen_amounts.add(a)
                    result.amounts_found.append(a)

    except Exception as exc:  # noqa: BLE001 - we want to record ANY failure and keep going
        logger.exception("Failed to extract %s", file_path)
        result.error = f"{type(exc).__name__}: {exc}"

    return result


def extract_folder(folder_path: Path, recursive: bool = False, name_filter: str = "proforma") -> list[PdfResult]:
    """
    Extract every PDF found directly inside (or recursively under) folder_path.

    - Only files whose *filename* contains `name_filter` (case-insensitive) are
      processed. Set name_filter="" to disable this filtering.
    - Duplicate files (identical content, e.g. the same invoice saved twice
      under different names) are skipped after the first occurrence, based on
      a content hash — not just the filename.
    """
    pattern = "**/*.pdf" if recursive else "*.pdf"
    all_pdfs = sorted(folder_path.glob(pattern))

    if name_filter:
        needle = name_filter.lower()
        candidates = [p for p in all_pdfs if needle in p.name.lower()]
        skipped_name = [p for p in all_pdfs if needle not in p.name.lower()]
        for p in skipped_name:
            logger.info("Skipping (filename does not contain '%s'): %s", name_filter, p.name)
    else:
        candidates = all_pdfs

    results: list[PdfResult] = []
    seen_hashes: dict[str, Path] = {}
    for pdf_file in candidates:
        try:
            file_hash = hashlib.sha256(pdf_file.read_bytes()).hexdigest()
        except Exception:
            file_hash = None

        if file_hash and file_hash in seen_hashes:
            logger.info(
                "Skipping duplicate (identical content to %s): %s",
                seen_hashes[file_hash].name, pdf_file.name,
            )
            continue
        if file_hash:
            seen_hashes[file_hash] = pdf_file

        logger.info("Extracting %s", pdf_file.name)
        results.append(extract_pdf(pdf_file))

    return results
