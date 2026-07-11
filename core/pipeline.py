"""
pipeline.py
-----------
Single entry point that ties extraction + Excel writing together, and
sets up rotating file logging. Used by main.py (GUI), cli.py and the
watch-folder service so all three modes behave identically.
"""

from __future__ import annotations

import logging
import logging.handlers
from datetime import datetime
from pathlib import Path

from .excel_writer import write_workbook
from .extractor import extract_folder
from .facture_parser import parse_facture

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("pdf_extractor")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

        file_handler = logging.handlers.RotatingFileHandler(
            LOG_DIR / "pdf_extractor.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(fmt)
        logger.addHandler(console_handler)
    return logger


def process_folder(folder: Path, output_dir: Path | None = None, recursive: bool = False) -> Path:
    """
    Extract every PDF in `folder`, write a timestamped Excel report into
    `output_dir` (defaults to `folder`), and return the report's path.
    """
    logger = setup_logging()
    folder = Path(folder)
    output_dir = Path(output_dir) if output_dir else folder

    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"Folder not found: {folder}")

    logger.info("Starting extraction for folder: %s", folder)
    results = extract_folder(folder, recursive=recursive)

    if not results:
        logger.warning("No PDF files found in %s", folder)

    factures = []
    for res in results:
        if res.error:
            continue
        fac = parse_facture(res.full_text, res.file_name)
        if fac is not None:
            factures.append(fac)
    if factures:
        logger.info("Recognized %d document(s) as 'Facture Proforma' template", len(factures))
        factures.sort(key=lambda f: (int(f.numero) if f.numero.isdigit() else 0, f.numero))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"PDF_Extraction_Report_{timestamp}.xlsx"
    write_workbook(results, output_path, source_folder=str(folder), factures=factures)

    ok = sum(1 for r in results if not r.error)
    failed = len(results) - ok
    logger.info("Done. %d file(s) OK, %d failed. Report: %s", ok, failed, output_path)

    return output_path
