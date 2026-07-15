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
import sys
from pathlib import Path

from .excel_writer import write_workbook
from .extractor import extract_folder
from .facture_parser import parse_facture

# When bundled by PyInstaller (--onefile), __file__ resolves inside a
# temporary extraction folder that's wiped after the process exits — logs
# written there would vanish, which is a problem for a background service
# that needs to survive reboots and be troubleshootable. Anchor LOG_DIR next
# to the actual .exe (or the source tree, when run as plain Python) instead.
if getattr(sys, "frozen", False):
    _APP_DIR = Path(sys.executable).resolve().parent
else:
    _APP_DIR = Path(__file__).resolve().parent.parent
LOG_DIR = _APP_DIR / "logs"


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


def process_folder(
    folder: Path,
    output_dir: Path | None = None,
    recursive: bool = False,
    name_filter: str = "proforma",
) -> Path:
    """
    Extract every PDF in `folder` whose filename contains `name_filter`
    (case-insensitive; set to "" to disable), skipping byte-identical
    duplicate files, and merge the results into `invoice_summary.xlsx` in
    `output_dir` (defaults to `folder`) — existing rows from previous runs
    are preserved; only genuinely new invoices are appended. Returns the
    report's path.
    """
    logger = setup_logging()
    folder = Path(folder)
    output_dir = Path(output_dir) if output_dir else folder

    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"Folder not found: {folder}")

    logger.info("Starting extraction for folder: %s (filter: '%s')", folder, name_filter)
    results = extract_folder(folder, recursive=recursive, name_filter=name_filter)

    if not results:
        logger.warning("No matching PDF files found in %s", folder)

    factures = []
    seen_numeros: dict[str, str] = {}
    for res in results:
        if res.error:
            continue
        fac = parse_facture(res.full_text, res.file_name, tables=res.tables)
        if fac is None:
            continue
        if fac.numero and fac.numero in seen_numeros:
            logger.info(
                "Skipping duplicate invoice N°%s (already seen in %s): %s",
                fac.numero, seen_numeros[fac.numero], fac.file_name,
            )
            continue
        if fac.numero:
            seen_numeros[fac.numero] = fac.file_name
        factures.append(fac)

    if factures:
        logger.info("Recognized %d document(s) as 'Facture Proforma' template", len(factures))
        factures.sort(key=lambda f: (int(f.numero) if f.numero.isdigit() else 0, f.numero))

    output_path = output_dir / "invoice_summary.xlsx"
    write_workbook(results, output_path, source_folder=str(folder), factures=factures)

    ok = sum(1 for r in results if not r.error)
    failed = len(results) - ok
    logger.info("Done. %d file(s) OK, %d failed. Report: %s", ok, failed, output_path)

    return output_path
