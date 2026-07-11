#!/usr/bin/env python3
"""
cli.py
------
Command-line / headless-service entry point (no GUI required).
Useful for running on a server, in Task Scheduler / cron, or as a
Windows Service / systemd unit.

Usage
-----
Run once, extract everything currently in the folder, then exit:
    python cli.py --folder "C:\\path\\to\\pdfs"

Run once, custom output folder:
    python cli.py --folder ./pdfs --output ./reports

Run as a persistent watch-folder service (auto-processes new PDFs):
    python cli.py --folder ./pdfs --watch

Include sub-folders too:
    python cli.py --folder ./pdfs --recursive
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from core.pipeline import process_folder, setup_logging
from core.watcher import FolderWatcher


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract data from PDFs into a professional Excel report.")
    parser.add_argument("--folder", "-f", required=True, help="Folder containing the PDF files (mandatory).")
    parser.add_argument("--output", "-o", default=None, help="Folder to save the Excel report (defaults to --folder).")
    parser.add_argument("--recursive", "-r", action="store_true", help="Also scan sub-folders.")
    parser.add_argument("--watch", "-w", action="store_true", help="Run continuously, watching the folder for new PDFs.")
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    output = Path(args.output).expanduser().resolve() if args.output else None

    logger = setup_logging()

    if not folder.exists() or not folder.is_dir():
        logger.error("Folder does not exist: %s", folder)
        print(f"ERROR: folder does not exist: {folder}", file=sys.stderr)
        return 1

    if args.watch:
        logger.info("Starting watch-folder service on: %s", folder)

        def _on_change():
            process_folder(folder, output, recursive=args.recursive)

        # Run once immediately on startup, then watch for changes.
        _on_change()
        watcher = FolderWatcher(folder, _on_change)
        watcher.run_forever()
        return 0

    report_path = process_folder(folder, output, recursive=args.recursive)
    print(f"Report generated: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
