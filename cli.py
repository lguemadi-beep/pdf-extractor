#!/usr/bin/env python3
"""
cli.py
------
Command-line / headless-service entry point (no GUI required).
Useful for running on a server, in Task Scheduler / cron, or as a
Windows Service / systemd unit.

The PDF folder only needs to be set up ONCE, with --configure. After
that, the service (or any later run) can be started with no arguments
at all — it reads the saved folder from config.json — so restarting
the laptop / the Windows service never requires re-entering the folder.

Usage
-----
One-time setup (writes config.json, does not run extraction):
    python cli.py --configure --folder "C:\\path\\to\\pdfs" --output "C:\\path\\to\\reports"

Run once, using the saved config:
    python cli.py

Run once, overriding the saved folder just for this run:
    python cli.py --folder ./pdfs --output ./reports

Run as a persistent watch-folder service (auto-processes new PDFs),
using the saved config — this is what the Windows service uses:
    python cli.py --watch

Include sub-folders too:
    python cli.py --recursive
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from core.config import AppConfig, load_config, save_config
from core.pipeline import process_folder, setup_logging
from core.watcher import FolderWatcher


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract data from PDFs into a professional Excel report.")
    parser.add_argument(
        "--configure", action="store_true",
        help="Save --folder/--output/etc. to config.json once, then exit (no extraction runs).",
    )
    parser.add_argument("--folder", "-f", default=None, help="Folder containing the PDF files. Only needed once (or to override the saved folder).")
    parser.add_argument("--output", "-o", default=None, help="Folder to save the Excel report (defaults to --folder).")
    parser.add_argument("--recursive", "-r", action="store_true", help="Also scan sub-folders.")
    parser.add_argument(
        "--name-filter", default=None,
        help='Only process PDFs whose filename contains this text (case-insensitive). Default: "proforma". Use "" to disable.',
    )
    parser.add_argument("--watch", "-w", action="store_true", help="Run continuously, watching the folder for new PDFs.")
    args = parser.parse_args()

    logger = setup_logging()
    saved = load_config()

    # ---------------------------------------------------------------- --configure
    if args.configure:
        if not args.folder:
            print("ERROR: --configure requires --folder (the PDF folder to remember).", file=sys.stderr)
            return 1
        cfg = AppConfig(
            pdf_folder=str(Path(args.folder).expanduser().resolve()),
            output_folder=str(Path(args.output).expanduser().resolve()) if args.output else "",
            name_filter=args.name_filter if args.name_filter is not None else "proforma",
            recursive=bool(args.recursive),
        )
        path = save_config(cfg)
        print(f"Saved configuration to {path}")
        print(f"  PDF folder:    {cfg.pdf_folder}")
        print(f"  Output folder: {cfg.output_folder or '(same as PDF folder)'}")
        print(f"  Name filter:   '{cfg.name_filter}'")
        print("You can now run the service with no arguments, e.g.:  cli.py --watch")
        return 0

    # ------------------------------------------------ Resolve folder/output/etc.
    # Priority: explicit CLI flag > saved config.json > error.
    folder_str = args.folder or (saved.pdf_folder if saved else None)
    if not folder_str:
        print(
            "ERROR: no PDF folder given, and none saved yet.\n"
            'Run this once first:  cli.py --configure --folder "C:\\path\\to\\pdfs"',
            file=sys.stderr,
        )
        return 1

    output_str = args.output or (saved.output_folder if saved else None) or None
    name_filter = args.name_filter if args.name_filter is not None else (saved.name_filter if saved else "proforma")
    recursive = args.recursive or (bool(saved.recursive) if saved else False)

    folder = Path(folder_str).expanduser().resolve()
    output = Path(output_str).expanduser().resolve() if output_str else None

    if not folder.exists() or not folder.is_dir():
        logger.error("Folder does not exist: %s", folder)
        print(f"ERROR: folder does not exist: {folder}", file=sys.stderr)
        return 1

    if args.watch:
        logger.info("Starting watch-folder service on: %s", folder)

        def _on_change():
            process_folder(folder, output, recursive=recursive, name_filter=name_filter)

        # Run once immediately on startup, then watch for changes.
        _on_change()
        watcher = FolderWatcher(folder, _on_change)
        watcher.run_forever()
        return 0

    report_path = process_folder(folder, output, recursive=recursive, name_filter=name_filter)
    print(f"Report generated: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
