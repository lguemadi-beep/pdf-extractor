"""
watcher.py
----------
Turns the extractor into a background "service": watches a folder and,
whenever a new PDF is dropped in (or the app starts up), automatically
extracts it and (re)generates the Excel report.

Uses the `watchdog` library for efficient OS-level file-system events,
with a debounce so a file being copied in isn't read half-written.
"""

from __future__ import annotations

import logging
import threading
import time
from pathlib import Path
from typing import Callable

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

logger = logging.getLogger("pdf_extractor")

DEBOUNCE_SECONDS = 2.0


class _PdfHandler(FileSystemEventHandler):
    def __init__(self, on_change: Callable[[], None]):
        self._on_change = on_change
        self._timer: threading.Timer | None = None
        self._lock = threading.Lock()

    def _schedule(self) -> None:
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
            self._timer = threading.Timer(DEBOUNCE_SECONDS, self._fire)
            self._timer.daemon = True
            self._timer.start()

    def _fire(self) -> None:
        try:
            self._on_change()
        except Exception:
            logger.exception("Error while handling folder change")

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".pdf"):
            logger.info("New PDF detected: %s", event.src_path)
            self._schedule()

    def on_modified(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".pdf"):
            self._schedule()


class FolderWatcher:
    """Watches `folder` and calls `on_change()` (debounced) whenever a PDF appears/changes."""

    def __init__(self, folder: Path, on_change: Callable[[], None]):
        self.folder = folder
        self.on_change = on_change
        self._observer: Observer | None = None

    def start(self) -> None:
        handler = _PdfHandler(self.on_change)
        self._observer = Observer()
        self._observer.schedule(handler, str(self.folder), recursive=False)
        self._observer.start()
        logger.info("Watching folder: %s", self.folder)

    def stop(self) -> None:
        if self._observer is not None:
            self._observer.stop()
            self._observer.join(timeout=5)
            logger.info("Stopped watching: %s", self.folder)

    def run_forever(self) -> None:
        """Blocking call for pure CLI/service usage (no GUI)."""
        self.start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop()
