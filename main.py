#!/usr/bin/env python3
"""
main.py
-------
Desktop GUI for the PDF -> Excel extraction service.

- Folder selection is mandatory (Extract / Start Service buttons are
  disabled until a folder is chosen).
- "Extract now" runs a one-off extraction of everything in the folder.
- "Run as background service" keeps watching the folder and
  automatically re-extracts whenever a PDF is added.
- Progress and log messages are shown live in the window; everything is
  also written to logs/pdf_extractor.log.
"""

from __future__ import annotations

import logging
import queue
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, Y, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from core.pipeline import process_folder, setup_logging
from core.watcher import FolderWatcher

APP_TITLE = "PDF Data Extractor"


class QueueLogHandler(logging.Handler):
    """Forwards log records into a thread-safe queue the GUI polls."""

    def __init__(self, log_queue: "queue.Queue[str]"):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        self.log_queue.put(self.format(record))


class PdfExtractorApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("720x520")
        self.root.minsize(640, 440)

        self.folder_var = StringVar(value="")
        self.output_var = StringVar(value="(same as source folder)")
        self.recursive_var = StringVar(value="0")
        self.status_var = StringVar(value="Select a folder to begin.")

        self.selected_folder: Path | None = None
        self.selected_output: Path | None = None

        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.watcher: FolderWatcher | None = None
        self.watching = False

        self._build_ui()
        self._wire_logging()
        self.root.after(150, self._drain_log_queue)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 8}

        header = ttk.Frame(self.root)
        header.pack(fill=X, **pad)
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Select a PDF folder, extract data, and get a professional Excel report.",
            foreground="#595959",
        ).pack(anchor="w")

        # ---- Folder selection (mandatory) ---------------------------------
        folder_frame = ttk.LabelFrame(self.root, text="1. Source folder (required)")
        folder_frame.pack(fill=X, **pad)
        row = ttk.Frame(folder_frame)
        row.pack(fill=X, padx=10, pady=8)
        ttk.Entry(row, textvariable=self.folder_var, state="readonly").pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row, text="Browse...", command=self._choose_folder).pack(side=LEFT, padx=(8, 0))

        # ---- Output folder (optional) --------------------------------------
        output_frame = ttk.LabelFrame(self.root, text="2. Output folder for the Excel report (optional)")
        output_frame.pack(fill=X, **pad)
        row2 = ttk.Frame(output_frame)
        row2.pack(fill=X, padx=10, pady=8)
        ttk.Entry(row2, textvariable=self.output_var, state="readonly").pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row2, text="Browse...", command=self._choose_output).pack(side=LEFT, padx=(8, 0))

        options_frame = ttk.Frame(self.root)
        options_frame.pack(fill=X, padx=12)
        ttk.Checkbutton(
            options_frame, text="Include sub-folders", variable=self.recursive_var, onvalue="1", offvalue="0"
        ).pack(anchor="w")

        # ---- Actions --------------------------------------------------------
        action_frame = ttk.LabelFrame(self.root, text="3. Run")
        action_frame.pack(fill=X, **pad)
        row3 = ttk.Frame(action_frame)
        row3.pack(fill=X, padx=10, pady=8)

        self.extract_btn = ttk.Button(row3, text="Extract now", command=self._extract_now, state="disabled")
        self.extract_btn.pack(side=LEFT)

        self.service_btn = ttk.Button(
            row3, text="Start background service", command=self._toggle_service, state="disabled"
        )
        self.service_btn.pack(side=LEFT, padx=(10, 0))

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.pack(fill=X, padx=10, pady=(0, 10))

        # ---- Status / log ----------------------------------------------------
        log_frame = ttk.LabelFrame(self.root, text="Activity log")
        log_frame.pack(fill=BOTH, expand=True, **pad)
        self.log_box = ttk.Treeview(log_frame, columns=("msg",), show="", height=8)
        self.log_box.pack(fill=BOTH, expand=True, side=LEFT, padx=(10, 0), pady=10)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_box.yview)
        scrollbar.pack(side=RIGHT, fill=Y, pady=10)
        self.log_box.configure(yscrollcommand=scrollbar.set)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", relief="sunken")
        status_bar.pack(fill=X, side="bottom")

    def _wire_logging(self) -> None:
        logger = setup_logging()
        handler = QueueLogHandler(self.log_queue)
        handler.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
        logger.addHandler(handler)

    def _drain_log_queue(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_box.insert("", END, values=(msg,))
                self.log_box.yview_moveto(1.0)
        except queue.Empty:
            pass
        self.root.after(150, self._drain_log_queue)

    # ------------------------------------------------------------ Handlers
    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing PDF files")
        if folder:
            self.selected_folder = Path(folder)
            self.folder_var.set(folder)
            self.status_var.set(f"Folder selected: {folder}")
            self.extract_btn.configure(state="normal")
            self.service_btn.configure(state="normal")

    def _choose_output(self) -> None:
        folder = filedialog.askdirectory(title="Select folder to save the Excel report")
        if folder:
            self.selected_output = Path(folder)
            self.output_var.set(folder)

    def _extract_now(self) -> None:
        if not self.selected_folder:
            messagebox.showwarning(APP_TITLE, "Please select a source folder first.")
            return
        self._set_busy(True, "Extracting PDFs...")

        def worker():
            try:
                report_path = process_folder(
                    self.selected_folder,
                    self.selected_output,
                    recursive=self.recursive_var.get() == "1",
                )
                self.root.after(0, lambda: self._extraction_done(report_path))
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda: self._extraction_failed(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _extraction_done(self, report_path: Path) -> None:
        self._set_busy(False, f"Done. Report saved: {report_path}")
        if messagebox.askyesno(APP_TITLE, f"Report generated:\n{report_path}\n\nOpen it now?"):
            self._open_file(report_path)

    def _extraction_failed(self, exc: Exception) -> None:
        self._set_busy(False, "Extraction failed. See log for details.")
        messagebox.showerror(APP_TITLE, f"Extraction failed:\n{exc}")

    def _toggle_service(self) -> None:
        if not self.watching:
            if not self.selected_folder:
                messagebox.showwarning(APP_TITLE, "Please select a source folder first.")
                return

            def on_change():
                process_folder(
                    self.selected_folder,
                    self.selected_output,
                    recursive=self.recursive_var.get() == "1",
                )

            self.watcher = FolderWatcher(self.selected_folder, on_change)
            self.watcher.start()
            self.watching = True
            self.service_btn.configure(text="Stop background service")
            self.status_var.set(f"Watching folder for new PDFs: {self.selected_folder}")
            threading.Thread(target=on_change, daemon=True).start()  # process what's already there
        else:
            if self.watcher:
                self.watcher.stop()
            self.watching = False
            self.service_btn.configure(text="Start background service")
            self.status_var.set("Background service stopped.")

    def _set_busy(self, busy: bool, message: str) -> None:
        self.status_var.set(message)
        if busy:
            self.extract_btn.configure(state="disabled")
            self.progress.start(12)
        else:
            self.extract_btn.configure(state="normal")
            self.progress.stop()

    @staticmethod
    def _open_file(path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["start", "", str(path)], shell=True)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass


def main() -> None:
    root = Tk()
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    PdfExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
