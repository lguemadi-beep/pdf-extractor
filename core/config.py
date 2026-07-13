"""
config.py
---------
Persistent configuration so the PDF folder (and output folder) only need
to be set up ONCE — via `--configure` on the CLI, or by using the app —
instead of being re-specified as a command-line argument every time the
background service starts or restarts.

The config file (config.json) lives next to the .exe when running as a
PyInstaller build, or next to the source tree otherwise, so it survives
reboots and re-installs of the Windows service.
"""

from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


CONFIG_PATH = _app_dir() / "config.json"


@dataclass
class AppConfig:
    pdf_folder: str = ""
    output_folder: str = ""
    name_filter: str = "proforma"
    recursive: bool = False


def load_config() -> AppConfig | None:
    if not CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return AppConfig(**{**asdict(AppConfig()), **data})
    except Exception:
        return None


def save_config(cfg: AppConfig) -> Path:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(asdict(cfg), indent=2, ensure_ascii=False), encoding="utf-8")
    return CONFIG_PATH
