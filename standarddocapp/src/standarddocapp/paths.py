"""Paths used by the GUI app (log directory, etc.)."""
from __future__ import annotations

import os
from pathlib import Path


def app_log_dir() -> Path:
    """Return the application log directory (created on demand).

    Uses %LOCALAPPDATA%\\StandardDocApp\\logs on Windows; falls back to
    ~/.standarddocapp/logs elsewhere.
    """
    if os.name == "nt":
        base_env = os.environ.get("LOCALAPPDATA")
        base = Path(base_env) if base_env else Path.home() / "AppData" / "Local"
        out = base / "StandardDocApp" / "logs"
    else:
        out = Path.home() / ".standarddocapp" / "logs"
    out.mkdir(parents=True, exist_ok=True)
    return out
