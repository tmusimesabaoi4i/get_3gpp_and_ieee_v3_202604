"""Small helpers: paths, datetime, relative links, HTML escape."""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def now_compact_date() -> str:
    return datetime.now().strftime("%Y%m%d")


def ensure_dir(p: Path) -> Path:
    p.mkdir(parents=True, exist_ok=True)
    return p


def rel_link(from_file: Path, target: Path | str) -> str:
    """POSIX relative URL from an HTML file to a target path."""
    target_p = Path(target) if not isinstance(target, Path) else target
    try:
        rel = os.path.relpath(target_p, start=from_file.parent)
    except ValueError:
        return target_p.resolve().as_uri()
    return rel.replace(os.sep, "/")
