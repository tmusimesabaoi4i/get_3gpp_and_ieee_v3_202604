"""Small helpers: filenames, paths, datetime, hashing, links."""
from __future__ import annotations

import hashlib
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote, urlsplit


_INVALID_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WS_COLLAPSE = re.compile(r"\s+")


def now_iso() -> str:
    """Timestamp used in Excel LastRunAt and logs."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def now_compact_date() -> str:
    return datetime.now().strftime("%Y%m%d")


def safe_filename(name: str, max_len: int = 120) -> str:
    """Return a filesystem-safe file or folder name.

    - Replaces characters forbidden on Windows.
    - Collapses whitespace.
    - Trims trailing dots/spaces (Windows rejects them).
    - Falls back to a hash when empty.
    """
    if not name:
        return "_"
    cleaned = _INVALID_FS_CHARS.sub("_", name)
    cleaned = _WS_COLLAPSE.sub(" ", cleaned).strip()
    cleaned = cleaned.rstrip(". ")
    if not cleaned:
        cleaned = hashlib.md5(name.encode("utf-8", "ignore")).hexdigest()[:10]
    if len(cleaned) > max_len:
        digest = hashlib.md5(cleaned.encode("utf-8", "ignore")).hexdigest()[:8]
        cleaned = cleaned[: max_len - 9].rstrip(". ") + "_" + digest
    return cleaned


def item_folder_name(seq: int, title: str) -> str:
    """Item folder like 001_R1-2409888_Xiaomi."""
    return f"{seq:03d}_{safe_filename(title, max_len=100)}"


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def human_size(size_bytes: int) -> str:
    """Human friendly size string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    kb = size_bytes / 1024.0
    if kb < 1024:
        return f"{kb:.1f} KB"
    mb = kb / 1024.0
    if mb < 1024:
        return f"{mb:.1f} MB"
    return f"{mb / 1024.0:.2f} GB"


def filename_from_url(url: str, fallback: str = "download.bin") -> str:
    """Best-effort file name extraction from a URL (for saving)."""
    try:
        path = urlsplit(url).path
        name = os.path.basename(unquote(path))
        name = safe_filename(name)
        return name or fallback
    except Exception:
        return fallback


def filename_from_content_disposition(header: str | None) -> str | None:
    """Parse filename from Content-Disposition response header, if present."""
    if not header:
        return None
    m = re.search(r'filename\*=UTF-8\'\'([^;]+)', header, re.IGNORECASE)
    if m:
        return safe_filename(unquote(m.group(1).strip().strip('"')))
    m = re.search(r'filename="?([^";]+)"?', header, re.IGNORECASE)
    if m:
        return safe_filename(m.group(1).strip())
    return None


def rel_link(from_file: Path, target: Path) -> str:
    """Produce a POSIX relative URL usable inside an HTML file."""
    try:
        rel = os.path.relpath(target, start=from_file.parent)
    except ValueError:
        # Different drive on Windows: fall back to absolute file:// link.
        return target.resolve().as_uri()
    return rel.replace(os.sep, "/")


def iter_files(root: Path) -> Iterable[Path]:
    for p in root.rglob("*"):
        if p.is_file():
            yield p


def is_windows() -> bool:
    return sys.platform.startswith("win")


def file_size(path: Path) -> int:
    try:
        return path.stat().st_size
    except OSError:
        return 0
