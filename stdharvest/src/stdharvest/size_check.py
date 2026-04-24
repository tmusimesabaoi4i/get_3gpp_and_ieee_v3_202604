"""File-size judgement (Min/Max) applied to raw downloads and unzipped files."""
from __future__ import annotations

from pathlib import Path
from typing import Tuple

from . import models
from .models import Settings
from .utils import file_size, human_size


def classify(path: Path, settings: Settings) -> Tuple[str, int]:
    """Return (status, size_bytes) where status is OK / TOO_SMALL / TOO_LARGE."""
    size = file_size(path)
    min_bytes = settings.min_file_size_kb * 1024
    max_bytes = settings.max_file_size_mb * 1024 * 1024
    if size < min_bytes:
        return models.SIZE_TOO_SMALL, size
    if size > max_bytes:
        return models.SIZE_TOO_LARGE, size
    return models.SIZE_OK, size


def describe_limits(settings: Settings) -> str:
    return (
        f"MinFileSizeKB={settings.min_file_size_kb}, "
        f"MaxFileSizeMB={settings.max_file_size_mb}"
    )


def size_message(path: Path, status: str, size: int, settings: Settings) -> str:
    if status == models.SIZE_TOO_SMALL:
        return (
            f"file size {human_size(size)} is smaller than "
            f"MinFileSizeKB={settings.min_file_size_kb}"
        )
    if status == models.SIZE_TOO_LARGE:
        return (
            f"file size {human_size(size)} is larger than "
            f"MaxFileSizeMB={settings.max_file_size_mb}"
        )
    return f"size {human_size(size)} ({path.name})"
