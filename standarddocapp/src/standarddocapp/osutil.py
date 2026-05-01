"""Cross-platform 'open file/folder in shell' helpers."""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def open_in_shell(path: Path) -> None:
    """Open a file or folder using the OS default handler."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(p)
    if os.name == "nt":
        os.startfile(str(p))  # type: ignore[attr-defined]
        return
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(p)])
        return
    subprocess.Popen(["xdg-open", str(p)])


def reveal_in_shell(path: Path) -> None:
    """Reveal a file in its parent folder. Falls back to opening it."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(p)
    if os.name == "nt":
        if p.is_file():
            subprocess.Popen(["explorer.exe", "/select,", str(p)])
            return
        os.startfile(str(p))  # type: ignore[attr-defined]
        return
    open_in_shell(p if p.is_dir() else p.parent)
