"""Kill Office / LibreOffice processes so PDF/HTML conversion stays stable."""
from __future__ import annotations

import logging
import shutil
import subprocess
from typing import List

from .utils import is_windows

logger = logging.getLogger(__name__)

WINDOWS_TARGETS: List[str] = [
    "WINWORD.EXE",
    "POWERPNT.EXE",
    "EXCEL.EXE",
    "soffice.exe",
    "soffice.bin",
    "oosplash.exe",
]

POSIX_TARGETS: List[str] = [
    "soffice.bin",
    "soffice",
    "oosplash",
]


def _taskkill(image_name: str) -> tuple[bool, str]:
    cmd = ["taskkill", "/F", "/T", "/IM", image_name]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except FileNotFoundError:
        return False, "taskkill command not available"
    except subprocess.TimeoutExpired:
        return False, "taskkill timed out"
    out = (proc.stdout + proc.stderr).strip()
    # taskkill returns 128 when process was not found; that's fine.
    return proc.returncode == 0, out


def _pkill(name: str) -> tuple[bool, str]:
    pkill = shutil.which("pkill")
    if not pkill:
        return False, "pkill not available"
    try:
        proc = subprocess.run(
            [pkill, "-f", name],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return False, "pkill timed out"
    return proc.returncode == 0, (proc.stdout + proc.stderr).strip()


def kill_office_apps(enabled: bool) -> List[str]:
    """Terminate Office-related processes. Returns a list of human messages.

    When ``enabled`` is False, this is a no-op. Missing processes are not
    considered errors.
    """
    messages: List[str] = []
    if not enabled:
        messages.append("KillOfficeAppsBeforeRun=no; skipped.")
        return messages

    targets = WINDOWS_TARGETS if is_windows() else POSIX_TARGETS
    for name in targets:
        if is_windows():
            killed, _ = _taskkill(name)
        else:
            killed, _ = _pkill(name)
        if killed:
            messages.append(f"killed {name}")
        else:
            messages.append(f"{name}: not running or already stopped")
    logger.info("Office killer: %s", "; ".join(messages))
    return messages
