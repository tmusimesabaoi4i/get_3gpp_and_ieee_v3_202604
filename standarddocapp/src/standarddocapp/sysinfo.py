"""Environment / runtime information for the About / settings tab."""
from __future__ import annotations

import os
import platform
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


@dataclass
class EnvInfo:
    app_name: str
    app_version: str
    python_version: str
    python_executable: str
    os_summary: str
    is_windows: bool
    office_available: bool
    office_detail: str
    soffice_path: Optional[str]
    soffice_detail: str
    log_dir: Path
    sample_paths: List[Path] = field(default_factory=list)
    repo_root: Optional[Path] = None
    readme_path: Optional[Path] = None


def _is_windows() -> bool:
    return os.name == "nt"


def _detect_office() -> tuple[bool, str]:
    """Detect Microsoft Office COM availability.

    Reuses the same probing strategy as stdharvest.pdf_converter._check_office_available
    but stays independent so the GUI can show diagnostics even when stdharvest fails to
    import (e.g. missing extras during a partial install).
    """
    if not _is_windows():
        return False, "Microsoft Office COM is only available on Windows."
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        return False, f"pywin32 not importable: {exc}"
    try:
        import win32com.client as _wc  # type: ignore
        for prog_id in ("Word.Application", "Excel.Application", "PowerPoint.Application"):
            try:
                app = _wc.Dispatch(prog_id)
                try:
                    app.Quit()
                except Exception:
                    pass
                return True, f"Detected {prog_id}"
            except Exception:
                continue
        return False, "win32com loaded but no Word/Excel/PowerPoint COM server responded."
    except Exception as exc:  # noqa: BLE001
        return False, f"COM probe failed: {exc}"


def _detect_soffice() -> tuple[Optional[str], str]:
    env = os.environ.get("STDHARVEST_SOFFICE")
    if env and Path(env).exists():
        return env, f"From STDHARVEST_SOFFICE: {env}"
    for candidate in ("soffice", "soffice.exe"):
        found = shutil.which(candidate)
        if found:
            return found, f"On PATH: {found}"
    if _is_windows():
        for pf in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(pf).exists():
                return pf, f"Detected installation: {pf}"
    return None, "LibreOffice (soffice) not found."


def _repo_root_guess() -> Optional[Path]:
    """Guess the repository root for sample file lookups.

    Walks upward from this file looking for stdharvest/samples or stdsearch/samples,
    so it works when running as a frozen exe living next to the source tree, when
    running via `python -m standarddocapp`, or when the package is installed.
    """
    here = Path(__file__).resolve()
    for ancestor in [here, *here.parents]:
        candidate = ancestor.parent
        if (candidate / "stdharvest" / "samples").exists():
            return candidate
        if (candidate / "stdsearch" / "samples").exists():
            return candidate
    cwd = Path.cwd()
    if (cwd / "stdharvest" / "samples").exists() or (cwd / "stdsearch" / "samples").exists():
        return cwd
    return None


def _sample_paths(repo_root: Optional[Path]) -> List[Path]:
    out: List[Path] = []
    if not repo_root:
        return out
    for rel in (
        Path("stdharvest") / "samples" / "sample_download.xlsx",
        Path("stdsearch") / "samples" / "sample_search.xlsx",
    ):
        p = repo_root / rel
        if p.exists():
            out.append(p)
    return out


def _readme_path(repo_root: Optional[Path]) -> Optional[Path]:
    if not repo_root:
        return None
    candidate = repo_root / "README.md"
    return candidate if candidate.exists() else None


def collect_env_info(app_name: str, app_version: str, log_dir: Path) -> EnvInfo:
    office_ok, office_detail = _detect_office()
    soffice, soffice_detail = _detect_soffice()
    repo_root = _repo_root_guess()
    return EnvInfo(
        app_name=app_name,
        app_version=app_version,
        python_version=sys.version.replace("\n", " "),
        python_executable=sys.executable,
        os_summary=f"{platform.system()} {platform.release()} ({platform.version()})",
        is_windows=_is_windows(),
        office_available=office_ok,
        office_detail=office_detail,
        soffice_path=soffice,
        soffice_detail=soffice_detail,
        log_dir=log_dir,
        sample_paths=_sample_paths(repo_root),
        repo_root=repo_root,
        readme_path=_readme_path(repo_root),
    )
