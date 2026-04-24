"""3GPP ZIP unpacking with Zip-Slip protection and parallel execution."""
from __future__ import annotations

import logging
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List

from . import models
from .models import DownloadRow, JobContext, Settings
from .utils import ensure_dir, item_folder_name, now_iso, safe_filename

logger = logging.getLogger(__name__)


def _safe_extract(zf: zipfile.ZipFile, target_dir: Path) -> List[Path]:
    """Extract members, flattening unsafe paths; returns extracted files."""
    produced: List[Path] = []
    for info in zf.infolist():
        if info.is_dir():
            continue
        raw_name = info.filename.replace("\\", "/")
        parts = [safe_filename(p) for p in raw_name.split("/") if p not in ("", ".", "..")]
        if not parts:
            continue
        dest = target_dir.joinpath(*parts)
        # Extra paranoia against Zip-Slip even though we already sanitised.
        try:
            dest_resolved = dest.resolve()
            target_resolved = target_dir.resolve()
            dest_resolved.relative_to(target_resolved)
        except ValueError:
            logger.warning("Skipping unsafe zip entry: %s", info.filename)
            continue
        ensure_dir(dest.parent)
        with zf.open(info) as src, open(dest, "wb") as dst:
            while True:
                chunk = src.read(1024 * 256)
                if not chunk:
                    break
                dst.write(chunk)
        produced.append(dest)
    return produced


def _unzip_row(row: DownloadRow, job: JobContext) -> None:
    if row.status != models.STATUS_DONE or row.raw_path is None:
        return
    if row.raw_path.suffix.lower() != ".zip":
        # Not a zip - treat the raw file as the single produced file later.
        return
    target = job.unpacked_dir / item_folder_name(row.seq, row.title)
    ensure_dir(target)
    try:
        with zipfile.ZipFile(row.raw_path, "r") as zf:
            produced = _safe_extract(zf, target)
    except zipfile.BadZipFile as exc:
        row.status = models.STATUS_ERROR
        row.message = f"zip open failed: {exc}"
        row.last_run_at = now_iso()
        return
    if not produced:
        row.status = models.STATUS_ERROR
        row.message = "zip was empty"
        row.last_run_at = now_iso()
        return
    # Stash produced filesystem paths on the row via a conventional attribute.
    row._unzipped_files = sorted(produced)  # type: ignore[attr-defined]
    row.message = f"downloaded, unzipped {len(produced)} files"


def unzip_all(rows: List[DownloadRow], job: JobContext, settings: Settings) -> None:
    """Parallel ZIP unpacking for 3GPP rows."""
    if job.source_type != models.SOURCE_3GPP:
        return
    if not rows:
        return
    ensure_dir(job.unpacked_dir)
    workers = max(1, settings.unzip_workers)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_unzip_row, r, job) for r in rows]
        for _ in as_completed(futures):
            pass
