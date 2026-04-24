"""Parallel HTTP downloader with retries, proxy support and SleepSec throttling."""
from __future__ import annotations

import logging
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List

import requests

from . import models
from .models import DownloadRow, JobContext, Settings
from .utils import (
    ensure_dir,
    filename_from_content_disposition,
    filename_from_url,
    item_folder_name,
    now_iso,
)

logger = logging.getLogger(__name__)


def _build_proxies(proxy_url: str) -> dict | None:
    if not proxy_url:
        return None
    return {"http": proxy_url, "https": proxy_url}


def _pick_extension(source_type: str, url: str, content_type: str | None) -> str:
    """Decide an extension for the saved file when we have to invent one."""
    url_lower = url.lower().split("?")[0]
    known_exts = (
        ".zip", ".docx", ".doc", ".docm", ".dot", ".dotx", ".dotm",
        ".pptx", ".ppt", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot",
        ".potx", ".potm", ".pdf", ".xls", ".xlsx", ".xlsm", ".csv",
        ".txt", ".rtf", ".odt", ".odp",
    )
    for ext in known_exts:
        if url_lower.endswith(ext):
            return ext
    if content_type:
        ct = content_type.lower()
        if "pdf" in ct:
            return ".pdf"
        if "zip" in ct:
            return ".zip"
        if "word" in ct or "msword" in ct or "officedocument.wordprocessingml" in ct:
            return ".docx"
        if "presentation" in ct or "powerpoint" in ct:
            return ".pptx"
        if "sheet" in ct or "excel" in ct:
            return ".xlsx"
    # fall back by source-type convention
    return ".zip" if source_type == models.SOURCE_3GPP else ".bin"


def _resolve_save_path(
    job: JobContext,
    row: DownloadRow,
    response: requests.Response,
) -> Path:
    folder = job.raw_dir / item_folder_name(row.seq, row.title)
    ensure_dir(folder)

    # Prefer: Content-Disposition -> URL path -> synthesized
    cd_name = filename_from_content_disposition(response.headers.get("Content-Disposition"))
    if cd_name:
        return folder / cd_name

    url_name = filename_from_url(row.url, fallback="")
    if url_name and "." in url_name:
        return folder / url_name

    ext = _pick_extension(job.source_type, row.url, response.headers.get("Content-Type"))
    stem = "original"
    return folder / f"{stem}{ext}"


def _download_one(
    row: DownloadRow,
    job: JobContext,
    settings: Settings,
    lock: threading.Lock,
) -> None:
    if not row.url:
        row.status = models.STATUS_ERROR
        row.message = "C列にハイパーリンクURLがありません"
        row.last_run_at = now_iso()
        return

    proxies = _build_proxies(settings.proxy_url)
    # Use browser-like headers to avoid bot detection (e.g. IEEE Mentor returns 418).
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
    }
    last_err: str = ""

    for attempt in range(1, settings.retry_count + 1):
        try:
            with requests.get(
                row.url,
                headers=headers,
                proxies=proxies,
                timeout=settings.timeout_sec,
                stream=True,
                allow_redirects=True,
            ) as resp:
                if resp.status_code >= 400:
                    last_err = f"HTTP {resp.status_code} {resp.reason}"
                    # 4xx/5xx: let retry loop decide whether to try again.
                    if resp.status_code in (401, 403, 404, 410):
                        break
                    raise requests.HTTPError(last_err)
                save_path = _resolve_save_path(job, row, resp)
                if save_path.exists() and not settings.overwrite_existing:
                    row.raw_path = save_path
                    row.saved_path = str(save_path)
                    row.status = models.STATUS_SKIPPED
                    row.message = "既存ファイルを再利用しました"
                    row.last_run_at = now_iso()
                    return

                tmp_path = save_path.with_suffix(save_path.suffix + ".part")
                with open(tmp_path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=1024 * 256):
                        if chunk:
                            f.write(chunk)
                os.replace(tmp_path, save_path)

                row.raw_path = save_path
                row.saved_path = str(save_path)
                row.status = models.STATUS_DONE  # tentative; PDF/HTML stages may override
                row.message = "downloaded"
                row.last_run_at = now_iso()
                return
        except Exception as exc:  # noqa: BLE001
            last_err = f"{type(exc).__name__}: {exc}"
            logger.warning("Download attempt %d failed for row %d: %s", attempt, row.row_no, last_err)
        finally:
            # SleepSec applies between requests (including retries).
            if settings.sleep_sec > 0:
                with lock:
                    time.sleep(settings.sleep_sec)

    row.status = models.STATUS_ERROR
    row.message = last_err or "download failed"
    row.last_run_at = now_iso()


def download_all(rows: List[DownloadRow], job: JobContext, settings: Settings) -> None:
    """Download every row in-place, respecting DownloadWorkers."""
    if not rows:
        return
    ensure_dir(job.raw_dir)
    throttle = threading.Lock()
    workers = max(1, settings.download_workers)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_download_one, r, job, settings, throttle) for r in rows]
        for _ in as_completed(futures):
            pass

    n_ok = sum(1 for r in rows if r.status in (models.STATUS_DONE, models.STATUS_SKIPPED))
    logger.info("Download finished: %d/%d succeeded", n_ok, len(rows))
