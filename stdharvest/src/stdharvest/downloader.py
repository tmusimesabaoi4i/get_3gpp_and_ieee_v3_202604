"""Parallel HTTP downloader with retries, proxy support and SleepSec throttling."""
from __future__ import annotations

import logging
import os
import re
import subprocess
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
from pathlib import Path
from typing import List
from urllib.parse import quote, urlsplit, urlunsplit

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


def normalize_proxy_url(proxy_url: str, user: str = "", password: str = "") -> str:
    """Return a requests-ready proxy URL with scheme and optional credentials.

    * Adds an ``http://`` scheme when the user only typed ``host:port``
      (requests/urllib3 reject a scheme-less proxy).
    * Injects ``user``/``password`` (URL-encoded so that domain backslashes,
      ``@`` or other special characters are handled) into the URL, replacing
      any credentials already present in ``proxy_url``.
    """
    proxy_url = (proxy_url or "").strip()
    if not proxy_url:
        return ""
    if "://" not in proxy_url:
        proxy_url = "http://" + proxy_url

    user = (user or "").strip()
    if user:
        parts = urlsplit(proxy_url)
        # Drop any credentials already embedded in the netloc.
        hostport = parts.netloc.rsplit("@", 1)[-1]
        userinfo = quote(user, safe="")
        if password:
            userinfo += ":" + quote(password, safe="")
        netloc = f"{userinfo}@{hostport}"
        proxy_url = urlunsplit((parts.scheme, netloc, parts.path, parts.query, parts.fragment))
    return proxy_url


# host:port (optionally prefixed by "http=" / "https=" as printed by netsh).
_PROXY_TOKEN_RE = re.compile(r"(?:(https?)\s*=\s*)?([A-Za-z0-9][A-Za-z0-9.\-]*:\d{1,5})")


def _pick_proxy_token(text: str) -> str:
    """Extract a ``host:port`` proxy from free-form text (netsh / registry).

    Locale-independent: scans for ``host:port`` tokens and prefers the entry
    explicitly tagged for ``https`` (then ``http``, then the first match).
    """
    https_val = http_val = first_val = ""
    for scheme, hostport in _PROXY_TOKEN_RE.findall(text or ""):
        if not first_val:
            first_val = hostport
        if scheme == "https" and not https_val:
            https_val = hostport
        elif scheme == "http" and not http_val:
            http_val = hostport
    return https_val or http_val or first_val


def _detect_proxy_via_netsh() -> str:
    """Return the WinHTTP proxy (``netsh winhttp show proxy``) or ""."""
    try:
        proc = subprocess.run(
            ["netsh", "winhttp", "show", "proxy"],
            capture_output=True,  # bytes: netsh uses the OEM codepage, which the
            timeout=10,           # locale codec (e.g. cp932) often cannot decode.
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:  # noqa: BLE001
        return ""
    raw = proc.stdout or b""
    # host:port tokens are ASCII, so a tolerant decode is enough to parse them.
    text = raw.decode("utf-8", "replace") if isinstance(raw, (bytes, bytearray)) else raw
    return _pick_proxy_token(text)


def _detect_proxy_via_wininet() -> str:
    """Return the per-user IE/WinINET proxy from the registry, or "".

    Corporate proxies are frequently configured only in IE/WinINET (not in
    WinHTTP), so this is a useful fallback when ``netsh`` reports direct access.
    """
    try:
        import winreg  # noqa: PLC0415 (Windows-only, imported lazily)
    except Exception:  # noqa: BLE001
        return ""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        )
        try:
            enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
            if not enable:
                return ""
            server, _ = winreg.QueryValueEx(key, "ProxyServer")
        finally:
            winreg.CloseKey(key)
    except Exception:  # noqa: BLE001
        return ""
    return _pick_proxy_token(str(server or ""))


@lru_cache(maxsize=1)
def detect_system_proxy() -> str:
    """Best-effort auto-detection of the system proxy on Windows.

    Tries ``netsh winhttp show proxy`` first, then the IE/WinINET registry
    setting. Returns a ``host:port`` string (no scheme) or "" when none is
    configured / not on Windows. Result is cached for the process lifetime.
    """
    if os.name != "nt":
        return ""
    return _detect_proxy_via_netsh() or _detect_proxy_via_wininet()


def _build_proxies(settings: Settings) -> dict | None:
    proxy_url = (settings.proxy_url or "").strip()
    if not proxy_url:
        detected = detect_system_proxy()
        if detected:
            logger.info(
                "ProxyURL未指定: システムのプロキシ設定を自動検出して使用します: %s",
                detected,
            )
            proxy_url = detected
    url = normalize_proxy_url(proxy_url, settings.proxy_user, settings.proxy_password)
    if not url:
        return None
    return {"http": url, "https": url}


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
    proxies: dict | None,
) -> None:
    if not row.url:
        row.status = models.STATUS_ERROR
        row.message = "C列にハイパーリンクURLがありません"
        row.last_run_at = now_iso()
        return

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
        except requests.exceptions.ProxyError as exc:
            # A 407 means the proxy itself rejected us: retrying with the same
            # (missing/invalid) credentials cannot help, so stop immediately and
            # surface an actionable message instead of spamming the log.
            if "407" in str(exc):
                if settings.proxy_user:
                    last_err = (
                        "プロキシ認証に失敗しました (407)。Sheet2 の ProxyUser / "
                        "ProxyPassword のユーザー名・パスワードを確認してください。"
                    )
                else:
                    last_err = (
                        "プロキシ認証が必要です (407)。Sheet2 の ProxyUser / ProxyPassword に "
                        "プロキシ用のユーザー名・パスワードを設定してください。"
                    )
                logger.error("Row %d: %s", row.row_no, last_err)
                break
            last_err = f"ProxyError: {exc}"
            logger.warning("Download attempt %d failed for row %d: %s", attempt, row.row_no, last_err)
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
    proxies = _build_proxies(settings)  # detect/normalize once for the whole run
    throttle = threading.Lock()
    workers = max(1, settings.download_workers)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_download_one, r, job, settings, throttle, proxies) for r in rows]
        for _ in as_completed(futures):
            pass

    n_ok = sum(1 for r in rows if r.status in (models.STATUS_DONE, models.STATUS_SKIPPED))
    logger.info("Download finished: %d/%d succeeded", n_ok, len(rows))
