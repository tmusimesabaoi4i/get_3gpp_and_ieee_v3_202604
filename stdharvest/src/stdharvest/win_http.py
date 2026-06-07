"""Windows-only HTTP GET via ``WinHttp.WinHttpRequest.5.1``.

This replicates what a browser (or any WinHTTP/WinINET based app) does on a
corporate network: it authenticates to the proxy using the **currently
logged-in Windows account** (NTLM / Negotiate single sign-on), so the user
does not have to store a proxy username/password anywhere.

We use this as the default download path when a proxy is in play and no
explicit proxy credentials were supplied, which is exactly the situation where
``requests`` returns ``407 Proxy Authentication Required`` (it has no built-in
integrated-auth support for proxies).
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Dict, Optional

logger = logging.getLogger(__name__)

# WinHttpRequestAutoLogonPolicy: 0 = Always send default credentials (even
# through a proxy). The COM default is "OnlyIfBypassProxy", which is why a
# plain proxied request fails with 407.
_AUTOLOGON_ALWAYS = 0

# HTTPREQUEST_PROXY_SETTING
_PROXY_SET = 2  # use the explicitly provided proxy "host:port"


@dataclass
class WinHttpResponse:
    status_code: int
    reason: str
    headers: Dict[str, str]  # keys lower-cased
    body: bytes

    def header(self, name: str) -> str:
        return self.headers.get(name.lower(), "")


def is_available() -> bool:
    """Return True if the WinHttpRequest COM object can be used."""
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def _parse_headers(raw: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for line in (raw or "").split("\r\n"):
        if ":" in line:
            key, _, value = line.partition(":")
            out[key.strip().lower()] = value.strip()
    return out


def get(
    url: str,
    *,
    proxy_hostport: str = "",
    timeout_sec: int = 60,
    headers: Optional[Dict[str, str]] = None,
) -> WinHttpResponse:
    """Perform a synchronous GET via WinHttpRequest.

    Raises on transport-level errors (DNS, connect, COM); HTTP status codes
    (including 4xx/5xx) are returned in the response so the caller can decide
    whether to retry.
    """
    import pythoncom
    import win32com.client

    # Each worker thread needs its own COM apartment.
    pythoncom.CoInitialize()
    try:
        # dynamic.Dispatch avoids gen_py/makepy caching, which is more robust
        # across threads and inside a PyInstaller-frozen exe.
        http = win32com.client.dynamic.Dispatch("WinHttp.WinHttpRequest.5.1")
        ms = max(1, int(timeout_sec)) * 1000
        http.SetTimeouts(ms, ms, ms, ms)
        if proxy_hostport:
            http.SetProxy(_PROXY_SET, proxy_hostport)
        http.SetAutoLogonPolicy(_AUTOLOGON_ALWAYS)
        http.Open("GET", url, False)
        for key, value in (headers or {}).items():
            try:
                http.SetRequestHeader(key, value)
            except Exception:  # noqa: BLE001
                pass
        http.Send()
        status = int(http.Status)
        reason = str(http.StatusText)
        parsed = _parse_headers(str(http.GetAllResponseHeaders()))
        body = b""
        if status < 400:
            raw = http.ResponseBody
            body = bytes(raw) if raw is not None else b""
        return WinHttpResponse(status, reason, parsed, body)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
