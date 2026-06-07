"""HTML integrity checker for stdharvest-generated job folders.

Verifies, for every HTML file produced by stdharvest:

* The file is well-formed enough to parse (BeautifulSoup html.parser).
* The structural skeleton exists (``<html>``, ``<head>``, ``<title>``,
  ``<body>``).
* Every ``<a href>`` / ``<img src>`` / ``<link href>`` / ``<script src>``
  pointing to a relative path resolves to a file that actually exists on
  disk.
* Every in-page anchor (``href="#id"``) has a matching ``id=`` somewhere
  in the same document.

The checker is read-only and never deletes or rewrites HTML.

It also exposes :func:`scan_job_folder` which walks a stdharvest job folder
and returns the categorized list of files to validate.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Tuple
from urllib.parse import unquote, urlsplit

logger = logging.getLogger(__name__)


# ----------------------------------------------------------------------- types

@dataclass
class FileReport:
    path: Path
    role: str  # "index" | "individual" | "combined" | "combine_full" | "other"
    parsed: bool = False
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    info: List[str] = field(default_factory=list)
    # link stats
    links_total: int = 0
    links_broken: int = 0
    anchors_missing: int = 0

    @property
    def ok(self) -> bool:
        return not self.errors


@dataclass
class JobReport:
    job_folder: Path
    files: List[FileReport] = field(default_factory=list)

    def summary(self) -> dict:
        total = len(self.files)
        ok = sum(1 for f in self.files if f.ok)
        with_warn = sum(1 for f in self.files if f.warnings and f.ok)
        broken = sum(1 for f in self.files if not f.ok)
        return {
            "total": total,
            "ok": ok,
            "ok_with_warnings": with_warn,
            "broken": broken,
            "links_broken": sum(f.links_broken for f in self.files),
            "anchors_missing": sum(f.anchors_missing for f in self.files),
        }


# ----------------------------------------------------------------------- I/O

def scan_job_folder(job_folder: Path) -> List[Tuple[Path, str]]:
    """Return (path, role) pairs for every HTML file we expect in a job folder.

    The folder layout follows stdharvest's spec: ``html/index.html``,
    ``html/files/<item>/index.html``, ``html/combined/*.html``,
    ``html/combine_full/*.html``.  Anything else under ``html/`` is reported
    as ``"other"``.
    """
    job_folder = Path(job_folder)
    html_root = job_folder / "html"
    if not html_root.exists():
        return []
    pairs: List[Tuple[Path, str]] = []

    index_path = html_root / "index.html"
    if index_path.exists():
        pairs.append((index_path, "index"))

    files_dir = html_root / "files"
    if files_dir.exists():
        for p in sorted(files_dir.rglob("*.html")):
            pairs.append((p, "individual"))

    combined_dir = html_root / "combined"
    if combined_dir.exists():
        for p in sorted(combined_dir.glob("*.html")):
            pairs.append((p, "combined"))

    combine_full_dir = html_root / "combine_full"
    if combine_full_dir.exists():
        for p in sorted(combine_full_dir.glob("*.html")):
            pairs.append((p, "combine_full"))

    seen = {p for p, _ in pairs}
    for p in sorted(html_root.rglob("*.html")):
        if p not in seen:
            pairs.append((p, "other"))

    return pairs


# ----------------------------------------------------------------------- core

_DOCTYPE_RE = re.compile(r"<!doctype\s+html", re.IGNORECASE)


def _resolve_local_target(href: str, base: Path) -> Optional[Path]:
    """Resolve a relative URL, ``file://`` URL or absolute disk path to a Path.

    Returns None for external URLs (``http://``, ``mailto:``, ``data:`` …),
    in-page anchors and empty strings.
    """
    if not href:
        return None
    if href.startswith("#"):
        return None
    parsed = urlsplit(href)
    if parsed.scheme in ("http", "https", "ftp", "mailto", "javascript", "data", "about"):
        return None
    if parsed.scheme == "file":
        # file:///c:/foo/bar.html  -> /c:/foo/bar.html (Windows-friendly)
        path_part = unquote(parsed.path)
        if path_part.startswith("/") and len(path_part) > 2 and path_part[2] == ":":
            path_part = path_part.lstrip("/")
        candidate = Path(path_part)
        return candidate
    raw_path = unquote(parsed.path)
    if not raw_path:
        return None
    candidate = Path(raw_path)
    if not candidate.is_absolute():
        candidate = (base.parent / candidate).resolve()
    return candidate


def _validate_file(
    path: Path,
    role: str,
    job_folder: Path,
) -> FileReport:
    report = FileReport(path=path, role=role)
    try:
        raw = path.read_text(encoding="utf-8", errors="strict")
    except UnicodeDecodeError as exc:
        report.errors.append(f"not valid UTF-8: {exc}")
        try:
            raw = path.read_text(encoding="utf-8", errors="replace")
        except Exception as exc2:  # noqa: BLE001
            report.errors.append(f"could not read file at all: {exc2}")
            return report
    except Exception as exc:  # noqa: BLE001
        report.errors.append(f"read failed: {exc}")
        return report

    if not raw.strip():
        report.errors.append("file is empty")
        return report

    if not _DOCTYPE_RE.search(raw[:512]):
        report.warnings.append("missing <!doctype html>")

    try:
        from bs4 import BeautifulSoup  # lazy import
    except ImportError:
        report.warnings.append(
            "beautifulsoup4 not installed; skipping structural validation"
        )
        return report

    try:
        soup = BeautifulSoup(raw, "html.parser")
    except Exception as exc:  # noqa: BLE001
        report.errors.append(f"parse failed: {exc}")
        return report
    report.parsed = True

    if soup.find("html") is None:
        report.warnings.append("no <html> root element")
    if soup.find("head") is None:
        report.warnings.append("no <head> element")
    title = soup.find("title")
    if title is None or not (title.string or "").strip():
        report.warnings.append("missing or empty <title>")
    body = soup.find("body")
    if body is None:
        report.errors.append("missing <body> element")
        return report

    # Collect every ``id`` defined in the document so anchor checks can
    # confirm a target exists.
    document_ids = {
        t.get("id") for t in soup.find_all(True) if t.has_attr("id")
    }
    document_ids.discard(None)

    link_attrs = (
        ("a", "href"),
        ("link", "href"),
        ("area", "href"),
        ("img", "src"),
        ("script", "src"),
        ("iframe", "src"),
        ("source", "src"),
        ("video", "poster"),
    )

    for tag_name, attr in link_attrs:
        for tag in soup.find_all(tag_name):
            val = tag.get(attr)
            if not val or not isinstance(val, str):
                continue
            report.links_total += 1
            if val.startswith("#"):
                target_id = val[1:]
                if target_id and target_id not in document_ids:
                    report.anchors_missing += 1
                    report.warnings.append(f"dangling anchor: {val}")
                continue
            target = _resolve_local_target(val, path)
            if target is None:
                continue
            if not target.exists():
                report.links_broken += 1
                report.errors.append(
                    f"broken link from <{tag_name} {attr}>: {val} -> {target}"
                )

    if role == "index":
        if soup.find("h1") is None:
            report.warnings.append("index.html has no <h1>")
        if not soup.find_all("a"):
            report.warnings.append("index.html has no links at all")

    if role in ("combined", "combine_full"):
        if not soup.find_all(["article", "section"]):
            report.warnings.append(
                f"{role} page has no <article>/<section> blocks"
            )

    report.info.append(
        f"links: total={report.links_total} broken={report.links_broken} "
        f"anchor_missing={report.anchors_missing}"
    )
    return report


def validate_job_html(
    job_folder: Path,
    progress: Optional[Callable[[int, int, FileReport], None]] = None,
) -> JobReport:
    """Validate every HTML file in ``job_folder``.

    Pass ``progress`` to receive ``(done, total, file_report)`` after each
    file is processed (handy to drive a Tk progress bar from a worker
    thread).
    """
    job_folder = Path(job_folder)
    pairs = scan_job_folder(job_folder)
    report = JobReport(job_folder=job_folder)
    total = len(pairs)
    for i, (path, role) in enumerate(pairs, start=1):
        try:
            fr = _validate_file(path, role, job_folder)
        except Exception as exc:  # noqa: BLE001
            fr = FileReport(path=path, role=role)
            fr.errors.append(f"validator crashed: {exc}")
        report.files.append(fr)
        if progress is not None:
            try:
                progress(i, total, fr)
            except Exception:
                pass
    return report


# ------------------------------------------------------------ rendering helpers

def format_report_lines(report: JobReport) -> Iterable[str]:
    """Yield human-readable text lines for a JobReport (for log panels)."""
    yield f"Job folder: {report.job_folder}"
    s = report.summary()
    yield (
        f"Files: {s['total']}  ok: {s['ok']}  ok+warn: {s['ok_with_warnings']}  "
        f"broken: {s['broken']}  broken-links: {s['links_broken']}  "
        f"missing-anchors: {s['anchors_missing']}"
    )
    yield ""
    for fr in report.files:
        rel = _safe_rel(fr.path, report.job_folder)
        if fr.errors:
            tag = "BROKEN"
        elif fr.warnings:
            tag = " WARN "
        else:
            tag = "  OK  "
        yield f"[{tag}] ({fr.role}) {rel}"
        for msg in fr.errors:
            yield f"        ERROR  : {msg}"
        for msg in fr.warnings:
            yield f"        warning: {msg}"
        for msg in fr.info:
            yield f"        info   : {msg}"


def _safe_rel(path: Path, base: Path) -> str:
    try:
        return str(path.relative_to(base))
    except ValueError:
        return str(path)
