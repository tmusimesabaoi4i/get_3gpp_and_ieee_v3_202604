"""CSV / JSON report writers (row_results, file_results, combined_html, manifest)."""
from __future__ import annotations

import csv
import json
import logging
from pathlib import Path
from typing import Iterable, List

from .models import CombinedBatch, DownloadRow, JobContext, ProcessedFile, Settings
from .utils import ensure_dir

logger = logging.getLogger(__name__)


def write_row_results(rows: Iterable[DownloadRow], logs_dir: Path) -> Path:
    ensure_dir(logs_dir)
    path = logs_dir / "row_results.csv"
    fields = ["row_no", "title", "url", "status", "saved_path", "message", "last_run_at"]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(
                {
                    "row_no": r.row_no,
                    "title": r.title,
                    "url": r.url,
                    "status": r.status,
                    "saved_path": r.saved_path,
                    "message": r.message,
                    "last_run_at": r.last_run_at,
                }
            )
    return path


def write_file_results(
    files: Iterable[ProcessedFile],
    logs_dir: Path,
    source_type: str,
) -> Path:
    ensure_dir(logs_dir)
    path = logs_dir / "file_results.csv"
    fields = [
        "seq", "row_no", "title", "source_type", "source_file",
        "file_size_bytes", "file_size_kb", "file_size_mb", "size_status",
        "pdf_path", "html_path", "status", "message",
    ]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for pf in files:
            w.writerow(
                {
                    "seq": pf.seq,
                    "row_no": pf.row.row_no,
                    "title": pf.row.title,
                    "source_type": source_type,
                    "source_file": str(pf.source_file),
                    "file_size_bytes": pf.size_bytes,
                    "file_size_kb": round(pf.size_bytes / 1024, 3) if pf.size_bytes else 0,
                    "file_size_mb": round(pf.size_bytes / (1024 * 1024), 4) if pf.size_bytes else 0,
                    "size_status": pf.size_status,
                    "pdf_path": str(pf.pdf_path) if pf.pdf_path else "",
                    "html_path": str(pf.html_path) if pf.html_path else "",
                    "status": pf.status,
                    "message": pf.message,
                }
            )
    return path


def write_combined_html_csv(batches: Iterable[CombinedBatch], logs_dir: Path) -> Path:
    ensure_dir(logs_dir)
    path = logs_dir / "combined_html.csv"
    fields = ["batch_no", "first_seq", "last_seq", "combined_html_path", "file_count"]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for b in batches:
            w.writerow(
                {
                    "batch_no": b.batch_no,
                    "first_seq": b.first_seq,
                    "last_seq": b.last_seq,
                    "combined_html_path": str(b.combined_html_path),
                    "file_count": b.file_count,
                }
            )
    return path


def write_manifest(
    job: JobContext,
    settings: Settings,
    rows: List[DownloadRow],
    files: List[ProcessedFile],
    batches: List[CombinedBatch],
    summary: dict,
    index_html: Path,
) -> Path:
    ensure_dir(job.logs_dir)
    path = job.logs_dir / "manifest.json"
    manifest = {
        "job_name": job.job_name,
        "source_type": job.source_type,
        "run_started_at": job.run_started_at,
        "job_folder": str(job.job_folder),
        "index_html": str(index_html),
        "summary": summary,
        "settings": {
            "timeout_sec": settings.timeout_sec,
            "retry_count": settings.retry_count,
            "sleep_sec": settings.sleep_sec,
            "overwrite_existing": settings.overwrite_existing,
            "min_file_size_kb": settings.min_file_size_kb,
            "max_file_size_mb": settings.max_file_size_mb,
            "on_too_small_file": settings.on_too_small_file,
            "on_too_large_file": settings.on_too_large_file,
            "kill_office_apps_before_run": settings.kill_office_apps_before_run,
            "download_workers": settings.download_workers,
            "unzip_workers": settings.unzip_workers,
            "pdf_workers": settings.pdf_workers,
            "html_workers": settings.html_workers,
            "combine_html_batch_size": settings.combine_html_batch_size,
        },
        "rows": [
            {
                "row_no": r.row_no,
                "title": r.title,
                "url": r.url,
                "status": r.status,
                "saved_path": r.saved_path,
                "message": r.message,
                "last_run_at": r.last_run_at,
            }
            for r in rows
        ],
        "files": [
            {
                "seq": pf.seq,
                "row_no": pf.row.row_no,
                "title": pf.row.title,
                "source_file": str(pf.source_file),
                "size_bytes": pf.size_bytes,
                "size_status": pf.size_status,
                "pdf_path": str(pf.pdf_path) if pf.pdf_path else "",
                "html_path": str(pf.html_path) if pf.html_path else "",
                "status": pf.status,
                "message": pf.message,
            }
            for pf in files
        ],
        "batches": [
            {
                "batch_no": b.batch_no,
                "first_seq": b.first_seq,
                "last_seq": b.last_seq,
                "combined_html_path": str(b.combined_html_path),
                "file_count": b.file_count,
            }
            for b in batches
        ],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return path
