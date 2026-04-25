"""Command-line entry point and pipeline orchestrator."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Tuple

from . import models
from .downloader import download_all
from .excel_io import read_excel, write_back
from .html_builder import (
    build_combine_full_html,
    build_combined_html,
    build_index_html,
    build_individual_html,
)
from .logger import (
    write_combined_html_csv,
    write_file_results,
    write_manifest,
    write_row_results,
)
from .models import (
    DownloadRow,
    JobContext,
    ProcessedFile,
    Settings,
)
from .office_killer import kill_office_apps
from .pdf_converter import CONVERTIBLE_EXTS, convert_all
from .size_check import classify, size_message
from .unzipper import unzip_all
from .utils import ensure_dir, item_folder_name, now_iso

logger = logging.getLogger("stdharvest")

# Extensions that should be included as ProcessedFile at all (vs. left as raw-only).
ALL_PROCESSABLE_EXTS = CONVERTIBLE_EXTS | {
    ".pdf", ".csv", ".txt",  # non-convertible but tracked
}


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


# ------------------------------------------------------------------ raw size
def _apply_raw_size_check(
    rows: List[DownloadRow],
    settings: Settings,
    source_type: str,
) -> None:
    """Classify the downloaded raw file for each row and react per settings."""
    for row in rows:
        if row.status != models.STATUS_DONE or row.raw_path is None:
            continue
        status, size = classify(row.raw_path, settings)
        row.raw_size_status = status
        msg = size_message(row.raw_path, status, size, settings)
        if status == models.SIZE_OK:
            continue
        if status == models.SIZE_TOO_SMALL:
            if settings.on_too_small_file == "error":
                row.status = models.STATUS_ERROR_TOO_SMALL
                row.message = msg
            else:
                row.status = models.STATUS_SKIPPED
                row.message = f"skipped: {msg}"
            row.last_run_at = now_iso()
        elif status == models.SIZE_TOO_LARGE:
            if source_type == models.SOURCE_3GPP and settings.on_too_large_file in ("skip", "keep_raw"):
                row.status = models.STATUS_DONE_WITH_SKIP
                row.message = f"raw {msg}; PDF/HTML skipped"
                row.last_run_at = now_iso()
            # For pdf_only and for IEEE, keep DONE here; per-file handling takes over.


# ------------------------------------------------------------------ collect files
def _collect_processed_files(
    rows: List[DownloadRow],
    job: JobContext,
) -> List[ProcessedFile]:
    """Turn rows (and their unzipped contents) into ProcessedFile objects.

    Sequence numbers are assigned globally in Excel row order, then file name
    order within a row (mainly affects 3GPP zips).
    """
    files: List[ProcessedFile] = []
    seq = 0
    for row in sorted(rows, key=lambda r: r.row_no):
        if row.status in (models.STATUS_ERROR, models.STATUS_ERROR_TOO_SMALL,
                          models.STATUS_SKIPPED):
            # Row won't produce files, but we still log the row itself.
            continue
        if row.raw_path is None:
            continue

        item_folder = item_folder_name(row.seq, row.title)

        if job.source_type == models.SOURCE_3GPP:
            if row.status == models.STATUS_DONE_WITH_SKIP:
                # zip too large and skip/keep_raw: no children
                continue
            unzipped = getattr(row, "_unzipped_files", None)
            if unzipped is None:
                # zip not extracted (e.g. raw was not a zip) - fall through to raw handling
                unzipped = [row.raw_path]
            for src in sorted(unzipped, key=lambda p: p.name.lower()):
                seq += 1
                files.append(
                    ProcessedFile(
                        seq=seq,
                        row=row,
                        source_file=src,
                        display_name=src.name,
                        item_folder=item_folder,
                    )
                )
        else:  # IEEE
            seq += 1
            files.append(
                ProcessedFile(
                    seq=seq,
                    row=row,
                    source_file=row.raw_path,
                    display_name=row.raw_path.name,
                    item_folder=item_folder,
                )
            )
    # Attach produced_files lists back to the rows (for logging / debugging).
    for pf in files:
        pf.row.produced_files.append(pf)
    return files


# ------------------------------------------------------------------ per-file size
def _apply_per_file_size_check(
    files: List[ProcessedFile],
    settings: Settings,
) -> None:
    for pf in files:
        status, size = classify(pf.source_file, settings)
        pf.size_status = status
        pf.size_bytes = size
        pf.message = size_message(pf.source_file, status, size, settings)
        if status == models.SIZE_TOO_SMALL:
            if settings.on_too_small_file == "error":
                pf.status = models.STATUS_ERROR_TOO_SMALL
            else:
                pf.status = models.STATUS_SKIPPED
        elif status == models.SIZE_TOO_LARGE:
            if settings.on_too_large_file in ("skip", "keep_raw"):
                pf.status = models.STATUS_SKIPPED
            # pdf_only leaves status empty so the PDF stage still runs.


def _pdf_targets(files: List[ProcessedFile], settings: Settings) -> List[ProcessedFile]:
    out: List[ProcessedFile] = []
    for pf in files:
        if pf.size_status == models.SIZE_TOO_LARGE and settings.on_too_large_file in ("skip", "keep_raw"):
            continue
        if pf.size_status == models.SIZE_TOO_SMALL:
            continue
        if pf.ext in CONVERTIBLE_EXTS or pf.ext == ".pdf":
            out.append(pf)
    return out


def _html_targets(files: List[ProcessedFile], settings: Settings) -> List[ProcessedFile]:
    out: List[ProcessedFile] = []
    for pf in files:
        if pf.size_status == models.SIZE_TOO_SMALL:
            continue
        if pf.size_status == models.SIZE_TOO_LARGE and settings.on_too_large_file in ("skip", "keep_raw", "pdf_only"):
            # spec: pdf_only -> HTMLを作らない ... but we still create a thin link page
            # to keep combined/index linking consistent. So don't filter out here;
            # html_builder will render the meta-only page.
            pass
        out.append(pf)
    return out


# ------------------------------------------------------------------ row finalize
def _finalize_rows(
    rows: List[DownloadRow],
    files: List[ProcessedFile],
    settings: Settings,
) -> None:
    """Re-derive row Status / Message from the per-file outcomes."""
    # Fill in STATUS_EMPTY for files that reached the end of the pipeline cleanly.
    for pf in files:
        if pf.status == models.STATUS_EMPTY:
            if pf.pdf_path or pf.html_path:
                pf.status = models.STATUS_DONE
            else:
                # non-convertible (e.g. .txt/.csv) - treat as done with no outputs
                pf.status = models.STATUS_DONE
                if not pf.message:
                    pf.message = "kept as raw (no PDF/HTML target)"

    by_row: dict[int, List[ProcessedFile]] = {}
    for pf in files:
        by_row.setdefault(pf.row.row_no, []).append(pf)

    for row in rows:
        if row.status in (models.STATUS_ERROR, models.STATUS_ERROR_TOO_SMALL, models.STATUS_SKIPPED):
            continue  # preserve decisive early error/skip
        group = by_row.get(row.row_no, [])
        if row.status == models.STATUS_DONE_WITH_SKIP:
            # already decided (e.g. raw zip too large)
            continue
        if not group:
            # downloaded ok but nothing to process (rare: only csv/txt)
            row.status = models.STATUS_DONE
            if not row.message:
                row.message = "downloaded"
            row.last_run_at = now_iso()
            continue

        n_ok = sum(1 for pf in group if pf.status == models.STATUS_DONE)
        n_skip = sum(1 for pf in group if pf.status in (models.STATUS_SKIPPED,))
        n_err = sum(1 for pf in group if pf.status.startswith("ERROR"))
        n_pdf = sum(1 for pf in group if pf.pdf_path and pf.pdf_path.exists())
        n_html = sum(1 for pf in group if pf.html_path and pf.html_path.exists())

        if n_err and n_ok == 0:
            row.status = models.STATUS_ERROR
        elif n_skip and n_ok:
            row.status = models.STATUS_DONE_WITH_SKIP
        elif n_skip and not n_ok:
            row.status = models.STATUS_SKIPPED
        elif n_err and n_ok:
            row.status = models.STATUS_DONE_WITH_SKIP
        else:
            row.status = models.STATUS_DONE

        parts: List[str] = []
        if row.message and "downloaded" in row.message:
            parts.append(row.message.split(";")[0])
        else:
            parts.append("downloaded")
        if row.raw_path and row.raw_path.suffix.lower() == ".zip":
            parts.append(f"unzipped {len(group)} files")
        parts.append(f"pdf {n_pdf}")
        parts.append(f"html {n_html}")
        if n_skip:
            parts.append(f"skipped {n_skip} files")
        if n_err:
            parts.append(f"errors {n_err}")
        row.message = ", ".join(parts)
        row.last_run_at = now_iso()


# ------------------------------------------------------------------ summary
def _summary(
    rows: List[DownloadRow],
    files: List[ProcessedFile],
) -> dict:
    total_rows = len(rows)
    success_rows = sum(1 for r in rows if r.status in (models.STATUS_DONE, models.STATUS_DONE_WITH_SKIP))
    error_rows = sum(1 for r in rows if r.status.startswith("ERROR"))
    size_skip_files = sum(
        1 for pf in files
        if pf.size_status in (models.SIZE_TOO_LARGE, models.SIZE_TOO_SMALL)
    )
    file_count = len(files)
    return {
        "総行数": total_rows,
        "処理件数": total_rows,
        "成功件数": success_rows,
        "エラー件数": error_rows,
        "対象ファイル数": file_count,
        "サイズ超過スキップ件数": size_skip_files,
    }


# ------------------------------------------------------------------ run
def run(excel_path: Path) -> int:
    logger.info("Loading Excel: %s", excel_path)
    data = read_excel(excel_path)
    job = data.job
    settings = data.settings
    logger.info("SourceType=%s JobName=%s", job.source_type, job.job_name)
    logger.info("Rows to process: %d (skipped: %d)", len(data.rows), len(data.skipped_rows))

    kill_office_apps(settings.kill_office_apps_before_run)

    ensure_dir(job.job_folder)
    ensure_dir(job.raw_dir)
    ensure_dir(job.pdf_dir)
    ensure_dir(job.html_dir)
    ensure_dir(job.html_files_dir)
    ensure_dir(job.html_combined_dir)
    ensure_dir(job.logs_dir)
    if job.source_type == models.SOURCE_3GPP:
        ensure_dir(job.unpacked_dir)

    logger.info("Downloading %d rows with %d workers", len(data.rows), settings.download_workers)
    download_all(data.rows, job, settings)

    logger.info("Classifying raw file sizes")
    _apply_raw_size_check(data.rows, settings, job.source_type)

    if job.source_type == models.SOURCE_3GPP:
        logger.info("Unzipping 3GPP downloads")
        unzip_all(
            [r for r in data.rows if r.status == models.STATUS_DONE],
            job,
            settings,
        )

    files = _collect_processed_files(data.rows, job)
    logger.info("Produced %d processed files", len(files))

    _apply_per_file_size_check(files, settings)

    pdf_targets = _pdf_targets(files, settings)
    logger.info("Converting %d files to PDF", len(pdf_targets))
    convert_all(pdf_targets, job, settings)

    html_targets = _html_targets(files, settings)
    logger.info("Generating %d individual HTML pages", len(html_targets))
    build_individual_html(html_targets, job, settings)

    logger.info("Generating combined HTML (batch size %d)", settings.combine_html_batch_size)
    batches = build_combined_html(files, job, settings)

    full_targets = [pf for pf in html_targets if pf.html_path and pf.html_path.exists()]
    combine_full_batches = build_combine_full_html(full_targets, job, settings)

    _finalize_rows(data.rows, files, settings)

    summary = _summary(data.rows, files)
    logger.info("Summary: %s", summary)

    index_path = build_index_html(
        job, files, batches, summary, combine_full_batches=combine_full_batches,
    )
    logger.info("index.html written: %s", index_path)

    write_row_results(data.rows + data.skipped_rows, job.logs_dir)
    write_file_results(files, job.logs_dir, job.source_type)
    write_combined_html_csv(batches, job.logs_dir)
    write_manifest(job, settings, data.rows, files, batches, summary, index_path)
    logger.info("Logs written under %s", job.logs_dir)

    logger.info("Writing back to Excel: %s", excel_path)
    # Set SavedPath for rows that only downloaded and didn't get populated by finalize.
    for r in data.rows:
        if not r.saved_path and r.raw_path:
            r.saved_path = str(r.raw_path)
    write_back(excel_path, data)

    logger.info("Done. Job folder: %s", job.job_folder)
    return 0


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="stdharvest",
        description="3GPP / IEEE standard document batch downloader & PDF/HTML converter",
    )
    sub = p.add_subparsers(dest="command", required=True)
    run_p = sub.add_parser("run", help="Run the pipeline for the given Excel file")
    run_p.add_argument(
        "--excel",
        required=True,
        type=Path,
        help="Path to the input workbook (Sheet1+Sheet2)",
    )
    run_p.add_argument("--verbose", action="store_true", help="Enable DEBUG logging")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    _configure_logging(getattr(args, "verbose", False))
    if args.command == "run":
        try:
            return run(args.excel.resolve())
        except Exception as exc:  # noqa: BLE001
            logger.exception("Fatal error: %s", exc)
            return 2
    parser.error("Unknown command")
    return 2  # pragma: no cover


if __name__ == "__main__":
    sys.exit(main())
