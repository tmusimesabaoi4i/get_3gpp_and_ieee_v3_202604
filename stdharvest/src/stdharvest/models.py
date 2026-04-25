"""Data classes shared across the pipeline."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


# ---------- Status constants (Sheet1!D) ----------
STATUS_EMPTY = ""
STATUS_DONE = "DONE"
STATUS_SKIPPED = "SKIPPED"
STATUS_ERROR = "ERROR"
STATUS_RETRY = "RETRY"
STATUS_ERROR_TOO_SMALL = "ERROR_TOO_SMALL"
STATUS_DONE_WITH_SKIP = "DONE_WITH_SKIP"


RETRY_STATUSES = {
    STATUS_EMPTY,
    STATUS_ERROR,
    STATUS_RETRY,
    STATUS_ERROR_TOO_SMALL,
}


# ---------- Source types ----------
SOURCE_3GPP = "3gpp"
SOURCE_IEEE = "ieee"
VALID_SOURCES = {SOURCE_3GPP, SOURCE_IEEE}


# ---------- Size judgement ----------
SIZE_OK = "OK"
SIZE_TOO_SMALL = "TOO_SMALL"
SIZE_TOO_LARGE = "TOO_LARGE"


@dataclass
class Settings:
    """Values stored in Sheet2 (Settings)."""

    proxy_url: str = ""
    timeout_sec: int = 60
    retry_count: int = 3
    sleep_sec: float = 0.5
    overwrite_existing: bool = False
    min_file_size_kb: int = 10
    max_file_size_mb: int = 100
    on_too_small_file: str = "error"  # error | skip
    on_too_large_file: str = "skip"   # skip | pdf_only | keep_raw
    kill_office_apps_before_run: bool = True
    download_workers: int = 8
    unzip_workers: int = 4
    pdf_workers: int = 2
    html_workers: int = 6
    combine_html_batch_size: int = 5


@dataclass
class JobContext:
    """Top-level settings read from Sheet1 and the computed job folder."""

    source_type: str
    output_root: Path
    job_name: str
    job_folder: Path
    run_started_at: str  # YYYY-MM-DD HH:MM:SS

    @property
    def raw_dir(self) -> Path:
        return self.job_folder / "raw"

    @property
    def unpacked_dir(self) -> Path:
        return self.job_folder / "unpacked"

    @property
    def pdf_dir(self) -> Path:
        return self.job_folder / "pdf"

    @property
    def html_dir(self) -> Path:
        return self.job_folder / "html"

    @property
    def html_files_dir(self) -> Path:
        return self.html_dir / "files"

    @property
    def html_combined_dir(self) -> Path:
        return self.html_dir / "combined"

    @property
    def html_combine_full_dir(self) -> Path:
        return self.html_dir / "combine_full"

    @property
    def logs_dir(self) -> Path:
        return self.job_folder / "logs"


@dataclass
class DownloadRow:
    """A row from Sheet1 (>=5) that we attempt to download."""

    row_no: int
    seq: int  # 1-based, assigned after filtering
    title: str
    url: str
    existing_status: str = ""

    status: str = STATUS_EMPTY
    saved_path: str = ""
    message: str = ""
    last_run_at: str = ""

    raw_path: Optional[Path] = None  # downloaded file (zip / docx / ...)
    raw_size_status: str = SIZE_OK
    produced_files: List["ProcessedFile"] = field(default_factory=list)


@dataclass
class ProcessedFile:
    """A concrete file to convert to PDF/HTML.

    For 3GPP: produced by unzipping. For IEEE: the downloaded file itself.
    """

    seq: int  # global sequence (1-based) over all processed files
    row: "DownloadRow"
    source_file: Path       # on-disk path of the actual file (unzipped or raw)
    display_name: str       # user-facing file name
    item_folder: str        # e.g. "001_R1-2409888_Xiaomi"
    size_bytes: int = 0
    size_status: str = SIZE_OK
    pdf_path: Optional[Path] = None
    html_path: Optional[Path] = None
    status: str = STATUS_EMPTY
    message: str = ""
    # PowerPoint extras: one entry per slide (index 0 = slide 1)
    slide_images: List[Path] = field(default_factory=list)
    slide_texts: List[str] = field(default_factory=list)

    @property
    def ext(self) -> str:
        return self.source_file.suffix.lower()


@dataclass
class CombinedBatch:
    batch_no: int
    first_seq: int
    last_seq: int
    combined_html_path: Path
    file_count: int


@dataclass
class CombineFullBatch:
    """A combine_full bundle: per-N inlined-body view (default N=5)."""

    batch_no: int
    first_seq: int
    last_seq: int
    combine_full_path: Path
    file_count: int
