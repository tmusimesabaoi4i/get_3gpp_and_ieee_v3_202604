"""PDF conversion via Microsoft Office (Word/PowerPoint/Excel) COM automation.

Falls back to LibreOffice headless if Office is unavailable.
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Optional

from . import models
from .models import JobContext, ProcessedFile, Settings
from .utils import ensure_dir, is_windows, now_iso

logger = logging.getLogger(__name__)


# Extensions that require a conversion pass.
WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".odt"}
PPT_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx", ".potm", ".odp"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm"}
CONVERTIBLE_EXTS = WORD_EXTS | PPT_EXTS | EXCEL_EXTS


# Thread-local COM objects (Word.Application, PowerPoint.Application, Excel.Application)
_thread_local = threading.local()


def _ensure_com_initialized():
    """Ensure COM is initialized for the current thread."""
    if not hasattr(_thread_local, "com_initialized"):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            _thread_local.com_initialized = True
        except Exception as e:
            logger.debug("CoInitialize failed: %s", e)
            _thread_local.com_initialized = False


def _uninitialize_com():
    """Uninitialize COM for the current thread."""
    if getattr(_thread_local, "com_initialized", False):
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _thread_local.com_initialized = False


def _get_word_app():
    """Get or create a Word.Application COM object for this thread."""
    _ensure_com_initialized()
    if not hasattr(_thread_local, "word_app"):
        try:
            import win32com.client
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            _thread_local.word_app = app
        except Exception as e:
            logger.debug("Failed to create Word.Application: %s", e)
            _thread_local.word_app = None
    return _thread_local.word_app


def _get_ppt_app():
    """Get or create a PowerPoint.Application COM object for this thread."""
    _ensure_com_initialized()
    if not hasattr(_thread_local, "ppt_app"):
        try:
            import win32com.client
            app = win32com.client.Dispatch("PowerPoint.Application")
            # PowerPoint doesn't allow Visible=False when opening files
            _thread_local.ppt_app = app
        except Exception as e:
            logger.debug("Failed to create PowerPoint.Application: %s", e)
            _thread_local.ppt_app = None
    return _thread_local.ppt_app


def _get_excel_app():
    """Get or create an Excel.Application COM object for this thread."""
    _ensure_com_initialized()
    if not hasattr(_thread_local, "excel_app"):
        try:
            import win32com.client
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            _thread_local.excel_app = app
        except Exception as e:
            logger.debug("Failed to create Excel.Application: %s", e)
            _thread_local.excel_app = None
    return _thread_local.excel_app


def _quit_thread_apps():
    """Quit any COM apps created in this thread."""
    for attr in ("word_app", "ppt_app", "excel_app"):
        app = getattr(_thread_local, attr, None)
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            setattr(_thread_local, attr, None)
    _uninitialize_com()


def _convert_word_to_pdf(input_path: Path, output_path: Path) -> tuple[bool, str]:
    """Convert Word document to PDF using Microsoft Word."""
    app = _get_word_app()
    if app is None:
        return False, "Word.Application not available"
    try:
        doc = app.Documents.Open(str(input_path), ReadOnly=True)
        doc.SaveAs(str(output_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(False)
        return True, "pdf converted (Word)"
    except Exception as e:
        return False, f"Word conversion failed: {e}"


def _convert_ppt_to_pdf(
    input_path: Path,
    output_path: Path,
    slides_dir: Optional[Path] = None,
    image_width: int = 1280,
    image_height: int = 720,
) -> tuple[bool, str, list[Path], list[str]]:
    """Convert PowerPoint presentation to PDF and extract per-slide image+text.

    Returns (ok, message, slide_images, slide_texts).
    - slide_images is a list of PNG paths (one per slide, in order).
    - slide_texts is a list of plain-text strings extracted from each slide's shapes.
    """
    app = _get_ppt_app()
    if app is None:
        return False, "PowerPoint.Application not available", [], []

    slide_images: list[Path] = []
    slide_texts: list[str] = []
    try:
        # msoFalse=0, msoTrue=-1
        pres = app.Presentations.Open(
            str(input_path), ReadOnly=True, Untitled=False, WithWindow=False
        )
        try:
            pres.SaveAs(str(output_path), FileFormat=32)  # 32 = ppSaveAsPDF

            if slides_dir is not None:
                ensure_dir(slides_dir)
                for idx, slide in enumerate(pres.Slides, start=1):
                    img_path = slides_dir / f"slide_{idx:03d}.png"
                    try:
                        slide.Export(str(img_path), "PNG", image_width, image_height)
                        if img_path.exists():
                            slide_images.append(img_path)
                    except Exception as exc:  # noqa: BLE001
                        logger.warning(
                            "Slide %d image export failed for %s: %s",
                            idx, input_path.name, exc,
                        )

                    # Text extraction from all shapes on the slide.
                    parts: list[str] = []
                    try:
                        for shape in slide.Shapes:
                            try:
                                if shape.HasTextFrame and shape.TextFrame.HasText:
                                    txt = shape.TextFrame.TextRange.Text
                                    if txt:
                                        parts.append(str(txt))
                            except Exception:
                                continue
                        # Include speaker notes if any.
                        try:
                            notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                            if notes and notes.strip():
                                parts.append("[Notes]\n" + str(notes))
                        except Exception:
                            pass
                    except Exception as exc:  # noqa: BLE001
                        logger.debug("Text extraction failed on slide %d: %s", idx, exc)
                    slide_texts.append("\n".join(parts).strip())
        finally:
            pres.Close()
        return True, "pdf converted (PowerPoint)", slide_images, slide_texts
    except Exception as e:
        return False, f"PowerPoint conversion failed: {e}", slide_images, slide_texts


def _convert_excel_to_pdf(input_path: Path, output_path: Path) -> tuple[bool, str]:
    """Convert Excel workbook to PDF using Microsoft Excel."""
    app = _get_excel_app()
    if app is None:
        return False, "Excel.Application not available"
    try:
        wb = app.Workbooks.Open(str(input_path), ReadOnly=True)
        wb.ExportAsFixedFormat(0, str(output_path))  # 0 = xlTypePDF
        wb.Close(False)
        return True, "pdf converted (Excel)"
    except Exception as e:
        return False, f"Excel conversion failed: {e}"


# ------------------------------------------------------------------ LibreOffice fallback

class _WorkerIdAllocator:
    """Hands out stable per-worker IDs for LibreOffice profiles."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._next = 0
        self._by_thread: dict[int, int] = {}

    def get(self) -> int:
        tid = threading.get_ident()
        with self._lock:
            if tid not in self._by_thread:
                self._next += 1
                self._by_thread[tid] = self._next
            return self._by_thread[tid]


_WORKER_IDS = _WorkerIdAllocator()


def _locate_soffice() -> Optional[str]:
    env = os.environ.get("STDHARVEST_SOFFICE")
    if env and Path(env).exists():
        return env
    for candidate in ("soffice", "soffice.exe"):
        found = shutil.which(candidate)
        if found:
            return found
    if is_windows():
        for pf in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(pf).exists():
                return pf
    return None


def _profile_uri(profile_dir: Path) -> str:
    uri = profile_dir.resolve().as_uri()
    return uri


def _run_soffice(
    soffice: str,
    input_file: Path,
    out_dir: Path,
    profile_dir: Path,
    timeout_sec: int,
) -> tuple[bool, str]:
    ensure_dir(profile_dir)
    ensure_dir(out_dir)
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--norestore",
        f"-env:UserInstallation={_profile_uri(profile_dir)}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(input_file),
    ]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            check=False,
        )
    except FileNotFoundError:
        return False, "soffice not found"
    except subprocess.TimeoutExpired:
        return False, f"soffice timed out after {timeout_sec}s"
    ok = proc.returncode == 0
    log = (proc.stdout + proc.stderr).strip()
    return ok, log


def _expected_pdf(out_dir: Path, input_file: Path) -> Path:
    return out_dir / (input_file.stem + ".pdf")


# ------------------------------------------------------------------ main conversion

def _check_office_available() -> bool:
    """Check if Microsoft Office COM automation is available."""
    if not is_windows():
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


_OFFICE_AVAILABLE: Optional[bool] = None


def _is_office_available() -> bool:
    global _OFFICE_AVAILABLE
    if _OFFICE_AVAILABLE is None:
        _OFFICE_AVAILABLE = _check_office_available()
    return _OFFICE_AVAILABLE


def _convert_file(
    pf: ProcessedFile,
    job: JobContext,
    settings: Settings,
    soffice: Optional[str],
    profile_root: Path,
) -> None:
    out_dir = job.pdf_dir / pf.item_folder
    ensure_dir(out_dir)

    # If already a PDF, just copy into pdf/ for a unified layout.
    if pf.ext == ".pdf":
        target = out_dir / pf.source_file.name
        if not target.exists() or target.resolve() != pf.source_file.resolve():
            try:
                shutil.copy2(pf.source_file, target)
            except Exception as exc:
                pf.status = models.STATUS_ERROR
                pf.message = f"pdf copy failed: {exc}"
                return
        pf.pdf_path = target
        pf.status = models.STATUS_DONE
        pf.message = "pdf copied"
        return

    if pf.ext not in CONVERTIBLE_EXTS:
        return

    pdf_path = out_dir / (pf.source_file.stem + ".pdf")

    # Try Microsoft Office first (Windows only)
    if _is_office_available():
        ok = False
        msg = ""
        if pf.ext in WORD_EXTS:
            ok, msg = _convert_word_to_pdf(pf.source_file, pdf_path)
        elif pf.ext in PPT_EXTS:
            slides_dir = out_dir / "slides"
            ok, msg, imgs, texts = _convert_ppt_to_pdf(
                pf.source_file, pdf_path, slides_dir=slides_dir,
            )
            pf.slide_images = imgs
            pf.slide_texts = texts
        elif pf.ext in EXCEL_EXTS:
            ok, msg = _convert_excel_to_pdf(pf.source_file, pdf_path)

        if ok and pdf_path.exists():
            pf.pdf_path = pdf_path
            pf.status = models.STATUS_DONE
            pf.message = msg
            return
        else:
            logger.warning("Office conversion failed for %s: %s", pf.source_file.name, msg)
            # Fall through to LibreOffice

    # Fallback to LibreOffice
    if not soffice:
        pf.status = models.STATUS_ERROR
        pf.message = "No PDF converter available (Office COM failed, LibreOffice not found)"
        return

    worker_id = _WORKER_IDS.get()
    profile_dir = profile_root / f"lo_profile_{worker_id:03d}"
    ok, log = _run_soffice(
        soffice=soffice,
        input_file=pf.source_file,
        out_dir=out_dir,
        profile_dir=profile_dir,
        timeout_sec=max(60, settings.timeout_sec * 5),
    )
    lo_pdf_path = _expected_pdf(out_dir, pf.source_file)
    if ok and lo_pdf_path.exists():
        pf.pdf_path = lo_pdf_path
        pf.status = models.STATUS_DONE
        pf.message = "pdf converted (LibreOffice)"
    else:
        pf.status = models.STATUS_ERROR
        pf.message = f"pdf conversion failed: {log[:200]}" if log else "pdf conversion failed"


def _worker_cleanup():
    """Called at end of each worker thread to release COM objects."""
    _quit_thread_apps()


def convert_all(files: List[ProcessedFile], job: JobContext, settings: Settings) -> None:
    """Convert every file (where applicable) to PDF."""
    if not files:
        return
    ensure_dir(job.pdf_dir)

    soffice = _locate_soffice()
    if not _is_office_available() and not soffice:
        logger.warning("Neither Microsoft Office nor LibreOffice found - PDF conversion will fail")

    profile_root = Path(tempfile.gettempdir()) / "stdharvest_lo_profiles"
    ensure_dir(profile_root)

    targets = [pf for pf in files if pf.ext in CONVERTIBLE_EXTS or pf.ext == ".pdf"]

    # For Office COM, use fewer workers to avoid COM threading issues
    workers = max(1, min(settings.pdf_workers, 2))

    def _convert_and_cleanup(pf: ProcessedFile):
        try:
            _convert_file(pf, job, settings, soffice, profile_root)
        finally:
            pass  # Don't quit apps between files for performance

    with ThreadPoolExecutor(max_workers=workers, initializer=None) as pool:
        futures = [pool.submit(_convert_and_cleanup, pf) for pf in targets]
        for _ in as_completed(futures):
            pass

    # Cleanup COM objects after all conversions
    _quit_thread_apps()

    _ = now_iso()
