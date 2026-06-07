"""HTML generation: per-file page, 5-file combined page, and index page."""
from __future__ import annotations

import base64
import html
import io
import logging
import os
import re
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Iterable, List, Optional, Tuple
from urllib.parse import urlsplit

from . import models
from .models import CombinedBatch, CombineFullBatch, JobContext, ProcessedFile, Settings
from .utils import ensure_dir, human_size, rel_link

logger = logging.getLogger(__name__)


# How many files to bundle into one combine_full_NNN.html.
COMBINE_FULL_CHUNK_SIZE = 5


# Extensions we try to embed as HTML body (best-effort; PDF is authoritative).
HTML_BODY_EXTS = {".docx"}
PPT_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx", ".potm", ".odp"}


_CSS = """
body { font-family: system-ui, "Segoe UI", sans-serif; margin: 2rem; color: #222; background: #fafafa; }
h1 { margin-bottom: 0.25rem; }
h2 { margin-top: 2.25rem; border-bottom: 1px solid #ddd; padding-bottom: 0.25rem; }
.meta { color: #555; font-size: 0.9rem; }
.card { background: white; border: 1px solid #e2e2e2; border-radius: 8px;
        padding: 1rem 1.25rem; margin: 1rem 0; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.badge { display: inline-block; padding: 0.1rem 0.55rem; border-radius: 999px;
         font-size: 0.8rem; margin-right: 0.25rem; background: #eef; color: #335; }
.badge.err { background: #fde; color: #822; }
.badge.ok { background: #e6f7ee; color: #1d6c3c; }
.badge.skip { background: #fff4db; color: #7a5d00; }
table.kv { border-collapse: collapse; margin-top: 0.5rem; }
table.kv td { padding: 0.15rem 0.75rem; vertical-align: top; }
table.kv td.k { color: #666; white-space: nowrap; }
a { color: #2a6df4; text-decoration: none; }
a:hover { text-decoration: underline; }
.small { font-size: 0.85rem; color: #666; }
.row-list li { margin-bottom: 0.35rem; }
.error-list { color: #a33; }
hr { border: none; border-top: 1px solid #ddd; margin: 2rem 0; }
article.doc { border: 1px solid #e2e2e2; border-radius: 8px; background: white; padding: 1rem 1.25rem; margin-bottom: 1.5rem; }
article.doc h3 { margin-top: 0; }
pre.docxbody { background: #fff; border: 1px solid #eee; padding: 0.75rem; white-space: pre-wrap; }

/* Word body images: never overflow the column; keep aspect ratio; an inline
   width (from the original Word size) may shrink small images appropriately. */
.docx { overflow-wrap: break-word; word-break: break-word; }
.docx img { max-width: 100%; height: auto; max-height: 85vh; }
.docx table { max-width: 100%; }
.docx table td, .docx table th { word-break: break-word; }

/* PowerPoint slides */
.slide { background: white; border: 1px solid #e2e2e2; border-radius: 8px;
         padding: 1rem 1.25rem; margin: 1rem 0; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
.slide h3 { margin: 0 0 0.5rem 0; color: #333; }
.slide .slide-image { margin: 0.5rem 0; text-align: center; }
.slide .slide-image img { max-width: 100%; height: auto; border: 1px solid #ddd;
                          border-radius: 4px; box-shadow: 0 2px 6px rgba(0,0,0,0.08); }
.slide .slide-text { background: #fafbfc; border: 1px solid #eee; border-radius: 4px;
                     padding: 0.75rem 1rem; margin-top: 0.75rem;
                     font-size: 0.92rem; line-height: 1.5; white-space: pre-wrap;
                     font-family: "Segoe UI", "Hiragino Sans", sans-serif; }
.slide .slide-text.empty { color: #999; font-style: italic; }
.slide-toc { display: flex; flex-wrap: wrap; gap: 0.5rem; margin: 0.5rem 0 1.5rem 0; }
.slide-toc a { display: inline-block; padding: 0.2rem 0.6rem; background: #eef3ff;
               color: #2a5bb4; border-radius: 4px; font-size: 0.85rem; }
"""


def _escape(value: object) -> str:
    return html.escape(str(value), quote=True)


def _status_badge(status: str) -> str:
    cls = "badge"
    if status == models.STATUS_DONE:
        cls += " ok"
    elif status in (models.STATUS_SKIPPED, models.STATUS_DONE_WITH_SKIP):
        cls += " skip"
    elif status.startswith("ERROR"):
        cls += " err"
    return f'<span class="{cls}">{_escape(status or "PENDING")}</span>'


# Image content-types browsers can render directly (embedded as-is).
_WEB_SAFE_IMAGE_TYPES = {
    "image/png", "image/jpeg", "image/jpg", "image/gif",
    "image/bmp", "image/webp", "image/svg+xml",
}

# DPI used when rasterising vector metafiles (EMF/WMF) to PNG. 150 keeps text
# crisp without bloating the embedded data URI too much.
_METAFILE_RENDER_DPI = 150


def _raster_to_png(data: bytes) -> Optional[bytes]:
    """Convert non-web image bytes (EMF/WMF/TIFF/...) to PNG via Pillow.

    Returns None when Pillow is unavailable or the image cannot be rendered
    (e.g. EMF/WMF rendering is only supported on Windows). The caller then
    falls back to embedding the original bytes unchanged.
    """
    try:
        from PIL import Image  # lazy import (Windows: renders EMF/WMF via GDI)
    except Exception:  # noqa: BLE001
        return None
    try:
        with Image.open(io.BytesIO(data)) as im:
            # WMF/EMF honour a dpi hint to control the rasterisation size.
            try:
                im.load(dpi=_METAFILE_RENDER_DPI)
            except TypeError:
                im.load()
            if im.mode not in ("RGB", "RGBA", "L"):
                im = im.convert("RGBA")
            out = io.BytesIO()
            im.save(out, "PNG")
            return out.getvalue()
    except Exception as exc:  # noqa: BLE001
        logger.warning("画像のPNG変換に失敗しました: %s", exc)
        return None


def _make_image_converter():
    """Return a mammoth image handler that web-safe-ifies embedded images.

    Word documents (especially 3GPP) frequently embed charts/equations as EMF
    or WMF metafiles. Browsers cannot display ``data:image/x-emf`` URIs, so we
    rasterise those to PNG; already-web-safe formats are embedded unchanged.
    """
    import mammoth  # caller guarantees mammoth is importable

    @mammoth.images.img_element
    def _convert(image):
        content_type = (getattr(image, "content_type", "") or "").lower()
        with image.open() as stream:
            data = stream.read()
        if content_type not in _WEB_SAFE_IMAGE_TYPES:
            png = _raster_to_png(data)
            if png is not None:
                data = png
                content_type = "image/png"
            elif not content_type:
                content_type = "application/octet-stream"
        encoded = base64.b64encode(data).decode("ascii")
        return {"src": f"data:{content_type};base64,{encoded}"}

    return _convert


# ---- Word intended image sizing (so big/small images match the original) ----

# OOXML namespaces used when reading the intended on-page image sizes.
_NS_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_NS_VML = "{urn:schemas-microsoft-com:vml}"
_NS_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_EMU_PER_PX = 9525  # 914400 EMU/inch ÷ 96 px/inch


def _vml_style_px(style: str, key: str) -> int:
    """Parse a VML ``style`` length (``width:123pt``) into CSS px (0 if absent)."""
    match = re.search(rf"{key}\s*:\s*([0-9.]+)pt", style or "")
    return round(float(match.group(1)) * 4 / 3) if match else 0  # 1pt = 4/3 px @96dpi


def _walk_image_dims(el: ET.Element, out: List[Tuple[int, int]]) -> None:
    """Collect (width_px, height_px) for each drawing, in document order.

    For ``mc:AlternateContent`` only the chosen branch is followed so a single
    logical image (DrawingML + VML fallback) is not counted twice.
    """
    tag = el.tag
    if tag == _NS_MC + "AlternateContent":
        target = el.find(_NS_MC + "Choice")
        if target is None:
            target = el.find(_NS_MC + "Fallback")
        if target is not None:
            for child in target:
                _walk_image_dims(child, out)
        return
    if tag == _NS_WP + "extent":
        cx = int(el.get("cx", "0") or "0")
        cy = int(el.get("cy", "0") or "0")
        out.append((round(cx / _EMU_PER_PX), round(cy / _EMU_PER_PX)))
        return
    if tag == _NS_VML + "shape" and el.find(_NS_VML + "imagedata") is not None:
        style = el.get("style", "")
        out.append((_vml_style_px(style, "width"), _vml_style_px(style, "height")))
    for child in el:
        _walk_image_dims(child, out)


def _extract_docx_image_dims(path: Path) -> List[Tuple[int, int]]:
    """Return the on-page display size (CSS px) of each image, in document order."""
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("word/document.xml")
        root = ET.fromstring(xml)
    except Exception:  # noqa: BLE001
        return []
    dims: List[Tuple[int, int]] = []
    _walk_image_dims(root, dims)
    return dims


def _apply_image_dims(body_html: str, dims: List[Tuple[int, int]]) -> str:
    """Set each <img> width to the size Word displays it at.

    Applied only when the image count matches exactly (all-or-nothing) so we
    never mis-assign a size. Width is set via inline style; CSS keeps the
    aspect ratio and caps oversized images to the container width.
    """
    try:
        from bs4 import BeautifulSoup  # lazy import
    except Exception:  # noqa: BLE001
        return body_html
    soup = BeautifulSoup(body_html, "html.parser")
    imgs = soup.find_all("img")
    if not imgs or len(imgs) != len(dims):
        return body_html
    for img, (w, _h) in zip(imgs, dims):
        if w and w > 0:
            existing = img.get("style", "")
            img["style"] = (existing + f"width:{w}px;").strip()
    return str(soup)


def _docx_to_html_body(path: Path) -> Optional[str]:
    try:
        import mammoth  # lazy import
    except ImportError:
        return None
    try:
        with open(path, "rb") as f:
            result = mammoth.convert_to_html(f, convert_image=_make_image_converter())
        body = result.value
        if not body:
            return None
        dims = _extract_docx_image_dims(path)
        if dims:
            body = _apply_image_dims(body, dims)
        return body
    except Exception as exc:  # noqa: BLE001
        logger.warning("mammoth failed on %s: %s", path, exc)
        return None


def _file_meta_table(
    pf: ProcessedFile,
    html_file: Path,
) -> str:
    pdf_link = (
        f'<a href="{_escape(rel_link(html_file, pf.pdf_path))}">PDFを開く</a>'
        if pf.pdf_path and pf.pdf_path.exists()
        else '<span class="small">PDFなし</span>'
    )
    raw_link = (
        f'<a href="{_escape(rel_link(html_file, pf.source_file))}">元ファイルを開く</a>'
        if pf.source_file.exists()
        else '<span class="small">元ファイルなし</span>'
    )
    return f"""
    <table class="kv">
      <tr><td class="k">SourceType</td><td>{_escape(_source_of(pf))}</td></tr>
      <tr><td class="k">元URL</td><td><a href="{_escape(pf.row.url)}">{_escape(pf.row.url)}</a></td></tr>
      <tr><td class="k">元ファイル名</td><td>{_escape(pf.display_name)}</td></tr>
      <tr><td class="k">ファイルサイズ</td><td>{_escape(human_size(pf.size_bytes))} ({pf.size_bytes:,} bytes)</td></tr>
      <tr><td class="k">PDFリンク</td><td>{pdf_link}</td></tr>
      <tr><td class="k">元ファイルリンク</td><td>{raw_link}</td></tr>
      <tr><td class="k">ステータス</td><td>{_status_badge(pf.status)}</td></tr>
      <tr><td class="k">メッセージ</td><td>{_escape(pf.message)}</td></tr>
    </table>
    """


def _source_of(pf: ProcessedFile) -> str:
    # The DownloadRow doesn't carry source_type directly; but the job does.
    # We stash source_type on processed files via the job, set by caller.
    return getattr(pf, "source_type", "")


def _page_shell(title: str, body: str) -> str:
    return f"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8"/>
<title>{_escape(title)}</title>
<style>{_CSS}</style>
</head>
<body>
{body}
</body>
</html>
"""


def _build_single_html(pf: ProcessedFile, job: JobContext) -> Path:
    target_dir = job.html_files_dir / pf.item_folder
    ensure_dir(target_dir)
    html_file = target_dir / "index.html"

    body_sections: List[str] = []
    body_sections.append(f"<h1>{_escape(pf.display_name)}</h1>")
    body_sections.append(
        f'<p class="meta">Row {pf.row.row_no} / Seq {pf.seq:03d} / '
        f'Job: {_escape(job.job_name)} ({_escape(job.source_type)})</p>'
    )
    body_sections.append('<div class="card">')
    body_sections.append(_file_meta_table(pf, html_file))
    body_sections.append("</div>")

    if pf.size_status == models.SIZE_OK and pf.ext in HTML_BODY_EXTS:
        body_text = _docx_to_html_body(pf.source_file)
        if body_text:
            body_sections.append("<h2>本文(簡易HTML)</h2>")
            body_sections.append(f'<div class="card docx">{body_text}</div>')
        else:
            body_sections.append(
                '<p class="small">HTML本文化に失敗したため、PDFリンクと元ファイルリンクのみ表示しています。</p>'
            )
    elif pf.ext in PPT_EXTS and pf.slide_images:
        body_sections.append(_render_ppt_slides(pf, html_file))
    elif pf.ext in PPT_EXTS:
        body_sections.append(
            '<p class="small">スライド画像が生成されなかったため、PDFリンクと元ファイルリンクを参照してください。</p>'
        )
    elif pf.ext == ".pdf":
        body_sections.append('<p class="small">PDF本文はPDFリンクから閲覧してください。</p>')
    else:
        body_sections.append(
            '<p class="small">このファイル種別はリンク表示中心です。必要に応じてPDFを参照してください。</p>'
        )

    html_file.write_text(_page_shell(pf.display_name, "\n".join(body_sections)), encoding="utf-8")
    pf.html_path = html_file
    return html_file


def _render_ppt_slides(pf: ProcessedFile, html_file: Path) -> str:
    """Render PowerPoint slide images + extracted text as HTML sections."""
    parts: List[str] = []
    parts.append(f'<h2>スライド ({len(pf.slide_images)} 枚)</h2>')

    # Slide table-of-contents (anchors).
    toc_items = [
        f'<a href="#slide-{i:03d}">#{i}</a>' for i in range(1, len(pf.slide_images) + 1)
    ]
    if toc_items:
        parts.append('<div class="slide-toc">' + "".join(toc_items) + "</div>")

    texts = pf.slide_texts or []
    for i, img in enumerate(pf.slide_images, start=1):
        text = texts[i - 1] if i - 1 < len(texts) else ""
        img_link = _escape(rel_link(html_file, img))
        parts.append(f'<section class="slide" id="slide-{i:03d}">')
        parts.append(f'<h3>Slide {i}</h3>')
        parts.append(
            f'<div class="slide-image"><a href="{img_link}" target="_blank">'
            f'<img src="{img_link}" alt="Slide {i}"/></a></div>'
        )
        if text:
            parts.append(f'<div class="slide-text">{_escape(text)}</div>')
        else:
            parts.append('<div class="slide-text empty">(テキストなし)</div>')
        parts.append("</section>")
    return "\n".join(parts)


def build_individual_html(
    files: List[ProcessedFile],
    job: JobContext,
    settings: Settings,
) -> None:
    if not files:
        return
    ensure_dir(job.html_files_dir)
    # Stamp source_type on each file for the meta table.
    for pf in files:
        setattr(pf, "source_type", job.source_type)

    targets = [pf for pf in files if pf.size_status != models.SIZE_TOO_LARGE
               or settings.on_too_large_file == "pdf_only"]  # pdf_only still gets HTML
    # per spec: TOO_LARGE skip/keep_raw => no HTML (we still build a minimal one
    # so the link from combined/index works). Decide:
    def _should_build(pf: ProcessedFile) -> bool:
        if pf.size_status == models.SIZE_TOO_LARGE and settings.on_too_large_file in ("skip", "keep_raw"):
            return True  # build a thin page that shows "skipped due to size"
        return True

    workers = max(1, settings.html_workers)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_build_single_html, pf, job) for pf in files if _should_build(pf)]
        for _ in as_completed(futures):
            pass


def _combined_header(job: JobContext, first_seq: int, last_seq: int) -> str:
    return (
        f'<h1>Combined HTML #{first_seq:03d}-{last_seq:03d}</h1>'
        f'<p class="meta">Job: {_escape(job.job_name)} ({_escape(job.source_type)}) — '
        f'{_escape(job.run_started_at)}</p>'
    )


def _combined_block(pf: ProcessedFile, combined_path: Path) -> str:
    individual = pf.html_path
    pdf_link = (
        f'<a href="{_escape(rel_link(combined_path, pf.pdf_path))}">PDF</a>'
        if pf.pdf_path and pf.pdf_path.exists()
        else '<span class="small">PDFなし</span>'
    )
    raw_link = (
        f'<a href="{_escape(rel_link(combined_path, pf.source_file))}">元ファイル</a>'
        if pf.source_file.exists()
        else '<span class="small">元ファイルなし</span>'
    )
    indiv_link = (
        f'<a href="{_escape(rel_link(combined_path, individual))}">個別HTML</a>'
        if individual and individual.exists()
        else '<span class="small">個別HTMLなし</span>'
    )
    size_info = human_size(pf.size_bytes) if pf.size_bytes else "-"
    reason = ""
    if pf.size_status == models.SIZE_TOO_LARGE:
        reason = f'<p class="small">サイズ超過のためPDF/HTML化をスキップ (size={size_info}).</p>'
    elif pf.size_status == models.SIZE_TOO_SMALL:
        reason = f'<p class="small">サイズ過小のため対象外 (size={size_info}).</p>'
    return f"""
    <article class="doc">
      <h3>{pf.seq:03d}. {_escape(pf.display_name)} {_status_badge(pf.status)}</h3>
      <p class="small">Row {pf.row.row_no} — {_escape(pf.row.title)}<br/>
         URL: <a href="{_escape(pf.row.url)}">{_escape(pf.row.url)}</a></p>
      <p>{pdf_link} / {raw_link} / {indiv_link}
         <span class="small">({_escape(size_info)})</span></p>
      {reason}
      <p class="small">{_escape(pf.message)}</p>
    </article>
    """


def build_combined_html(
    files: List[ProcessedFile],
    job: JobContext,
    settings: Settings,
) -> List[CombinedBatch]:
    if not files:
        return []
    ensure_dir(job.html_combined_dir)
    batch_size = max(1, settings.combine_html_batch_size)
    ordered = sorted(files, key=lambda pf: (pf.row.row_no, pf.seq))
    batches: List[CombinedBatch] = []
    for batch_no, start in enumerate(range(0, len(ordered), batch_size), start=1):
        group = ordered[start : start + batch_size]
        first_seq = group[0].seq
        last_seq = group[-1].seq
        path = job.html_combined_dir / f"combined_{first_seq:03d}_{last_seq:03d}.html"
        blocks = [_combined_block(pf, path) for pf in group]
        body = _combined_header(job, first_seq, last_seq) + "\n" + "\n".join(blocks)
        path.write_text(_page_shell(f"Combined {first_seq:03d}-{last_seq:03d}", body), encoding="utf-8")
        batches.append(
            CombinedBatch(
                batch_no=batch_no,
                first_seq=first_seq,
                last_seq=last_seq,
                combined_html_path=path,
                file_count=len(group),
            )
        )
    return batches


def build_index_html(
    job: JobContext,
    files: List[ProcessedFile],
    batches: List[CombinedBatch],
    summary: dict,
    combine_full_batches: Optional[List[CombineFullBatch]] = None,
) -> Path:
    ensure_dir(job.html_dir)
    index_path = job.html_dir / "index.html"

    def _li_for_file(pf: ProcessedFile) -> str:
        individual = pf.html_path
        indiv_link = (
            f'<a href="{_escape(rel_link(index_path, individual))}">{_escape(pf.display_name)}</a>'
            if individual and individual.exists()
            else _escape(pf.display_name)
        )
        return (
            f'<li>{pf.seq:03d} · {indiv_link} {_status_badge(pf.status)} '
            f'<span class="small">({_escape(human_size(pf.size_bytes))})</span></li>'
        )

    def _li_for_batch(b: CombinedBatch) -> str:
        return (
            f'<li><a href="{_escape(rel_link(index_path, b.combined_html_path))}">'
            f'combined_{b.first_seq:03d}_{b.last_seq:03d}.html</a> '
            f'<span class="small">({b.file_count} files)</span></li>'
        )

    errors: List[ProcessedFile] = [pf for pf in files if pf.status.startswith("ERROR")]

    body: List[str] = []
    body.append(f"<h1>{_escape(job.job_name)}</h1>")
    body.append(
        f'<p class="meta">SourceType: {_escape(job.source_type)} · '
        f'実行日時: {_escape(job.run_started_at)}</p>'
    )

    body.append('<div class="card"><h2>サマリ</h2><table class="kv">')
    for k, v in summary.items():
        body.append(f'<tr><td class="k">{_escape(k)}</td><td>{_escape(v)}</td></tr>')
    body.append("</table></div>")

    body.append('<h2>個別HTML一覧</h2><ul class="row-list">')
    body.extend(_li_for_file(pf) for pf in sorted(files, key=lambda p: p.seq))
    body.append("</ul>")

    body.append('<h2>Combine HTML一覧</h2><ul class="row-list">')
    body.extend(_li_for_batch(b) for b in batches)
    body.append("</ul>")

    if combine_full_batches:
        def _li_for_full(b: CombineFullBatch) -> str:
            return (
                f'<li><a href="{_escape(rel_link(index_path, b.combine_full_path))}">'
                f'{_escape(b.combine_full_path.name)}</a> '
                f'<span class="small">({b.first_seq}〜{b.last_seq} 件目 / {b.file_count} files)</span></li>'
            )
        body.append('<h2>まとめ全文HTML (combine_full)</h2>')
        body.append(
            '<p class="small">クリック不要で本文と画像をまとめて閲覧できます '
            f'(1ページ {COMBINE_FULL_CHUNK_SIZE} 件ずつ)。</p>'
        )
        body.append('<ul class="row-list">')
        body.extend(_li_for_full(b) for b in combine_full_batches)
        body.append("</ul>")

    body.append('<h2>エラー一覧</h2>')
    if errors:
        body.append('<ul class="row-list error-list">')
        for pf in errors:
            body.append(
                f'<li>seq={pf.seq:03d} row={pf.row.row_no} · {_escape(pf.display_name)} · '
                f'{_escape(pf.status)} · {_escape(pf.message)}</li>'
            )
        body.append("</ul>")
    else:
        body.append('<p class="small">エラーはありません。</p>')

    index_path.write_text(_page_shell(job.job_name, "\n".join(body)), encoding="utf-8")
    return index_path


def iter_errors(files: Iterable[ProcessedFile]) -> Iterable[ProcessedFile]:
    return (pf for pf in files if pf.status.startswith("ERROR"))


# ===================================================================
# combine_full: per-N HTML with each document's body inlined (no clicks)
# ===================================================================

_COMBINE_FULL_CSS = """
body {
    font-family: "Segoe UI", "Meiryo", "Hiragino Sans", sans-serif;
    background: #f5f6f8;
    color: #222;
    line-height: 1.7;
    margin: 0;
    padding: 0;
}
header.cf-header {
    background: #ffffff;
    padding: 24px 40px;
    border-bottom: 1px solid #ddd;
}
header.cf-header h1 { margin: 0 0 4px 0; }
header.cf-header p { margin: 0; color: #555; }

nav.cf-toc {
    max-width: 1100px;
    margin: 24px auto;
    padding: 16px 24px;
    background: #ffffff;
    border: 1px solid #ddd;
    border-radius: 10px;
}
nav.cf-toc strong { display: block; margin-bottom: 8px; color: #333; }
nav.cf-toc a {
    display: inline-block;
    margin: 4px 8px 4px 0;
    padding: 4px 10px;
    background: #eef3ff;
    color: #2a5bb4;
    border-radius: 4px;
    text-decoration: none;
    font-size: 0.9em;
}
nav.cf-toc a:hover { background: #d9e4ff; }

main.cf-main {
    max-width: 1100px;
    margin: 32px auto;
    padding: 0 24px;
}

.document-block {
    background: #ffffff;
    border: 1px solid #ddd;
    border-radius: 12px;
    padding: 28px;
    margin-bottom: 40px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.document-block h2 {
    margin-top: 0;
    padding-bottom: 12px;
    border-bottom: 1px solid #e5e5e5;
}
.document-meta {
    font-size: 0.9em;
    color: #555;
    margin-bottom: 20px;
    background: #fafbfc;
    border: 1px solid #eee;
    border-radius: 6px;
    padding: 10px 14px;
}
.document-meta a { color: #2a6df4; text-decoration: none; margin-right: 12px; }
.document-meta a:hover { text-decoration: underline; }
.document-meta .badge {
    display: inline-block;
    padding: 0.1rem 0.55rem;
    border-radius: 999px;
    font-size: 0.8rem;
    margin-right: 0.25rem;
    background: #eef;
    color: #335;
}
.document-meta .badge.ok { background: #e6f7ee; color: #1d6c3c; }
.document-meta .badge.skip { background: #fff4db; color: #7a5d00; }
.document-meta .badge.err { background: #fde; color: #822; }

.document-content {
    overflow-wrap: break-word;
    word-break: break-word;
}
.document-content img,
.document-content svg {
    max-width: 100%;
    height: auto;
    max-height: 85vh;
    display: block;
    margin: 16px auto;
    border: 1px solid #ddd;
    border-radius: 6px;
}
.document-content table { max-width: 100%; border-collapse: collapse; }
.document-content table td, .document-content table th { border: 1px solid #ddd; padding: 4px 8px; }
.document-content pre, .document-content code { white-space: pre-wrap; word-break: break-word; }

/* PowerPoint slide cards reused from individual HTML */
.document-content .slide {
    background: #fafbfc;
    border: 1px solid #e2e2e2;
    border-radius: 8px;
    padding: 1rem 1.25rem;
    margin: 1rem 0;
}
.document-content .slide h3 { margin: 0 0 0.5rem 0; }
.document-content .slide-image { text-align: center; }
.document-content .slide-text {
    background: #fff;
    border: 1px solid #eee;
    border-radius: 4px;
    padding: 0.75rem 1rem;
    margin-top: 0.75rem;
    white-space: pre-wrap;
}
.document-content .slide-toc { display: flex; flex-wrap: wrap; gap: 0.4rem; margin: 0.5rem 0 1rem 0; }
.document-content .slide-toc a {
    display: inline-block;
    padding: 0.15rem 0.55rem;
    background: #eef3ff;
    color: #2a5bb4;
    border-radius: 4px;
    font-size: 0.85rem;
    text-decoration: none;
}

.error-box {
    background: #fff4f4;
    border: 1px solid #f4c0c0;
    color: #822;
    padding: 12px 16px;
    border-radius: 6px;
}

.back-to-top {
    margin-top: 24px;
    text-align: right;
    font-size: 0.9em;
}
.back-to-top a { color: #2a6df4; text-decoration: none; }

footer.cf-footer {
    text-align: center;
    color: #888;
    font-size: 0.85em;
    padding: 24px 0 40px 0;
}
"""


def _is_external_url(url: str) -> bool:
    if not url:
        return True
    if url.startswith("#"):
        return True
    if url.startswith("data:"):
        return True
    if url.startswith("mailto:") or url.startswith("javascript:"):
        return True
    parsed = urlsplit(url)
    return bool(parsed.scheme)


def _rewrite_relative_path(
    url: str,
    src_html_file: Path,
    dst_html_file: Path,
) -> str:
    """Convert a relative URL inside `src_html_file` to one usable from `dst_html_file`.

    Anchors, absolute URLs and data: URIs are returned unchanged.
    """
    if _is_external_url(url):
        return url
    # Split off optional fragment / query for safe path resolution.
    parsed = urlsplit(url)
    path_part = parsed.path
    suffix = ""
    if parsed.query:
        suffix += "?" + parsed.query
    if parsed.fragment:
        suffix += "#" + parsed.fragment
    try:
        absolute = (src_html_file.parent / path_part).resolve()
    except OSError:
        return url
    try:
        rel = os.path.relpath(absolute, start=dst_html_file.parent)
    except ValueError:
        return absolute.as_uri() + suffix
    return rel.replace(os.sep, "/") + suffix


def _scope_anchor(href: str, doc_index: int) -> str:
    """If href is `#xxx`, prefix with `d{N}-` so multiple inlined docs don't collide."""
    if href and href.startswith("#") and len(href) > 1:
        return f"#d{doc_index}-{href[1:]}"
    return href


def _scope_id(value: str, doc_index: int) -> str:
    if not value:
        return value
    return f"d{doc_index}-{value}"


def _extract_inlined_body(
    pf: ProcessedFile,
    combine_full_path: Path,
    doc_index: int,
) -> str:
    """Read pf.html_path, return the inlined <body> content (HTML string)."""
    try:
        from bs4 import BeautifulSoup  # lazy import
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("beautifulsoup4 is required for combine_full") from exc

    if not pf.html_path or not pf.html_path.exists():
        raise FileNotFoundError(pf.html_path)
    raw = pf.html_path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(raw, "html.parser")
    body = soup.body
    if body is None:
        return ""

    # Strip the duplicated header that the individual page already shows
    # (we render our own document-meta above the inlined content).
    first_h1 = body.find("h1")
    if first_h1 is not None:
        first_h1.decompose()
    first_meta = body.find("p", class_="meta")
    if first_meta is not None:
        first_meta.decompose()
    first_card = body.find("div", class_="card")
    if first_card is not None:
        # Only drop the meta-table card (the one without docx body content).
        if "docx" not in (first_card.get("class") or []):
            first_card.decompose()

    src = pf.html_path
    # Rewrite all relative paths and prefix IDs / fragment links per doc.
    for tag in body.find_all(True):
        # Scope IDs so internal anchors don't collide across documents.
        if tag.has_attr("id"):
            tag["id"] = _scope_id(tag["id"], doc_index)

        for attr in ("src", "href", "data-src", "poster"):
            if not tag.has_attr(attr):
                continue
            value = tag[attr]
            if not isinstance(value, str):
                continue
            # In-page anchors get scoped to this document's namespace.
            if value.startswith("#"):
                tag[attr] = _scope_anchor(value, doc_index)
                continue
            try:
                tag[attr] = _rewrite_relative_path(value, src, combine_full_path)
            except Exception as exc:  # noqa: BLE001
                logger.warning("failed to rewrite image path: %s (%s)", value, exc)

    # Convert children to HTML (skip the body tag itself).
    return "".join(str(child) for child in body.children)


def _document_block_html(
    pf: ProcessedFile,
    combine_full_path: Path,
    doc_index: int,
) -> str:
    """Render one <section class='document-block'> for a ProcessedFile."""
    title = f"{pf.seq}. {_escape(pf.display_name)}"
    section_id = f"doc-{doc_index}"

    pdf_link = (
        f'<a href="{_escape(rel_link(combine_full_path, pf.pdf_path))}">PDFを開く</a>'
        if pf.pdf_path and pf.pdf_path.exists()
        else '<span class="small">PDFなし</span>'
    )
    raw_link = (
        f'<a href="{_escape(rel_link(combine_full_path, pf.source_file))}">元ファイル</a>'
        if pf.source_file.exists()
        else '<span class="small">元ファイルなし</span>'
    )
    indiv_link = (
        f'<a href="{_escape(rel_link(combine_full_path, pf.html_path))}">個別HTML</a>'
        if pf.html_path and pf.html_path.exists()
        else '<span class="small">個別HTMLなし</span>'
    )

    meta = (
        f'<div class="document-meta">'
        f'<div>元ファイル名: <strong>{_escape(pf.display_name)}</strong> '
        f'· Row {pf.row.row_no} · Seq {pf.seq:03d} '
        f'· {_status_badge(pf.status)} '
        f'· {_escape(human_size(pf.size_bytes))}</div>'
        f'<div>{pdf_link} · {raw_link} · {indiv_link}</div>'
        f'</div>'
    )

    try:
        inlined = _extract_inlined_body(pf, combine_full_path, doc_index)
    except FileNotFoundError:
        logger.warning("combine_full failed to include: %s (no individual HTML)", pf.display_name)
        inlined = '<div class="error-box">この文書のHTML読み込みに失敗しました（個別HTMLが存在しません）。</div>'
    except Exception as exc:  # noqa: BLE001
        logger.warning("combine_full failed to include: %s (%s)", pf.display_name, exc)
        inlined = (
            f'<div class="error-box">この文書のHTML読み込みに失敗しました: '
            f'{_escape(exc)}</div>'
        )

    back = '<p class="back-to-top"><a href="#cf-top">▲ 上部へ戻る</a></p>'

    return (
        f'<section id="{section_id}" class="document-block">'
        f'<h2>{title}</h2>'
        f'{meta}'
        f'<div class="document-content">{inlined}</div>'
        f'{back}'
        f'</section>'
    )


def _combine_full_page(
    job: JobContext,
    files: List[ProcessedFile],
    out_path: Path,
    batch_no: int,
    total_batches: int,
) -> None:
    title = out_path.stem  # combine_full_001
    first_seq = files[0].seq
    last_seq = files[-1].seq
    toc_links = "".join(
        f'<a href="#doc-{i + 1}">{pf.seq}. {_escape(pf.display_name)}</a>'
        for i, pf in enumerate(files)
    )
    blocks = "\n".join(
        _document_block_html(pf, out_path, doc_index=i + 1)
        for i, pf in enumerate(files)
    )
    body = (
        f'<a id="cf-top"></a>'
        f'<header class="cf-header">'
        f'<h1>{_escape(title)}</h1>'
        f'<p>{first_seq}〜{last_seq} 件目の文書をまとめて表示 '
        f'(batch {batch_no}/{total_batches}, {len(files)} files)</p>'
        f'<p class="small">Job: {_escape(job.job_name)} ({_escape(job.source_type)})'
        f' · {_escape(job.run_started_at)}</p>'
        f'</header>'
        f'<nav class="cf-toc"><strong>目次</strong>{toc_links}</nav>'
        f'<main class="cf-main">{blocks}</main>'
        f'<footer class="cf-footer">stdharvest combine_full</footer>'
    )
    full = (
        '<!DOCTYPE html>\n'
        '<html lang="ja">\n<head>\n<meta charset="UTF-8"/>\n'
        f'<title>{_escape(title)}</title>\n'
        f'<style>{_COMBINE_FULL_CSS}</style>\n'
        '</head>\n<body>\n'
        f'{body}\n'
        '</body>\n</html>\n'
    )
    out_path.write_text(full, encoding="utf-8")


def build_combine_full_html(
    files: List[ProcessedFile],
    job: JobContext,
    settings: Settings,  # noqa: ARG001  (kept for API symmetry)
    chunk_size: int = COMBINE_FULL_CHUNK_SIZE,
) -> List[CombineFullBatch]:
    """Generate `html/combine_full/combine_full_NNN.html` (N files per page).

    Each page inlines the body of each individual HTML, with image / link paths
    rewritten so assets render correctly from the combine_full location.
    Errors on a single file do not stop the whole job.
    """
    if not files:
        return []
    logger.info("combine_full generation started")
    ensure_dir(job.html_combine_full_dir)

    chunk = max(1, int(chunk_size))
    ordered = sorted(files, key=lambda pf: (pf.seq, pf.row.row_no))
    batches: List[CombineFullBatch] = []
    total = (len(ordered) + chunk - 1) // chunk
    for batch_no, start in enumerate(range(0, len(ordered), chunk), start=1):
        group = ordered[start : start + chunk]
        out_path = job.html_combine_full_dir / f"combine_full_{batch_no:03d}.html"
        try:
            _combine_full_page(job, group, out_path, batch_no, total)
            logger.info("%s created: %d documents", out_path.name, len(group))
        except Exception as exc:  # noqa: BLE001
            logger.warning("failed to write %s: %s", out_path.name, exc)
            continue
        batches.append(
            CombineFullBatch(
                batch_no=batch_no,
                first_seq=group[0].seq,
                last_seq=group[-1].seq,
                combine_full_path=out_path,
                file_count=len(group),
            )
        )
    logger.info("combine_full generation completed (%d page(s))", len(batches))
    return batches
