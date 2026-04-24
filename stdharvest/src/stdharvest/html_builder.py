"""HTML generation: per-file page, 5-file combined page, and index page."""
from __future__ import annotations

import html
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Iterable, List, Optional

from . import models
from .models import CombinedBatch, JobContext, ProcessedFile, Settings
from .utils import ensure_dir, human_size, rel_link

logger = logging.getLogger(__name__)


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


def _docx_to_html_body(path: Path) -> Optional[str]:
    try:
        import mammoth  # lazy import
    except ImportError:
        return None
    try:
        with open(path, "rb") as f:
            result = mammoth.convert_to_html(f)
        body = result.value
        return body or None
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
