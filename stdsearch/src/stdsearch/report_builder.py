"""Generate search_results.csv / search_results.html / search_summary.json."""
from __future__ import annotations

import csv
import html
import json
import logging
from collections import Counter
from pathlib import Path
from typing import Dict, List

from .models import SearchHit, SearchProject, SearchRule
from .utils import now_iso, rel_link

logger = logging.getLogger(__name__)


CSV_COLUMNS = [
    "project_name",
    "rule_name",
    "match_type",
    "scope",
    "terms",
    "source_type",
    "doc_type",
    "source_seq",
    "row_no",
    "document_title",
    "location_type",
    "location_no",
    "heading_context",
    "matched_text",
    "before_text",
    "after_text",
    "pdf_path",
    "html_path",
    "html_anchor",
    "source_file",
]


def _esc(v: object) -> str:
    return html.escape(str(v or ""), quote=True)


def write_csv(out_path: Path, project: SearchProject, hits: List[SearchHit]) -> None:
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_COLUMNS)
        for hit in hits:
            writer.writerow([
                project.project_name,
                hit.rule.rule_name,
                hit.rule.match_type,
                hit.rule.scope,
                " | ".join(hit.rule.terms),
                hit.doc.source_type,
                hit.doc.doc_type,
                hit.doc.seq,
                hit.doc.row_no,
                hit.doc.title,
                hit.location_type,
                hit.location_no,
                hit.heading_context,
                hit.matched_text,
                hit.before_text,
                hit.after_text,
                hit.doc.pdf_path,
                str(hit.doc.html_path),
                hit.html_anchor,
                hit.doc.source_file,
            ])


def _rule_summary(hits: List[SearchHit], rules: List[SearchRule]) -> List[Dict]:
    counts: Counter = Counter(h.rule.rule_name for h in hits)
    out = []
    for r in rules:
        if not r.use:
            continue
        out.append({
            "rule_name": r.rule_name,
            "match_type": r.match_type,
            "scope": r.scope,
            "terms": r.terms,
            "hits": int(counts.get(r.rule_name, 0)),
        })
    return out


def _doc_summary(hits: List[SearchHit]) -> List[Dict]:
    buckets: Dict[tuple, Dict] = {}
    for h in hits:
        key = (h.doc.seq, h.doc.title)
        slot = buckets.setdefault(key, {
            "seq": h.doc.seq,
            "row_no": h.doc.row_no,
            "title": h.doc.title,
            "doc_type": h.doc.doc_type,
            "html_path": str(h.doc.html_path),
            "pdf_path": h.doc.pdf_path,
            "hits": 0,
        })
        slot["hits"] += 1
    return sorted(buckets.values(), key=lambda d: (d["seq"], d["title"]))


def write_json(
    out_path: Path,
    project: SearchProject,
    hits: List[SearchHit],
    documents_count: int,
) -> None:
    payload = {
        "project_name": project.project_name,
        "job_folder": str(project.job_folder),
        "searched_at": now_iso(),
        "rules_count": len([r for r in project.rules if r.use]),
        "documents_count": documents_count,
        "hits_count": len(hits),
        "rules": _rule_summary(hits, project.rules),
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


_CSS = """
body { font-family: system-ui, "Segoe UI", sans-serif; margin: 2rem; color: #222; background: #fafafa; }
h1 { margin-bottom: 0.25rem; }
h2 { margin-top: 2rem; border-bottom: 1px solid #ddd; padding-bottom: 0.25rem; }
.meta { color: #555; font-size: 0.9rem; }
table { border-collapse: collapse; margin: 0.5rem 0 1.25rem 0; }
table.kv td { padding: 0.15rem 0.75rem; vertical-align: top; }
table.kv td.k { color: #666; white-space: nowrap; }
table.summary { width: 100%; background: white; border: 1px solid #e2e2e2; border-radius: 6px; }
table.summary th, table.summary td { padding: 0.35rem 0.75rem; text-align: left; border-bottom: 1px solid #eee; }
.hit { background: white; border: 1px solid #e2e2e2; border-radius: 6px; padding: 0.75rem 1rem; margin: 0.75rem 0; }
.hit .header { display: flex; gap: 0.5rem; flex-wrap: wrap; align-items: baseline; }
.hit .rule { display: inline-block; padding: 0.1rem 0.5rem; border-radius: 999px;
             background: #eef; color: #335; font-size: 0.8rem; }
.hit .doc-type { font-size: 0.8rem; color: #666; }
.hit .location { font-size: 0.85rem; color: #444; }
.hit .matched { margin: 0.5rem 0; padding: 0.5rem 0.75rem; background: #fff8d6; border-left: 3px solid #d4a700;
                white-space: pre-wrap; line-height: 1.5; }
.hit .context { color: #555; font-size: 0.9rem; }
.hit .heading { font-size: 0.85rem; color: #2a6df4; margin-bottom: 0.25rem; }
.hit .links a { margin-right: 0.75rem; }
mark { background: #ffd24d; padding: 0 0.1rem; }
a { color: #2a6df4; text-decoration: none; }
a:hover { text-decoration: underline; }
.small { color: #666; font-size: 0.85rem; }
"""


def _highlight(text: str, terms: List[str]) -> str:
    if not text:
        return ""
    escaped = _esc(text)
    # Use simple case-insensitive replacement at token level.
    import re
    for term in sorted([t for t in terms if t], key=len, reverse=True):
        if not term:
            continue
        pattern = re.compile(re.escape(_esc(term)), re.IGNORECASE)
        escaped = pattern.sub(lambda m: f"<mark>{m.group(0)}</mark>", escaped)
    return escaped


def _link(from_file: Path, target: str, label: str) -> str:
    if not target:
        return f'<span class="small">{_esc(label)}なし</span>'
    p = Path(target)
    if not p.exists():
        return f'<span class="small">{_esc(label)}なし</span>'
    return f'<a href="{_esc(rel_link(from_file, p))}">{_esc(label)}</a>'


def _jump_link(from_file: Path, hit: SearchHit) -> str:
    if not hit.doc.html_path or not hit.doc.html_path.exists():
        return ""
    base = rel_link(from_file, hit.doc.html_path)
    href = base + (hit.html_anchor or "")
    label = "ジャンプ" if hit.html_anchor else "個別HTML先頭へ"
    return f'<a href="{_esc(href)}">{_esc(label)}</a>'


def write_html(
    out_path: Path,
    project: SearchProject,
    hits: List[SearchHit],
    documents_count: int,
) -> None:
    rule_rows = []
    for row in _rule_summary(hits, project.rules):
        rule_rows.append(
            f"<tr><td>{_esc(row['rule_name'])}</td><td>{_esc(row['match_type'])}</td>"
            f"<td>{_esc(row['scope'])}</td><td>{_esc(' | '.join(row['terms']))}</td>"
            f"<td>{row['hits']}</td></tr>"
        )

    doc_rows = []
    for row in _doc_summary(hits):
        html_link = _link(out_path, row["html_path"], "個別HTML")
        pdf_link = _link(out_path, row["pdf_path"], "PDF")
        doc_rows.append(
            f"<tr><td>{row['seq']:03d}</td><td>{_esc(row['title'])}</td>"
            f"<td>{_esc(row['doc_type'])}</td><td>{row['hits']}</td>"
            f"<td>{html_link} / {pdf_link}</td></tr>"
        )

    hit_blocks: List[str] = []
    for hit in hits:
        heading = f'<div class="heading">{_esc(hit.heading_context)}</div>' if hit.heading_context else ""
        location = f"{_esc(hit.location_type)} {_esc(hit.location_no)}".strip()
        hit_blocks.append(f"""
<div class="hit">
  <div class="header">
    <span class="rule">{_esc(hit.rule.rule_name)}</span>
    <span class="doc-type">[{_esc(hit.doc.doc_type)}] {_esc(hit.doc.title)}</span>
    <span class="location">· {location}</span>
  </div>
  {heading}
  <div class="context small">{_highlight(hit.before_text, hit.rule.terms)}</div>
  <div class="matched">{_highlight(hit.matched_text, hit.rule.terms)}</div>
  <div class="context small">{_highlight(hit.after_text, hit.rule.terms)}</div>
  <div class="links small">
    {_link(out_path, hit.doc.pdf_path, 'PDF')}
    {_link(out_path, str(hit.doc.html_path), '個別HTML')}
    {_jump_link(out_path, hit)}
  </div>
</div>
""")

    body = f"""
<h1>stdsearch: {_esc(project.project_name)}</h1>
<p class="meta">JobFolder: {_esc(project.job_folder)}<br/>検索日時: {_esc(now_iso())}</p>

<h2>検索条件</h2>
<table class="summary">
  <thead><tr><th>RuleName</th><th>MatchType</th><th>Scope</th><th>Terms</th><th>Hits</th></tr></thead>
  <tbody>
  {''.join(rule_rows) if rule_rows else '<tr><td colspan="5" class="small">(検索条件なし)</td></tr>'}
  </tbody>
</table>

<h2>文書別ヒット件数</h2>
<table class="summary">
  <thead><tr><th>Seq</th><th>Title</th><th>DocType</th><th>Hits</th><th>Links</th></tr></thead>
  <tbody>
  {''.join(doc_rows) if doc_rows else '<tr><td colspan="5" class="small">(ヒットなし)</td></tr>'}
  </tbody>
</table>

<h2>検索結果 ({len(hits)} hits / {documents_count} documents)</h2>
{''.join(hit_blocks) if hit_blocks else '<p class="small">(ヒットなし)</p>'}
"""

    out_path.write_text(f"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8"/>
<title>stdsearch: {_esc(project.project_name)}</title>
<style>{_CSS}</style>
</head>
<body>
{body}
</body>
</html>
""", encoding="utf-8")
