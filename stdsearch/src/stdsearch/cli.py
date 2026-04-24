"""CLI entry point for stdsearch.

Usage:
    python -m stdsearch run --excel samples/sample_search.xlsx
    stdsearch run --excel samples/sample_search.xlsx
"""
from __future__ import annotations

import argparse
import logging
import sys
import traceback
from pathlib import Path
from typing import Dict, List

from . import __version__
from .excel_io import read_search_excel
from .html_parser import ParsedDocument, parse_html
from .manifest_loader import collect_target_documents
from .matcher import match_rule
from .models import SearchHit, SearchProject, SearchRule
from .report_builder import write_csv, write_html, write_json
from .splitter import split_document
from .utils import ensure_dir, now_compact_date

logger = logging.getLogger("stdsearch")


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def _resolve_output_dir(project: SearchProject) -> Path:
    if project.output_folder is not None:
        base = project.output_folder
    else:
        base = project.job_folder / "search"
    folder = base / f"{now_compact_date()}_{project.project_name}"
    return ensure_dir(folder)


def _iter_rules(project: SearchProject) -> List[SearchRule]:
    return [r for r in project.rules if r.use and r.terms]


def run(excel_path: Path) -> int:
    logger.info("stdsearch %s", __version__)
    logger.info("Loading Excel: %s", excel_path)
    project = read_search_excel(excel_path)

    active_rules = _iter_rules(project)
    logger.info(
        "Project=%s JobFolder=%s Rules(active)=%d",
        project.project_name, project.job_folder, len(active_rules),
    )
    if not project.job_folder.exists():
        raise FileNotFoundError(f"JobFolder not found: {project.job_folder}")
    if not active_rules:
        logger.warning("No active rules. Output will be empty.")

    docs = collect_target_documents(project.job_folder)
    if not docs:
        logger.warning("No HTML targets found in manifest.")
    else:
        logger.info("Target HTMLs: %d", len(docs))

    # Parse each HTML once; split lazily per rule scope.
    parsed_cache: Dict[Path, ParsedDocument] = {}
    for doc in docs:
        parsed = parse_html(doc)
        parsed_cache[doc.html_path] = parsed
        logger.debug(
            "Parsed %s (doc_type=%s, blocks=%d)",
            doc.html_path.name, doc.doc_type, len(parsed.blocks),
        )

    all_hits: List[SearchHit] = []
    # For efficiency, group rules by scope so we split each doc only once per scope.
    scope_groups: Dict[str, List[SearchRule]] = {}
    for rule in active_rules:
        scope_groups.setdefault(rule.scope, []).append(rule)

    for scope, rules in scope_groups.items():
        logger.info("Scope=%s  rules=%d", scope, len(rules))
        for parsed in parsed_cache.values():
            units = split_document(parsed, scope)
            if not units:
                continue
            for rule in rules:
                hits = match_rule(rule, units)
                if hits:
                    logger.debug(
                        "Rule=%s doc=%s hits=%d",
                        rule.rule_name, parsed.doc.html_path.name, len(hits),
                    )
                    all_hits.extend(hits)

    # Stable ordering of hits for output.
    all_hits.sort(key=lambda h: (
        h.rule.rule_name,
        h.doc.seq,
        h.location_no,
    ))

    output_dir = _resolve_output_dir(project)
    csv_path = output_dir / "search_results.csv"
    html_path = output_dir / "search_results.html"
    json_path = output_dir / "search_summary.json"

    write_csv(csv_path, project, all_hits)
    write_html(html_path, project, all_hits, documents_count=len(docs))
    write_json(json_path, project, all_hits, documents_count=len(docs))

    logger.info("Results:")
    logger.info("  CSV : %s", csv_path)
    logger.info("  HTML: %s", html_path)
    logger.info("  JSON: %s", json_path)
    logger.info("Total hits: %d  (rules=%d, documents=%d)",
                len(all_hits), len(active_rules), len(docs))
    return 0


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="stdsearch",
        description="Search tool for stdharvest-generated HTML documents.",
    )
    p.add_argument("--version", action="version", version=f"stdsearch {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    run_p = sub.add_parser("run", help="Run a search job from Search.xlsx.")
    run_p.add_argument(
        "--excel", required=True, type=Path,
        help="Path to Search.xlsx (Sheet1 carries config + rules)",
    )
    run_p.add_argument("-v", "--verbose", action="store_true", help="Verbose logging.")
    return p


def main(argv: List[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    _setup_logging(getattr(args, "verbose", False))
    try:
        if args.command == "run":
            return run(args.excel.resolve())
        parser.print_help()
        return 2
    except Exception as exc:  # noqa: BLE001
        logger.error("Fatal error: %s", exc)
        logger.debug("Traceback:\n%s", traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
