"""Read manifest.json produced by stdharvest and build the list of target HTMLs."""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import List

from .models import TargetDocument

logger = logging.getLogger(__name__)


def load_manifest(job_folder: Path) -> dict:
    manifest_path = job_folder / "logs" / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"manifest.json not found under {manifest_path}")
    with open(manifest_path, "r", encoding="utf-8") as f:
        return json.load(f)


def collect_target_documents(job_folder: Path) -> List[TargetDocument]:
    manifest = load_manifest(job_folder)
    source_type = str(manifest.get("source_type", "") or "").lower()

    targets: List[TargetDocument] = []
    for entry in manifest.get("files", []):
        html_path_str = entry.get("html_path") or ""
        if not html_path_str:
            continue
        html_path = Path(html_path_str)
        if not html_path.exists():
            alt = job_folder / html_path_str
            if alt.exists():
                html_path = alt
            else:
                logger.warning("html_path not found on disk: %s", html_path_str)
                continue
        targets.append(
            TargetDocument(
                seq=int(entry.get("seq") or 0),
                row_no=int(entry.get("row_no") or 0),
                title=str(entry.get("title") or ""),
                source_type=source_type,
                source_file=str(entry.get("source_file") or ""),
                pdf_path=str(entry.get("pdf_path") or ""),
                html_path=html_path,
                size_bytes=int(entry.get("size_bytes") or 0),
                status=str(entry.get("status") or ""),
            )
        )
    logger.info("Loaded %d target document(s) from manifest.", len(targets))
    return targets
