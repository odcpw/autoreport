#!/usr/bin/env python3
"""Backfill missing findingText fields in a project_sidecar.json from the seed knowledge base.

Default behavior:
- Only fills rows that look like "tiroir" questions (row.customer.items length > 1).
- Only fills when workstate.findingText is empty/whitespace.
- Uses the repo's seed files in AutoBericht/data/seed based on the sidecar's project locale.
- Creates a timestamped backup in ./backup/ next to the sidecar file.

Usage:
  python3 tools/backfill_sidecar_findings.py /path/to/project_sidecar.json

Options:
  --all-empty     Fill *all* rows with empty findingText (not just tiroir rows).
  --dry-run       Print what would change without writing.
  --seed-dir DIR  Use seed directory containing knowledge_base_*.json (defaults to repo seed dir).
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Tuple


def _resolve_locale_key(locale: Optional[str]) -> str:
    base = str(locale or "de-CH").lower()
    if base.startswith("fr"):
        return "fr"
    if base.startswith("it"):
        return "it"
    return "de"


def _to_text(value: Any) -> str:
    if isinstance(value, list):
        return "\n".join("" if v is None else str(v) for v in value)
    if value is None:
        return ""
    return str(value)


def _extract_report_project(doc: Any) -> Optional[Dict[str, Any]]:
    if not isinstance(doc, dict):
        return None
    report = doc.get("report")
    if isinstance(report, dict):
        project = report.get("project")
        if isinstance(project, dict) and isinstance(project.get("chapters"), list):
            return project
    if isinstance(doc.get("chapters"), list):
        return doc
    return None


def _iter_project_rows(project: Dict[str, Any]) -> Iterable[Dict[str, Any]]:
    chapters = project.get("chapters")
    if not isinstance(chapters, list):
        return
    for chapter in chapters:
        if not isinstance(chapter, dict):
            continue
        rows = chapter.get("rows")
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                continue
            if row.get("kind") == "section":
                continue
            yield row


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def _now_iso() -> str:
    return _dt.datetime.now(tz=_dt.timezone.utc).isoformat().replace("+00:00", "Z")


def _timestamp_slug() -> str:
    # Match the app's backup format (replace : and .)
    return _dt.datetime.now(tz=_dt.timezone.utc).isoformat().replace("+00:00", "Z").replace(":", "-").replace(".", "-")


def _create_backup(sidecar_path: Path) -> Path:
    backup_dir = sidecar_path.parent / "backup"
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_name = f"project_sidecar_{_timestamp_slug()}.json"
    backup_path = backup_dir / backup_name
    shutil.copy2(sidecar_path, backup_path)
    return backup_path


def _build_library_map(knowledge_base: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    library = knowledge_base.get("library")
    entries = []
    if isinstance(library, dict):
        entries = library.get("entries") or []
    if not isinstance(entries, list):
        entries = []
    out: Dict[str, Dict[str, Any]] = {}
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        entry_id = entry.get("id")
        if entry_id is None:
            continue
        out[str(entry_id)] = entry
    return out


def backfill_sidecar_findings(
    sidecar_doc: Dict[str, Any],
    knowledge_base: Dict[str, Any],
    *,
    tiroir_only: bool,
) -> Tuple[int, int]:
    """Returns (filled_count, considered_count)."""

    project = _extract_report_project(sidecar_doc)
    if not project:
        raise ValueError("Could not find report project in sidecar.")

    lib = _build_library_map(knowledge_base)

    filled = 0
    considered = 0

    for row in _iter_project_rows(project):
        row_id = row.get("id")
        if row_id is None:
            continue
        customer = row.get("customer") if isinstance(row.get("customer"), dict) else {}
        items = customer.get("items") if isinstance(customer.get("items"), list) else []
        is_tiroir = len(items) > 1
        if tiroir_only and not is_tiroir:
            continue

        ws = row.get("workstate")
        if not isinstance(ws, dict):
            ws = {}
            row["workstate"] = ws

        existing = _to_text(ws.get("findingText")).strip()
        if existing:
            continue

        entry = lib.get(str(row_id))
        if not entry:
            continue
        seed_finding = _to_text(entry.get("finding"))
        if not seed_finding.strip():
            continue

        ws["findingText"] = seed_finding
        filled += 1
        considered += 1

    # Update sidecar meta timestamp, similar to app behavior
    meta = sidecar_doc.get("meta")
    if not isinstance(meta, dict):
        meta = {}
        sidecar_doc["meta"] = meta
    meta["updatedAt"] = _now_iso()

    return filled, considered


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Backfill empty sidecar findings from seed.")
    parser.add_argument("sidecar", type=Path, help="Path to project_sidecar.json")
    parser.add_argument("--all-empty", action="store_true", help="Fill all empty findings (not just tiroir rows).")
    parser.add_argument("--dry-run", action="store_true", help="Do not write; just report counts.")
    parser.add_argument(
        "--seed-dir",
        type=Path,
        default=None,
        help="Directory containing knowledge_base_de/fr/it.json (defaults to repo AutoBericht/data/seed).",
    )
    args = parser.parse_args(argv)

    sidecar_path: Path = args.sidecar
    if sidecar_path.is_dir():
        sidecar_path = sidecar_path / "project_sidecar.json"

    if not sidecar_path.exists():
        print(f"ERROR: sidecar not found: {sidecar_path}", file=sys.stderr)
        return 2

    sidecar_doc = _load_json(sidecar_path)
    project = _extract_report_project(sidecar_doc)
    locale = None
    if project and isinstance(project.get("meta"), dict):
        locale = project["meta"].get("locale")

    locale_key = _resolve_locale_key(locale)

    repo_root = Path(__file__).resolve().parents[1]
    seed_dir = args.seed_dir or (repo_root / "AutoBericht" / "data" / "seed")
    seed_path = seed_dir / f"knowledge_base_{locale_key}.json"
    if not seed_path.exists():
        print(f"ERROR: seed not found: {seed_path}", file=sys.stderr)
        return 2

    knowledge_base = _load_json(seed_path)

    tiroir_only = not args.all_empty

    before_doc = None
    if args.dry_run:
        before_doc = json.dumps(sidecar_doc, sort_keys=True)

    filled, considered = backfill_sidecar_findings(sidecar_doc, knowledge_base, tiroir_only=tiroir_only)

    if args.dry_run:
        after_doc = json.dumps(sidecar_doc, sort_keys=True)
        changed = before_doc != after_doc
        print(f"seed={seed_path}")
        print(f"locale={locale or 'unknown'} (key={locale_key})")
        print(f"mode={'tiroir-only' if tiroir_only else 'all-empty'}")
        print(f"filled={filled}")
        print(f"changed={changed}")
        return 0

    backup_path = _create_backup(sidecar_path)
    _write_json(sidecar_path, sidecar_doc)

    print(f"seed={seed_path}")
    print(f"backup={backup_path}")
    print(f"mode={'tiroir-only' if tiroir_only else 'all-empty'}")
    print(f"filled={filled}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
