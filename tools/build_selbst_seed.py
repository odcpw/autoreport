#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

OUTPUT_DIR = Path("AutoBericht/data/seed")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

NS_SHEET = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

ID_RE = re.compile(r"^\d+(?:\.\d+)+(?:\.[a-z])?\.?$", re.IGNORECASE)
HEADER_RE = re.compile(r"^\d+(?:\.\d+)?\.?$", re.IGNORECASE)
CHAPTER_PLACEHOLDER_RANGE = range(11, 15)
STRUCTURE_VERSION = "2026-01-07"

SPECIAL_COLLAPSE = {
    ("9", "5"): {"2", "3", "4", "5", "6", "7"},
    ("9", "9"): {"2", "3", "4", "5", "6", "7", "8", "9"},
}

SHEET_PATTERNS = [
    "selbst",
    "autoévaluation",
    "autoevaluation",
    "autoéval",
    "autovalutazione",
    "autovalut",
    "auto-evaluation",
    "auto",
]


def normalize_id(raw: str) -> str:
    return raw.strip().rstrip(".")


def group_id_for(item_id: str) -> str:
    if item_id.endswith(tuple("abcdefghijklmnopqrstuvwxyz")):
        return item_id.rsplit(".", 1)[0]
    return item_id


def chapter_for(item_id: str) -> int | None:
    parts = item_id.split(".")
    if not parts:
        return None
    try:
        return int(parts[0])
    except ValueError:
        return None


def collapse_id(item_id: str) -> str:
    cleaned = group_id_for(item_id)
    parts = [p for p in cleaned.split(".") if p]
    if len(parts) > 3:
        parts = parts[:3]
        cleaned = ".".join(parts)

    if len(parts) >= 3:
        key = (parts[0], parts[1])
        targets = SPECIAL_COLLAPSE.get(key)
        if targets and parts[2] in targets:
            return f"{parts[0]}.{parts[1]}.1"

    return cleaned


def split_id_parts(item_id: str) -> list[str]:
    return [part for part in str(item_id).split(".") if part]


def choose_sheet(sheets: list[str]) -> str:
    lowered = [(name, name.lower()) for name in sheets]
    for pattern in SHEET_PATTERNS:
        for name, lower in lowered:
            if pattern in lower:
                return name
    return sheets[0]


def load_sheet_rows(xlsx: Path, sheet_name: str | None) -> tuple[str, dict[int, dict[str, str]]]:
    with zipfile.ZipFile(xlsx) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = []
        for sheet in wb.findall("main:sheets/main:sheet", NS_SHEET):
            name = sheet.attrib.get("name")
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            sheets.append((name, rid))

        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("rel:Relationship", NS_REL)}

        if sheet_name:
            target_rid = next(rid for name, rid in sheets if name == sheet_name)
            target_sheet = sheet_name
        else:
            target_sheet = choose_sheet([name for name, _ in sheets])
            target_rid = next(rid for name, rid in sheets if name == target_sheet)

        target_path = "xl/" + rid_map[target_rid]

        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sst.findall("main:si", NS_SHEET):
                text_parts = [t.text or "" for t in si.findall(".//main:t", NS_SHEET)]
                shared.append("".join(text_parts))

        sheet_xml = ET.fromstring(z.read(target_path))

        def cell_value(cell: ET.Element) -> str:
            t = cell.attrib.get("t")
            v = cell.find("main:v", NS_SHEET)
            if v is None:
                return ""
            raw = v.text or ""
            if t == "s":
                try:
                    return shared[int(raw)]
                except Exception:
                    return raw
            return raw

        rows: dict[int, dict[str, str]] = {}
        for row in sheet_xml.findall(".//main:row", NS_SHEET):
            r_idx = int(row.attrib.get("r"))
            row_map: dict[str, str] = {}
            for c in row.findall("main:c", NS_SHEET):
                ref = c.attrib.get("r")
                col = "".join(ch for ch in ref if ch.isalpha())
                row_map[col] = cell_value(c)
            rows[r_idx] = row_map

        return target_sheet, rows


def build_items(rows: dict[int, dict[str, str]]) -> list[dict[str, object]]:
    current_chapter_label = ""
    current_section_label = ""
    items: list[dict[str, object]] = []

    for row_idx in sorted(rows):
        row = rows[row_idx]
        raw_id = row.get("A", "").strip()
        label = row.get("B", "").strip()
        question = row.get("C", "").strip()

        if not raw_id:
            continue

        header_match = HEADER_RE.match(raw_id)
        item_match = ID_RE.match(raw_id)
        parts = [p for p in normalize_id(raw_id).split(".") if p]

        if header_match and len(parts) == 1 and label:
            current_chapter_label = f"{parts[0]} {label}".strip()
            current_section_label = ""
            continue

        if header_match and len(parts) == 2 and label:
            label_id = raw_id.strip()
            if not label_id.endswith("."):
                label_id = f"{label_id}."
            current_section_label = f"{label_id} {label}".strip()
            continue

        if not item_match or len(parts) < 3 or not question:
            continue

        item_id = normalize_id(raw_id)
        chapter = chapter_for(item_id)
        items.append(
            {
                "id": item_id,
                "groupId": group_id_for(item_id),
                "collapsedId": collapse_id(item_id),
                "chapter": chapter,
                "chapterLabel": current_chapter_label,
                "question": question,
                "sectionLabel": current_section_label,
            }
        )

    def sort_key(item: dict[str, object]) -> list[tuple[int, str]]:
        parts = str(item.get("id", "")).split(".")
        key = []
        for part in parts:
            if not part:
                continue
            match = re.match(r"^(\d+)([a-z]*)$", part, re.IGNORECASE)
            if match:
                key.append((int(match.group(1)), match.group(2).lower()))
            else:
                key.append((0, part.lower()))
        return key

    items.sort(key=sort_key)
    return items


def build_structure(items: list[dict[str, object]], source_name: str) -> dict[str, object]:
    group_map: dict[str, list[str]] = {}
    collapsed_map: dict[str, list[str]] = {}
    presentation_tags: set[str] = set()
    chapter_tags: set[str] = set()

    for item in items:
        group_id = str(item.get("groupId") or "")
        collapsed_id = str(item.get("collapsedId") or "")
        item_id = str(item.get("id") or "")
        if not group_id or not collapsed_id or not item_id:
            continue
        group_map.setdefault(group_id, []).append(item_id)
        collapsed_map.setdefault(collapsed_id, []).append(item_id)
        parts = split_id_parts(collapsed_id)
        if len(parts) >= 2:
            presentation_tags.add(".".join(parts[:2]))
        if parts:
            chapter_tags.add(parts[0])

    for group_id in group_map:
        group_map[group_id].sort()
    for collapsed_id in collapsed_map:
        collapsed_map[collapsed_id].sort()

    def sort_key(value: str) -> list[int]:
        return [int(p) if p.isdigit() else 0 for p in value.split(".")]

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    return {
        "meta": {
            "version": STRUCTURE_VERSION,
            "generatedAt": timestamp,
            "sources": {
                "selbstbeurteilung": source_name,
            },
        },
        "rules": {
            "collapse": {
                "tiroirSuffix": "a-z",
                "fourthLevelToThird": True,
                "special": {
                    "9.5": sorted(SPECIAL_COLLAPSE.get(("9", "5"), [])),
                    "9.9": sorted(SPECIAL_COLLAPSE.get(("9", "9"), [])),
                },
            },
            "presentationTagDepth": 2,
        },
        "items": items,
        "groupMap": group_map,
        "collapsedMap": collapsed_map,
        "presentationTags": sorted(presentation_tags, key=sort_key),
        "chapterTags": sorted(chapter_tags, key=sort_key),
    }


def build_placeholders(items: list[dict[str, object]], source_name: str) -> dict[str, object]:
    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    payload: dict[str, object] = {
        "meta": {
            "source": source_name,
            "importedAt": timestamp,
            "notes": "Placeholder chapters 11-14 derived from Selbstbeurteilung (raw IDs).",
        },
        "chapters": {},
    }
    for item in items:
        item_id = str(item.get("id") or "")
        chapter = chapter_for(item_id)
        if chapter in CHAPTER_PLACEHOLDER_RANGE:
            chapter_key = str(chapter)
            chapter_entries = payload["chapters"].setdefault(chapter_key, [])
            chapter_entries.append(
                {
                    "id": item_id,
                    "collapsedId": item.get("collapsedId"),
                    "question": item.get("question", ""),
                }
            )

    for chapter_key, entries in payload["chapters"].items():
        entries.sort(key=lambda item: item.get("id", ""))

    return payload


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Selbstbeurteilung seed JSON from an xlsx.")
    parser.add_argument("--xlsx", required=True, type=Path, help="Path to the Selbstbeurteilung Excel file")
    parser.add_argument("--sheet", help="Optional sheet name override")
    parser.add_argument("--locale", help="Optional locale tag for meta")
    parser.add_argument("--out", type=Path, help="Output path for selbstbeurteilung_ids JSON")
    parser.add_argument("--structure-out", type=Path, help="Output path for structure_manifest JSON")
    parser.add_argument("--placeholders-out", type=Path, help="Output path for placeholders JSON")
    args = parser.parse_args()

    if not args.xlsx.exists():
        raise SystemExit(f"Missing {args.xlsx}")

    target_sheet, rows = load_sheet_rows(args.xlsx, args.sheet)
    items = build_items(rows)

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    meta = {
        "source": args.xlsx.name,
        "importedAt": timestamp,
        "sheet": target_sheet,
    }
    if args.locale:
        meta["locale"] = args.locale

    payload = {
        "meta": meta,
        "items": items,
    }

    out_path = args.out
    if not out_path:
        suffix = f"_{args.locale}" if args.locale else ""
        out_path = OUTPUT_DIR / f"selbstbeurteilung_ids{suffix}.json"
    out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    if args.structure_out or args.placeholders_out:
        structure_payload = build_structure(items, args.xlsx.name)
        if args.locale:
            structure_payload["meta"]["locale"] = args.locale
        structure_path = args.structure_out
        if not structure_path:
            suffix = f"_{args.locale}" if args.locale else ""
            structure_path = OUTPUT_DIR / f"structure_manifest{suffix}.json"
        structure_path.write_text(json.dumps(structure_payload, indent=2, ensure_ascii=False), encoding="utf-8")

        placeholders_payload = build_placeholders(items, args.xlsx.name)
        if args.locale:
            placeholders_payload["meta"]["locale"] = args.locale
        placeholders_path = args.placeholders_out
        if not placeholders_path:
            suffix = f"_{args.locale}" if args.locale else ""
            placeholders_path = OUTPUT_DIR / f"placeholders{suffix}.json"
        placeholders_path.write_text(json.dumps(placeholders_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
