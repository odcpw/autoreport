#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

OPUS_DOCX = Path("fromWork/extracted/2025-10-02 OPUS 4.1 recommendations.docx")
SELBST_XLSX = Path("fromWork/extracted/Selbstbeurteilung Integrierte Sichereheit d.V14.xlsx")
OUTPUT_DIR = Path("data/seed")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

NS_WORD = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
NS_SHEET = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

ID_RE = re.compile(r"^\d+(?:\.\d+)+(?:\.[a-z])?\.?$", re.IGNORECASE)
LEVEL_RE = re.compile(r"^level\s*([1-4])$", re.IGNORECASE)
CHAPTER_PLACEHOLDER_RANGE = range(11, 15)
SPECIAL_COLLAPSE = {
    ("9", "5"): {"2", "3", "4", "5", "6", "7"},
    ("9", "9"): {"2", "3", "4", "5", "6", "7", "8", "9"},
}


@dataclass
class OpusEntry:
    id: str
    group_id: str
    finding: str
    levels: dict[str, list[str]]


def normalize_id(raw: str) -> str:
    cleaned = raw.strip().rstrip(".")
    return cleaned


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


def parse_docx_paragraphs(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    paras: list[str] = []
    for p in root.findall(".//w:p", NS_WORD):
        texts = [t.text for t in p.findall(".//w:t", NS_WORD) if t.text]
        line = "".join(texts).strip()
        if line:
            paras.append(line)
    return paras


def parse_opus(docx: Path) -> list[OpusEntry]:
    paras = parse_docx_paragraphs(docx)
    entries: list[OpusEntry] = []
    current: OpusEntry | None = None
    expecting_finding = False
    current_level: int | None = None
    auto_level = 1
    marker_seen = False

    for line in paras:
        if ID_RE.match(line):
            item_id = normalize_id(line)
            # Skip chapter headings like "1.1" that are not actual questions
            parts = [p for p in item_id.split(".") if p]
            if len(parts) < 3:
                current = None
                expecting_finding = False
                current_level = None
                auto_level = 1
                marker_seen = False
                continue

            current = OpusEntry(
                id=item_id,
                group_id=group_id_for(item_id),
                finding="",
                levels={"1": [], "2": [], "3": [], "4": []},
            )
            entries.append(current)
            expecting_finding = True
            current_level = None
            auto_level = 1
            marker_seen = False
            continue

        if current is None:
            continue

        if expecting_finding:
            current.finding = line
            expecting_finding = False
            continue

        level_match = LEVEL_RE.match(line)
        if level_match:
            marker_seen = True
            current_level = int(level_match.group(1))
            continue

        if marker_seen:
            if current_level is None:
                current_level = 1
            current.levels[str(current_level)].append(line)
            continue

        # No level markers present: assign each paragraph to next level.
        if auto_level > 4:
            auto_level = 4
        current.levels[str(auto_level)].append(line)
        auto_level += 1

    return entries


def parse_selbstbeurteilung_ids(xlsx: Path) -> dict[str, dict[str, str]]:
    with zipfile.ZipFile(xlsx) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = []
        for sheet in wb.findall("main:sheets/main:sheet", NS_SHEET):
            name = sheet.attrib.get("name")
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            sheets.append((name, rid))

        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("rel:Relationship", NS_REL)}

        target_name = next((name for name, _ in sheets if "selbst" in name.lower()), sheets[0][0])
        target_rid = next(rid for name, rid in sheets if name == target_name)
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

        results: dict[str, dict[str, str]] = {}
        for row_map in rows.values():
            item_id = row_map.get("A", "").strip()
            question = row_map.get("C", "").strip()
            if not item_id or not question:
                continue
            cleaned = normalize_id(item_id)
            if not ID_RE.match(item_id.strip()):
                # skip headers like "1." or "1.1."
                continue
            results[cleaned] = {
                "question": question,
                "chapterLabel": row_map.get("B", "").strip(),
            }
        return results


def build_coverage(entries: list[OpusEntry], selbst_ids: dict[str, dict[str, str]]) -> dict[str, object]:
    opus_ids = {e.id for e in entries}
    selbst_set = set(selbst_ids.keys())

    missing_in_opus = sorted(selbst_set - opus_ids)
    extra_in_opus = sorted(opus_ids - selbst_set)

    # tiroir groups
    tiroir_ids = sorted([i for i in selbst_set if i.endswith(tuple("abcdefghijklmnopqrstuvwxyz"))])
    group_map = {}
    for item_id in tiroir_ids:
        group_map.setdefault(group_id_for(item_id), []).append(item_id)

    opus_groups = {e.group_id for e in entries}
    missing_groups = sorted([gid for gid in group_map if gid not in opus_groups])

    # check missing levels
    missing_levels = []
    for entry in entries:
        for level in ("1", "2", "3", "4"):
            if not entry.levels.get(level):
                missing_levels.append({"id": entry.id, "level": level})

    def split_scope(items: set[str]) -> dict[str, list[str]]:
        in_scope = []
        placeholders = []
        other = []
        for item in sorted(items):
            chapter = chapter_for(item)
            if chapter is None:
                other.append(item)
            elif chapter in CHAPTER_PLACEHOLDER_RANGE:
                placeholders.append(item)
            elif 1 <= chapter <= 10:
                in_scope.append(item)
            else:
                other.append(item)
        return {
            "in_scope": in_scope,
            "placeholders": placeholders,
            "other": other,
        }

    # Collapsed coverage: merge tiroir sub-IDs and 4th-level IDs to 3rd-level.
    opus_collapsed = {collapse_id(item_id) for item_id in opus_ids}
    selbst_collapsed = {collapse_id(item_id) for item_id in selbst_set}
    missing_collapsed = sorted(selbst_collapsed - opus_collapsed)
    extra_collapsed = sorted(opus_collapsed - selbst_collapsed)

    return {
        "opus_count": len(opus_ids),
        "selbst_count": len(selbst_set),
        "missing_in_opus": missing_in_opus,
        "extra_in_opus": extra_in_opus,
        "tiroir_groups": group_map,
        "missing_tiroir_groups_in_opus": missing_groups,
        "missing_levels": missing_levels,
        "scope_raw": split_scope(selbst_set),
        "scope_opus_raw": split_scope(opus_ids),
        "collapsed": {
            "opus_count": len(opus_collapsed),
            "selbst_count": len(selbst_collapsed),
            "missing_in_opus": missing_collapsed,
            "extra_in_opus": extra_collapsed,
            "scope_selbst": split_scope(selbst_collapsed),
            "scope_opus": split_scope(opus_collapsed),
        },
    }


def write_outputs(entries: list[OpusEntry], selbst_ids: dict[str, dict[str, str]], coverage: dict[str, object]):
    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    output = {
        "meta": {
            "source": OPUS_DOCX.name,
            "importedAt": timestamp,
        },
        "entries": [
            {
                "id": e.id,
                "groupId": e.group_id,
                "finding": e.finding,
                "levels": e.levels,
            }
            for e in entries
        ],
    }

    library_path = OUTPUT_DIR / "library_master.json"
    library_path.write_text(json.dumps(output, indent=2, ensure_ascii=False), encoding="utf-8")

    report_path = OUTPUT_DIR / "coverage-report.md"
    collapsed = coverage["collapsed"]
    scope_raw = coverage["scope_raw"]
    scope_opus_raw = coverage["scope_opus_raw"]
    scope_selbst_collapsed = collapsed["scope_selbst"]
    scope_opus_collapsed = collapsed["scope_opus"]
    report_lines = [
        "# OPUS Coverage Report",
        "",
        f"Source: `{OPUS_DOCX}`",
        f"Generated: `{timestamp}`",
        "",
        f"- OPUS entries: **{coverage['opus_count']}**",
        f"- Selbstbeurteilung IDs: **{coverage['selbst_count']}**",
        f"- Missing in OPUS: **{len(coverage['missing_in_opus'])}**",
        f"- Extra in OPUS: **{len(coverage['extra_in_opus'])}**",
        "",
        "Scope: chapters 1-10 = in-scope, 11-14 = placeholders.",
        "Collapsed coverage merges tiroir sub-IDs (a/b/...) and 4th-level IDs to 3rd-level.",
        "Custom collapse: 9.5.2-9.5.7 -> 9.5.1, 9.9.2-9.9.9 -> 9.9.1.",
        f"Placeholder export: `{OUTPUT_DIR / 'placeholders.json'}`",
        "",
        "## Scope Summary (raw IDs)",
        f"- Selbstbeurteilung in-scope: **{len(scope_raw['in_scope'])}**",
        f"- Selbstbeurteilung placeholders: **{len(scope_raw['placeholders'])}**",
        f"- Selbstbeurteilung other: **{len(scope_raw['other'])}**",
        f"- OPUS in-scope: **{len(scope_opus_raw['in_scope'])}**",
        f"- OPUS placeholders: **{len(scope_opus_raw['placeholders'])}**",
        f"- OPUS other: **{len(scope_opus_raw['other'])}**",
        "",
        "## Missing in OPUS (present in Selbstbeurteilung)",
    ]
    if coverage["missing_in_opus"]:
        report_lines += [f"- {item}" for item in coverage["missing_in_opus"]]
    else:
        report_lines.append("- None")

    report_lines += ["", "## Extra in OPUS (not in Selbstbeurteilung)"]
    if coverage["extra_in_opus"]:
        report_lines += [f"- {item}" for item in coverage["extra_in_opus"]]
    else:
        report_lines.append("- None")

    report_lines += ["", "## Missing in OPUS (in-scope 1-10)"]
    if scope_raw["in_scope"]:
        report_lines += [f"- {item}" for item in scope_raw["in_scope"] if item in coverage["missing_in_opus"]]
        if not any(item in coverage["missing_in_opus"] for item in scope_raw["in_scope"]):
            report_lines.append("- None")
    else:
        report_lines.append("- None")

    report_lines += ["", "## Missing in OPUS (placeholders 11-14)"]
    if scope_raw["placeholders"]:
        report_lines += [f"- {item}" for item in scope_raw["placeholders"] if item in coverage["missing_in_opus"]]
        if not any(item in coverage["missing_in_opus"] for item in scope_raw["placeholders"]):
            report_lines.append("- None")
    else:
        report_lines.append("- None")

    report_lines += ["", "## Tiroir Groups (from Selbstbeurteilung)"]
    for group_id, members in coverage["tiroir_groups"].items():
        report_lines.append(f"- {group_id}: {', '.join(members)}")

    report_lines += ["", "## Missing Tiroir Groups in OPUS"]
    if coverage["missing_tiroir_groups_in_opus"]:
        report_lines += [f"- {gid}" for gid in coverage["missing_tiroir_groups_in_opus"]]
    else:
        report_lines.append("- None")

    report_lines += ["", "## Missing Recommendation Levels"]
    if coverage["missing_levels"]:
        for item in coverage["missing_levels"]:
            report_lines.append(f"- {item['id']} (Level {item['level']})")
    else:
        report_lines.append("- None")

    report_lines += [
        "",
        "## Collapsed Coverage Summary (letters + 4th-level -> 3rd-level)",
        f"- OPUS collapsed entries: **{collapsed['opus_count']}**",
        f"- Selbstbeurteilung collapsed IDs: **{collapsed['selbst_count']}**",
        f"- Missing in OPUS (collapsed): **{len(collapsed['missing_in_opus'])}**",
        f"- Extra in OPUS (collapsed): **{len(collapsed['extra_in_opus'])}**",
        "",
        "## Missing in OPUS (collapsed, in-scope 1-10)",
    ]
    if scope_selbst_collapsed["in_scope"]:
        report_lines += [
            f"- {item}"
            for item in scope_selbst_collapsed["in_scope"]
            if item in collapsed["missing_in_opus"]
        ]
        if not any(item in collapsed["missing_in_opus"] for item in scope_selbst_collapsed["in_scope"]):
            report_lines.append("- None")
    else:
        report_lines.append("- None")

    report_lines += ["", "## Missing in OPUS (collapsed, placeholders 11-14)"]
    if scope_selbst_collapsed["placeholders"]:
        report_lines += [
            f"- {item}"
            for item in scope_selbst_collapsed["placeholders"]
            if item in collapsed["missing_in_opus"]
        ]
        if not any(item in collapsed["missing_in_opus"] for item in scope_selbst_collapsed["placeholders"]):
            report_lines.append("- None")
    else:
        report_lines.append("- None")

    report_path.write_text("\n".join(report_lines), encoding="utf-8")

    placeholder_payload: dict[str, object] = {
        "meta": {
            "source": SELBST_XLSX.name,
            "importedAt": timestamp,
            "notes": "Placeholder chapters 11-14 derived from Selbstbeurteilung (raw IDs).",
        },
        "chapters": {},
    }
    for item_id, meta in selbst_ids.items():
        chapter = chapter_for(item_id)
        if chapter in CHAPTER_PLACEHOLDER_RANGE:
            chapter_key = str(chapter)
            chapter_entries = placeholder_payload["chapters"].setdefault(chapter_key, [])
            chapter_entries.append(
                {
                    "id": item_id,
                    "collapsedId": collapse_id(item_id),
                    "question": meta.get("question", ""),
                }
            )

    for chapter_key, entries in placeholder_payload["chapters"].items():
        entries.sort(key=lambda item: item.get("id", ""))

    placeholders_path = OUTPUT_DIR / "placeholders.json"
    placeholders_path.write_text(json.dumps(placeholder_payload, indent=2, ensure_ascii=False), encoding="utf-8")


def main():
    if not OPUS_DOCX.exists():
        raise SystemExit(f"Missing {OPUS_DOCX}")
    if not SELBST_XLSX.exists():
        raise SystemExit(f"Missing {SELBST_XLSX}")

    entries = parse_opus(OPUS_DOCX)
    selbst_ids = parse_selbstbeurteilung_ids(SELBST_XLSX)
    coverage = build_coverage(entries, selbst_ids)
    write_outputs(entries, selbst_ids, coverage)

    print(f"Wrote {OUTPUT_DIR / 'library_master.json'}")
    print(f"Wrote {OUTPUT_DIR / 'coverage-report.md'}")
    print(f"Wrote {OUTPUT_DIR / 'placeholders.json'}")


if __name__ == "__main__":
    main()
