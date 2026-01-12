#!/usr/bin/env python3
"""
Preflight validator for Word ribbon embedding.

Checks:
- customUI part exists exactly once (customUI/customUI.xml)
- [Content_Types].xml has exactly one Override for /customUI/customUI.xml
- root _rels/.rels has exactly one ui/extensibility relationship pointing to customUI/customUI.xml (no leading slash)
- word/_rels/document.xml.rels must not contain a ui/extensibility relationship (avoids duplicates)
- customUI root namespace matches 2006/01 schema (not customUI14)
"""

from __future__ import annotations
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


def parse_xml(z: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(z.read(name))


def main():
    docm_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ProjectTemplate/Vorlage IST-Aufnahme-Bericht d.V01.docm")
    if not docm_path.exists():
        print(f"ERROR: file not found: {docm_path}", file=sys.stderr)
        return 1

    errors: list[str] = []
    warnings: list[str] = []

    with zipfile.ZipFile(docm_path) as z:
        names = z.namelist()

        custom_parts = [n for n in names if n.lower().endswith("customui/customui.xml")]
        if not custom_parts:
            errors.append("customUI part missing (expected customUI/customUI.xml)")
        elif len(custom_parts) > 1:
            errors.append(f"duplicate customUI parts: {custom_parts}")
        else:
            custom_part = custom_parts[0]

        # Content types
        ct_root = parse_xml(z, "[Content_Types].xml")
        overrides = [
            o for o in ct_root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")
            if o.attrib.get("PartName", "").lower() == "/customui/customui.xml"
        ]
        if len(overrides) != 1:
            errors.append(f"[Content_Types].xml Override count for /customUI/customUI.xml = {len(overrides)} (expected 1)")

        # Root relationships
        rels_root = parse_xml(z, "_rels/.rels")
        ui_rels = [
            r for r in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
            if r.attrib.get("Type") == "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
        ]
        if len(ui_rels) != 1:
            errors.append(f"_rels/.rels ui/extensibility relationship count = {len(ui_rels)} (expected 1)")
        else:
            target = ui_rels[0].attrib.get("Target", "")
            if target not in ("customUI/customUI.xml", "customui/customUI.xml"):
                errors.append(f"_rels/.rels ui/extensibility Target is '{target}' (expected customUI/customUI.xml without leading slash)")

        # Document rels should NOT contain ui/extensibility for customUI (avoids duplicate loading/precedence issues)
        if "word/_rels/document.xml.rels" in names:
            doc_rels = parse_xml(z, "word/_rels/document.xml.rels")
            doc_ui = [
                r for r in doc_rels.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
                if r.attrib.get("Type") == "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
            ]
            if doc_ui:
                warnings.append(f"word/_rels/document.xml.rels contains ui/extensibility relationship(s): {len(doc_ui)} (remove to rely on root-level customUI part)")

        # Namespace sanity on customUI part
        if custom_parts:
            ui_root = parse_xml(z, custom_part)
            if not ui_root.tag.endswith("customUI"):
                errors.append(f"customUI root tag unexpected: {ui_root.tag}")
            ns = ui_root.tag.split("}")[0].strip("{")
            if ns != "http://schemas.microsoft.com/office/2006/01/customui":
                errors.append(f"customUI namespace is {ns}, expected http://schemas.microsoft.com/office/2006/01/customui")

    if errors:
        print("Ribbon preflight FAILED:")
        for e in errors:
            print(f"- ERROR: {e}")
        for w in warnings:
            print(f"- WARN: {w}")
        return 1

    print("Ribbon preflight OK")
    if warnings:
        for w in warnings:
            print(f"- WARN: {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
