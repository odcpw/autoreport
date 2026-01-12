#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from datetime import datetime, timezone
from pathlib import Path

OUTPUT_DIR = Path("AutoBericht/data/seed")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

FOLLOWUP_FR = re.compile(r"\s*(?:,|;|\.|\?)?\s*si\s+oui\b.*$", re.IGNORECASE)
FOLLOWUP_IT = re.compile(r"\s*(?:,|;|\.|\?)?\s*se\s+s[ìi]\b.*$", re.IGNORECASE)

VOWELS_FR = set("aeiouyàâäæéèêëîïôöùûüÿh")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def strip_question_mark(text: str) -> str:
    return text.rstrip().rstrip("?")


def strip_followup(text: str, locale: str) -> str:
    if locale == "fr":
        return FOLLOWUP_FR.sub("", text).strip()
    if locale == "it":
        return FOLLOWUP_IT.sub("", text).strip()
    return text


def add_period(text: str) -> str:
    text = text.strip()
    if not text:
        return text
    if text.endswith("."):
        return text
    return f"{text}."


def fr_negate(question: str) -> str:
    text = normalize_text(question)
    text = strip_followup(text, "fr")
    text = strip_question_mark(text)

    # Y a-t-il / Existe-t-il
    match = re.match(r"^y a[- ]t-il\s+(.*)$", text, re.IGNORECASE)
    if match:
        rest = match.group(1).strip()
        return add_period(f"Il n'y a pas {rest}")

    match = re.match(r"^existe[- ]t-il\s+(.*)$", text, re.IGNORECASE)
    if match:
        rest = match.group(1).strip()
        return add_period(f"Il n'existe pas {rest}")

    # Inversion with hyphen: <subject> <verb>-t-il/elle/ils/elles
    match = re.match(r"^(?P<subject>.+?)\s+(?P<lemma>[A-Za-zÀ-ÿ'’]+)-(?:t-)?(?P<pron>il|elle|ils|elles)\b\s*(?P<rest>.*)$", text)
    if match:
        subject = match.group("subject").strip()
        verb = match.group("lemma").strip()
        rest = match.group("rest").strip()
        neg_prefix = "n'" if verb[:1].lower() in VOWELS_FR else "ne "
        phrase = f"{subject} {neg_prefix}{verb} pas"
        if rest:
            phrase = f"{phrase} {rest}"
        return add_period(phrase)

    # Subject + verb (no inversion)
    match = re.match(r"^(?P<subject>(?:L'|L’|La |Le |Les |L'entreprise|L’entreprise|La société|L'établissement|L’etablissement|L'organisation|L’organisation|La direction|Le personnel|Les collaborateurs|Les employés|Les salariés|Les travailleurs)[^ ]*)\s+(?P<lemma>[A-Za-zÀ-ÿ'’]+)\b\s*(?P<rest>.*)$", text)
    if match:
        subject = match.group("subject").strip()
        verb = match.group("lemma").strip()
        rest = match.group("rest").strip()
        neg_prefix = "n'" if verb[:1].lower() in VOWELS_FR else "ne "
        phrase = f"{subject} {neg_prefix}{verb} pas"
        if rest:
            phrase = f"{phrase} {rest}"
        return add_period(phrase)

    # Fallback
    return add_period(f"Il n'est pas vrai que {text}")


def it_negate(question: str) -> str:
    text = normalize_text(question)
    text = strip_followup(text, "it")
    text = strip_question_mark(text)

    # Leading verb forms
    for prefix, replacement in [
        (r"^esiste\b", "Non esiste"),
        (r"^esistono\b", "Non esistono"),
        (r"^c['’]è\b", "Non c'è"),
        (r"^ci sono\b", "Non ci sono"),
        (r"^vi è\b", "Non vi è"),
        (r"^è\b", "Non è"),
        (r"^e'\b", "Non è"),
        (r"^sono\b", "Non sono"),
        (r"^viene\b", "Non viene"),
        (r"^vengono\b", "Non vengono"),
    ]:
        if re.match(prefix, text, re.IGNORECASE):
            rest = re.sub(prefix, "", text, flags=re.IGNORECASE).strip()
            return add_period(f"{replacement} {rest}".strip())

    # Subject + verb
    subjects = [
        "L'azienda",
        "L’azienda",
        "L'impresa",
        "L’impresa",
        "La società",
        "La direzione",
        "La direzione aziendale",
        "Il personale",
        "I dipendenti",
        "I lavoratori",
        "I collaboratori",
        "I superiori",
        "Gli addetti",
        "Le persone",
    ]
    subject_pattern = "|".join(re.escape(s) for s in subjects)
    match = re.match(rf"^(?P<subject>{subject_pattern})\s+(?P<lemma>[A-Za-zÀ-ÿ'’]+)\b\s*(?P<rest>.*)$", text)
    if match:
        subject = match.group("subject").strip()
        verb = match.group("lemma").strip()
        rest = match.group("rest").strip()
        phrase = f"{subject} non {verb}"
        if rest:
            phrase = f"{phrase} {rest}"
        return add_period(phrase)

    # Fallback
    return add_period(f"Non è vero che {text}")


def sort_key(value: str) -> list[tuple[int, str]]:
    parts = str(value).split(".")
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Build library_master from selbstbeurteilung questions.")
    parser.add_argument("--selbst", required=True, type=Path, help="Path to selbstbeurteilung_ids JSON")
    parser.add_argument("--locale", required=True, choices=["fr", "it"], help="Locale for negation rules")
    parser.add_argument("--out", type=Path, help="Output path for library_master JSON")
    args = parser.parse_args()

    if not args.selbst.exists():
        raise SystemExit(f"Missing {args.selbst}")

    data = json.loads(args.selbst.read_text(encoding="utf-8"))
    items = data.get("items", [])

    groups: dict[str, list[dict[str, object]]] = {}
    for item in items:
        group_id = str(item.get("groupId") or item.get("id") or "").strip()
        if not group_id:
            continue
        groups.setdefault(group_id, []).append(item)

    entries = []
    for group_id, group_items in groups.items():
        group_items.sort(key=lambda item: sort_key(item.get("id", "")))
        preferred = next((item for item in group_items if str(item.get("id", "")) == group_id), None)
        item = preferred or group_items[0]
        question = str(item.get("question", "")).strip()
        if not question:
            continue
        finding = fr_negate(question) if args.locale == "fr" else it_negate(question)
        entries.append(
            {
                "id": group_id,
                "groupId": group_id,
                "finding": finding,
                "levels": {"1": [], "2": [], "3": [], "4": []},
                "children": [],
            }
        )

    entries.sort(key=lambda entry: sort_key(entry.get("id", "")))

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    source_name = str(data.get("meta", {}).get("source", "")) or args.selbst.name
    payload = {
        "meta": {
            "source": source_name,
            "importedAt": timestamp,
        },
        "entries": entries,
    }

    out_path = args.out
    if not out_path:
        suffix = "_f" if args.locale == "fr" else "_i"
        out_path = OUTPUT_DIR / f"library_master{suffix}.json"
    out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
