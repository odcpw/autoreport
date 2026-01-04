# Unified Project Schema Blueprint

## Overview
- Excel / VBA will merge the legacy *master report* and *Selbstbeurteilung* sources into one `project.json`.
- AutoBericht ingests this single file and acts as the working copy for findings, overrides, photos, and export flags.
- Export produces a fresh `project.json` plus downstream artefacts (PDF, PPTX) without mutating the original master templates.

## JSON Schema

Machine-readable schema: `docs/reference/project.schema.json`

## Data Model
```json
{
  "version": 1,
  "meta": {
    "projectId": "2025-ACME-001",
    "company": "ACME AG",
    "createdAt": "2025-02-11T08:45:00Z",
    "locale": "de-CH",
    "author": "Consultant Name"
  },
  "lists": {
    "photo.bericht": [
      {
        "value": "1.1.3",
        "label": "1.1.3",
        "labels": { "de": "1.1.3", "fr": "1.1.3", "it": "1.1.3", "en": "1.1.3" },
        "group": "bericht",
        "sortOrder": 10,
        "chapterId": "1.1"
      }
    ],
    "photo.seminar": [{ "value": "PSA Basics", "label": "PSA Basics" }],
    "photo.topic": [{ "value": "Gefährdung", "label": "Gefährdung" }]
  },
  "photos": {
    "site/leitbild.jpg": {
      "notes": "Neue Leitbild-Tafel im Empfang",
      "tags": {
        "bericht": ["1.1.3"],
        "seminar": ["Führungskräftetraining"],
        "topic": ["Gefährdung"]
      }
    }
  },
  "chapters": [
    {
      "id": "1",
      "parentId": "",
      "orderIndex": 1,
      "title": {
        "de": "Leitbild, Sicherheitsziele und Strategien",
        "fr": "",
        "it": "",
        "en": ""
      },
      "pageSize": 5,
      "isActive": true,
      "rows": [
        {
          "id": "1.1.1",
          "chapterId": "1",
          "titleOverride": "Unternehmensleitbild, Führungsgrundsätze",
          "master": {
            "finding": "Das Unternehmen verfügt nicht über ein Unternehmensleitbild.",
            "levels": {
              "1": "Eine Sicherheitscharta als ersten Schritt etablieren …",
              "2": "Die vorhandenen Werte …",
              "3": "Die dokumentierte Charta …",
              "4": "Die Charta als lebendiges Instrument …"
            }
          },
          "overrides": {
            "finding": { "text": "", "enabled": false },
            "levels": {
              "1": { "text": "", "enabled": false },
              "2": { "text": "", "enabled": false },
              "3": { "text": "", "enabled": false },
              "4": { "text": "", "enabled": false }
            }
          },
          "customer": {
            "answer": 0,
            "remark": "Leitbild in Vorbereitung, Umsetzung 2025",
            "priority": null
          },
          "workstate": {
            "selectedLevel": 4,
            "includeFinding": true,
            "includeRecommendation": true,
            "overwriteMode": "append",
            "done": false,
            "notes": "",
            "lastEditedBy": "consultant@firma.ch",
            "lastEditedAt": "2025-02-11T09:05:00Z",
            "findingOverride": "",
            "useFindingOverride": false,
            "levelOverrides": {"1": "", "2": "", "3": "", "4": ""},
            "useLevelOverride": {"1": false, "2": false, "3": false, "4": false}
          }
        }
      ]
    },
    {
      "id": "4.8",
      "parentId": "4",
      "orderIndex": 999,
      "title": { "de": "Audit Beobachtungen", "fr": "", "it": "", "en": "" },
      "pageSize": 5,
      "isActive": true,
      "rows": [
        {
          "id": "4.8.base-001",
          "chapterId": "4.8",
          "titleOverride": "Audit Beobachtung",
          "master": {
            "finding": "",
            "levels": { "1": "", "2": "", "3": "", "4": "" }
          },
          "customer": {
            "answer": 0,
            "remark": "",
            "priority": null
          },
          "workstate": {
            "selectedLevel": null,
            "includeFinding": true,
            "includeRecommendation": true,
            "overwriteMode": "append",
            "done": false,
            "notes": "",
            "findingOverride": "Neues Thema: Lagerbereich unzureichend beleuchtet.",
            "useFindingOverride": true,
            "levelOverrides": {"1": "", "2": "", "3": "", "4": "Beleuchtung innerhalb von 30 Tagen auf LED umrüsten."},
            "useLevelOverride": {"1": false, "2": false, "3": false, "4": false}
          }
        }
      ]
    }
  ]
}
```

### Field Notes
- `version`: Increment when schema changes; enables backward-compatible upgrades.
- `meta`: Excel owns IDs and timestamps; AutoBericht updates `author` and `createdAt` on export if needed.
- `lists`: Tag/button libraries; keys are list names from the `Lists` sheet (e.g. `photo.bericht`, `photo.seminar`, `photo.topic`).
- `photos`: Only metadata (notes, logical tags for Bericht/seminar/topic). Physical files remain on disk; VBA handles moving/renaming post-export.
- `chapters`: Ordered list matching the master structure. `title` is a locale map (from `defaultTitle_{de,fr,it,en}`).
- `rows.id`: Canonical identifier from the master (e.g., `1.3.10`). Newly added UI rows use a generated suffix (`4.8.custom-001`) so Excel/VBA can track them.
- `master`: Raw template content; level text lives under `master.levels["1".."4"]`.
- `customer`: Selbstbeurteilung answer (`1`/`0`), remark, and optional priority.
- `overrides`: Explicit override blocks with `{text, enabled}` for finding + per-level recommendations (mirrors the structured sheet override columns).
- `workstate`: UI-driven settings — level selection, include toggles, completion flag, overwrite/append preference, free-form notes, plus `useFindingOverride` / `useLevelOverride` maps (duplicating override enablement for convenience/compatibility).

## Special Handling
- **Chapter 4.8 (Audit Beobachtungen):** Imported rows come from the master, but the UI may clone or add additional rows with generated IDs. Only rows marked for inclusion are renumbered into `4.8.1`, `4.8.2`, … during export.
- **Renumbering:** Happens only at export time. UI always shows canonical IDs so consultants keep the master reference; the export payload records the renumbered sequence for downstream Word/PPT/photo alignment.
- **Photos:** Tagging happens here; VBA reads the `photos` block and rearranges files on disk according to the renumbered IDs produced in the export payload.

## Next Steps
1. Treat the VBA exporter/loader as canonical and keep this document aligned with `macros/modABProjectExport.bas`.
2. Update AutoBericht importer/state to hydrate directly from this payload and write edits back into the same shape.
