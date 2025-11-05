# Unified Project Schema Blueprint

## Overview
- Excel / VBA will merge the legacy *master report* and *Selbstbeurteilung* sources into one `project.json`.
- AutoBericht ingests this single file and acts as the working copy for findings, overrides, photos, and export flags.
- Export produces a fresh `project.json` plus downstream artefacts (PDF, PPTX) without mutating the original master templates.

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
    "berichtList": [{ "value": "1.1.3", "label": "1.1.3" }],
    "seminarList": ["PSA Basics", "Führungskräftetraining"],
    "topicList": ["Gefährdung", "Brandschutz"]
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
      "title": "Leitbild, Sicherheitsziele und Strategien",
      "pageSize": 5,
      "rows": [
        {
          "id": "1.1.1",
          "title": "Unternehmensleitbild, Führungsgrundsätze",
          "master": {
            "finding": "Das Unternehmen verfügt nicht über ein Unternehmensleitbild.",
            "recommendations": {
              "level1": "Eine Sicherheitscharta als ersten Schritt etablieren …",
              "level2": "Die vorhandenen Werte …",
              "level3": "Die dokumentierte Charta …",
              "level4": "Die Charta als lebendiges Instrument …"
            }
          },
          "customer": {
            "answer": 0,
            "remark": "Leitbild in Vorbereitung, Umsetzung 2025",
            "priority": null
          },
          "workstate": {
            "selectedLevel": 4,
            "useFindingOverride": false,
            "useLevelOverride": {"1": false, "2": false, "3": false, "4": false},
            "findingOverride": "",
            "levelOverrides": {"1": "", "2": "", "3": "", "4": ""},
            "includeFinding": true,
            "includeRecommendation": true,
            "overwriteMode": "append",
            "done": false,
            "notes": "",
            "lastEditedBy": "consultant@firma.ch",
            "lastEditedAt": "2025-02-11T09:05:00Z"
          },
          "assets": {
            "photos": ["site/leitbild.jpg"],
            "slides": [],
            "documents": []
          },
          "exportHints": {
            "renumberedId": null,
            "includeInReport": true
          }
        }
      ]
    },
    {
      "id": "4.8",
      "title": "Audit Beobachtungen",
      "pageSize": 5,
      "rows": [
        {
          "id": "4.8.base-001",
          "title": "Audit Beobachtung",
          "master": {
            "finding": "",
            "recommendations": {}
          },
          "customer": {
            "answer": 0,
            "remark": "",
            "priority": null
          },
          "workstate": {
            "selectedLevel": null,
            "useFindingOverride": true,
            "useLevelOverride": {"1": false, "2": false, "3": false, "4": false},
            "findingOverride": "Neues Thema: Lagerbereich unzureichend beleuchtet.",
            "levelOverrides": {"1": "", "2": "", "3": "", "4": "Beleuchtung innerhalb von 30 Tagen auf LED umrüsten."},
            "includeFinding": true,
            "includeRecommendation": true,
            "overwriteMode": "append",
            "done": false,
            "notes": ""
          },
          "assets": {
            "photos": ["photos/lager/licht.jpg"],
            "slides": [],
            "documents": []
          },
          "exportHints": {}
        }
      ]
    }
  ]
}
```

### Field Notes
- `version`: Increment when schema changes; enables backward-compatible upgrades.
- `meta`: Excel owns IDs and timestamps; AutoBericht updates `author` and `createdAt` on export if needed.
- `lists`: Free-text tag libraries shared between PhotoSorter and report rows.
- `photos`: Only metadata (notes, logical tags for Bericht/seminar/topic). Physical files remain on disk; VBA handles moving/renaming post-export.
- `chapters`: Ordered list matching the master structure. `pageSize` controls “rows per page” navigation.
- `rows.id`: Canonical identifier from the master (e.g., `1.3.10`). Newly added UI rows use a generated suffix (`4.8.custom-001`) so Excel/VBA can track them.
- `master`: Raw template content, never edited in place.
- `customer`: Selbstbeurteilung answer (`1`/`0`), remark, and optional priority.
- `workstate`: UI-driven settings — level selection, include toggles, completion flag, overwrite/append preference, free-form notes, plus `useFindingOverride` / `useLevelOverride` maps controlling whether adjusted texts replace the master versions.
- `assets`: References to photos/slides by filename only; renumbering at export updates captions but not file keys.
- `exportHints`: Populated during export (e.g., `renumberedId` → `1.3.7`). AutoBericht writes this to the output JSON so VBA can align assets and numbering.

## Special Handling
- **Chapter 4.8 (Audit Beobachtungen):** Imported rows come from the master, but the UI may clone or add additional rows with generated IDs. Only rows marked for inclusion are renumbered into `4.8.1`, `4.8.2`, … during export.
- **Renumbering:** Happens only at export time. UI always shows canonical IDs so consultants keep the master reference; the export payload records the renumbered sequence for downstream Word/PPT/photo alignment.
- **Photos:** Tagging happens here; VBA reads the `photos` block and rearranges files on disk according to the renumbered IDs produced in the export payload.

## Next Steps
1. Finish the Excel macro to emit `project.json` conforming to this structure.
2. Update AutoBericht importer/state to hydrate directly from the unified payload.
3. Implement UI actions for level toggles, overrides, inclusion filters, and ad-hoc row creation within chapter 4.8.
4. Add export routine that writes a fresh `project.json` plus renumbered lists for Word/PPT and photo relocation.
