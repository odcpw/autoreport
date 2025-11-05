# project.json Schema Reference

Complete technical specification for the AutoBericht unified data format.

## Overview

`project.json` is the canonical data interchange format between Excel VBA and the web UI. It represents the complete state of a safety assessment project.

## Schema Version

Current version: **1**

Version increments when breaking changes are introduced. The system supports backward-compatible upgrades.

## Root Structure

```json
{
  "version": 1,
  "meta": { },
  "chapters": [ ],
  "photos": { },
  "lists": { },
  "history": [ ]
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `version` | integer | Yes | Schema version (currently 1) |
| `meta` | object | Yes | Project metadata |
| `chapters` | array | Yes | Hierarchical chapter structure with rows |
| `photos` | object | No | Photo metadata keyed by filename |
| `lists` | object | No | Tag vocabularies and button definitions |
| `history` | array | No | Append-only override change log |

## `meta` Object

Project-level metadata.

```json
{
  "meta": {
    "projectId": "2025-ACME-001",
    "company": "ACME AG",
    "createdAt": "2025-02-11T08:45:00Z",
    "locale": "de-CH",
    "author": "Consultant Name"
  }
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `projectId` | string | Yes | Unique project identifier |
| `company` | string | Yes | Company name for reports |
| `createdAt` | string (ISO 8601) | Yes | Project creation timestamp |
| `locale` | string | No | Default locale (de-CH, fr-CH, it-CH, en-CH) |
| `author` | string | No | Primary consultant/author name |

**Excel mapping**: `Meta` sheet (key/value pairs)

## `chapters` Array

Hierarchical report structure with nested rows.

```json
{
  "chapters": [
    {
      "id": "1",
      "parentId": "",
      "orderIndex": 1,
      "title": {
        "de": "Leitbild, Sicherheitsziele",
        "fr": "Politique de sécurité",
        "it": "Politica di sicurezza",
        "en": "Safety Policy"
      },
      "pageSize": 5,
      "isActive": true,
      "rows": [ ]
    }
  ]
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `id` | string | Yes | Chapter ID (e.g., "1", "1.1", "1.1.1") |
| `parentId` | string | No | Parent chapter ID (empty for root) |
| `orderIndex` | integer | Yes | Sort order within parent |
| `title` | object | Yes | Localized titles (de, fr, it, en) |
| `pageSize` | integer | No | Items per page (UI pagination hint) |
| `isActive` | boolean | No | Whether chapter is active (default: true) |
| `rows` | array | Yes | Finding rows within this chapter |

**Excel mapping**: `Chapters` sheet

### `chapters[].rows` Array

Individual findings and recommendations.

```json
{
  "rows": [
    {
      "id": "1.1.1",
      "chapterId": "1.1",
      "titleOverride": "",
      "master": { },
      "customer": { },
      "workstate": { },
      "assets": { },
      "exportHints": { }
    }
  ]
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `id` | string | Yes | Row ID (e.g., "1.1.1", "4.8.custom-001") |
| `chapterId` | string | Yes | Parent chapter ID |
| `titleOverride` | string | No | Custom title (overrides master) |
| `master` | object | Yes | Template content from Word |
| `customer` | object | No | Self-assessment data |
| `workstate` | object | No | Consultant edits and UI state |
| `assets` | object | No | Linked resources (photos, slides) |
| `exportHints` | object | No | Export-time metadata |

**Excel mapping**: `Rows` sheet

### `master` Object

Template content from master findings document.

```json
{
  "master": {
    "finding": "Das Unternehmen verfügt nicht über...",
    "recommendations": {
      "level1": "Eine Sicherheitscharta etablieren...",
      "level2": "Die vorhandenen Werte dokumentieren...",
      "level3": "Die Charta als lebendiges Instrument...",
      "level4": "Regelmässige Überprüfung und Anpassung..."
    }
  }
}
```

Alternative format (legacy compatibility):
```json
{
  "master": {
    "finding": "...",
    "levels": {
      "1": "...",
      "2": "...",
      "3": "...",
      "4": "..."
    }
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `finding` | string | Master finding text |
| `recommendations.level1-4` | string | Recommendation for each severity level |

**Excel mapping**: `Rows` sheet columns `masterFinding`, `masterLevel1-4`

### `customer` Object

Customer self-assessment data.

```json
{
  "customer": {
    "answer": 2,
    "remark": "Leitbild in Vorbereitung",
    "priority": null
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `answer` | integer | Self-assessment score (0-4, or null) |
| `remark` | string | Customer notes/comments |
| `priority` | string | Optional priority flag |

**Excel mapping**: `Rows` sheet columns `customerAnswer`, `customerRemark`, `customerPriority`

### `workstate` Object

Consultant edits and UI state.

```json
{
  "workstate": {
    "selectedLevel": 2,
    "useFindingOverride": false,
    "useLevelOverride": { "1": false, "2": false, "3": false, "4": false },
    "findingOverride": "",
    "levelOverrides": { "1": "", "2": "", "3": "", "4": "" },
    "includeFinding": true,
    "includeRecommendation": true,
    "overwriteMode": "append",
    "done": false,
    "notes": "",
    "lastEditedBy": "consultant@firma.ch",
    "lastEditedAt": "2025-02-11T09:05:00Z"
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `selectedLevel` | integer | Chosen recommendation level (1-4, or null) |
| `useFindingOverride` | boolean | Whether to use custom finding text |
| `useLevelOverride` | object | Per-level override flags (keys: "1"-"4") |
| `findingOverride` | string | Custom finding text (if enabled) |
| `levelOverrides` | object | Custom recommendation per level (if enabled) |
| `includeFinding` | boolean | Include finding in exports |
| `includeRecommendation` | boolean | Include recommendation in exports |
| `overwriteMode` | string | "append" or "replace" (for future use) |
| `done` | boolean | Consultant marked complete |
| `notes` | string | Internal notes (not exported) |
| `lastEditedBy` | string | Last editor identifier |
| `lastEditedAt` | string (ISO 8601) | Last edit timestamp |

**Excel mapping**: `Rows` sheet columns (various override and state fields)

### `assets` Object

Linked resources for the row.

```json
{
  "assets": {
    "photos": ["site/leitbild.jpg", "photos/entrance.jpg"],
    "slides": [],
    "documents": []
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `photos` | array of strings | Photo filenames (keys in `photos` object) |
| `slides` | array | Reserved for future use |
| `documents` | array | Reserved for future use |

### `exportHints` Object

Export-time metadata (populated during export).

```json
{
  "exportHints": {
    "renumberedId": "1.3.7",
    "includeInReport": true
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `renumberedId` | string | Export-time renumbered ID (for chapter 4.8) |
| `includeInReport` | boolean | Whether row was included in export |

## `photos` Object

Photo metadata keyed by filename.

```json
{
  "photos": {
    "site/leitbild.jpg": {
      "displayName": "Leitbild im Empfang",
      "notes": "Neue Tafel, gut sichtbar",
      "tags": {
        "bericht": ["1.1.3"],
        "seminar": ["Führungskräftetraining"],
        "topic": ["Gefährdung"]
      },
      "preferredLocale": "de-CH",
      "capturedAt": "2025-02-11T10:30:00Z"
    }
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `displayName` | string | Human-readable photo name |
| `notes` | string | Photo description/notes |
| `tags` | object | Tag assignments by dimension |
| `tags.bericht` | array | Bericht chapter IDs (e.g., ["1.1.3", "2.4.1"]) |
| `tags.seminar` | array | Seminar tags (e.g., ["PSA Basics"]) |
| `tags.topic` | array | Topic tags (e.g., ["Gefährdung"]) |
| `preferredLocale` | string | Locale hint for display |
| `capturedAt` | string (ISO 8601) | Photo capture timestamp |

**Excel mapping**: `Photos` sheet

## `lists` Object

Tag vocabularies and button definitions.

```json
{
  "lists": {
    "berichtList": [
      {
        "value": "1.1.3",
        "label": "1.1.3 — Leitbild",
        "labels": {
          "de": "1.1.3 — Leitbild",
          "fr": "1.1.3 — Vision",
          "it": "1.1.3 — Visione",
          "en": "1.1.3 — Mission"
        },
        "group": "bericht",
        "sortOrder": 1,
        "chapterId": "1.1.3"
      }
    ],
    "seminarList": [
      {
        "value": "psa-basics",
        "label": "PSA Basics",
        "labels": {
          "de": "PSA Grundlagen",
          "fr": "Bases EPI",
          "it": "Basi DPI",
          "en": "PPE Basics"
        },
        "group": "seminar",
        "sortOrder": 1,
        "chapterId": ""
      }
    ],
    "topicList": [
      {
        "value": "gefaehrdung",
        "label": "Gefährdung",
        "labels": {
          "de": "Gefährdung",
          "fr": "Danger",
          "it": "Pericolo",
          "en": "Hazard"
        },
        "group": "topic",
        "sortOrder": 1,
        "chapterId": ""
      }
    ]
  }
}
```

| Field | Type | Description |
|-------|------|-------------|
| `value` | string | Internal identifier |
| `label` | string | Display label (default locale) |
| `labels` | object | Localized labels (de, fr, it, en) |
| `group` | string | Button group (bericht, seminar, topic) |
| `sortOrder` | integer | Display order |
| `chapterId` | string | Optional chapter association |

**Common list names**:
- `berichtList`: Bericht chapter buttons
- `seminarList`: Seminar assignment buttons
- `topicList`: Topic buttons
- `photo.topic`: Legacy topic folders (deprecated)

**Excel mapping**: `Lists` sheet

## `history` Array

Append-only log of override changes.

```json
{
  "history": [
    {
      "timestamp": "2025-02-11T09:05:00Z",
      "rowId": "1.1.1",
      "field": "findingOverride",
      "oldValue": "",
      "newValue": "Neues Leitbild wurde erstellt",
      "user": "consultant@firma.ch"
    }
  ]
}
```

| Field | Type | Description |
|-------|------|-------------|
| `timestamp` | string (ISO 8601) | Change timestamp |
| `rowId` | string | Affected row ID |
| `field` | string | Field name that changed |
| `oldValue` | string | Previous value |
| `newValue` | string | New value |
| `user` | string | User identifier |

**Excel mapping**: `OverridesHistory` sheet

## Special Cases

### Chapter 4.8 (Audit Observations)

Custom rows not in master template:

```json
{
  "id": "4.8",
  "title": { "de": "Audit Beobachtungen" },
  "rows": [
    {
      "id": "4.8.custom-001",
      "master": {
        "finding": "",
        "recommendations": {}
      },
      "workstate": {
        "useFindingOverride": true,
        "findingOverride": "Neues Thema: ...",
        "useLevelOverride": { "4": true },
        "levelOverrides": { "4": "Empfehlung: ..." }
      }
    }
  ]
}
```

**Renumbering at export**:
- `4.8.custom-001` → `4.8.1`
- `4.8.custom-002` → `4.8.2`
- Original ID preserved in `exportHints.renumberedId`

### Empty Master Rows

Rows can exist with empty master content:

```json
{
  "id": "new-item",
  "master": {
    "finding": "",
    "recommendations": {}
  },
  "workstate": {
    "useFindingOverride": true,
    "findingOverride": "Custom content...",
    "selectedLevel": null
  }
}
```

## Validation Rules

### Required Fields

- `version`, `meta.projectId`, `meta.company`, `meta.createdAt`
- Each chapter: `id`, `title`, `rows`
- Each row: `id`, `chapterId`, `master`

### ID Formats

**Chapter IDs**:
- Root: "1", "2", "3", "4"
- Level 1: "1.1", "1.2", "4.8"
- Level 2: "1.1.1", "1.1.2"
- Level 3: "1.1.1.1" (rare)

**Row IDs**:
- Standard: Match chapter ID pattern ("1.1.1", "2.3.5")
- Custom: `{chapterId}.custom-{nnn}` (e.g., "4.8.custom-001")

### Tag References

- `photos.*.tags.chapters` must reference existing chapter IDs
- List item `chapterId` should reference existing chapters (soft constraint)

### Timestamps

All timestamps must be ISO 8601 format: `YYYY-MM-DDTHH:MM:SSZ`

## Excel Sheet Mapping

| JSON Path | Excel Sheet | Columns |
|-----------|-------------|---------|
| `meta` | Meta | key, value |
| `chapters` | Chapters | chapterId, parentId, orderIndex, defaultTitle_*, pageSize, isActive |
| `chapters[].rows` | Rows | rowId, chapterId, titleOverride, master*, customer*, workstate*, ... |
| `photos` | Photos | fileName, displayName, notes, tag*, preferredLocale, capturedAt |
| `lists` | Lists | listName, value, label_*, group, sortOrder, chapterId |
| `history` | OverridesHistory | timestamp, rowId, fieldName, oldValue, newValue, user |

## Example: Minimal Valid Project

```json
{
  "version": 1,
  "meta": {
    "projectId": "2025-TEST-001",
    "company": "Test Company",
    "createdAt": "2025-01-01T12:00:00Z"
  },
  "chapters": [
    {
      "id": "1",
      "parentId": "",
      "orderIndex": 1,
      "title": { "de": "Kapitel 1" },
      "pageSize": 5,
      "isActive": true,
      "rows": []
    }
  ],
  "photos": {},
  "lists": {},
  "history": []
}
```

## Example: Complete Row

```json
{
  "id": "1.1.1",
  "chapterId": "1.1",
  "titleOverride": "",
  "master": {
    "finding": "Das Unternehmen verfügt nicht über ein Leitbild.",
    "recommendations": {
      "level1": "Sicherheitscharta etablieren",
      "level2": "Vorhandene Werte dokumentieren",
      "level3": "Charta als lebendiges Instrument nutzen",
      "level4": "Regelmässige Überprüfung etablieren"
    }
  },
  "customer": {
    "answer": 2,
    "remark": "Leitbild in Arbeit, Fertigstellung Q2/2025",
    "priority": "high"
  },
  "workstate": {
    "selectedLevel": 3,
    "useFindingOverride": false,
    "useLevelOverride": { "1": false, "2": false, "3": true, "4": false },
    "findingOverride": "",
    "levelOverrides": {
      "1": "",
      "2": "",
      "3": "Die Charta soll aktiv kommuniziert und in Meetings besprochen werden.",
      "4": ""
    },
    "includeFinding": true,
    "includeRecommendation": true,
    "overwriteMode": "append",
    "done": true,
    "notes": "Kunde sehr engagiert, Follow-up in 3 Monaten",
    "lastEditedBy": "consultant@example.com",
    "lastEditedAt": "2025-02-11T14:30:00Z"
  },
  "assets": {
    "photos": ["leitbild_001.jpg", "leitbild_002.jpg"],
    "slides": [],
    "documents": []
  },
  "exportHints": {
    "renumberedId": null,
    "includeInReport": true
  }
}
```

## Version History

### Version 1 (Current)
- Initial schema
- Support for hierarchical chapters
- Multi-locale titles
- Photo tagging (Bericht, seminar, topic)
- Override system with enable flags
- Append-only history log

### Future Versions
- Version 2 (planned): Add `assets.documents`, enhanced export hints
- Version 3 (planned): Multimedia support (video, audio)

## Related Documentation

- [Data Model Overview](../architecture/data-model.md) - Conceptual explanation
- [VBA Modules Reference](../../macros/README.md) - Implementation details
- [System Overview](../architecture/system-overview.md) - Integration

---

**See also**:
- [Getting Started Guide](../guides/getting-started.md)
- [VBA Workflow](../guides/vba-workflow.md)
- [HTML Workflow](../guides/html-workflow.md)
