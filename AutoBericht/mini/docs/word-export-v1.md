# Word Export v1

This document describes the current no-VBA DOCX export contract used by `AutoBericht/mini/shared/word-export.js`.

## Scope

- Export target: `.docx` template selected by the user at export time.
- Output file: `Outputs/YYYY-MM-DD_AutoBericht_NoVBA.docx`.
- Data source: current project sidecar state in memory.

## Export Rule

Rows are exported only when both flags are true:

- `workstate.includeFinding === true`
- `workstate.done === true`

In other words: report-ready rows are `includeFinding && done`.

## Placeholder Contract

### Text markers

- `NAME$$`
- `COMPANY$$`
- `COMPANY_ID$$`
- `ADDRESS$$`
- `AUTHOR$$`
- `MOD$$`
- `CO$$`
- `DATE$$`

### Chapter payload markers

- `CHAPTER0$$`, `CHAPTER1$$`, `CHAPTER2$$`, ... (by chapter id)
- Chapter `0`: exported as lettered list paragraphs.
- Other chapters: exported as table payloads.

### Image markers

- `LOGO_BIG$$` (document body, centered large logo)
- `LOGO_SMALL$$` (headers, small logo)
- `SPIDER$$` (rendered spider chart image)

### Thermometer markers

- `THERMO1$$` ... `THERMO14$$` for numeric chapters as used in template.
- No thermo for chapter `0`.
- No thermo for `4.8` subsection chapter.

## Module Map

- `shared/word-export.js`
  - Export orchestration and template mutation flow.
- `shared/report-rows.js`
  - Shared row projection, renumbering, and text/priority resolution.
- `shared/word-docx-zip.js`
  - Browser ZIP unpack/pack helpers for DOCX parts.
- `shared/word-docx-xml.js`
  - OOXML-safe marker replacement and relationship utilities.
- `shared/spider.js`
  - Spider score computation (`computeSpider`).
- `shared/spider-chart.js`
  - Shared spider drawing used by UI and export.

## High-Level Flow

1. User picks template `.docx`.
2. Template ZIP parts are unpacked.
3. Text markers are replaced in `word/document.xml` and headers.
4. Chapter markers are replaced with generated XML payloads.
5. Logos are injected if available from project Outputs paths.
6. Spider scores are computed and injected as image + thermos.
7. Updated parts are repacked and written to `Outputs/`.

## Notes for Template Editing

- Marker tokens should stay inside normal paragraph text where possible.
- The XML helper is paragraph-aware and tolerates split runs, but simple marker placement is safer.
- If Word rewrites runs while editing, markers can still be detected as long as full token text remains in the paragraph.

## TOC and Field Updates

- Export enforces `w:updateFields=true` in `word/settings.xml` so Word can refresh fields/TOC on open.
- This is safe to keep in both places:
  - in template defaults, and
  - enforced by exporter (guaranteed even if template variant misses it).
- TOC content is still style-driven:
  - use built-in Heading styles (`Heading 1`, `Heading 2`, `Heading 3`, ...)
  - configure TOC levels in template to include only the heading levels you want.
