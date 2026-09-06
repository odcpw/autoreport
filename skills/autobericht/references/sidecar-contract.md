# File contract and safe execution

Verified against the local AutoBericht code at revision `6d6e32b5f9533103e006512514f42ca0776142e7` on 6 September 2026. This is a compatibility snapshot, not a promise about every future version. The supplied project is the authority for its actual IDs and keys.

## Data ownership

| Object | Meaning and rule |
|---|---|
| `sidecar.report.project` | Report project, with `meta` and `chapters[].rows[]` |
| `row.id` with chapter ID | Existing report destination; preserve identity and order |
| `row.master.finding/recommendation` | Starting library text; keep intact when drafting workstate |
| `row.customer` / `customer.items[]` | Imported answers, comments and evidence; preserve all source fields |
| `row.workstate.findingText` | Case finding, lightly adapted from the generic text |
| `row.workstate.recommendationText` | Selected and assembled case recommendation |
| `selectedLevel` | Integer 1..4; current display 0/33/66/100 |
| `scoreTouched` | True for adopted manual/dictated assessment; prevents self-assessment defaults from overwriting it |
| `includeFinding`, `includeRecommendation` | Inclusion controls; independent of review status |
| `done` | Report row reviewed/ready; final export requires includeFinding AND done |
| `priority` | Existing 0..4 control; do not infer from selectedLevel |
| `libraryAction`, `findingLibraryAction` | Off/Append/Replace queues for native library update; not proof that the text is generalised |
| `sidecar.photos.photos[path]` | Photo record keyed by the actual project-relative path |
| `tags.report`, `tags.observations`, `tags.training` | Arrays of exact option values, not invented translated labels |
| `sidecar.photos.photoTagOptions` | Available options with values/labels; preserve them |
| `sidecar.photos.photoRoot`, `meta`, `spider`, other branches | Preserve unless explicitly part of the task |

The double `photos.photos` nesting is intentional. The sidecar contains tags/notes, not the image pixels. Standard rows may group several self-assessment items; inspect `id`, `originalId`, `groupId` and `collapsedId` where present. A statement about one sub-item must not be assigned to every sibling automatically.

Field observations are normally in chapter 4.8 and have type `field_observation`. Match them by `tag` against current options/library values. Do not assume a numeric `4.8.x` ID identifies the same category in every project. Summary rows and observations receive no implementation score.

## Work on a copy

Read the complete JSON programmatically and preserve unknown data. Record its SHA-256. Apply only intended changes. If the source changed while drafting, re-read and reconcile before writing; do not overwrite another editor’s updates. Keep original and output under different names. The supplied helpers refuse to overwrite an existing output.

The draft editor switches the relevant library action to `off` when changing a finding/recommendation so the new project text cannot be queued accidentally for library export. This does not discard the separate closeout library workflow. It sets `scoreTouched` and aligns `autoScoreLevel` for an explicit level edit. Material changes reset Done unless the edit plan explicitly records the consultant’s validation.

## External edit plan

This is a helper format only. Do not insert it into `project_sidecar.json` or represent it as the app’s native patch API.

```json
{
  "format": "autobericht-editor-patch/1",
  "sourceSha256": "SHA256_OF_THE_ACTUAL_INPUT_FILE",
  "rows": [
    {
      "chapterId": "EXISTING_CHAPTER_ID",
      "rowId": "EXISTING_ROW_ID",
      "changes": {
        "findingText": "Case-specific finding",
        "recommendationText": "Assembled recommendation with applicable links",
        "selectedLevel": 2,
        "includeFinding": true,
        "includeRecommendation": true,
        "done": false
      }
    }
  ],
  "photos": [
    {"path": "photos/resized/ACTUAL_FILENAME.jpg", "observationsAdd": ["EXISTING_CATEGORY_VALUE"]}
  ]
}
```

Replace the illustrative IDs/hash/path with actual values. Omit fields not being changed and omit `selectedLevel` for observations/summaries. Photo changes also support `reportAdd`, `reportRemove`, `observationsRemove`, `trainingAdd`, `trainingRemove` and replacement `notes`; preserve the original notes unless replacement is intended. The helper cannot prove that a named image exists: verify against the supplied sidecar or permitted text-only filename manifest. Verify actual image resolution later in the local app; do not request image uploads.

Run from the skill folder, or resolve these paths relative to the installed skill:

```sh
python3 scripts/sidecar_tool.py inspect project_sidecar.json > project_index.json
python3 scripts/sidecar_tool.py apply project_sidecar.json edits.json project_sidecar_draft.json
python3 scripts/sidecar_tool.py validate project_sidecar.json project_sidecar_draft.json
```

The helper rejects stale hashes, unknown rows, changes outside the allowed fields, unknown new tag values, invalid levels and accidental changes to customer/master data. It supports the nested current format. It does not add chapters, add observation categories, migrate flat legacy sidecars or calculate the spider chart. For those tasks, use the current app code or an app-exported migrated file, preserving the source.

## Library additions

```json
{
  "format": "library-additions/1",
  "sourceSha256": "SHA256_OF_CURRENT_LIBRARY",
  "additions": [
    {
      "kind": "observation",
      "key": "EXISTING_CATEGORY_VALUE",
      "text": "Reviewed, generalised recommendation with its relevant references",
      "reviewed": true,
      "generalised": true
    }
  ]
}
```

Use `kind: "entry"` with the real library entry ID for question recommendations. The true flags record the assistant’s completed review; the script cannot establish them automatically.

```sh
python3 scripts/library_tool.py library_current.json additions.json library_next.json
```

This appends to existing targets and preserves the generic findings. It checks exact duplicate blocks; the assistant must also resolve semantic overlap and contextual scope.

## Repository verification, when available

For the originating repository `odcpw/autoreport`, use the user-provided accessible checkout, connector or source archive. Do not assume permission to the repository from this name alone. Read:

- `docs/autobericht/workflow.md` and `README.md` for app operations;
- `AutoBericht/mini/shared/import-self.js` for customer mapping;
- `state.js`, `normalize.js`, `seeds.js` in that shared directory for scores, row types and library imports;
- `report-rows.js` for inclusion/export gates;
- `io-sidecar.js` and `sidecar-storage.js` for file merging and library updates;
- `AutoBericht/mini/photosorter/io-sidecar.js`, `photos.js`, `tags.js` for photo storage;
- the relevant tests and `AutoBericht/tests/helpers.js` for actual import/normalisation checks.

Do not treat `experiments/ai-voice-report` or `experiments/ai-evidence-lab` as a finished main-app dictation pipeline; the inspected documentation identifies them as isolated experiments.

Report verification in layers: valid JSON and preserved unrelated data; actual application normalisation/import if run; interactive load/export if performed. One level does not prove the next. The returned sidecar should be complete and downloadable even if final app import remains for the consultant to verify. Do not call an unvalidated file a finished compatible sidecar.

The optional Node.js helper exercises those modules without changing the repo:

```sh
node scripts/verify_app.cjs /path/to/autoreport library_next.json project_sidecar_draft.json
```

Omit the sidecar argument when checking only a library. A failure after app evolution requires investigation; do not weaken the checks to make an incompatible file pass.

The checker also reports recommendation coverage. The current importer looks up standard library entries using each question group's collapsed ID. Text stored only under individual sub-item IDs may pass schema validation yet never reach a new project's rows. A coverage failure lists the affected IDs and exits unsuccessfully. Resolve their mapping in a library copy; do not silently discard those passages or treat schema validity as complete import coverage. Section headers may share IDs with actual rows and must be excluded from row lookup, while remaining unchanged in the file.
