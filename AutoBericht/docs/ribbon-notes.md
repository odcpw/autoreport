# Ribbon embedding (customUI) – quick checklist

Use this to avoid the “found unreadable content, repair?” prompt when adding our ribbon to DOCM/PPTM files.

## Minimal structure
- Add a folder at zip root: `customUI/`.
- Place the ribbon markup in `customUI/customUI.xml`.
- Root element must use namespace `http://schemas.microsoft.com/office/2006/01/customui` and exactly one `<customUI>` part is allowed per package. All controls must sit inside a `<group>` which sits inside a `<tab>` under `<ribbon>`; you cannot place controls or separators directly under `<tab>`. citeturn0search0
- Relationship in `_rels/.rels` (root rels) pointing to the part:
  ```xml
  <Relationship
    Id="rIdCustomUI"
    Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    Target="customUI/customUI.xml"/>
  ```
  (Target is relative, no leading slash.)
- `[Content_Types].xml` needs an Override:
  ```xml
  <Override PartName="/customUI/customUI.xml"
            ContentType="application/vnd.ms-office.customui+xml"/>
  ```
  Ensure only one Override for this part to avoid duplicate‑part corruption.
- If you need Backstage (2007/10) features use `customUI2` with namespace `http://schemas.microsoft.com/office/2007/10/customui`; only one of customUI/customUI2 is honored, with customUI2 taking precedence. citeturn0search1

## Common corruption causes
- Multiple `customUI/customUI.xml` entries or duplicate Overrides in `[Content_Types].xml`.
- Relationship Target written as `/customUI/customUI.xml` plus another entry without slash (duplicates).
- Using the `customUI14` namespace without matching relationship/content type.
- Invalid XML (e.g., separator outside a `<group>`, missing required attributes, duplicate IDs).
- Custom images referenced without corresponding image part + relationship.

## Vertical separators
- Allowed only inside a `<group>`: `<separator id="sep1"/>`.
- Not allowed directly under `<tab>` or `<ribbon>`; doing so breaks XML validation and triggers repair.
- All controls on a tab must be inside a `<group>`; separator is a valid child of `<group>` along with buttons, menus, etc. citeturn0search0

## Recommended embedding workflow
1) Open the DOCM/PPTM with the “Office Custom UI Editor” or an OPC‑aware zip tool.
2) Add/replace `customUI/customUI.xml` with our `AutoBericht/vba/ribbon.xml`.
3) Verify `_rels/.rels` contains exactly one `ui/extensibility` relationship pointing to `customUI/customUI.xml`.
4) Verify `[Content_Types].xml` has a single Override for `/customUI/customUI.xml` with `application/vnd.ms-office.customui+xml`.
5) Validate XML (the Custom UI Editor can validate against the 2006/01 schema).
6) Save and reopen in Word to confirm no repair prompt; then re‑sign macros if required.

## Current project status
- Ribbon XML lives in `AutoBericht/vba/ribbon.xml`.
- Templates now reside in `ProjectTemplate/` root (`Vorlage IST-Aufnahme-Bericht d.V01.docm`, `Vorlage AutoBericht.pptx`).
- If a repair prompt reappears, re-embed the ribbon using the steps above to eliminate duplicate overrides/relationships left by manual zip edits.
 - Jan 2026: fixed DOCM `_rels/.rels` target from `/customUI/customUI.xml` to `customUI/customUI.xml` (leading slash caused repair prompt).

## 2026-01 refresh (what works now)
- Replaced PowerPoint-only `imageMso` ids with Word-safe icons: `FileOpen`, `TextToTableDialog`, `FormatPainter`, `MailMergeInsertFields`, `PictureInsertFromFile`, `ChartRadar`, `FileSaveAsPdfOrXps`.
- Added in-group separators in Markdown and Export groups; separators are only valid inside `<group>`.
- Deduped `customUI/customUI.xml` inside `ProjectTemplate/Vorlage IST-Aufnahme-Bericht d.V01.docm` and ensured a single override + relationship (2006/01 schema).
- Keep using the 2006/01 namespace/part. If you add `customUI14.xml`, switch namespace to `http://schemas.microsoft.com/office/2009/07/customui` and add the correct root relationship.
- Post-fix verification (2026-01-12): the DOCM must physically contain `customUI/customUI.xml` in the zip. A missing part (even with correct override/relationship) hides the tab. Quick check:
  - `python - <<'PY'\nimport zipfile\nz=zipfile.ZipFile('ProjectTemplate/Vorlage IST-Aufnahme-Bericht d.V01.docm')\nprint([n for n in z.namelist() if 'customui' in n.lower()])\nPY`
  - Expect: `['customUI/customUI.xml']`. If empty, re-add the part from `AutoBericht/vba/ribbon.xml` and ensure only one override + one relationship remain.

## Field guide (from “don’t get burned” note)
- Match namespace to part: `customUI` → 2006/01; `customUI14` → 2009/07.
- Use macro-enabled docs (`.docm`/`.dotm`) in a trusted location; enable “Show add-in user interface errors” in Word options when debugging.
- Validate XML: unique ids, escaped `&`, controls live inside a `<group>`.
- Restart Word fully when testing; UI state can cache.

## 2026-01 refresh (what works now)
- Replaced PowerPoint-only `imageMso` ids with Word-safe icons: `FileOpen`, `TextToTableDialog`, `FormatPainter`, `MailMergeInsertFields`, `PictureInsertFromFile`, `ChartRadar`, `FileSaveAsPdfOrXps`.
- Added in-group separators in Markdown and Export groups; separators are only valid inside `<group>`.
- Deduped `customUI/customUI.xml` inside `ProjectTemplate/Vorlage IST-Aufnahme-Bericht d.V01.docm` and ensured a single override + relationship (2006/01 schema).
- Keep using the 2006/01 namespace/part. If you add `customUI14.xml`, switch namespace to `http://schemas.microsoft.com/office/2009/07/customui` and add the correct root relationship.

## Field guide (from “don’t get burned” note)
- Match namespace to part: `customUI` → 2006/01; `customUI14` → 2009/07.
- Use macro-enabled docs (`.docm`/`.dotm`) in a trusted location; enable “Show add-in user interface errors” in Word options when debugging.
- Validate XML: unique ids, escaped `&`, controls live inside a `<group>`.
- Restart Word fully when testing; UI state can cache.
