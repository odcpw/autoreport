# Ribbon embedding (customUI) – quick checklist

Use this to avoid the “found unreadable content, repair?” prompt when adding our ribbon to DOCM/PPTM files.

## Minimal structure
- Add a folder at zip root: `customUI/`.
- Place the ribbon markup in `customUI/customUI.xml`.
- Root element must use namespace `http://schemas.microsoft.com/office/2006/01/customui` and exactly one `<customUI>` part is allowed per package.
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

## Common corruption causes
- Multiple `customUI/customUI.xml` entries or duplicate Overrides in `[Content_Types].xml`.
- Relationship Target written as `/customUI/customUI.xml` plus another entry without slash (duplicates).
- Using the `customUI14` namespace without matching relationship/content type.
- Invalid XML (e.g., separator outside a `<group>`, missing required attributes, duplicate IDs).
- Custom images referenced without corresponding image part + relationship.

## Vertical separators
- Allowed only inside a `<group>`: `<separator id="sep1"/>`.
- Not allowed directly under `<tab>` or `<ribbon>`; doing so breaks XML validation and triggers repair.

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
- If a repair prompt reappears, re‑embed the ribbon using the steps above to eliminate duplicate overrides/relationships left by manual zip edits.

