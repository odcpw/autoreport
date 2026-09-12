# Consolidate past reports into a usable Masterbericht

Read this when the user requests a Masterbericht, Megabericht, consolidated report or reusable Word report containing material from several past reports. It is a working report the author can shorten, strike through or edit in place for a new visit. Put the reusable content in the report itself, in its normal chapters and finding/recommendation tables. A detached paragraph catalogue, appended dump or numbered collection of “Fallvarianten” does not meet that request unless explicitly requested.

For full author setup, combine this route with [onboarding.md](onboarding.md) for the application library and [style-calibration.md](style-calibration.md) for the personal guide. Use one source inventory and content review for all requested outputs. These outputs serve different purposes: the Word master supports editing in place, the JSON library supplies selectable passages to new Sidecars, and the guide controls the author's voice. None replaces the others.

## Establish the base and read the originals

Start from the report the user designates, such as “consolidate everything into report 4”. Work on a copy. If no base is designated, inspect the available originals and choose the most complete suitable layout; state the choice. Ask only if materially incompatible templates or languages leave the intended result unclear.

Read every supplied readable report, preserving chapter, table, row and cell roles. Follow the input privacy and coverage rules in [onboarding.md](onboarding.md); site photos stay on the work computer, including photos embedded in old reports. Retain existing anonymised text and usable formatting. Inspect the native Word structure and, where available, its rendered layout. Text extraction alone cannot establish which recommendation belongs to which finding or whether table formatting survived.

Keep a private working inventory of source units: chapter/topic, finding, paired recommendations, scope/conditions, applicable references and destination. Distinguish useful additions, repeated wording, complementary detail, context-specific alternatives and contradictions. “Everything” means accounting for all useful authored material, not duplicating repeated text or retaining old company particulars. Record exclusions and reasons outside the report.

Use an available document editing tool or library that preserves native DOCX structures; use `ooxml-cli` when available and suitable, checking its actual help before using commands. The bundled `extract_docx.py` is an extraction aid, not a lossless Word editor. Do not claim the skill itself installs an editor or that writing Markdown produces the requested Word file. If editing is unavailable, complete the content mapping and identify the missing execution capability; do not present that intermediate work as the finished Masterbericht.

## Consolidate by meaning and keep the pairs intact

For each source unit, find its actual topic in the chosen base. Keep existing suitable wording, then integrate useful additions beside it. Correct grammar and smooth joins in the author's style with minimal changes to meaning. Use the original authored reports and confirmed corrections as style evidence; the generated consolidation must not become its own proof of the author's voice.

- Merge repeated statements and genuinely complementary measures into a coherent entry without discarding useful reasoning, practical examples, conditions or links.
- Keep distinct situations separately editable in the relevant chapter. A manual pallet truck and an electric truck may require different measures even if their topic is related.
- Preserve an alternative as a concise, clearly scoped normal report entry or paragraph. Do not combine mutually exclusive company situations into one asserted finding or invent a fictional company that has every issue. Use a short descriptive topic/condition where needed, following the template's normal style rather than introducing an elaborate variant numbering system.
- Keep each finding aligned with the recommendations addressing that finding. If the source has paired columns, add or clone paired rows. If the template uses a different paired structure, preserve that structure. Do not create two independent long lists whose entries stop corresponding as they grow.
- A new recommendation that fits an existing finding belongs in that recommendation cell. A substantially different finding needs its own paired entry in the appropriate chapter. Review a source cell containing several topics rather than assuming all its paragraphs share one destination.
- Generalise company-specific facts for reuse while retaining the technical circumstances that make a passage valid. Do not turn a local arrangement or a one-off frequency into a universal requirement. Keep contradictions or unresolved technical scope in the external review notes until they can be resolved.

For introductions, positive observations and management summaries, preserve the base layout and useful reusable language. Do not stitch the histories and achievements of several companies into one factual account. Keep conditional alternatives clearly distinguishable and remove old identities. Include only the functions requested or already within the consolidation's scope.

The reusable library's generic negative findings remain unchanged unless separately requested. Editable findings in the Word master may retain the authored situations and useful alternatives; do not overwrite the application question structure merely to make it resemble the Word layout.

## Preserve the actual Word layout

Use the base document's section settings, headings, table widths, column grid, cell margins, borders, merged-cell relationships, paragraph and character styles, list indentation, numbering and header/footer layout. Clone suitable nearby structural elements for additional entries rather than pasting all content into one cell. When copying content from another DOCX, preserve or correctly remap the styles, numbering and hyperlink relationships it uses; relationship identifiers are local to each package.

Keep finding and recommendation content on the intended sides of each pair. Retain relevant emphasis and working links. Remove old company data from headers, footers and fields in the reusable copy as well as the body. Do not carry restricted source images into an uploaded output. Preserve source files unchanged.

Longer content can create unattractive page breaks. When the user accepts manual pagination cleanup, treat those breaks as cosmetic; do not shorten useful content, shrink all text or change the table structure to hide them. Pairing, readable cell contents and a valid editable document still matter. Distinguish a split row continuing across pages from a finding aligned with the wrong recommendation.

## Check the deliverable, not just the extraction

Before delivery:

1. Reconcile the source inventory: every readable useful unit is retained, merged without meaning loss, kept as a scoped alternative or explicitly set aside. A word count or paragraph count alone cannot prove coverage.
2. Re-extract the finished DOCX and inspect every changed finding/recommendation pair against its intended destination. Check for orphaned measures, duplicated blocks, lost conditions, wrong chapter placement and accidental loss of base content.
3. Check DOCX package integrity and internal relationships with available tools. Render or reopen the document when possible and inspect changed tables, headings, long entries and cross-page continuations. If rendering is unavailable, state that the visual layout remains unverified; do not equate XML validity with a visual check.
4. Read the final prose in the author's voice. Keep source citations, coverage notes, drafting instructions and calibration commentary outside the client-facing report. Retain applicable technical references in the author's convention.

Deliver the editable consolidated DOCX using a stable active name such as `Masterbericht_<author>_<locale>.docx`, or the name the user requested. Preserve the prior version in an archive. Include the personal library and guide when the full setup is requested, with a short external coverage/limitations note. Do not substitute a PDF, review HTML or text catalogue for the requested Word master.

## Connect the master to the next report

Build the library from the same reviewed source units, matched to actual application IDs/categories. Do not assume Word row numbers are library IDs. Preserve the intended coverage across Word and JSON even though their structures differ; do not force every Word passage into an unrelated generic question. Any unmapped useful material stays visible in the working coverage notes for resolution.

The consultant loads the adopted library when creating a new project. The new Sidecar embeds it. A normal recording session then needs the recordings, that Sidecar, the personal guide and this skill; neither the Word master nor a duplicate standalone library is a routine additional upload. Follow [dictation-and-writing.md](dictation-and-writing.md) through to the full updated Sidecar and clean preview.

After the consultant's review, follow [library-cycle.md](library-cycle.md) to preserve useful new wording. Update the Word master too when requested; a library update does not silently rewrite an existing Word master or a live project.
