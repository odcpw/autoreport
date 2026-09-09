# Build a colleague’s personal library from past work

## Begin with the reports

A colleague may start by uploading all available anonymised reports in manageable batches. Inventory the whole collection: filenames, type, language, author, whether final/reviewed, duplicates and readable status. Read representative full reports before editing; then process the remainder with a coverage register. Do not claim the whole corpus was covered after sampling it for style.

The intended input excludes company-specific information and site photos that must stay on the work computer. Prepare text-only or image-stripped report copies locally before upload; blurring a company name does not make embedded site images eligible to leave the computer. Ask for anonymised copies when originals have not yet been uploaded. If identifiable originals are already present, keep originals unchanged and make a sanitised working copy before building shareable assets. Remove company/person names, contact details, exact site addresses, signatures, logos, revealing photo labels and document metadata where applicable. Do not describe text extraction or name replacement alone as complete anonymisation. Preserve technical context needed for the recommendation. This is a user-requested input constraint, not permission to erase their originals.

Start useful reading even if there is no library. To produce an importable library, later obtain the actual base library or export from AutoBericht in the right locale, including its generic findings and question structure. If necessary, guide them to create a project, select its language and export its library. Do not invent question IDs or reconstruct the entire self-assessment from memory.

## Read Word, PDF and JSON correctly

Word reports often put findings and recommendations in table cells. Preserve table, row and cell boundaries and reading order. `extract_docx.py` emits these plus hyperlink targets; it cannot read scanned images and does not render the page. Inspect the source visually if cell roles or relationships are ambiguous. For PDF, use text extraction and page inspection; scanned reports need OCR. Label unreadable sections and continue with readable ones.

A report is evidence of the author’s previous wording, not evidence about the next company. Existing library JSON may mix bootstrap text, authored replacements and appended text. A sidecar may contain drafts, final edits and explicit library actions. Review its workstate rather than assuming every stored paragraph is approved.

Translations need two sources of judgment: the original-language authored text for meaning and the author’s original writing in the target language for phrasing. A translated JSON need not sound like their own original reports. Establish which version is the original before comparing styles. Translate unique recommendations faithfully if requested; otherwise keep language variants distinct and ask only where the choice matters.

## Separate authored material and bootstrap

Use an identified bootstrap baseline where available. Compare recommendations exactly and by normalised whitespace before classifying them. Remove an untouched bootstrap block from the reconstructed recommendation library. Retain an authored replacement or addition, including edited/appended material that the author confirms they accepted. Do not delete text merely because it sounds generic, and do not attribute AI seed text to the author solely because it appears in a JSON.

When a retained addition depends on a removed seed sentence, preserve the authored meaning through minimal grammatical repair. Do not silently add back the full AI paragraph. Keep uncertain origin and incomplete dependencies in a short internal review list. If the baseline is unavailable, state that bootstrap removal cannot be proven exhaustively and use the reports as the stronger authoring source.

Generic negative findings and structure remain unchanged by default. The reconstruction concerns recommendations. Do not fill empty recommendation categories with invented content to achieve apparent completeness.

## Reconstruct the recommendation library

Maintain an internal candidate inventory with source location, extracted text, technical scope, candidate question/category and inclusion status. These locations need not appear in report prose.

For each candidate:

1. Identify the actual problem and action, preserving useful reasoning and practical examples.
2. Match the real question/observation category by meaning and scope. Use actual IDs and mappings from the supplied structure; several candidates may belong to one question, and a reusable passage may legitimately serve more than one category.
3. Correct language lightly. Remove old company particulars and unsupported carryover of numbers or frequencies. Keep technical conditions, alternatives and modality.
4. Compare with retained variants. Merge exact duplicates, retain meaningful alternatives and do not erase useful detail merely to shorten the library.
5. Keep the applicable technical links. Verify new/replaced checklist links against the primary publisher and correct language edition. A general guide must not be called a dedicated checklist. Do not make legal or technical upgrades under the guise of style correction.
6. Place variants in the application-compatible recommendation string, separated clearly (normally blank line, `---`, blank line). Keep a small external selection index if the corpus needs it; do not add an unrecognised schema inside the library.

Do not assume every paragraph in a given table cell addresses the same category. Conversely, do not split a condition from the action it limits.

## Learn the writing profile

Use `assets/author-profile-template.md` and the evidence order in [report-authoring.md](report-authoring.md). Record actionable choices with supporting examples, not only adjectives such as “professional” or “concise”. Assess style separately for findings, recommendations, management summary and positive observations. Look at opening verbs, typical sentence length, how the author involves supervisors/employees, modality, useful questions, explanations, terminology and reference formatting. Frequency across several authored reports is stronger evidence than one occurrence.

Describe whether differences are functional (summary versus detailed measure), linguistic (translation), chronological or inconsistent editing. Keep typos out of the profile. Distinguish confirmed preferences from working hypotheses. A harmonised experimental copy does not automatically become ground truth for the author’s voice.

Provide a small before/after sample and ask for only the stylistic choices that materially change the profile. Continue building the library while optional preferences are pending. Retain confirmed examples separately from generated proposals. If no style examples are supplied for a new report, follow the library’s wording and use restrained language while stating the profile is provisional.

## Package the author's reusable inputs

The shared repository distributes how AutoBericht works, not each consultant's personal style. Keep a private author pack for use alongside the shared skill:

- `author-profile.md`: identity, locale, version/date, the evidenced writing conventions and approved examples. Examples may be embedded here or kept in a referenced file.
- The author's current library JSON, using its real application schema and locale.

Use a folder or one ZIP when file tools permit; separate existing files are equally valid. Keep client Sidecars, recordings, complete reports and unapproved generated drafts outside this reusable pack. A pack is personal input data, not another skill that needs installation. Do not commit it to the shared repository. Record which profile/library version was used in the project's working notes.

For first use, inspect the consultant's reports/library and build these resources as part of the work. For subsequent visits, reuse the latest available pack with the visit's Sidecar and recording/transcript. Do not make the consultant rebuild their profile or upload both the shared ZIP and its equivalent portable Markdown. If a profile or library is already available in the conversation or accessible private storage, use it. A Sidecar's embedded library can supply report passages when no separate library was provided; do not overwrite its project edits to synchronise with a newer pack.

At a requested closeout, update the private pack from approved corrections using [library-cycle.md](library-cycle.md). Keep the prior version and carry the new version into the next project. Improve common behaviour in the shared skill separately, without copying personal examples or client facts into it.

## Onboarding delivery and coverage

Deliver a private author pack containing the new `library_user_<author>_<locale>.json` and `author-profile.md` with confirmed examples. Keep the short homogeneity assessment and coverage summary separately: reports read, sections unreadable, paragraphs retained/merged, empty categories, uncertain mappings and bootstrap limitations. Keep the source version and record the new file’s hash. Validate against the app’s library importer when available. Do not bundle the colleague's corpus into the reusable skill itself.

“Complete” means every supplied readable report has been accounted for and its usable recommendations matched or explicitly set aside. It does not mean every generic question has a recommendation or that no future report can add useful material.
