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

Follow [style-calibration.md](style-calibration.md) and populate the common numbered [author-profile-template.md](../assets/author-profile-template.md). Give every colleague the same structure and depth of analysis while deriving their rules and examples from their own writing. Use the evidence order in [report-authoring.md](report-authoring.md).

Explain meaningful variation, turn observations into actionable rules, and run the internal drafting/comparison trial before calling the guide calibrated. Preserve the difference between authored examples, confirmed corrections and generated proposals. Keep a harmonised experiment separate from the source writing. If evidence is insufficient for a function, mark that part provisional and continue with supported material; do not require the author to fill gaps in a questionnaire.

## Reuse the style guide with a new project

The shared repository distributes the workflow. Each consultant keeps their personal style guide privately. “Author profile” and “style guide” refer to the same file. It records the writing conventions and may include a few approved examples extracted from the supplied reports; the consultant should not have to assemble an additional example collection.

For an ordinary report, the inputs are the MP3 recording(s), project Sidecar, shared skill and style guide. The library has already been loaded into the Sidecar by the application. Do not create or require an author pack, duplicate library upload or repeated onboarding. Record the guide version and identifiable embedded library provenance in the working notes without making these extra user tasks.

After review, retain explicit approved style corrections in a new version of the guide. A reusable library update is a separate requested operation; follow [library-cycle.md](library-cycle.md). Improve common behaviour in the shared skill separately, without copying personal examples or client facts into it.

## Onboarding delivery and coverage

Deliver the personal style guide and, when library creation was requested, the new `library_user_<author>_<locale>.json`. Embed useful confirmed examples in the guide rather than requiring another upload. Keep the short homogeneity assessment and coverage summary separately: reports read, sections unreadable, paragraphs retained/merged, empty categories, uncertain mappings and bootstrap limitations. Preserve source versions. Validate a newly created library against the app's importer when available; the consultant loads it in the application for new projects. Do not bundle their corpus into the reusable skill itself.

“Complete” means every supplied readable report has been accounted for and its usable recommendations matched or explicitly set aside. It does not mean every generic question has a recommendation or that no future report can add useful material.
