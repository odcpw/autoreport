---
name: autobericht
description: Consolidate an SST consultant's past reports into an editable Word Masterbericht, personal AutoBericht library and style guide; turn visit recordings and the project Sidecar into a complete review-ready Sidecar in their voice; retain reviewed improvements for later projects.
---

# AutoBericht — from professional judgment to a report

The product goal is a structured report the consultant can use in their own voice after substantive review, without rewriting unwanted prose. Treat author-style fidelity, evidence fidelity and Sidecar compatibility as separate requirements. Help the consultant spend their time observing, evaluating and reviewing. Take responsibility for selecting, assembling and writing the appropriate report content in their voice. Guide a new colleague through the process; do not expect them to know the file structure or select paragraphs themselves.

Keep execution pragmatic. Start from the files available, make routine editorial decisions and deliver the requested files. Use the supporting procedures only where they help this job; do not turn them into a questionnaire, approval sequence or separate exercise the consultant must supervise. Reuse existing extraction, mappings and style evidence. Keep only the working notes needed to avoid omissions or resume long work, preferably together; no routine certificates, scores or extra reports about the process. Show the result and briefly identify real unresolved decisions. Stop when the output is usable and the relevant checks pass, rather than polishing the workflow itself.

The normal report workflow has four inputs: **MP3 recording(s), the project Sidecar, this skill and the consultant's style guide**. The application loads the library into the Sidecar when the project is created. An “empty” Sidecar is not yet filled for this visit; it still contains the question structure, library passages and any imported customer/photo metadata. Use that embedded library as the starting content and the supplied style guide as the writing authority. Do not request a duplicate library, a separate examples collection or an author-pack ZIP for an ordinary report.

This is an author-neutral workflow. Each consultant keeps their own style guide privately; “author profile” in these instructions means that same guide, not another required file. It may include representative approved passages, but examples are not an additional upload requirement. No previous consultant's library, client data or personal writing profile is bundled in this skill. Respond in the user's language and write the report in the project locale.

The shared skill, helpers and workflow are maintained in the AutoBericht repository. Each consultant's personal library, profile and reports remain in their chosen private storage. **Site photos stay on the work computer.** The consultant views them locally and dictates filenames/categories and descriptions. Work receives permitted audio/text and sidecar metadata, never photo uploads, thumbnails, contact sheets or screen captures. Generate text and photo associations from that testimony; the local app reconnects the returned sidecar to the image files and produces the illustrated report. Do not request cloud image access as a prerequisite or treat a local desktop agent as offline image analysis.

## Authoring contract

Before writing report content, read [report-authoring.md](references/report-authoring.md). Establish the author’s voice and structure from their current instructions, confirmed examples and supplied library; assemble useful existing wording with minimal adaptation. Keep report text and its clean preview free of assistant narration and working annotations. Put source mappings, review questions and system feedback in separate files. Check every changed client-facing passage before delivery; a valid JSON file alone is not success.

## Choose the starting point from the files already available

| Route | Starting material | Requested result |
|---|---|---|
| Build the author's reusable material | Anonymised past reports, a chosen Word base and actual app structure when creating JSON | Editable Masterbericht, personal library and personal style guide; produce the requested subset or all three for full setup |
| Prepare a new visit report | Recording(s), project Sidecar with embedded library, personal guide | Full processed draft Sidecar and matching clean preview, with transcript and working questions separately |
| Learn from a reviewed visit | Reviewed Sidecar/corrections and current personal library/guide | Enriched reusable library and approved style updates; refresh the Word master when requested |

- **New visit, Sidecar and style guide:** read [dictation-and-writing.md](references/dictation-and-writing.md), the supplied guide, then [sidecar-contract.md](references/sidecar-contract.md). Use the library already embedded in the Sidecar; begin without a separate onboarding exercise.
- **Masterbericht / Megabericht / Word consolidation:** read [master-report.md](references/master-report.md) and [onboarding.md](references/onboarding.md). Consolidate within a copy of the designated report, preserving its layout and finding/recommendation pairing. For full author setup, also build the library and calibrated guide from the same source review.
- **Style-guide or library creation from old reports:** read [onboarding.md](references/onboarding.md). This is preparation for an author who needs it, not a repeated prerequisite for each report.
- **Reviewed report or sidecar returning from a project:** read [library-cycle.md](references/library-cycle.md). Extract useful reusable changes and prepare the next library version.
- **Questions about recording, uploads or repository access:** read [start-and-capabilities.md](references/start-and-capabilities.md). Check the actual tools available; instructions do not themselves grant audio, file, repository or network access.

Read [examples.md](references/examples.md) for cross-topic interviews, paragraph selection, uncertainty and library updates. It contains synthetic examples, not source material for a real company.

At the start, inspect the recordings, Sidecar and style guide and state the next useful step. Reuse an available installed skill or attached workflow; the ZIP and portable skill document are alternatives, not two required inputs. Ask only for missing information that affects the work: typically an absent Sidecar, unresolved report language or a decision that changes a technical recommendation. If the Sidecar genuinely lacks the expected library/structure, identify the specific missing data before requesting another export. Do not mistake untouched generic workstate text for a case finding or an approved writing example. If only anonymised reports are provided for onboarding, start reading them instead of demanding a complete project first.

## Use current application context when useful

The maintained application and editable skill are at https://github.com/odcpw/autoreport (`skill/source/`). When current tag-to-row behaviour, sidecar compatibility, import/export or app testing matters, propose using that repository for context; see [repository-context.md](references/repository-context.md). Reuse an accessible checkout or GitHub connection, or offer to clone it when shell/network access permits. Continue with the supplied files when repository access is unavailable or unnecessary. Repository code explains application behaviour; the consultant’s own library and profile determine report wording.

## Working agreement

1. **Preserve the author’s judgment.** Distinguish observations, self-assessment answers, statements in interviews, interpretation, proposed measures and decisions. Keep uncertainty and explicit corrections. Do not turn a rough impression into several proven failings.
2. **Use the library actively.** Select paragraphs or individual sentences across applicable variants; combine them into one coherent recommendation. The consultant need not dictate each measure. Keep useful explanations and practical detail; remove duplication, irrelevant remedies and old company particulars. Do not deliver a menu of paragraphs unless asked.
3. **Adapt findings lightly.** Start from the matching generic negative finding. Adjust attribution, scope and partial implementation, for example “Selon nos discussions, …”. Do not retain unsupported parts of a compound finding. Keep generic findings unchanged in the reusable library unless expressly requested.
4. **Allow one theme to inform several sections.** Map each supported aspect to the actual project questions. Distinguish a main treatment, a complementary angle and an unresolved hypothesis. Avoid repeated paragraphs, invented row IDs and blanket score changes. The hierarchy/employee-voice example is in `examples.md`.
5. **Use natural assessment language.** Interpret “barely”, “more or less”, “mostly”, “almost” in context using the existing four levels. Preserve imported customer answers/comments. A manual level sets `scoreTouched=true`; do not score field observations or summaries or confuse implementation with priority.
6. **Prepare inclusion; leave review visible.** Clear issues described for the report become drafted, included rows with the relevant recommendations and photos. New or changed draft rows remain `done=false` until the consultant’s review or explicit validation. “Done” means the report item is ready, not that the company has implemented the measure.
7. **Carry technical links with selected text.** Keep only applicable references under “Voir aussi :” or the author’s locale equivalent; deduplicate them after assembly. Do not attach citations to old reports in the final prose. Internal working mappings stay outside the report.
8. **Keep project and library separate.** A sidecar contains the current case. The library carries reusable writing to the next case. At a requested project closeout, prepare an enriched library copy from reviewed text without requiring the consultant to select every sentence. Do not learn from unreviewed generated drafts or import company facts into future defaults.
9. **Deliver a real file when tools permit.** Edit a copy of the full input JSON, preserving unrelated data. Validate it, return it as a downloadable artifact, and describe checks actually run. Never replace a full sidecar with a partial JSON snippet. If file tools are absent, provide the completed draft/patch and state that the executable file still needs to be produced.

## End-to-end execution

Inventory and authoring contract → full transcript and photo identities → corrections and topic/evidence mapping → light finding adaptation and paragraph assembly → existing-level assessment → sidecar copy → clean preview → evidence, style and application checks with internal repair → one delivery for consultant review/Done → approved library/profile update at closeout.

Work through a long recording in resumable batches. Keep an external coverage register of processed time ranges and unresolved items. Reconcile topics and corrections across all batches before claiming completeness. Try the supplied recording first when transcription is supported; chunk only when limits or reliability require it. Never silently substitute a summary for the full transcript.

When asked to create a style guide, follow [style-calibration.md](references/style-calibration.md) and populate the common numbered [author-profile-template.md](assets/author-profile-template.md). Every colleague receives the same depth of analysis and drafting checks, with personal rules and examples drawn from their actual reports and confirmed edits. The guide structure is shared; sentence forms and report prose remain the author's own. Test the guide on available passages before claiming calibration and label unsupported functions provisional. When library reconstruction is requested, use an actual exported library/assessment structure. Remove identifiable untouched AI bootstrap recommendations while retaining authored additions and useful variants; keep ambiguous origin in an internal review list rather than guessing. These onboarding operations are separate from drafting a visit with an already loaded Sidecar and supplied guide.

## Supporting tools

The Python helpers use the standard library; audio splitting additionally needs `ffmpeg` and `ffprobe`. The optional app checker uses Node.js. They make no network requests and do not transcribe audio or generate prose.

- `scripts/sidecar_tool.py`: inspect, apply a narrow external edit plan to a copy, or validate preservation against the original. Supports the current nested sidecar; legacy inputs need adaptation first. Read the contract before use.
- `scripts/report_preview.py`: generate a clean row-based preview directly from the Sidecar, with possible process-language leakage reported separately; does not certify style or evidence.
- `scripts/library_tool.py`: apply reviewed, generalised recommendation additions to a library copy while preserving findings and other data. Detects exact normalised duplicate blocks; semantic deduplication remains the assistant’s responsibility.
- `scripts/verify_app.cjs`: optional Node.js check against a supplied AutoBericht checkout for actual library import and sidecar normalisation.
- `scripts/audio_chunks.py`: split an oversized recording, preserving source offsets and overlap in a manifest. Use only if necessary.
- `scripts/extract_docx.py`: extract paragraphs/tables and link targets from Word reports for onboarding. This is text extraction, not anonymisation, OCR or a formatting review.

When repository access is available, inspect the relevant current modules and run the repository’s normalisation/import checks. The portable scripts and documented contract allow a compatible existing sidecar to be prepared without fetching the whole repository. Never claim app import was tested when only JSON-level checks ran.

## Expected delivery

For full author setup: an editable consolidated Word Masterbericht, personal library built against actual application structure and calibrated personal style guide. For a narrower onboarding request, deliver the requested subset. Keep coverage/gaps and unapproved example candidates in separate working material. For a visit: full updated Sidecar copy and clean preview with identical report wording; transcript and working review/coverage/validation material separately; system-improvement Markdown when needed. Do not stop at transcription or a proposed mapping when the requested outcome is the processed Sidecar. Ready for review means the supported scope has been drafted, included and checked, with Done left for the consultant; it does not mean silently approving the report. For requested closeout: next library copy and approved style-guide updates, with a concise change list; update the Word master when requested. The user-facing interaction should remain light even when the internal work is thorough.
