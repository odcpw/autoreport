---
name: autobericht
description: Guide an SST consultant from anonymised past reports to a personal AutoBericht library and writing profile, then turn visit dictation, interviews, self-assessment and photos into a review-ready project sidecar and carry reviewed recommendations into the next library.
---

# AutoBericht — from professional judgment to a report

Help the consultant spend their time observing, evaluating and reviewing. Take responsibility for selecting, assembling and writing the appropriate report content in their voice. Guide a new colleague through the process; do not expect them to know the file structure or select paragraphs themselves.

This is an author-neutral workflow. The consultant supplies their own library and style profile, or anonymised past reports from which to build them. No previous consultant’s library, client data or personal writing profile is bundled. Respond in the user’s language and write the report in the project locale.

The shared skill, helpers and workflow are maintained in the AutoBericht repository. Each consultant's personal library, profile and reports remain in their chosen private storage. **Site photos stay on the work computer.** The consultant views them locally and dictates filenames/categories and descriptions. Work receives permitted audio/text and sidecar metadata, never photo uploads, thumbnails, contact sheets or screen captures. Generate text and photo associations from that testimony; the local app reconnects the returned sidecar to the image files and produces the illustrated report. Do not request cloud image access as a prerequisite or treat a local desktop agent as offline image analysis.

## Choose the starting point from the files already available

- **First-time author, with old reports:** read [onboarding.md](references/onboarding.md). Build their personal library and profile before or alongside a first new report.
- **New visit, existing library:** read [dictation-and-writing.md](references/dictation-and-writing.md), the author’s profile, then [sidecar-contract.md](references/sidecar-contract.md) before editing a sidecar.
- **Reviewed report or sidecar returning from a project:** read [library-cycle.md](references/library-cycle.md). Extract useful reusable changes and prepare the next library version.
- **Questions about recording, uploads or repository access:** read [start-and-capabilities.md](references/start-and-capabilities.md). Check the actual tools available; instructions do not themselves grant audio, file, repository or network access.

Read [examples.md](references/examples.md) for cross-topic interviews, paragraph selection, uncertainty and library updates. It contains synthetic examples, not source material for a real company.

At the start, inspect supplied files and state the next useful step. Ask only for missing information that affects the work: typically report language, the absent project/library, or a decision that changes a technical recommendation. Do not repeat questions answered by the files. If only anonymised reports are provided, start reading them instead of demanding a complete project first.

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

Inventory → full transcript and photo identities → topic/evidence mapping → light finding adaptation → paragraph selection and assembly → existing-level assessment → sidecar copy → validation → consultant review/Done → reusable library update → next project.

Work through a long recording in resumable batches. Keep an external coverage register of processed time ranges and unresolved items. Reconcile topics and corrections across all batches before claiming completeness. Try the supplied recording first when transcription is supported; chunk only when limits or reliability require it. Never silently substitute a summary for the full transcript.

For a new author, build their profile from their actual reports and confirmed edits. Use `assets/author-profile-template.md` as a starting structure, not a fixed style. Reconstruct the library against an actual exported library/assessment structure. Remove identifiable untouched AI bootstrap recommendations while retaining authored additions and useful variants; keep ambiguous origin in an internal review list rather than guessing.

## Supporting tools

The Python helpers use the standard library; audio splitting additionally needs `ffmpeg` and `ffprobe`. The optional app checker uses Node.js. They make no network requests and do not transcribe audio or generate prose.

- `scripts/sidecar_tool.py`: inspect, apply a narrow external edit plan to a copy, or validate preservation against the original. Supports the current nested sidecar; legacy inputs need adaptation first. Read the contract before use.
- `scripts/library_tool.py`: apply reviewed, generalised recommendation additions to a library copy while preserving findings and other data. Detects exact normalised duplicate blocks; semantic deduplication remains the assistant’s responsibility.
- `scripts/verify_app.cjs`: optional Node.js check against a supplied AutoBericht checkout for actual library import and sidecar normalisation.
- `scripts/audio_chunks.py`: split an oversized recording, preserving source offsets and overlap in a manifest. Use only if necessary.
- `scripts/extract_docx.py`: extract paragraphs/tables and link targets from Word reports for onboarding. This is text extraction, not anonymisation, OCR or a formatting review.

When repository access is available, inspect the relevant current modules and run the repository’s normalisation/import checks. The portable scripts and documented contract allow a compatible existing sidecar to be prepared without fetching the whole repository. Never claim app import was tested when only JSON-level checks ran.

## Expected delivery

For onboarding: personal library copy, author profile, short style assessment, coverage/gaps list and reviewed-example candidates. For a visit: full updated sidecar copy, readable draft, short unresolved list, coverage record and validation result. For closeout: next library copy plus a concise change list. The user-facing interaction should remain light even when the internal work is thorough.
