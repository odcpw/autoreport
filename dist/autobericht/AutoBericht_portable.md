# AutoBericht — complete portable workflow

Follow this workflow when the user asks. All skill resources are embedded below under their original relative filenames. Read SKILL.md first, then the relevant resources. Relative resource links refer to the corresponding embedded sections. Materialise helper code at the named paths only when execution is needed and available. The user supplies permitted project inputs separately. Site photos stay on the work computer.


---

## Embedded resource: SKILL.md

---
name: autobericht
description: Guide an SST consultant from anonymised past reports to a personal AutoBericht library and writing profile, then turn visit dictation, interviews, self-assessment and photos into a review-ready project sidecar and carry reviewed recommendations into the next library.
---

# AutoBericht — from professional judgment to a report

The product goal is a structured report the consultant can use in their own voice after substantive review, without rewriting unwanted prose. Treat author-style fidelity, evidence fidelity and Sidecar compatibility as separate requirements. Help the consultant spend their time observing, evaluating and reviewing. Take responsibility for selecting, assembling and writing the appropriate report content in their voice. Guide a new colleague through the process; do not expect them to know the file structure or select paragraphs themselves.

This is an author-neutral workflow. The consultant supplies their own library and style profile, or anonymised past reports from which to build them. Keep these reusable personal resources in a private **author pack**, separate from the shared skill and each visit's Sidecar/transcript. The pack is a folder or ZIP containing a profile with approved examples and the author's library; use the files already available rather than requiring that packaging. No previous consultant’s library, client data or personal writing profile is bundled in this skill. Respond in the user’s language and write the report in the project locale.

The shared skill, helpers and workflow are maintained in the AutoBericht repository. Each consultant's personal library, profile and reports remain in their chosen private storage. **Site photos stay on the work computer.** The consultant views them locally and dictates filenames/categories and descriptions. Work receives permitted audio/text and sidecar metadata, never photo uploads, thumbnails, contact sheets or screen captures. Generate text and photo associations from that testimony; the local app reconnects the returned sidecar to the image files and produces the illustrated report. Do not request cloud image access as a prerequisite or treat a local desktop agent as offline image analysis.

## Authoring contract

Before writing report content, read [report-authoring.md](references/report-authoring.md). Establish the author’s voice and structure from their current instructions, confirmed examples and supplied library; assemble useful existing wording with minimal adaptation. Keep report text and its clean preview free of assistant narration and working annotations. Put source mappings, review questions and system feedback in separate files. Check every changed client-facing passage before delivery; a valid JSON file alone is not success.

## Choose the starting point from the files already available

- **First-time author, with old reports:** read [onboarding.md](references/onboarding.md). Build their personal library and profile before or alongside a first new report.
- **New visit, existing library:** read [dictation-and-writing.md](references/dictation-and-writing.md), the author’s profile, then [sidecar-contract.md](references/sidecar-contract.md) before editing a sidecar.
- **Reviewed report or sidecar returning from a project:** read [library-cycle.md](references/library-cycle.md). Extract useful reusable changes and prepare the next library version.
- **Questions about recording, uploads or repository access:** read [start-and-capabilities.md](references/start-and-capabilities.md). Check the actual tools available; instructions do not themselves grant audio, file, repository or network access.

Read [examples.md](references/examples.md) for cross-topic interviews, paragraph selection, uncertainty and library updates. It contains synthetic examples, not source material for a real company.

At the start, inspect supplied files, including any author pack, and state the next useful step. Reuse an available installed skill or attached workflow; the ZIP and portable skill document are alternatives, not two required inputs. Ask only for missing information that affects the work: typically report language, the absent project/library, or a decision that changes a technical recommendation. Do not repeat questions answered by the files or ask for a separate library when suitable embedded text is already available. If only anonymised reports are provided, start reading them instead of demanding a complete project first.

## Use current application context when useful

The maintained application and editable skill are at https://github.com/odcpw/autoreport (`skills/autobericht/`). When current tag-to-row behaviour, sidecar compatibility, import/export or app testing matters, propose using that repository for context; see [repository-context.md](references/repository-context.md). Reuse an accessible checkout or GitHub connection, or offer to clone it when shell/network access permits. Continue with the supplied files when repository access is unavailable or unnecessary. Repository code explains application behaviour; the consultant’s own library and profile determine report wording.

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

For a new author, build their profile from their actual reports and confirmed edits. Use `assets/author-profile-template.md` as a starting structure, not a fixed style. Reconstruct the library against an actual exported library/assessment structure. Remove identifiable untouched AI bootstrap recommendations while retaining authored additions and useful variants; keep ambiguous origin in an internal review list rather than guessing.

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

For onboarding: a private author pack containing the personal library and author profile with confirmed examples; keep the short style assessment, coverage/gaps and unapproved example candidates in separate working material. For a visit: full updated sidecar copy and clean preview with identical report wording; transcript and working review/coverage/validation material separately; system-improvement Markdown when needed. Deliver the requested scope without invented findings or an unsolicited full-company synthesis. For closeout: next library/profile version in the author pack plus a concise change list. The user-facing interaction should remain light even when the internal work is thorough.


---

## Embedded resource: agents/openai.yaml

````yaml
interface:
  display_name: "AutoBericht"
  short_description: "From visit dictation to report and reusable library"
  default_prompt: "Use $autobericht to guide me from my reports or visit dictation to a reviewed sidecar and an enriched personal library."
````


---

## Embedded resource: assets/author-profile-template.md

# Author profile

Populate from the consultant’s own reports and confirmed feedback. This template sets no personal style by default.

- Author and report locale:
- Profile version/date and associated library filename/version:
- Original writing languages; translations and their original-language source:
- Reports/sections actually examined:
- Confirmed preferences versus provisional observations:

## Output and voice contract

- Report perspective and recommendation form, with an actual example:
- Required field/paragraph structure and where it applies:
- Desired detail: explanations, concrete examples and practical steps to retain:
- Technical reference convention:
- Evidence for each preference (approved example, explicit instruction, or provisional inference):

Use this as a working aid, not a questionnaire the author must complete before a first report. Keep client facts out of reusable examples.

## Findings

Typical attribution, negative formulation, partial implementation, scope and degree of certainty. Add representative authored examples.

## Recommendations

Typical opening verbs, paragraph length, actors, explanations, questions, concrete examples, modality, alternatives and technical vocabulary. Add representative authored examples and explain when each style is used.

## Management summary and positives

Expected concision, synthesis, tone and treatment of strengths. Do not assume the detailed recommendation style applies unchanged.

## References

Preferred language, publisher editions, link format and placement. Distinguish technical references from internal report-source tracking.

## Avoid

Document the author’s confirmed dislikes. Correct grammatical mistakes without adopting them as a signature. Preserve technical meaning and useful reasoning.

## Confirmed corrections

Before/after examples the author has actually reviewed, with the precise lesson and scope. Keep generated proposals distinct.

For each example, retain the original observation, the author-approved finding/recommendation, and the specific correction. Mark generated alternatives as unapproved. Do not promote them because the author continued to another task. Do not invent a style-match percentage.


---

## Embedded resource: references/dictation-and-writing.md

# From rambling to a review-ready report

Apply [report-authoring.md](report-authoring.md) before composing. It governs author-style evidence, minimal adaptation, clean output boundaries and the final semantic/style pass.

## Capture the judgment before editing the prose

Use the imported self-assessment as the starting point: yes/no, comments, evidence and question groupings. Keep those customer statements intact. The consultant’s visit assessment belongs in the workstate and may contradict the customer’s answer.

Transcribe before drafting in the report language. For Swiss German, a faithful Standard German transcript is acceptable; preserve dialect meaning and uncertainty rather than demanding dialect spelling. Maintain timestamps and stable photo identifiers where available. The author may revisit topics, correct themselves, read a checklist aloud, describe an interview, offer an idea or dictate a decision. Preserve those distinctions. Do not convert a quoted question, a hypothetical “if they do not…” or a future plan into an observed failure.

Correct transcription uncertainties that can be resolved from context, but do not guess critical negations, equipment names, numbers, responsible persons or dates. Mark recognition uncertainties in the transcript/working review file; retain substantive interview uncertainty in natural report wording. Handle the clear parts without waiting for all uncertainties to be resolved.

## Build a topic map that can span chapters

Maintain an external working map, not fields injected into the sidecar. Useful fields are: topic ID, source time/statement, attribution, scope/site, certainty, photos, candidate row IDs, each row’s distinct angle, score evidence, proposed actions, unresolved points and review state.

One statement may support several sections. “Decisions are very top-down and workers have no voice” can concern leadership behaviour, the organisation of consultation and employee participation. It does not by itself establish absent training, undefined job descriptions, psychological illness, a statutory breach or a failure on every leadership question.

Read the real questions. For every destination, articulate the different supported point. Choose a primary location for the fullest treatment. In complementary locations, use a short relevant finding or a distinct action. A management summary may synthesise the pattern once. Do not duplicate the full text or mechanically lower all associated scores.

Separate a confirmed theme from potential links. If the consultant reports a broad impression without examples, preserve “Selon nos entretiens…” or “Les échanges donnent l’impression…” as appropriate, and leave unsupported mappings in the questions list. Ask for a concrete example only if it changes inclusion or the remedy. See `examples.md`.

## Findings

Use the closest applicable generic negative finding and change only what the case requires: attribution, scope, partial implementation, frequency and qualification. Examples of attribution include “Selon l’autoévaluation…”, “Selon nos discussions…”, “Selon les entretiens avec les collaborateurs…” and “Lors de la visite…”. Use the author’s language conventions.

Check every factual claim in the selected library wording against the current testimony. Remove unsupported clauses, including the generic sentence's main diagnosis when necessary: overdue measures establish delay, but do not alone establish missing follow-up. Occasional audits are not no audits. Lack of visible evidence in one photo is not proof an item does not exist elsewhere. Positive observations are legitimate; do not force a negative finding into a fully satisfied question. If no generic wording fits, write a short case-specific finding and record any wording/mapping exception in the separate working review file, without altering the generic library finding.

## Recommendation selection is the assistant’s job

The consultant describes the situation and judgment. They do not have to know the library or dictate every measure. Read applicable variants and select paragraphs or sentences that address the actual issue. Use neighbouring categories for reusable measures where appropriate, while preserving the actual finding’s category and technical scope.

Select useful immediate correction, longer-term organisation, employee involvement, instruction or checks only when justified. These are possible elements, not a mandatory five-part template. The user’s stated preferences and exclusions prevail.

Combine complementary passages into one coherent recommendation. Prefer existing good wording, retain practical explanations and questions, harmonise terminology and grammar, and remove overlap. A `---` separator is a variant boundary in the library, not a requirement to copy the whole block into the report. Preserve conditions and alternatives when extracting a sentence.

The assistant may propose a suitable library measure even when the consultant only described the problem. Write it as a recommendation in the author’s voice, leaving Done false for review; do not claim it was agreed with the company. Questions are for facts that change technical applicability or decisions, not for sentence order or equivalent phrasings.

Do not import company names, site details, old dates, headcounts, frequencies or promises. Do not invent obligations or technical specifications. When a missing detail prevents a sound measure, record that decision in the working questions file and continue drafting the supported remainder. New technical claims or links need primary-source verification; stylistic rephrasing of supplied text does not require fresh web research.

## Natural-language evaluation

Use the application’s four existing levels, whose current display values are 0/33/66/100. The consultant need not say percentages.

| Meaning in context | Working level |
|---|---|
| Effectively absent, nothing implemented | 1 |
| Barely started, a few elements exist but much is missing | 2 |
| More or less in place | 2 or 3 according to the described extent; unresolved if not enough evidence |
| Mostly/almost implemented, some gaps remain | 3 |
| Fully implemented and working, with no relevant reservation | 4 |

“Barely started” and “barely anything left to finish” differ. Future implementation is not present implementation. An explicit correction replaces the earlier judgment. A yes/no default is not a consultant assessment. Preserve the nuance in the finding and set `scoreTouched=true` for an adopted dictated level. Do not invent a score for an inferred secondary topic; do not use implementation level as urgency or priority. No score applies to field observations or summaries.

## Photos and inclusion

The consultant views the photos on the work computer; the assistant does not receive or inspect their pixels. Resolve spoken identifiers against the sidecar or a permitted text-only manifest, without needing access to the image files. Preserve paths so the local app can reconnect them after import. Never request images, thumbnails, contact sheets or screen sharing as a workaround. A phrase such as “photo 19” must be resolved from an explicit numbering map; do not guess browser order. Resolve spoken identifiers to the actual file path. Match category labels/aliases against the current library and photo options. A category may have several photos; a photo may support several categories. Keep equipment/site distinctions and existing tags/notes. A category named alone permits classification, not invention of a defect. Text-only interview findings require no fabricated photo.

Prepare clear described report issues with the adapted finding, assembled recommendation, relevant photos and inclusion flags. Leave `done=false` on new or materially changed drafts, unless the consultant explicitly validates that item. A previously approved row changed by new information needs re-review. Keep explicit “off the record / not in report” material out of the delivered report.

Retain only links relevant to the selected sentences. Deduplicate at the end of the recommendation under “Voir aussi :” / “Siehe auch:” / the author’s convention. A guide’s own local procedures are not universal requirements. Do not attach old-report source citations to the prose.

## Process long input without losing its end

Keep coverage by recording range and by topic, including excluded and unresolved items. A later “correction” applies to earlier material even when it arrives in another audio chunk. Deduplicate overlap by timestamps and meaning; repeated emphasis is not necessarily a new issue. Reading existing report/library text aloud must be distinguished from adopting it for the current company.

Before delivery, sweep the entire transcript for topics never mapped, corrections not applied, conditions dropped and contradictions between yes/no, commentary, level and finding. When the requested scope includes a management summary, prepare it after the topic sweep from the current case. Do not replace an existing summary or write a full-company diagnosis merely because this is a first photo batch. Do not treat a large transcript as permission to output only its most salient themes.

## Review and outputs

Provide a complete Sidecar copy and a clean draft grouped by chapter, using the same field text. Keep decisions still needed and validation/coverage results in separate working material, and software feedback in a separate system-improvement file. The draft preview should include prepared rows even though the current final exporter requires both `includeFinding=true` and `done=true`. Keep the preview separate; never mark everything Done merely to obtain a complete Word export.

The target is that the consultant reviews the substance and mainly checks Done. The delivery can include unresolved items in its separate working file; a final approved report cannot pretend they were resolved. After review, use the library cycle to keep worthwhile new wording for later projects.


---

## Embedded resource: references/examples.md

# Behavioural examples

All examples are synthetic. They illustrate decisions, not facts about a real company or a universal author style. Resolve actual row IDs from each supplied project.

## One interview theme, several report areas

Dictation: “From our discussions, management is very hierarchical. The workers say decisions arrive from above and they have nothing to say. During the safety meetings they are told what to do, but their suggestions are not discussed. I want the supervisors to involve them when defining practical measures.”

Break this into the supported aspects rather than assigning one large paragraph everywhere:

| Destination | Distinct angle | Possible treatment |
|---|---|---|
| Employee participation | Suggestions are not discussed; workers have limited input into measures | Main finding based on the interview and a recommendation to involve employees in defining measures and give feedback on suggestions |
| Leadership | Supervisors’ safety meetings mainly transmit instructions | Complementary point about how supervisors conduct the discussion, with a practical facilitation/involvement recommendation |
| Organisation | How suggestions reach decisions and get a response | Include if the account supports an organisational gap; otherwise leave the existence of a defined process to confirm |
| Management summary | A common pattern across the supported findings | One short synthesis about limited employee involvement and the intended development |

A French participation finding could adapt the relevant generic wording to “Selon les entretiens avec les collaborateurs, leurs propositions sont peu prises en compte dans la définition des mesures de sécurité.” A leadership finding could concern the limited place given to discussion during safety meetings. These are proposals whose wording must be checked against the author’s library.

Do not infer that roles are undocumented, no risk assessments exist, all supervisors lack training or a law has been violated. Do not lower every leadership/organisation/participation score from this single theme. Determine each level from the actual positive question and available evidence. If the only dictation were “the management seems hierarchical”, retain it as an attributed impression until the specific impact is clear.

If the project has no suitable leadership row, use a supported chapter synthesis/summary field if available, or flag the mapping. Do not invent a question or overwrite an unrelated one just to achieve three destinations.

## Select parts of a recommendation

Library variants include: replace missing rack pins and damaged uprights; arrange checks during supervisors’ visits; ask workers to report damage; replace condemned racks.

Dictation: “There are missing pins. Otherwise the racks look fine. Have the supervisors check them during visits and ask people to report anything missing.”

Select the missing-pin phrase, the relevant check and the reporting sentence. A possible result is:

“Ajouter les goupilles de sécurité manquantes sur les rayonnages. Contrôler leur présence lors des visites de sécurité réalisées par les supérieurs. Encourager les collaborateurs à signaler immédiatement les dommages ou les éléments manquants.”

Keep the applicable rack checklist link once. Omit damaged uprights and wholesale replacement. No question to the consultant is needed to choose these paragraphs. If they had only described missing pins, a suitable library recommendation to replace them could still be prepared without asking them to dictate it.

## Reading a question is not answering it

Audio: “Do employees receive periodic instruction? … That’s the next question. I still need to ask them. Photo 19 shows the notice board.”

Do not score the instruction question or write an absence finding. The photo may be classified if its identity/category is clear; its presence does not answer whether instruction occurs. Record the question as not yet assessed.

## A correction at the end of a long recording

At 08:12: “They have no inspection register.” At 74:05: “Correction on the register: the maintenance manager showed it to me later. It exists; it is the follow-up of overdue actions that is missing.”

Apply the later correction to every affected draft. Remove the false absence finding and associated register-creation recommendation. Select only the supported follow-up material. Revisit any level based on the initial statement. Do not count the two statements as independent evidence for two defects.

## Natural assessment and customer answer

Customer answer: yes. Consultant: “Only the new hires get it. Nothing systematic afterward; it’s barely in place.”

Preserve the customer’s yes/comment. Use a partial consultant level if supported by the actual question; do not default to completely absent. Adapt the finding to the missing ongoing element and choose recommendations for that gap. Set scoreTouched. If the same sentence is mapped to another question, do not copy the score automatically.

## A photo with two issues

“IMG_0042, loading area: the exit is blocked by pallets and the extinguisher behind them is inaccessible.”

The same file may have the existing escape-route and fire-protection categories. Prepare the distinct relevant findings. Avoid repeating general housekeeping advice in both recommendations when a concise cross-topic treatment suffices. Keep links applicable to each issue. Do not add unstable stacking or forklift training defects unless described/observed.

## Useful accretion without copying the company

A reviewed case recommendation adds a useful method: team members photograph a recurring handling difficulty, discuss options with the supervisor and trial a suitable arrangement before standardising it. The case also names a site and a dated implementation meeting.

Retain the reusable participatory method under the matching question/category, remove the site/date, and preserve its context. Compare with the current library before appending. If already expressed adequately, avoid a new duplicate. The profile may learn to retain such practical explanations only if the author’s feedback supports that preference.

A new automated draft marked Done by a script is not a reviewed example. Do not use it to build the author’s style or restore an AI bootstrap library.

## No library yet

Colleague: “Here are my anonymised reports. Build the system for me.”

Inventory/read the reports, extract their writing patterns and reusable passages, and begin a profile. Guide them to supply/export the target AutoBericht structure when exact mapping is needed. Produce a library against that structure. Do not block initial reading because they do not know what a sidecar is. Do not invent missing IDs or clone another consultant’s personal profile.


---

## Embedded resource: references/library-cycle.md

# The library grows from project to project

## Purpose

The library is the consultant’s accumulated reusable writing and practical advice. It should absorb worthwhile new material from reviewed project sidecars. This is part of the workflow, not an exceptional manual rewriting exercise.

Keep three versions conceptually distinct: the library used to start the project, the current project sidecar, and the next reusable library. A current library may already contain updates from another project; merge against that current version rather than overwriting it with an older project’s starting library.

## At a requested project closeout

When the consultant asks to finish the report cycle or enrich the library, inspect the reviewed sidecar and prepare a new library copy. They need not choose every reusable sentence. Use explicit approval, a reviewed final report, and the item’s Done state in context to identify material suitable for learning. A stale or imported Done flag alone is not proof of authorship or style approval; if the provenance is unclear, keep candidates provisional.

Compare project `workstate.recommendationText` with the starting library/master and with the current library. Classify changes as:

- already covered or exact duplicate;
- better wording of the same idea;
- useful additional explanation or implementation detail;
- distinct alternative for a particular context;
- company-specific facts only;
- unreviewed draft or unclear technical change.

Generalise the useful text: remove names, old site references, event dates and company-specific assumptions. Retain the technical circumstances that make it valid. A requirement for one equipment type must not become a default for all machinery. A measure adopted in a single company can become a scoped alternative, not a universal prescription.

Prefer adding a useful variant. Replace existing text only when the intent is to supersede it and no useful meaning is lost. Deduplicate both exact wording and overlapping ideas. Retain complementary detail. Keep generic negative findings unchanged unless the user expressly requests a findings-library revision.

## Native AutoBericht update route

The current app stores `workstate.libraryAction` and `findingLibraryAction` as `off`, `append` or `replace`. These queue actions; the library changes when **Generate / Update Library** is run. Done and these actions are different controls.

Native Append/Replace uses the row’s current text. It does not automatically anonymise or generalise it. Standard-row hash markers reduce repeated processing, but the inspected observation update path can append the same text again on repeated runs. Do not rely on those flags as semantic deduplication.

If using the native route, ensure the exact queued text is appropriate for reuse. Do not rewrite a reviewed project’s case-specific recommendation solely to make its library export generic. For generalised variants, prefer a separate updated library copy built from the reviewed sidecar and current library. Keep a small merge record outside both files.

## Portable update route

The assistant prepares a `library-additions/1` plan containing only reviewed, generalised recommendation text, exact target IDs/category values and the source library hash. `library_tool.py` appends these to a copy while preserving findings, structure, metadata and unrelated entries. It skips exact normalised duplicate blocks and does not make semantic decisions or redact company data. The assistant performs that review before calling it.

This helper intentionally handles additions to existing targets only. Replacement, a new category, or a schema change needs a deliberate edit plus full import validation; do not force such changes through a guessed ID.

Return the next library and a concise change list. The user can adopt it for the next project without a separate confirmation for every harmless addition when they already requested the update. Preserve the previous library. Do not upload or publish the author’s library to other people without their direction.

## Start the next project

Use the latest adopted library when creating the new AutoBericht project. Bring its author profile and confirmed examples. A file beside an already-created sidecar does not automatically refresh the master text embedded in that sidecar. Do not overwrite existing project edits to refresh its library.

Keep the updated library, profile and approved examples together in the consultant's private author pack (folder or ZIP), preserving the prior version. The shared skill stays in the repository; personal style updates travel with the pack. Reuse available files on the next visit instead of asking the consultant to reconstruct their style or reinstall the common workflow.

Teach only confirmed style preferences. An accepted case-specific measure is not automatically a preference to include that measure in all reports. A reviewed phrase becomes an example in the appropriate function—finding, recommendation or summary. Do not train the profile on the assistant’s own unreviewed drafts.


---

## Embedded resource: references/onboarding.md

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


---

## Embedded resource: references/report-authoring.md

# Write the author's report, ready for substantive review

Read this before drafting any report text, including a first trial. AutoBericht saves time when the consultant reviews the judgment and can use the wording. A technically valid Sidecar that requires a prose rewrite has not met that goal. A rough recording is input to professional writing, not permission to deliver rough report text.

## Establish the authoring contract from available evidence

Use the current user's instructions and corrections first. Then use their confirmed profile and approved examples, their own finished reports in the target language, and the supplied library. Treat unreviewed workstate, generated rewrites and translated bootstrap text as weaker evidence. An assistant's suggestion is not an approved example merely because the conversation continued. Do not assign a percentage of style fidelity or claim a perfect match without a defined evaluation.

Before composing, establish these choices in a short **internal** working note:

- Report locale and the author's voice: impersonal, first-person plural, direct instructions or another evidenced form.
- Findings: usual attribution, sentence structure, degree of detail and treatment of partial implementation.
- Recommendations: opening verbs, modality, actors, paragraph structure, explanations and practical examples worth retaining.
- Output structure: actual chapter/row identities, finding versus recommendation fields, summary/positive conventions, reference placement and requested scope.
- Explicit dislikes or corrections, with the source that supports each preference.

Do not ask the consultant to fill in that note or choose a paragraph menu. Use the embedded library if a standalone profile is missing; begin with a provisional profile drawn from the supplied authored material. Read several relevant examples rather than treating the first seed paragraph as the author's style. Ask for one missing example only if competing conventions would materially change the result and the available material cannot resolve them. Do not turn optional calibration into a prerequisite for processing the recording.

Preserve each author's choices. French infinitives, “Nous recommandons”, a four-paragraph structure or a particular attribution are not universal defaults. Preserve requested structures when provided; do not force every short finding or recommendation into an invented template.

## Compose from evidence and existing writing

For each report item, settle the supported situation and the intended action before writing. Choose the real destination question/category. Keep the evidence mapping and source locations outside the report.

1. **Keep useful authored wording.** Start with the relevant finding and recommendation passages. Check each claim against the current evidence before keeping it: a library diagnosis is a wording candidate, not evidence that this company has that problem. Adjust scope, attribution, partial implementation, grammatical joins and case-specific details. Remove unsupported claims even if they are central to the generic sentence. Do not paraphrase good text merely to make it look newly generated.
2. **Assemble the recommendation.** Select the relevant sentences across variants and, where appropriate, neighbouring categories. Keep conditions, useful reasoning, examples, implementation detail and the author's preferred sequence. Remove duplication and irrelevant remedies; do not compress everything into generic “clarify, communicate, monitor” advice.
3. **Fill only a real gap.** If the library has no suitable wording, write the missing passage using the same actors, verbs, register and level of detail. An empty library entry does not excuse omitting a clear, supportable recommendation. It also does not authorize inventing a technical requirement.
4. **Keep professional meaning.** Preserve who said what, the extent of the issue and what remains uncertain. Correct speech disfluencies and grammar without copying verbal filler or turning an impression into a proven failure. Keep observations and measures in their respective fields.
5. **Read the result as the author.** Would this person put the paragraph in a client report without deleting framing, translating abstract wording into concrete language, or restoring useful detail? Revise it now if not. Do not ask the user to diagnose avoidable prose defects.

Use the consultant's reporting perspective when authorized by the supplied observations. “Selon nos discussions…” attributes interview evidence without pretending the model attended the visit. Do not wrap that evidence in “le consultant décrit”, “selon la dictée”, “les échanges dictés montrent” or “the transcript indicates”. Similarly, write the recommendation itself in the author's voice, not “the consultant recommends that…”. An explicit request for third-person reporting takes precedence.

Plain wording is a default when examples do not settle a choice: name the object, actor, action and concrete difference. Do not replace “les personnes interrogées donnent des indications différentes” with an abstract account of “la traduction des résultats en modalités opérationnelles”. Do not add an explanatory final sentence simply to give every paragraph the same shape.

### Example: adapt the supplied voice, not a universal voice

Synthetic input: an inspection register exists, but several overdue actions remain open. The author provides these relevant library sentences:

- Finding: “Les mesures définies ne sont pas systématiquement suivies jusqu'à leur réalisation.”
- Recommendation: “Définir avec les responsables les délais de réalisation. Examiner les mesures ouvertes lors des séances de secteur.”

The supported finding is: “Un registre des contrôles est disponible. Plusieurs mesures restent ouvertes après les délais convenus.” Overdue actions establish delay; by themselves they do not establish a lack of follow-up. Use the generic finding about follow-up only if the current testimony supports that additional claim. Keep or lightly adapt the supplied recommendation where it fits. Do not invent absence of the register, rewrite the delay as “le consultant relève un déficit de traçabilité”, or add a new reporting system.

If another author's approved recommendation instead says “Nous recommandons de convenir des échéances avec les responsables et de reprendre les actions ouvertes lors des réunions de secteur”, retain that voice. A helper must not mechanically turn every recommendation into infinitives.

## Keep output boundaries explicit

| Destination | Belongs here | Does not belong here |
|---|---|---|
| Report fields and clean preview | Finished findings, recommendations, scoped synthesis/positives and relevant technical references | Transcript timestamps, confidence labels, source IDs appended as prose, assistant narration, reviewer instructions, explanation of drafting choices |
| Working review file | Time-to-row/photo mapping, uncertain transcription, unresolved decisions, wording provenance and checks actually run | A competing version of the client report |
| System-improvement file | Comments about AutoBericht, tag naming, UI behaviour, import/export, workflow or the skill | Findings about the client's SST practices |
| Transcript | Faithful recognized speech, language changes and marked recognition uncertainties | Silent replacement with a summary or the polished report |

These boundaries cover `findingText`, `recommendationText`, chapter positives/front matter, summaries, captions and any photo notes that the application may display or export. Photo notes should describe the item when needed; do not use them as a hidden evidence database. Preserve existing unrelated notes and source/customer fields.

A clean preview contains the **same report text** as the Sidecar, with headings appropriate to the actual structure. Do not improve only the Markdown while leaving the rejected wording in JSON. Do not append “Repère de revue”, “Source de la dictée”, “à vérifier sur la photo …” or a mini-audit below each paragraph. The working review file may group these details by row instead.

Substantive uncertainty belongs in natural report language when it is itself a supported finding: for example, “Selon nos échanges, l'attribution du contrôle n'est pas clairement définie.” An unresolved transcription or the assistant's inability to inspect a photo is a processing limitation, not a client failing. If only a question was dictated, put it in the working questions list and leave the corresponding finding/score unchanged. If part of a finding is clear, draft that part and keep the unresolved extension outside it. Do not hide a known limitation by making the report sound certain.

## One delivery after internal preparation and repair

“One shot” means the user should not have to supervise the internal passes. Complete the available work before returning it:

- Reconcile the full transcript, later corrections, withdrawn statements, positive evidence and system comments.
- Map and assemble the requested scope in the author's voice. A first batch is not an instruction to manufacture a full-company diagnosis or replace an existing management summary; update a summary only when requested or within the agreed report scope.
- Apply changes to a full Sidecar copy. Validate preservation and actual application behaviour when relevant code is available.
- Generate the clean preview from the finished fields. `scripts/report_preview.py` can do this for row-based drafts and report possible process-language leakage separately. It does not write prose, render photos, validate facts or certify style.
- Read **all changed client-facing text**, not only a sample, against the authoring contract and matching authored passages. Check field roles, scope, modality, concrete actors/actions, unnecessary abstraction, duplication, and whether useful detail was lost.
- Resolve delivery problems found by these checks, regenerate affected outputs, then verify those changes. A clean keyword scan is not a substitute for this semantic/style review.

Stop when the requested material is accounted for, technical/structural checks appropriate to this edit pass, and no identified prose or output-boundary defect remains unresolved. Do not keep rewriting already suitable paragraphs to pursue an invented style score. Do not claim this finite review guarantees error-free transcription or perfect imitation.

Deliver a full draft Sidecar and clean preview, with the transcript and compact working review material separately. Create a separate system-improvement Markdown when such comments occur. Keep detailed coverage/validation records available without making the final response another report. If uncertainty cannot be resolved, name the affected decision in the working file and deliver the supported remainder. `done=false` means substantive review is still needed; it is not permission to leave avoidable prose cleanup to the author.

After user review, add only genuinely approved corrections to their private profile/examples. Record the reusable lesson and its scope; do not copy client facts into the shared skill or infer approval from silence. Use those corrections on the next run so the same editorial work is not requested again.


---

## Embedded resource: references/repository-context.md

# Get current AutoBericht application context

Use the maintained repository https://github.com/odcpw/autoreport when a colleague needs current application behaviour, photo-tag/report-row integration, schema compatibility or import/export checks. Suggest this option when it would improve the current task; do not make every report depend on a checkout.

## Choose an available route

- Reuse a user-provided checkout or source archive first. Read any applicable `AGENTS.md` instructions and note the revision/version.
- With an authorized GitHub connector, read the relevant files at one commit. Connector access does not automatically authenticate terminal `git` or `gh` commands.
- With permitted shell and network access, offer to clone the repository into a separate working directory. If the user already requested repo context, proceed without asking again. For a fresh checkout:

  ```sh
  git clone https://github.com/odcpw/autoreport.git autoreport-context
  git -C autoreport-context rev-parse HEAD
  ```

  If the destination exists, inspect its remote, revision and working-tree status before using it. Do not overwrite it or pull over local work just to obtain context. A separate checkout is sufficient.
- If access is unavailable, explain the specific limitation and continue with the supplied sidecar and portable helpers when compatible. An uploaded source archive is another option. Do not ask a colleague to paste passwords or access tokens into a chat.

## Read the relevant implementation

Start with `README.md`, `docs/autobericht/workflow.md` and `docs/autobericht/mini/`. Use `AutoBericht/mini/` for the current app; `legacy/` and `experiments/` are not evidence of current production behaviour. Locate current modules rather than assuming paths in an older skill still apply.

For photo tags and report rows, trace the PhotoSorter tag changes through sidecar saving/loading and report-row normalization. A category missing from an exported row list alone does not prove the app lacks automatic row creation. For report wording, use the colleague’s own library and writing profile; app code and repository examples do not replace them.

Inspect the relevant tests and documented commands before running them. Work on a copy of the supplied sidecar and state whether verification covered JSON preservation, actual app normalization, or interactive import/export. Record the inspected commit in a separate technical note, not in report prose.

Reading/cloning for context does not authorize pushing changes or modifying the colleague’s project originals. If the user also requests implementation, follow that authorization and the repository’s instructions. Keep client recordings, photos, reports, sidecars and personal libraries out of source commits and skill distribution packages.


---

## Embedded resource: references/sidecar-contract.md

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

## Clean preview and authoring review

After applying and validating the edit plan, generate the preview from the same fields:

```sh
python3 scripts/report_preview.py project_sidecar_draft.json report_draft.md > preview_checks.json
```

This copy-only helper includes rows with `includeFinding=true` even when Done is false, respects the recommendation flag and observation/summary order, and preserves their field text. It does not fall back to generic library text if a case field is missing. Identifiers shown in headings are persistent row IDs, not a claim about the final export's display numbering.

Inspect its separate diagnostics before delivery. `reviewCandidates` locates possible process annotations or outside narration; resolve these against the authoring contract, without blindly removing legitimate quotations or an explicitly requested third-person voice. A successful exit means the preview was generated, not that prose or evidence passed review. `chapterTextNotRendered` lists existing chapter positives/front matter: include them through a suitable current-app/template route when in scope; never claim this row preview covers them, photos or Word layout. Review other changed client-facing fields too.

Keep diagnostics and unresolved decisions outside `report_draft.md` and report fields. Apply wording corrections to the Sidecar, then regenerate the preview. See [report-authoring.md](report-authoring.md) for the required semantic/style review; keyword matching cannot perform it.

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


---

## Embedded resource: references/start-and-capabilities.md

# Starting in ChatGPT Work, ChatGPT or Codex

## What a colleague provides

First-time setup can start with anonymised past reports (Word preferred; PDF or existing JSON/sidecars also useful). The assistant guides them to provide a base AutoBericht library or exported project when mapping requires its actual question IDs. They do not need to arrive with a writing profile already prepared.

For a new visit, provide:

- this skill package or the complete portable instruction document;
- the latest `project_sidecar.json` with imported self-assessment when available;
- their current `library_user_*.json` and author profile;
- the visit recording or full transcript;
- photo references already in the sidecar, or a permitted text-only filename manifest linking spoken identifiers to project-relative paths. The image files stay on the work computer.

A sidecar stores photo metadata, not the photo image bytes. Use the consultant’s descriptions as the evidence; do not request images or claim to inspect them. Do not substitute thumbnails, contact sheets, embedded report photos or screen sharing for uploads. Check that an export intended for Work contains no embedded image bytes. Reopen the returned sidecar on the work computer, where the app resolves its existing photo paths. A sidecar may contain some library text, but it is not always the complete latest reusable library.

## Minimal first message from the colleague

“Use the AutoBericht instructions I attached. Guide me through the process. If I have not yet built my library/profile, start from my anonymised reports. Otherwise use my library, sidecar and recording to prepare a complete sidecar for review. The photos stay on my work computer; use my descriptions and the sidecar’s filenames. Choose and combine the appropriate recommendation passages yourself. Keep customer answers and unrelated project data intact. Leave Done for my review. Tell me only what is missing to take the next step.”

## Capability check, without ceremony

Inspect the actual session for file reading/writing, ZIP extraction, audio transcription, code execution, model-download access and repository access; image access is not required. Do not ask the consultant to diagnose tools that you can inspect yourself. Use available capabilities and state only material gaps.

**Audio file accepted and transcription available:** try the original recording. Verify duration, language and processed coverage. If it fails because of size or runtime, split and resume. An MP3 attachment being accepted does not prove its entire speech has been transcribed.

**Audio transcription absent:** first check whether the session can prepare and run a multilingual transcription engine itself; see [transcription-runtime.md](transcription-runtime.md). Handle setup and processing when the tools permit it. If execution or necessary resources are absent, use the complete transcript from the user’s recorder, ChatGPT Record where available, or a transcription tool they choose. Preserve the audio for verification. A meeting summary is insufficient when details, negation or later corrections matter. Do not initiate a separately billed API service without the user’s authorization.

**Files can be read but not written:** produce the text and external edit plan, then explain that a tool-enabled session is required to return an edited sidecar. Do not claim the file was created.

**Repository context:** when current implementation or app-level validation would help, suggest using https://github.com/odcpw/autoreport. Reuse an accessible checkout or GitHub connection; offer a clone when execution and network access permit. Follow [repository-context.md](repository-context.md). A compatible existing sidecar can still be edited using this package alone. If the schema differs, inspect its source/docs or obtain an app-exported current sidecar before applying incompatible edits.

## Official product guidance checked on 6 September 2026

Capabilities and account availability may change; recheck these official pages when setup differs.

- ChatGPT supports reusable skills containing instructions, examples and code. Eligible workspaces can upload and share them through Plugins → Skills; workspace permissions apply. A normal attachment does not prove a skill was installed. If the uploader does not accept the supplied archive, use the complete portable document with “Create with chat”, or supply it as explicit task instructions. [Skills in ChatGPT](https://help.openai.com/en/articles/20001066)
- Work is intended for longer tasks and finished deliverables. Desktop Work can use granted local files; web/mobile Work requires uploads or connected sources. Voice in Work/Codex is documented for eligible desktop accounts. Local execution can still send task context to the cloud; it does not satisfy the photo restriction by itself. Keep image contents out of the task. [ChatGPT Work and Codex](https://help.openai.com/en/articles/20001275/)
- The generic upload guide lists document, spreadsheet and text formats. It does not establish universal transcription of arbitrary long MP3 attachments in every Work session. [Supported file types](https://help.openai.com/en/articles/8983675-what-types-of-files-are-supported)
- Record provides recording/transcription. The documentation checked lists supported paid workspaces, macOS availability and a four-hour session cap. This is a recording feature, not evidence that any uploaded MP3 will be handled identically. [ChatGPT Record](https://help.openai.com/en/articles/11487532-chatgpt-record)
- The separately used transcription API accepts MP3 and other audio formats, with a 25 MB file limit at the date checked. This is an API limit, not the ChatGPT attachment limit. Larger input can be split; API use needs the appropriate tools/account. [File transcription API](https://developers.openai.com/api/docs/guides/speech-to-text)
- The GitHub app retrieves repository content to which the colleague has granted access. Sharing this skill does not share repository permissions or credentials. Reading repository code is different from executing it or editing a local checkout. [Connecting GitHub](https://help.openai.com/en/articles/11145903)

## Long recordings

The consultant may make one long MP3 while thinking, reading and commenting. They need not pre-organise the report. Prefer natural spoken markers: “photo IMG_0123”, “question on instruction”, “from the interview”, “my assessment”, “correction”, “keep this out of the report”.

If splitting becomes necessary, a practical starting point is 20-minute chunks with 5 seconds of overlap. These are processing choices, not mandatory recording rules or product limits. `audio_chunks.py` provides offsets. Shorten chunks further for the actual service limit. Reconcile duplicated overlap and apply late corrections to earlier topics using original recording time.

Keep a progress file with source filename/hash, duration, chunks processed, failed intervals, current topic mappings and outstanding corrections. Resume unfinished ranges after interruptions. Distinguish silence, reading from a checklist, actual visit evidence and a proposed recommendation. Reading a negative generic finding aloud does not mean the author confirms that defect.

## Reuse and sharing

Use a dedicated ChatGPT project or reusable skill where available, but supply the current library version explicitly. Keep each company’s live case separate. Anonymised onboarding examples and generalised reusable language may be shared as authorized; the colleague package itself contains neither a personal corpus nor a client project. Installing in local Codex does not establish installation or synchronization in ChatGPT Work.


---

## Embedded resource: references/transcription-runtime.md

# Obtain the transcript inside the available environment

The consultant supplies permitted audio; the assistant should handle transcription mechanics when the session allows it. The photos remain on the work computer throughout. Audio language and final report language are separate choices: someone may speak Swiss German and receive a French report in their established style.

## Capability and execution

Use an existing audio transcription tool first. Otherwise inspect the actual code/shell runtime, Python, ffmpeg, package/model availability, permitted network access, disk, memory and execution limits. ChatGPT Work can expose code/shell execution, with separate network controls; availability is configuration-dependent. Do not assume an unrestricted workstation, GPU or installed speech model. [Work execution and network controls](https://learn.chatgpt.com/docs/enterprise/chatgpt-work-cloud-security)

If the user has asked to process audio and task-environment installation is permitted, prepare an isolated environment, install a multilingual speech engine and download its model weights. Avoid system-wide installation or requesting administrator access for an optional route. Do not make the user run commands that the session can run itself. If the actual runtime blocks a dependency/download, report that specific limitation and choose an available alternative. Do not endlessly retry or call a separately billed API without authorization.

Open-source Whisper is one executable option. It needs Python dependencies, ffmpeg and downloaded model weights; larger models require more resources. Installing it does not require an OpenAI API key. Use a multilingual model rather than a model ending in `.en`. Adapt this illustrative recipe to the actual environment and capacity; it is not an instruction to install software just because the skill was attached. [Official Whisper setup](https://github.com/openai/whisper)

```sh
python3 -m venv .venv-transcription
.venv-transcription/bin/python -m pip install openai-whisper
.venv-transcription/bin/whisper visit.mp3 --model turbo --language German --task transcribe --output_format all --output_dir transcript
```

This is a Unix example; on Windows the environment executables are under `Scripts`. Check ffmpeg separately. For French speech use `--language French`. For mixed languages, avoid forcing the entire recording into one language; inspect natural segments and use an engine that handles the combination or transcribe segments with the correct hint. The command downloads weights if not cached. Test execution on a short excerpt before spending the runtime on the full recording; this does not require the consultant to re-record or pre-chunk anything. Then process the full original, splitting only if needed. Keep the raw transcript and structured segment timestamps.

## German, French and Swiss German

German and French are listed Whisper languages. Swiss German is not a separate language option in that list; `German` is a starting hint, not a guarantee of dialect accuracy. Do not invent a `gsw` or `de-CH` option unless the chosen engine documents it. [Whisper language definitions](https://github.com/openai/whisper/blob/main/whisper/tokenizer.py)

For Swiss German, aim to preserve meaning in a Standard German transcript before drafting in the requested report language. Do not require phonetic dialect spelling. Check a representative passage for negations, partial implementation, specialist SST terms, photo numbers and corrections. If recognition is poor, try an available stronger or dialect-suitable engine, or ask only about the uncertain passages; do not disguise guessed speech as a clean transcript. For a first dialect trial, report quality as untested until there is actual audio evidence. Chunking addresses size and runtime, not dialect recognition accuracy.

Preserve language changes, speaker attribution, reading-aloud versus assessment, and original timing. Transcription is not report drafting. The final French/German wording comes from the author's library and profile after the content has been understood. For Whisper, `--task translate` means translation into English and is inappropriate as a shortcut to a French report. [Whisper usage](https://github.com/openai/whisper)

If audio itself must also remain on the work computer, use an authorized local transcription process and provide only the permitted transcript to Work. A local file path opened by a cloud-connected assistant is not evidence that its contents remain offline. No photo access is needed for any of these transcription routes.


---

## Embedded resource: scripts/audio_chunks.py

````python
#!/usr/bin/env python3
"""Split only when needed; keep original-time offsets. Requires ffmpeg/ffprobe."""
import argparse,hashlib,json,math,shutil,subprocess,sys
from pathlib import Path

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('output_dir');p.add_argument('--minutes',type=float,default=20);p.add_argument('--overlap',type=float,default=5);args=p.parse_args()
 if not shutil.which('ffmpeg') or not shutil.which('ffprobe'):raise ValueError('ffmpeg and ffprobe must be available')
 size=args.minutes*60
 if not math.isfinite(size) or not math.isfinite(args.overlap) or size<=0 or not 0<=args.overlap<size:raise ValueError('Require positive chunk length and 0 <= overlap < chunk length')
 src=Path(args.source).resolve();out=Path(args.output_dir)
 probe=subprocess.run(['ffprobe','-v','error','-show_entries','format=duration','-of','json',str(src)],capture_output=True,text=True,check=True)
 duration=float(json.loads(probe.stdout)['format']['duration'])
 if not math.isfinite(duration) or duration<=0:raise ValueError('Unknown/nonpositive duration')
 out.mkdir(parents=True,exist_ok=False)
 digest=hashlib.sha256()
 with src.open('rb') as f:
  for block in iter(lambda:f.read(1024*1024),b''):digest.update(block)
 manifest={'sourceFile':src.name,'sourceSha256':digest.hexdigest(),'sourceDurationSeconds':duration,'overlapSeconds':args.overlap,'complete':False,'chunks':[],'note':'Chunk transcript timestamps + sourceStartSeconds = original recording time. Deduplicate overlap; this script does not transcribe.'}
 mp=out/'audio_manifest.json'
 def save():mp.write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
 save();start=0.0;index=1
 while start<duration:
  end=min(start+size,duration);name=f'chunk_{index:04d}.mp3';target=out/name
  subprocess.run(['ffmpeg','-nostdin','-v','error','-n','-ss',str(start),'-i',str(src),'-t',str(end-start),'-map','0:a:0','-vn','-ac','1','-ar','16000','-c:a','libmp3lame','-b:a','64k','-map_metadata','-1',str(target)],check=True)
  if not target.stat().st_size:raise ValueError('Empty audio chunk')
  manifest['chunks'].append({'file':name,'sourceStartSeconds':round(start,6),'sourceEndSeconds':round(end,6),'bytes':target.stat().st_size,'transcriptionStatus':'not_started'});save()
  if end>=duration:break
  start=end-args.overlap;index+=1
 manifest['complete']=True;save();print(json.dumps({'manifest':str(mp),'chunks':len(manifest['chunks']),'sourceUnchanged':True}))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,KeyError,subprocess.CalledProcessError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
````


---

## Embedded resource: scripts/build_distribution.py

````python
#!/usr/bin/env python3
"""Package this skill only; never include adjacent user libraries or projects."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import zipfile


REPO_ROOT = Path(__file__).resolve().parents[3]
DISTRIBUTION = REPO_ROOT / 'dist' / 'autobericht'


def build(output):
    root = Path(__file__).resolve().parents[1]
    output = output.resolve()
    if output == root or root in output.parents:
        raise ValueError('Choose an output directory outside the skill source')
    files = sorted(p for p in root.rglob('*') if p.is_file()
                   and '__pycache__' not in p.parts and p.suffix != '.pyc')
    for p in files:
        if p.suffix == '.md':
            for target in re.findall(r'\]\(([^)]+)\)', p.read_text(encoding='utf-8')):
                if '://' not in target and not target.startswith('#'):
                    if not (p.parent / target.split('#')[0]).exists():
                        raise ValueError(f'Missing resource: {p.name}: {target}')
    output.mkdir(parents=True, exist_ok=True)
    parts = ['# AutoBericht — complete portable workflow\n\n'
             'Follow this workflow when the user asks. All skill resources are '
             'embedded below under their original relative filenames. Read '
             'SKILL.md first, then the relevant resources. Relative resource '
             'links refer to the corresponding embedded sections. Materialise '
             'helper code at the named paths only when execution is needed and '
             'available. The user supplies permitted project inputs separately. '
             'Site photos stay on the work computer.\n']
    ordered = [root / 'SKILL.md'] + [p for p in files if p != root / 'SKILL.md']
    for p in ordered:
        parts.append('\n\n---\n\n## Embedded resource: '
                     + p.relative_to(root).as_posix() + '\n\n')
        text = p.read_text(encoding='utf-8')
        if p.suffix == '.md':
            parts.append(text)
        else:
            lang = {'.py': 'python', '.cjs': 'javascript', '.yaml': 'yaml'}.get(p.suffix, 'text')
            parts.append('````' + lang + '\n' + text.rstrip() + '\n````\n')
    portable = output / 'AutoBericht_portable.md'
    portable.write_text(''.join(parts), encoding='utf-8', newline='\n')
    archive = (REPO_ROOT if output == DISTRIBUTION else output) / 'autobericht-skill.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in files:
            # Fixed metadata makes identical source produce identical ZIP bytes.
            info = zipfile.ZipInfo('autobericht/' + p.relative_to(root).as_posix(),
                                   date_time=(1980, 1, 1, 0, 0, 0))
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, p.read_bytes())
    with zipfile.ZipFile(archive) as z:
        if z.testzip() is not None:
            raise ValueError('Archive integrity check failed')
        for p in files:
            if z.read('autobericht/' + p.relative_to(root).as_posix()) != p.read_bytes():
                raise ValueError('Archive content mismatch')
    manifest = {p.name: {'bytes': p.stat().st_size,
                        'sha256': hashlib.sha256(p.read_bytes()).hexdigest()}
                for p in [portable, archive]}
    (output / 'package_manifest.json').write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8', newline='\n')
    if output == DISTRIBUTION:
        # Preserve the existing portable-document download URL.
        (REPO_ROOT / 'docs/autobericht/AutoBericht_portable.md').write_bytes(portable.read_bytes())
    print(json.dumps({'sourceFiles': len(files), 'artifacts': manifest}, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, default=DISTRIBUTION,
                        help='Output directory (default: repo dist/autobericht)')
    build(parser.parse_args().output)
````


---

## Embedded resource: scripts/extract_docx.py

````python
#!/usr/bin/env python3
"""Extract DOCX paragraph/table boundaries and links; no OCR or anonymisation."""
import argparse,json,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def extract(source):
 with zipfile.ZipFile(source) as z:
  if sum(x.file_size for x in z.infolist())>200*1024*1024:raise ValueError('DOCX exceeds extraction size allowance')
  rels={}
  if 'word/_rels/document.xml.rels' in z.namelist():
   for x in ET.fromstring(z.read('word/_rels/document.xml.rels')):rels[x.get('Id')]=x.get('Target')
  root=ET.fromstring(z.read('word/document.xml'));body=root.find(W+'body');blocks=[]
  def paragraph(el,where):
   t=''.join(n.text or '' if n.tag==W+'t' else '\t' if n.tag==W+'tab' else '\n' if n.tag in (W+'br',W+'cr') else '' for n in el.iter())
   links=[{'text':''.join(x.itertext()),'target':rels.get(x.get(R+'id')),'anchor':x.get(W+'anchor')} for x in el.iter(W+'hyperlink')]
   images=sum(1 for n in el.iter() if n.tag in (W+'drawing',W+'pict'))
   blocks.append({'location':where,'text':t,'links':links,'imageContainers':images})
  def walk(parent,where):
   for i,x in enumerate(parent,1):
    loc=where+'/'+str(i)
    if x.tag==W+'p':paragraph(x,loc)
    elif x.tag==W+'tbl':
     for ri,row in enumerate(x.findall(W+'tr'),1):
      for ci,cell in enumerate(row.findall(W+'tc'),1):walk(cell,loc+f'/table-row-{ri}/cell-{ci}')
    elif x.tag in (W+'sdt',W+'sdtContent'):walk(x,loc)
  if body is None:raise ValueError('Missing document body')
  walk(body,'body')
  return {'sourceFile':Path(source).name,'blocks':blocks,'limitations':['Main document body only; headers, footers, comments, tracked-deletion text and metadata require separate review.','Images are counted, not interpreted. Table cell relationships/layout may require visual inspection.','This output has not been anonymised.']}
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('output');a=p.parse_args();data=extract(a.source)
 with Path(a.output).open('x',encoding='utf-8') as f:json.dump(data,f,ensure_ascii=False,indent=2);f.write('\n')
 print(json.dumps({'blocks':len(data['blocks']),'output':a.output}))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,KeyError,zipfile.BadZipFile,ET.ParseError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
````


---

## Embedded resource: scripts/library_tool.py

````python
#!/usr/bin/env python3
"""Append reviewed recommendation variants to an existing library copy."""
import argparse,copy,hashlib,json,re,sys
from pathlib import Path

def normal(text):return re.sub(r'\s+',' ',text).strip()
def merge(source,plan,sha):
 if plan.get('format')!='library-additions/1' or plan.get('sourceSha256')!=sha:raise ValueError('Wrong format or stale source hash')
 if set(plan)-{'format','sourceSha256','additions'}:raise ValueError('Unsupported plan field')
 if not isinstance(source.get('library'),dict):raise ValueError('Missing library')
 out=copy.deepcopy(source);applied=[];skipped=[]
 for item in plan.get('additions',[]):
  if set(item)!={'kind','key','text','reviewed','generalised'}:raise ValueError('Addition requires kind,key,text,reviewed,generalised')
  if item['reviewed'] is not True or item['generalised'] is not True:raise ValueError('Only reviewed, generalised text may be merged')
  if item['kind'] not in ('entry','observation'):raise ValueError('Unknown target kind')
  group='entries' if item['kind']=='entry' else 'observations';field='id' if group=='entries' else 'value'
  found=[x for x in out['library'].get(group,[]) if str(x.get(field))==str(item['key'])]
  if len(found)!=1:raise ValueError('Target must already exist uniquely')
  text=item['text']
  if not isinstance(text,str) or not text.strip():raise ValueError('Non-empty recommendation text required')
  current=found[0].get('recommendation','')
  if not isinstance(current,str):raise ValueError('Expected recommendation string')
  blocks=re.split(r'\n\s*---\s*\n',current)
  if normal(text) in {normal(x) for x in blocks}:
   skipped.append(item['key']);continue
  found[0]['recommendation']=(current.rstrip()+'\n\n---\n\n' if current.strip() else '')+text.strip()
  applied.append(item['key'])
 # Prove all data other than the targeted recommendation strings stayed intact.
 stripped=copy.deepcopy(out)
 for group in ('entries','observations'):
  for old,new in zip(source['library'].get(group,[]),stripped['library'].get(group,[])):
   if 'recommendation' in old:new['recommendation']=old['recommendation']
   else:new.pop('recommendation',None)
 if stripped!=source:raise ValueError('Non-recommendation library data changed')
 return out,{'valid':True,'applied':applied,'exactDuplicatesSkipped':skipped,'genericFindingsPreserved':True,'semanticReviewPerformedByScript':False}
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('plan');p.add_argument('output');args=p.parse_args()
 raw=Path(args.source).read_bytes();source=json.loads(raw.decode('utf-8-sig'));plan=json.loads(Path(args.plan).read_text(encoding='utf-8-sig'))
 out,report=merge(source,plan,hashlib.sha256(raw).hexdigest())
 with Path(args.output).open('x',encoding='utf-8') as f:json.dump(out,f,ensure_ascii=False,indent=2);f.write('\n')
 print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,TypeError,KeyError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
````


---

## Embedded resource: scripts/report_preview.py

````python
#!/usr/bin/env python3
"""Render included draft rows verbatim; emit review diagnostics outside the preview.

No prose generation, style score, photo rendering, or app-export equivalence claim.
"""
import argparse
import json
import re
import sys
from pathlib import Path
import sidecar_tool as sc

LABELS = {
    'fr': ('Rapport', 'Constat', 'Recommandation'),
    'de': ('Bericht', 'Feststellung', 'Empfehlung'),
    'it': ('Rapporto', 'Constatazione', 'Raccomandazione'),
    'en': ('Report', 'Finding', 'Recommendation'),
}
# Review candidates, not a universal list of forbidden words. Never rewrite text
# automatically: explicit third-person reporting or a literal quotation can be valid.
PATTERNS = {
    'working_annotation': r'rep[eè]re\s+de\s+revue|source\s+de\s+la\s+dict[eé]e|review\s+(?:note|marker)|Pr[uü]fvermerk|nota\s+di\s+revisione',
    'processing_narration': r'selon\s+la\s+dict[eé]e|[eé]changes\s+dict[eé]s|according\s+to\s+the\s+transcript|laut\s+(?:dem\s+)?Transkript|secondo\s+la\s+trascrizione',
    'outside_narrator': r'le\s+consultant\s+(?:recommande|d[eé]crit|rel[eè]ve)|the\s+consultant\s+(?:recommends|describes)|der\s+Berater\s+empfiehlt|il\s+consulente\s+raccomanda',
}


def review_candidates(text, chapter_id, row_id, field):
    return [dict(chapterId=chapter_id, rowId=row_id, field=field,
                 kind=kind, excerpt=match.group(0))
            for kind, pattern in PATTERNS.items()
            for match in re.finditer(pattern, text, re.IGNORECASE)]


def heading(value, locale):
    if isinstance(value, dict):
        value = value.get(locale) or value.get(locale.split('-')[0]) or next(iter(value.values()), '')
    return ' '.join(str(value or '').split())


def ordered_rows(chapter):
    rows = chapter['rows']
    if str(chapter['id']) not in ('0', '4.8'):
        return rows
    by_id = {str(r.get('id')): r for r in rows if r.get('kind') != 'section'}
    result = []
    seen = set()
    for rid in chapter.get('meta', {}).get('order', []):
        rid = str(rid)
        if rid in by_id and rid not in seen:
            result.append(by_id[rid])
            seen.add(rid)
    result.extend(r for r in rows if str(r.get('id')) not in seen)
    return result


def render(doc):
    project = sc.project(doc)
    sc.row_map(doc)  # Reject ambiguous identities before writing a preview.
    locale = str(project.get('meta', {}).get('locale', ''))
    language = locale.split('-')[0].lower()
    if language not in LABELS:
        raise ValueError('Unsupported/missing report locale; render with the author\'s headings explicitly')
    title, finding_label, recommendation_label = LABELS[language]
    parts = ['# ' + title, '']
    report = {'rowsRendered': [], 'reviewCandidates': [], 'chapterTextNotRendered': [],
              'validationLevel': 'Row preview and review candidates only; semantic/style review required',
              'photosRendered': False, 'appExportTested': False}
    for chapter in project['chapters']:
        cid = str(chapter['id'])
        content = []
        for field in ('positivesText', 'frontMatterText'):
            text = chapter.get('meta', {}).get(field, '')
            if text:
                report['chapterTextNotRendered'].append({'chapterId': cid, 'field': field})
                report['reviewCandidates'].extend(review_candidates(str(text), cid, None, field))
        for row in ordered_rows(chapter):
            if row.get('kind') == 'section':
                continue
            sc.check_ws(row)
            ws = row.get('workstate', {})
            if ws.get('includeFinding') is not True:
                continue
            rid = str(row['id'])
            # Do not silently replace absent case text with a generic library finding.
            if not isinstance(ws.get('findingText'), str):
                raise ValueError(f'Missing case findingText: {cid}/{rid}')
            finding = ws['findingText']
            recommendation = ''
            if ws.get('includeRecommendation', True):
                if not isinstance(ws.get('recommendationText'), str):
                    raise ValueError(f'Missing case recommendationText: {cid}/{rid}')
                recommendation = ws['recommendationText']
            if not finding.strip() and not recommendation.strip():
                raise ValueError(f'Included row has no report text: {cid}/{rid}')
            label = heading(row.get('titleOverride') or row.get('sectionLabel') or row.get('tag'), locale)
            content.extend(['### ' + rid + (' — ' + label if label else ''), ''])
            for field, label, text in [('findingText', finding_label, finding),
                                       ('recommendationText', recommendation_label, recommendation)]:
                if text.strip():
                    content.extend(['**' + label + '**', '', text, ''])
                    report['reviewCandidates'].extend(review_candidates(text, cid, rid, field))
            report['rowsRendered'].append({'chapterId': cid, 'rowId': rid})
        if content:
            label = heading(chapter.get('title'), locale)
            parts.extend(['## ' + cid + (' — ' + label if label else ''), ''] + content)
    return '\n'.join(parts), report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('sidecar')
    parser.add_argument('preview')
    args = parser.parse_args()
    doc, _ = sc.read(args.sidecar)
    text, report = render(doc)
    # Copy-only: never overwrite the Sidecar, an earlier preview or another file.
    with Path(args.preview).open('x', encoding='utf-8', newline='\n') as output:
        output.write(text)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, TypeError, KeyError) as error:
        print('ERROR: ' + str(error), file=sys.stderr)
        sys.exit(1)
````


---

## Embedded resource: scripts/sidecar_tool.py

````python
#!/usr/bin/env python3
"""Narrow copy-only editor for the nested AutoBericht sidecar; no AI or network."""
import argparse,copy,hashlib,json,sys
from pathlib import Path

WS_FIELDS={'findingText','recommendationText','selectedLevel','includeFinding','includeRecommendation','done','priority'}
DERIVED={'scoreTouched','autoScoreLevel','findingLibraryAction','libraryAction'}
GROUPS=('report','observations','training')

def fail(message):raise ValueError(message)
def unique_pairs(pairs):
 d={}
 for k,v in pairs:
  if k in d:fail('Duplicate JSON key: '+k)
  d[k]=v
 return d
def read(path):
 data=Path(path).read_bytes()
 return json.loads(data.decode('utf-8-sig'),object_pairs_hook=unique_pairs),hashlib.sha256(data).hexdigest()
def write_new(path,doc):
 p=Path(path)
 with p.open('x',encoding='utf-8') as f:json.dump(doc,f,ensure_ascii=False,indent=2);f.write('\n')
def project(doc):
 p=doc.get('report',{}).get('project')
 if not isinstance(p,dict) or not isinstance(p.get('chapters'),list):fail('Expected report.project.chapters; legacy/unknown schema needs adaptation')
 return p
def row_map(doc):
 out={};chapters=set()
 for c in project(doc)['chapters']:
  cid=str(c.get('id',''))
  if not cid or cid in chapters:fail('Missing or duplicate chapter ID')
  chapters.add(cid)
  if not isinstance(c.get('rows'),list):fail('Chapter rows must be a list')
  for r in c['rows']:
   if r.get('kind')=='section':continue
   rid=str(r.get('id',''));key=(cid,rid)
   if not rid or key in out:fail('Missing or duplicate row ID within chapter')
   out[key]=r
 return out
def options(doc,group):
 opts=doc.get('photos',{}).get('photoTagOptions',{}).get(group,[])
 return {str(x.get('value','')) if isinstance(x,dict) else str(x) for x in opts}
def check_ws(row):
 ws=row.get('workstate',{})
 if not isinstance(ws,dict):fail('workstate must be an object')
 for key in ['includeFinding','includeRecommendation','done','scoreTouched']:
  if key in ws and type(ws[key]) is not bool:fail(key+' must be boolean')
 if 'selectedLevel' in ws and (type(ws['selectedLevel']) is not int or ws['selectedLevel'] not in range(1,5)):fail('selectedLevel must be an integer 1..4')
 if 'priority' in ws and (type(ws['priority']) is not int or ws['priority'] not in range(5)):fail('priority must be an integer 0..4')
 for key in ['findingText','recommendationText']:
  if key in ws and not isinstance(ws[key],str):fail(key+' must be a string')
def validate(before,after):
 oldrows=row_map(before);newrows=row_map(after)
 if oldrows.keys()!=newrows.keys():fail('Row identities changed')
 stripped=copy.deepcopy(after);strippedrows=row_map(stripped);changed=[]
 for key,new in newrows.items():
  old=oldrows[key];check_ws(new)
  if new==old:continue
  if old.get('kind')=='section':fail('Section rows cannot be edited')
  ow=old.get('workstate',{});nw=new.get('workstate',{})
  deltas={k for k in set(ow)|set(nw) if ow.get(k)!=nw.get(k) or (k in ow)!=(k in nw)}
  if deltas-(WS_FIELDS|DERIVED):fail('Unsupported workstate change: '+str(deltas))
  for k in ('findingLibraryAction','libraryAction'):
   if k in deltas and nw.get(k)!='off':fail('Draft editor cannot queue library actions')
  if 'selectedLevel' in deltas:
   if old.get('type') in ('field_observation','summary'):fail('No score for observation/summary')
   if nw.get('scoreTouched') is not True:fail('Changed assessment requires scoreTouched=true')
  copyrow=strippedrows[key]
  if 'workstate' in old:copyrow['workstate']=copy.deepcopy(ow)
  else:copyrow.pop('workstate',None)
  if copyrow!=old:fail('Row data outside workstate changed')
  changed.append({'chapterId':key[0],'rowId':key[1],'fields':sorted(deltas)})
 oldphotos=before.get('photos',{}).get('photos',{})
 newphotos=after.get('photos',{}).get('photos',{})
 photo_changes=[]
 for path,photo in newphotos.items():
  old=oldphotos.get(path)
  if photo==old:continue
  if not isinstance(photo,dict):fail('Photo entry must be an object')
  for k in set(photo)|(set(old) if old else set()):
   if k not in ('tags','notes') and (old is None or photo.get(k)!=old.get(k)):fail('Unsupported photo field change')
  if old is None and set(photo)-{'tags','notes'}:fail('New photo entries may contain only tags and notes')
  if 'notes' in photo and not isinstance(photo['notes'],str):fail('Photo notes must be text')
  tags=photo.get('tags',{});oldtags=(old or {}).get('tags',{})
  if set(tags)-set(GROUPS):fail('Unknown photo tag group')
  for group in GROUPS:
   vals=tags.get(group,[])
   if not isinstance(vals,list) or any(not isinstance(v,str) for v in vals):fail('Photo tags must be string lists')
   if len(vals)!=len(set(vals)):fail('Duplicate photo tags')
   if set(vals)-set(oldtags.get(group,[]))-options(before,group):fail('Unknown new tag value in '+group)
  photo_changes.append(path)
 if set(oldphotos)-set(newphotos):fail('Photo records removed')
 if photo_changes:
  if 'photos' not in before:fail('Photo branch absent; initialize through app first')
  stripped['photos']['photos']=copy.deepcopy(oldphotos)
 if stripped!=before:fail('Unrelated sidecar data changed')
 return {'valid':True,'rowChanges':changed,'photoChanges':photo_changes,'validationLevel':'JSON contract and preservation; not app/UI import'}
def apply(doc,plan,sha):
 if plan.get('format')!='autobericht-editor-patch/1':fail('Unknown patch format')
 if plan.get('sourceSha256')!=sha:fail('Source hash mismatch; regenerate against the latest sidecar')
 if set(plan)-{'format','sourceSha256','rows','photos'}:fail('Unsupported patch keys')
 out=copy.deepcopy(doc);rows=row_map(out);seen=set()
 for edit in plan.get('rows',[]):
  if set(edit)!={'chapterId','rowId','changes'}:fail('Row edit keys must be chapterId,rowId,changes')
  key=(str(edit['chapterId']),str(edit['rowId']))
  if key in seen or key not in rows:fail('Duplicate or unknown row target '+str(key))
  seen.add(key);row=rows[key]
  if row.get('kind')=='section':fail('Cannot edit section row')
  changes=edit['changes']
  if not isinstance(changes,dict) or set(changes)-WS_FIELDS:fail('Unsupported row edit field')
  ws=row.setdefault('workstate',{});material=any(ws.get(k)!=v for k,v in changes.items() if k!='done')
  if 'selectedLevel' in changes:
   if row.get('type') in ('field_observation','summary'):fail('No score for observation/summary')
   ws['scoreTouched']=True;ws['autoScoreLevel']=changes['selectedLevel']
  if 'findingText' in changes:ws['findingLibraryAction']='off'
  if 'recommendationText' in changes:ws['libraryAction']='off'
  ws.update(changes)
  if material and 'done' not in changes:ws['done']=False
  check_ws(row)
 seenphotos=set()
 for edit in plan.get('photos',[]):
  allowed={'path','notes'}|{g+s for g in GROUPS for s in ('Add','Remove')}
  if set(edit)-allowed:fail('Unsupported photo edit key')
  path=edit.get('path')
  if not isinstance(path,str) or not path or path in seenphotos:fail('Missing or duplicate photo path')
  if path.startswith(('/', '\\')) or '..' in path.replace('\\','/').split('/'):fail('Photo path must be project-relative')
  seenphotos.add(path)
  if 'photos' not in out:fail('Initialize photo branch through app first')
  ph=out['photos'].setdefault('photos',{}).setdefault(path,{'notes':'','tags':{g:[] for g in GROUPS}})
  tags=ph.setdefault('tags',{})
  for g in GROUPS:
   if g+'Add' not in edit and g+'Remove' not in edit:continue
   adds=edit.get(g+'Add',[]);removes=edit.get(g+'Remove',[])
   if not isinstance(adds,list) or not isinstance(removes,list) or any(not isinstance(x,str) for x in adds+removes):fail('Tag changes must be string lists')
   if set(adds)&set(removes):fail('Same tag added and removed')
   if set(adds)-options(doc,g):fail('Unknown tag value in '+g)
   values=[x for x in tags.get(g,[]) if x not in removes]
   tags[g]=list(dict.fromkeys(values+adds))
  if 'notes' in edit:ph['notes']=edit['notes']
 validate(doc,out)
 return out

def main():
 parser=argparse.ArgumentParser(description=__doc__);sub=parser.add_subparsers(dest='command',required=True)
 p=sub.add_parser('inspect');p.add_argument('source')
 p=sub.add_parser('apply');p.add_argument('source');p.add_argument('plan');p.add_argument('output')
 p=sub.add_parser('validate');p.add_argument('source');p.add_argument('result')
 args=parser.parse_args();doc,sha=read(args.source)
 if args.command=='inspect':
  rows=row_map(doc)
  print(json.dumps({'sourceSha256':sha,'locale':project(doc).get('meta',{}).get('locale'),'rows':[{'chapterId':key[0],'rowId':key[1],**row} for key,row in rows.items()],'photoTagOptions':doc.get('photos',{}).get('photoTagOptions',{}),'photoPaths':list(doc.get('photos',{}).get('photos',{}))},ensure_ascii=False,indent=2))
 elif args.command=='apply':
  plan,_=read(args.plan);out=apply(doc,plan,sha);report=validate(doc,out);write_new(args.output,out);print(json.dumps(report,ensure_ascii=False,indent=2))
 else:
  result,_=read(args.result);print(json.dumps(validate(doc,result),ensure_ascii=False,indent=2))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,TypeError,KeyError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
````


---

## Embedded resource: scripts/test_helpers.py

````python
#!/usr/bin/env python3
"""Synthetic behaviour checks for the portable helpers, without client data."""
import copy,hashlib,json,shutil,subprocess,sys,tempfile,unittest,zipfile
from pathlib import Path
import sidecar_tool as sc
import library_tool as lib
import extract_docx as dx

def fixture():
 return {'report':{'project':{'meta':{'locale':'fr-CH'},'chapters':[{'id':'1','rows':[{'id':'1.1','type':'standard','master':{'finding':'Generic','recommendation':'Base'},'customer':{'answer':1,'items':[{'id':'1.1','comment':'Customer statement','answer':1}]},'workstate':{'selectedLevel':4,'scoreTouched':False,'findingText':'Generic','recommendationText':'Base','done':True,'includeFinding':True,'libraryAction':'append'}}]},{'id':'4.8','rows':[{'id':'4.8.1','type':'field_observation','tag':'Rayonnages','master':{'finding':'Observation','recommendation':'Base observation'},'customer':{},'workstate':{'selectedLevel':1,'done':False}}]}]}},'photos':{'photoRoot':'photos','photoTagOptions':{'observations':[{'value':'Rayonnages','label':'Rayonnages'}],'report':['1.1'],'training':['Training']},'photos':{'photos/a.jpg':{'tags':{'report':['1.1'],'observations':[],'training':['Training']},'notes':'Keep me'}}},'spider':{'unrelated':[1,2]},'unknownFutureField':{'keep':True}}

class Helpers(unittest.TestCase):
 def test_case_edit_preserves_source_and_review_gate(self):
  source=fixture();source['report']['project']['chapters'][0]['rows'].append({'kind':'section','title':'Section heading without row ID'});snapshot=copy.deepcopy(source)
  plan={'format':'autobericht-editor-patch/1','sourceSha256':'test','rows':[{'chapterId':'1','rowId':'1.1','changes':{'selectedLevel':2,'recommendationText':'A case-specific draft'}}],'photos':[{'path':'photos/a.jpg','observationsAdd':['Rayonnages']}]}
  out=sc.apply(source,plan,'test');self.assertEqual(source,snapshot)
  r=out['report']['project']['chapters'][0]['rows'][0]
  self.assertFalse(r['workstate']['done']);self.assertTrue(r['workstate']['scoreTouched']);self.assertEqual(r['workstate']['libraryAction'],'off')
  self.assertEqual(r['customer'],source['report']['project']['chapters'][0]['rows'][0]['customer'])
  self.assertEqual(r['master'],source['report']['project']['chapters'][0]['rows'][0]['master'])
  self.assertEqual(out['photos']['photos']['photos/a.jpg']['notes'],'Keep me');self.assertEqual(out['photos']['photos']['photos/a.jpg']['tags']['training'],['Training'])
  self.assertEqual(out['spider'],source['spider']);self.assertTrue(sc.validate(source,json.loads(json.dumps(out)))['valid'])
 def test_rejects_stale_unknown_invalid_and_unrelated_changes(self):
  source=fixture();base={'format':'autobericht-editor-patch/1','sourceSha256':'test','rows':[]}
  with self.assertRaises(ValueError):sc.apply(source,base,'different')
  for change in [{'selectedLevel':80},{'selectedLevel':True},{'priority':7},{'invented':1}]:
   p={**base,'rows':[{'chapterId':'1','rowId':'1.1','changes':change}]}
   with self.assertRaises(ValueError):sc.apply(source,p,'test')
  p={**base,'photos':[{'path':'photos/a.jpg','observationsAdd':['Unknown category']} ]}
  with self.assertRaises(ValueError):sc.apply(source,p,'test')
  p={**base,'rows':[{'chapterId':'4.8','rowId':'4.8.1','changes':{'selectedLevel':2}}]}
  with self.assertRaises(ValueError):sc.apply(source,p,'test')
  changed=copy.deepcopy(source);changed['report']['project']['chapters'][0]['rows'][0]['customer']['answer']=0
  with self.assertRaises(ValueError):sc.validate(source,changed)
  with self.assertRaises(ValueError):sc.row_map({'chapters':[]})
 def test_explicit_done_and_no_overwrite(self):
  source=fixture();plan={'format':'autobericht-editor-patch/1','sourceSha256':'test','rows':[{'chapterId':'1','rowId':'1.1','changes':{'findingText':'Reviewed new text','done':True}}]}
  out=sc.apply(source,plan,'test');self.assertTrue(out['report']['project']['chapters'][0]['rows'][0]['workstate']['done'])
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'out.json';sc.write_new(p,out)
   with self.assertRaises(FileExistsError):sc.write_new(p,out)
 def test_library_addition_and_repeat(self):
  source={'schemaVersion':1,'structure':{'items':[]},'library':{'entries':[{'id':'1.1','finding':'Generic','recommendation':'First'}],'observations':[]},'tags':{'keep':True}}
  item={'kind':'entry','key':'1.1','text':'Second useful variant','reviewed':True,'generalised':True}
  plan={'format':'library-additions/1','sourceSha256':'test','additions':[item,item]}
  out,report=lib.merge(source,plan,'test');self.assertEqual(report['applied'],['1.1']);self.assertEqual(report['exactDuplicatesSkipped'],['1.1']);self.assertEqual(out['library']['entries'][0]['finding'],'Generic')
  self.assertEqual(source['library']['entries'][0]['recommendation'],'First')
  _,repeat=lib.merge(out,plan,'test');self.assertEqual(repeat['applied'],[])
  with self.assertRaises(ValueError):lib.merge(source,{**plan,'additions':[{**item,'reviewed':False}]},'test')
 def test_docx_table_and_hyperlink(self):
  xml='''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Introduction</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Finding</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Recommendation </w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>reference</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'''
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'synthetic.docx'
   with zipfile.ZipFile(p,'w') as z:
    z.writestr('word/document.xml',xml);z.writestr('word/_rels/document.xml.rels','<Relationships><Relationship Id="rId1" Target="https://example.org/reference"/></Relationships>')
   data=dx.extract(p);self.assertEqual([b['text'] for b in data['blocks']],['Introduction','Finding','Recommendation reference']);self.assertIn('/cell-2/',data['blocks'][2]['location']);self.assertEqual(data['blocks'][2]['links'][0]['target'],'https://example.org/reference')
 @unittest.skipUnless(shutil.which('ffmpeg') and shutil.which('ffprobe'),'ffmpeg/ffprobe unavailable')
 def test_audio_chunk_offsets_and_original(self):
  with tempfile.TemporaryDirectory() as temp:
   root=Path(temp);src=root/'synthetic.wav'
   subprocess.run(['ffmpeg','-v','error','-f','lavfi','-i','sine=frequency=400:duration=4','-c:a','pcm_s16le',str(src)],check=True)
   original=hashlib.sha256(src.read_bytes()).hexdigest()
   subprocess.run([sys.executable,str(Path(__file__).with_name('audio_chunks.py')),str(src),str(root/'chunks'),'--minutes','0.03','--overlap','0.2'],capture_output=True,text=True,check=True)
   m=json.loads((root/'chunks/audio_manifest.json').read_text());self.assertTrue(m['complete']);self.assertEqual(m['sourceSha256'],original);self.assertEqual(hashlib.sha256(src.read_bytes()).hexdigest(),original)
   self.assertEqual(m['chunks'][0]['sourceStartSeconds'],0);self.assertEqual(m['chunks'][-1]['sourceEndSeconds'],4)
   for a,b in zip(m['chunks'],m['chunks'][1:]):self.assertAlmostEqual(a['sourceEndSeconds']-b['sourceStartSeconds'],0.2)
if __name__=='__main__':unittest.main(verbosity=2)
````


---

## Embedded resource: scripts/test_report_preview.py

````python
"""Check draft-field fidelity, output separation, selection and copy-only behaviour."""
import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import report_preview as preview
from test_helpers import fixture


class ReportPreview(unittest.TestCase):
    def test_unreviewed_included_text_is_preserved_without_private_context(self):
        source = fixture()
        row = source['report']['project']['chapters'][0]['rows'][0]
        row['workstate'].update(done=False, findingText='Selon nos discussions, les délais ne sont pas suivis.',
                               recommendationText='Nous recommandons de reprendre les actions ouvertes.\n\nPréserver ce détail utile.',
                               includeRecommendation=True)
        row['master']['finding'] = 'Le consultant recommande une formule à ne pas importer.'
        row['customer']['remark'] = 'Repère de revue : conserver dans les données sources.'
        source['photos']['photos']['photos/a.jpg']['notes'] = 'Unrelated photo note'
        snapshot = copy.deepcopy(source)
        text, report = preview.render(source)
        self.assertIn(row['workstate']['findingText'], text)
        self.assertIn(row['workstate']['recommendationText'], text)
        self.assertNotIn(row['master']['finding'], text)
        self.assertNotIn(row['customer']['remark'], text)
        self.assertNotIn('Unrelated photo note', text)
        self.assertEqual(report['reviewCandidates'], [])
        self.assertEqual(source, snapshot)
        self.assertFalse(row['workstate']['done'])

    def test_leakage_is_reported_separately_without_silent_rewriting(self):
        source = fixture()
        ws = source['report']['project']['chapters'][0]['rows'][0]['workstate']
        ws['findingText'] = 'Les échanges dictés montrent un écart.\n\nRepère de revue : 02:00.'
        ws['recommendationText'] = 'Le consultant recommande un contrôle.'
        text, report = preview.render(source)
        self.assertIn(ws['findingText'], text)
        self.assertIn(ws['recommendationText'], text)
        self.assertEqual({x['kind'] for x in report['reviewCandidates']},
                         {'processing_narration', 'working_annotation', 'outside_narrator'})
        self.assertNotIn('reviewCandidates', text)
        self.assertNotIn('validationLevel', text)

    def test_disabled_recommendation_and_excluded_rows_do_not_leak(self):
        source = fixture()
        ws = source['report']['project']['chapters'][0]['rows'][0]['workstate']
        ws.update(includeRecommendation=False, recommendationText='Repère de revue : hidden')
        hidden = source['report']['project']['chapters'][1]['rows'][0]
        hidden['workstate'].update(findingText='Do not include this observation', includeFinding=False)
        text, report = preview.render(source)
        self.assertNotIn('hidden', text)
        self.assertNotIn('Do not include this observation', text)
        self.assertEqual(len(report['rowsRendered']), 1)
        self.assertEqual(report['reviewCandidates'], [])

    def test_observation_order_and_unrendered_chapter_content_are_explicit(self):
        source = fixture()
        chapter = source['report']['project']['chapters'][1]
        first = chapter['rows'][0]
        first['workstate'].update(includeFinding=True, findingText='Premier constat.', includeRecommendation=False)
        second = copy.deepcopy(first)
        second['id'] = '4.8.2'
        second['workstate']['findingText'] = 'Second constat.'
        chapter['rows'].append(second)
        chapter['meta'] = {'order': ['4.8.2', '4.8.1'], 'positivesText': 'Texte de chapitre existant.'}
        text, report = preview.render(source)
        self.assertLess(text.index('Second constat.'), text.index('Premier constat.'))
        self.assertEqual(report['chapterTextNotRendered'], [{'chapterId': '4.8', 'field': 'positivesText'}])

    def test_missing_case_text_fails_instead_of_falling_back_to_seed(self):
        source = fixture()
        del source['report']['project']['chapters'][0]['rows'][0]['workstate']['findingText']
        with self.assertRaisesRegex(ValueError, 'Missing case findingText'):
            preview.render(source)

    def test_cli_writes_preview_and_separate_diagnostics_without_overwrite(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / 'sidecar.json'
            target = Path(temp) / 'report.md'
            source.write_text(json.dumps(fixture()))
            before = source.read_bytes()
            cmd = [sys.executable, str(Path(preview.__file__)), str(source), str(target)]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            self.assertEqual(len(json.loads(result.stdout)['rowsRendered']), 1)
            self.assertNotIn('validationLevel', target.read_text())
            self.assertEqual(source.read_bytes(), before)
            existing = target.read_bytes()
            repeat = subprocess.run(cmd, capture_output=True, text=True)
            self.assertNotEqual(repeat.returncode, 0)
            self.assertEqual(target.read_bytes(), existing)


if __name__ == '__main__':
    unittest.main(verbosity=2)
````


---

## Embedded resource: scripts/verify_app.cjs

````javascript
#!/usr/bin/env node
// Optional validation against a supplied AutoBericht checkout. No network.
const fs=require('node:fs');const path=require('node:path');const vm=require('node:vm');const assert=require('node:assert/strict');
const [repo,libraryPath,sidecarPath]=process.argv.slice(2);
if(!repo||!libraryPath){console.error('Usage: node verify_app.cjs REPO LIBRARY_JSON [SIDECAR_JSON]');process.exit(2);}
const hooks={};const ctx={console,structuredClone,TextEncoder,TextDecoder,URL,setTimeout,clearTimeout,__AUTO_BERICHT_TEST__:hooks,document:{getElementById(){return null;},addEventListener(){},visibilityState:'visible'},addEventListener(){},AutoReportDebug:{logLine(){}},AutoBerichtI18n:{t:(k,f)=>f||k,tf:(k,f)=>f||k,setLocale(){},resolveSpellcheckLang:l=>l||'fr-CH'},AutoBerichtFsHandle:{saveHandle:async()=>{},loadHandle:async()=>null,requestHandlePermission:async()=>false}};ctx.window=ctx;ctx.globalThis=ctx;vm.createContext(ctx);
for(const file of ['dependencies.js','state.js','normalize.js','seeds.js','report-rows.js']){const p=path.join(repo,'AutoBericht/mini/shared',file);vm.runInContext(fs.readFileSync(p,'utf8'),ctx,{filename:p});}
const makerPath=path.join(repo,'AutoBericht/mini/librarymaker.js');vm.runInContext(fs.readFileSync(makerPath,'utf8'),ctx,{filename:makerPath});
const library=JSON.parse(fs.readFileSync(libraryPath,'utf8'));ctx.AutoBerichtSeeds.validateKnowledgeBase(library);
const maker=hooks.librarymaker?.normalizeKnowledgeBaseForMaker;assert.equal(typeof maker,'function','LibraryMaker validation hook unavailable in this revision');
const normal=maker(library);
for(const group of ['entries','observations']){const key=group==='entries'?'id':'value';for(const item of library.library[group]||[]){const actual=normal.library[group].find(x=>x[key]===item[key]);assert(actual);assert.equal(actual.finding,item.finding||'');assert.equal(actual.recommendation,item.recommendation||'');}}
const project=ctx.AutoBerichtSeeds.buildProjectFromKnowledgeBase(library);const text=JSON.stringify(project);
// The importer addresses standard entries by collapsedId/groupId, not every raw
// questionnaire sub-item ID. A schema-valid library can therefore lose text at
// project creation. Report coverage separately and fail rather than hiding loss.
const missingRecommendations=[];
for(const item of [...(library.library.entries||[]),...(library.library.observations||[])]){if(item.recommendation&&!text.includes(JSON.stringify(item.recommendation).slice(1,-1)))missingRecommendations.push(item.id||item.value);}
let checked=0;
if(sidecarPath){
 const sidecar=JSON.parse(fs.readFileSync(sidecarPath,'utf8'));assert(sidecar.report?.project?.chapters,'Expected nested sidecar');
 const copy=JSON.parse(JSON.stringify(sidecar));ctx.AutoBerichtNormalize.normalizeProject(copy.report.project);ctx.AutoBerichtNormalize.syncObservationChapterRows(copy.report.project,copy);
 for(const c of sidecar.report.project.chapters){for(const row of c.rows||[]){if(row.kind==='section')continue;const actual=copy.report.project.chapters.find(x=>String(x.id)===String(c.id))?.rows.find(x=>x.kind!=='section'&&x.id===row.id);assert(actual,`Row lost after normalisation: ${row.id}`);assert.deepEqual(actual.customer,row.customer);
  for(const key of ['findingText','recommendationText','selectedLevel','includeFinding','includeRecommendation','done']){if(Object.hasOwn(row.workstate||{},key))assert.equal(actual.workstate[key],row.workstate[key],`Changed ${row.id}/${key}`);}
  if(row.type==='field_observation'||row.type==='summary')assert.equal(ctx.AutoBerichtState.calculateScore(actual),null);checked++;
 }}
 assert.deepEqual(copy.photos,sidecar.photos);
}
console.log(JSON.stringify({librarySchema:'passed',libraryMaker:'passed',projectImport:'passed',recommendationCoverage:missingRecommendations.length?'failed':'passed',missingRecommendations,sidecarRowsChecked:checked,interactiveImportTested:false},null,2));
if(missingRecommendations.length)process.exitCode=1;
````
