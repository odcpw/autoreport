# From rambling to a review-ready report

Apply [report-authoring.md](report-authoring.md) before composing. It governs author-style evidence, minimal adaptation, clean output boundaries and the final semantic/style pass.

Begin with the MP3 recording(s), project Sidecar, this skill and the supplied style guide. Use the library already loaded in the Sidecar; a separate library or examples upload is not part of this report workflow.

Own the whole requested conversion: recording → full transcript → reconciled evidence and corrections → existing question/observation mapping → lightly adapted findings and assembled recommendations → assessment/photo assignments → full Sidecar copy → matching preview and validation. A transcript, topic map or prose suggestion is an intermediate result, not completion of a request for a ready Sidecar. Check the actual transcription/file capabilities using [start-and-capabilities.md](start-and-capabilities.md); process the available work and report a material capability gap honestly if execution cannot finish.

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

The consultant views the photos on the work computer; the assistant does not receive or inspect their pixels. Resolve spoken identifiers against the sidecar or a permitted text-only manifest, without needing access to the image files. Preserve paths so the local app can reconnect them after import. Never request images, thumbnails, contact sheets or screen sharing as a workaround. A phrase such as “photo 19” resolves to the record whose persistent `photoNumber` is 19 in `sidecar.photos.photos`. Preserve that number and its file path. For an older sidecar without these numbers, use an explicit supplied numbering map or request clarification; do not guess browser order or a filtered position. Resolve spoken identifiers to the actual file path. Match category labels/aliases against the current library and photo options. A category may have several photos; a photo may support several categories. Keep equipment/site distinctions and existing tags/notes. A category named alone permits classification, not invention of a defect. Text-only interview findings require no fabricated photo.

Prepare clear described report issues with the adapted finding, assembled recommendation, relevant photos and inclusion flags. Leave `done=false` on new or materially changed drafts, unless the consultant explicitly validates that item. A previously approved row changed by new information needs re-review. Keep explicit “off the record / not in report” material out of the delivered report.

Retain only links relevant to the selected sentences. Deduplicate at the end of the recommendation under “Voir aussi :” / “Siehe auch:” / the author’s convention. A guide’s own local procedures are not universal requirements. Do not attach old-report source citations to the prose.

## Process long input without losing its end

Keep coverage by recording range and by topic, including excluded and unresolved items. A later “correction” applies to earlier material even when it arrives in another audio chunk. Deduplicate overlap by timestamps and meaning; repeated emphasis is not necessarily a new issue. Reading existing report/library text aloud must be distinguished from adopting it for the current company.

Before delivery, sweep the entire transcript for topics never mapped, corrections not applied, conditions dropped and contradictions between yes/no, commentary, level and finding. When the requested scope includes a management summary, prepare it after the topic sweep from the current case. Do not replace an existing summary or write a full-company diagnosis merely because this is a first photo batch. Do not treat a large transcript as permission to output only its most salient themes.

## Review and outputs

Apply the supported decisions to a copy of the actual input through [sidecar-contract.md](sidecar-contract.md), preserving its full structure, customer data, library masters and unrelated fields. Set both Include flags for drafted items and newly assigned observations; leave Done false unless explicitly validated. Respect explicit exclusions and review choices. Resolve spoken photo numbers to existing records and preserve their identities. Do not insert invented row IDs or upload image content to complete the mapping.

Provide a complete Sidecar copy and a clean draft grouped by chapter, using the same field text. Keep decisions still needed and validation/coverage results in separate working material, and software feedback in a separate system-improvement file. The draft preview should include prepared rows even though the current final exporter requires both `includeFinding=true` and `done=true`. Keep the preview separate; never mark everything Done merely to obtain a complete Word export.

Before calling the Sidecar ready for review, account for the full recording, apply later corrections, check every changed passage against evidence and personal style, validate JSON and preservation, and run actual app checks when available. Report which checks ran. Isolate genuinely unresolved decisions outside the report rather than leaving avoidable prose cleanup to the consultant. An unread audio interval or incomplete mapping must remain visible as unfinished coverage, even when the supported remainder is delivered.

The target is that the consultant reviews the substance and mainly checks Done. The delivery can include unresolved items in its separate working file; a final approved report cannot pretend they were resolved. After review, use the library cycle to keep worthwhile new wording for later projects.
