# Write the author's report, ready for substantive review

Read this before drafting any report text, including a first trial. AutoBericht saves time when the consultant reviews the judgment and can use the wording. A technically valid Sidecar that requires a prose rewrite has not met that goal. A rough recording is input to professional writing, not permission to deliver rough report text.

## Establish the authoring contract from available evidence

Use the current user's instructions and corrections first. Then use their confirmed profile and approved examples, their own finished reports in the target language, and the supplied library. Treat unreviewed workstate, generated rewrites and translated bootstrap text as weaker evidence. An assistant's suggestion is not an approved example merely because the conversation continued. Do not assign a percentage of style fidelity or claim a perfect match without a defined evaluation.

For a normal visit, read the supplied style guide as that profile and use the library already embedded in the Sidecar. The guide controls expression; the embedded library supplies candidate content, whose claims still need current evidence. Approved examples mean existing authored or confirmed passages when available, often within the guide; do not ask for a separate examples collection or library as a routine fifth input.

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
- Apply the personal guide's specific rules for each report function and inspect joins between reused and newly written sentences. Then read across the report for unintended shifts of voice. Fix the cause of a recurring mismatch instead of asking the consultant to rewrite each occurrence; retain legitimate differences between findings, recommendations and summaries.
- Resolve delivery problems found by these checks, regenerate affected outputs, then verify those changes. A clean keyword scan is not a substitute for this semantic/style review.

Stop when the requested material is accounted for, technical/structural checks appropriate to this edit pass, and no identified prose or output-boundary defect remains unresolved. Do not keep rewriting already suitable paragraphs to pursue an invented style score. Do not claim this finite review guarantees error-free transcription or perfect imitation.

Deliver a full draft Sidecar and clean preview, with the transcript and compact working review material separately. Create a separate system-improvement Markdown when such comments occur. Keep detailed coverage/validation records available without making the final response another report. If uncertainty cannot be resolved, name the affected decision in the working file and deliver the supported remainder. `done=false` means substantive review is still needed; it is not permission to leave avoidable prose cleanup to the author.

After user review, add only genuinely approved corrections to their private profile/examples. Record the reusable lesson and its scope; do not copy client facts into the shared skill or infer approval from silence. Use those corrections on the next run so the same editorial work is not requested again.
