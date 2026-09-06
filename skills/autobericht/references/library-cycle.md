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

Teach only confirmed style preferences. An accepted case-specific measure is not automatically a preference to include that measure in all reports. A reviewed phrase becomes an example in the appropriate function—finding, recommendation or summary. Do not train the profile on the assistant’s own unreviewed drafts.
