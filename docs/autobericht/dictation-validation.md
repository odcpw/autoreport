# Validation

## Authoring workflow — 9 September 2026

- Skill structure passed the skill-creator validator, both from source and from the extracted root ZIP.
- All 12 portable helper tests passed from source and from the extracted ZIP. The six new preview checks cover exact draft wording with Done false, process-language diagnostics kept separately, inclusion flags, observation order and explicit chapter-content omissions, refusal to substitute missing case text with library defaults, and source/output preservation.
- The distribution builder verified every archived resource against the source. The portable Markdown copies and root ZIP were rebuilt together; their hashes are recorded in `dist/autobericht/package_manifest.json`.
- CI now runs the helper tests as well as checking that generated distribution files match the skill source.
- Two independent synthetic report tasks exercised contrasting confirmed profiles: French first-person recommendations and German impersonal recommendations. Both retained useful library detail, applied later corrections, separated software feedback and unresolved questions, preserved scores/source branches and produced previews matching the Sidecar fields.
- Review of the first French run caught an unsupported generic diagnosis: overdue actions had become a claim of missing follow-up. The authoring rule and synthetic example were corrected, then the same input was rerun in a fresh context. The new finding stated only the supported delay and kept the positive photo correction without a recommendation. This is a bounded regression exercise, not a percentage of style fidelity.
- Native app normalisation was also exercised on the initial synthetic inputs/outputs. It changed an untouched fixture row's automatic score and filled missing defaults even for the original input. The checks reported that limitation separately and did not write those changes into the returned Sidecars. This was not a successful full import/export certification.

These checks establish file behaviour and package consistency. They do not certify an author's style, factual accuracy, photo rendering or interactive Word export. The authoring workflow therefore also requires a semantic review of every changed client-facing passage against the supplied author evidence.

## Earlier integration baseline — 6 September 2026

- Skill structure: passed the skill-creator validator.
- Six helper tests passed: sidecar preservation and review state; rejected stale hashes/unknown fields/invalid levels or tags; explicit Done and output-overwrite protection; reviewed library additions and duplicate handling; Word table/hyperlink extraction; audio splitting with time offsets and original-file preservation.
- Actual AutoBericht code at revision `6d6e32b5f9533103e006512514f42ca0776142e7`: the author's rebuilt v3 library passed schema validation, LibraryMaker normalisation, project construction and recommendation coverage. That personal library was a local validation input only and is not distributed here.
- A synthetic edited sidecar built using the app was normalised through the real modules: 351 non-section rows checked, customer data and the checked workstate fields preserved; observation/summary scores remained null and photo metadata stayed unchanged. Separate copy-preservation checks cover unrelated fields.
- Negative coverage check: the repo's original French seed has 11 recommendation entries under sub-item IDs whose exact text does not reach a constructed project. The checker reports these and exits unsuccessfully. This is the app's collapsed-question lookup behaviour; it did not affect the rebuilt v3 library. No app or seed changes were made.

These are programmatic checks, not an interactive browser import/export test. The audio test used a short synthetic sound to verify splitting, not speech-recognition accuracy. No real MP3-to-final-sidecar run or colleague ChatGPT installation has been tested yet. The next practical test is to provide the permitted recording, sidecar and library/profile while photos remain on the work computer in the intended session, try the full recording, then chunk if required and compare the resulting draft with the consultant's review.

The package contains no personal report corpus, client sidecar, author's library or private author profile. All examples are synthetic. The shared source is versioned in the repository; private validation inputs are excluded. The photo boundary and multilingual runtime instructions are documentation changes; actual German/French/Swiss German recognition remains to be tested with speech.
