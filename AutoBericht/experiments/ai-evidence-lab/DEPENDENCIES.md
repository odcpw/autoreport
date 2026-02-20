# AI Evidence Lab - Dependency Graph and Parallel Lanes

This file maps task dependencies and defines the 5-lane parallel execution strategy.

## Dependency Backbone

The build is a DAG with these major dependency gates:

1. `G1 Foundation`: T001-T010
2. `G2 Runtime + IO`: T011-T032 (depends on G1)
3. `G3 Contracts + Orchestration`: T033-T045 (depends on G1, partly on G2)
4. `G4 Extraction`: T046-T066 (depends on G2 + G3)
5. `G5 Indexing`: T067-T082 (depends on G4)
6. `G6 Retrieval/UI`: T083-T097 (depends on G5 + G2)
7. `G7 Generation/Export`: T098-T116 (depends on G6 + G3)
8. `G8 Quality Layers`: T117-T152 (cross-cutting; most depend on G2-G7)
9. `G9 Docs + DoD`: T153-T165 (depends on all prior gates)

## Task-Level Critical Dependencies

- T011-T017 depend on T001-T003.
- T018-T025 depend on T001-T003 and T011.
- T026-T032 depend on T021 plus T001-T003.
- T036-T038 depend on T033-T035.
- T039-T045 depend on T018-T025 and T004.
- T046-T052 depend on T039-T045.
- T053-T058 depend on T039-T045.
- T059-T066 depend on T039-T045 and T046-T050.
- T067-T074 depend on T046-T066.
- T075-T082 depend on T067-T074.
- T083-T090 depend on T075-T082 and T026-T032.
- T091-T097 depend on T083-T090 and T031-T032.
- T098-T104 depend on T083-T097.
- T105-T110 depend on T033-T038 and T083-T104.
- T111-T116 depend on T033-T038 and T026-T032 and T105-T110.
- T117-T123 begin at G1 but complete at G7.
- T124-T129 depend on T039-T045 and T075-T090.
- T130-T135 depend on T075-T090 and T124-T129.
- T136-T142 depend on T046-T116.
- T143-T147 depend on T011-T017 and T075-T110.
- T148-T152 depend on T105-T116.
- T153-T158 depend on stabilized implementation from G1-G8.
- T159-T165 depend on successful G1-G8 and docs.

## 5 Parallel Lanes (Subagent-Style Workstreams)

Lane A - Shell + Runtime
- Primary tasks: T001-T017
- Outputs: app shell, capability detection, startup diagnostics

Lane B - File IO + Sidecar Adapter
- Primary tasks: T018-T032
- Outputs: folder pick, inventory, row picker, selected-row panel

Lane C - Contracts + Export Plumbing
- Primary tasks: T033-T038, T105-T116
- Outputs: schemas, validation helpers, proposal export, patch preview

Lane D - Ingestion + Extraction Engine
- Primary tasks: T039-T074
- Outputs: worker protocol, ingestion lifecycle, text/OCR/chunk pipeline

Lane E - Retrieval + UX + Ops
- Primary tasks: T075-T165
- Outputs: retrieval, evidence UI, generation, logs, QA, docs, DoD

## Parallel Waves

Wave 1 (high parallelism)
- Lane A: T001-T017
- Lane B: T018-T032
- Lane C: T033-T036
- Lane D: T039-T041
- Lane E: T117-T122

Wave 2
- Lane C: T037-T038 + T105-T110
- Lane D: T042-T066
- Lane E: T083-T097

Wave 3
- Lane C: T111-T116
- Lane D: T067-T082 (if not complete)
- Lane E: T098-T104 + T124-T165

## Current Execution Snapshot

Completed in this pass (Wave 1 core):
- Lane A: T001-T017
- Lane B: T018-T032
- Lane C: T033-T036
- Lane D: T039-T041, T044-T045
- Lane E: T117, T122

In progress next:
- Lane C: T037-T038
- Lane D: T042-T043
- Lane E: T118-T121, T123

