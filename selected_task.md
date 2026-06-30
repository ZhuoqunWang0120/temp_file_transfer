# Selected Task

Selected task ID:

- `stage2_informed_task_15b_rootwords_clarified`

Source task path:

- `original_task_search_20260629_161516_PDT/stage2_informed_task_15b_rootwords_clarified/`

Why selected:

- This is a clarified variant of the original Stage-2-informed Task 15.
- It combines cross-sheet lookup, semicolon alias extraction, complete-word keyword matching, row-wise aggregation, duplicate collapse, suppression logic, spend semantics, and strict target localization.
- Local expected-output smoke passed.
- The variant explicitly clarifies how to map `Applications.Category` to `RootWords.Applies To Category`.

Stage 2 inspirations:

- `58701`: cross-sheet lookup using two source fields.
- `51431`: semicolon-separated extraction.
- `41265`: complete-word conditional matching and sum.
- `52532`: win-count aggregation across game columns.
- `42354`: explicit target sheet/range and side-effect preservation.
- `58687`: condition/range discipline and runtime robustness.

Input workbook:

- `test_files/vendor_onboarding_hold_queue.xlsx`

Expected output source:

- Embedded directly in `test_by_code.py`.

Checked sheet/range:

- `HoldQueue!A2:I11`

Output contract:

- final literal values only
- no formulas in the scored output range
- preserve all source sheets and protected notes

Known limitations:

- This is a smoke/diagnostic platform package, not benchmark-comparable scoring.
- The task is intentionally dense; it is designed to expose failure modes rather than be easy.
- The validator checks concrete expected values and side-effect preservation but does not inspect model reasoning.
