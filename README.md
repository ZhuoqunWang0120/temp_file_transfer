# DeepSeek Platform Task 15 Package

This package is a platform smoke/diagnostic task, not a benchmark-comparable score.

Selected task:

- `stage2_informed_task_15b_rootwords_clarified`

## Upload Files

Upload the files in:

- `test_files/vendor_onboarding_hold_queue.xlsx`

Paste `task_prompt.md` into the platform prompt box.

Use `test_by_code.py` as the platform test code.

The platform should run:

```bash
cd /eval && python3 test_by_code.py
```

The validator expects:

- uploaded files under `/test_files`
- model-generated workbook under `/workspace`
- result written to `/eval/code_result.json`

Expected correct result:

```json
{"resolved": true, "score": 1.0, "reason": "..."}
```

Caveat:

- The task requires final literal values in `HoldQueue!A2:I11`.
- Formula outputs in the scored range are intentionally rejected.
- This package differs from the earlier Task 15 package only by clarifying the `RootWords` mapping in the prompt.
