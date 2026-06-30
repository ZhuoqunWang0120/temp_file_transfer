from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SOURCE_TASK = ROOT.parent / "original_task_search_20260629_161516_PDT" / "stage2_informed_task_15b_rootwords_clarified"


def run_case(candidate_source: Path | None) -> dict:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        test_files = tmp_path / "test_files"
        workspace = tmp_path / "workspace"
        eval_dir = tmp_path / "eval"
        test_files.mkdir()
        workspace.mkdir()
        eval_dir.mkdir()
        shutil.copy2(ROOT / "test_files" / "vendor_onboarding_hold_queue.xlsx", test_files / "vendor_onboarding_hold_queue.xlsx")
        if candidate_source is not None:
            shutil.copy2(candidate_source, workspace / "vendor_onboarding_hold_queue_completed.xlsx")
        env = dict(os.environ)
        env["TEST_FILES_DIR"] = str(test_files)
        env["WORKSPACE_DIR"] = str(workspace)
        env["EVAL_DIR"] = str(eval_dir)
        subprocess.run([sys.executable, str(ROOT / "test_by_code.py")], check=True, env=env)
        return json.loads((eval_dir / "code_result.json").read_text(encoding="utf-8"))


def main() -> None:
    good = run_case(SOURCE_TASK / "expected_correct_output.xlsx")
    missing = run_case(None)
    print(json.dumps({"known_correct": good, "missing_output": missing}, indent=2))
    if not good["resolved"] or good["score"] < 0.99:
        raise SystemExit("known-correct workbook did not pass")
    if missing["score"] != 0.0:
        raise SystemExit("missing-output case should score 0")


if __name__ == "__main__":
    main()
