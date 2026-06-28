import json
import math
import os
from pathlib import Path

from openpyxl import load_workbook


EXPECTED = {
    "A1": 11604.67,
    "B1": 11604.67,
    "C1": 390.68,
    "D1": 390.68,
    "E1": 0.28135,
    "F1": 0.28135,
    "G1": 16.1921,
    "H1": 16.19,
}

ORIGINAL_INPUT_NAME = "1_585-41_init.xlsx"
TOLERANCE = 1e-6
RESOLVED_THRESHOLD = 0.85


def platform_path(env_name, default):
    return Path(os.environ.get(env_name, default))


def write_result(resolved, score, reason):
    eval_dir = platform_path("EVAL_DIR", "/eval")
    eval_dir.mkdir(parents=True, exist_ok=True)
    result_path = eval_dir / "code_result.json"
    payload = {
        "resolved": bool(resolved),
        "score": float(score),
        "reason": str(reason),
    }
    result_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def list_xlsx_files(root):
    if not root.exists():
        return []
    return [
        path
        for path in root.rglob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]


def choose_workbook():
    workspace = platform_path("WORKSPACE_DIR", "/workspace")
    test_files = platform_path("TEST_FILES_DIR", "/test_files")

    workspace_files = list_xlsx_files(workspace)
    if workspace_files:
        likely_names = ("answer", "output", "result", "solved", "final")

        def rank(path):
            lower = path.name.lower()
            likely = any(token in lower for token in likely_names)
            return (1 if likely else 0, path.stat().st_mtime)

        return max(workspace_files, key=rank), "workspace"

    fallback_files = [
        path for path in list_xlsx_files(test_files)
        if path.name != ORIGINAL_INPUT_NAME
    ]
    if fallback_files:
        return max(fallback_files, key=lambda path: path.stat().st_mtime), "test_files_fallback"

    return None, "missing"


def is_numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value))


def close_enough(actual, expected):
    return is_numeric(actual) and abs(float(actual) - float(expected)) <= TOLERANCE


def evaluate_workbook(path):
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as exc:
        return 0.0, f"could not open workbook {path.name}: {type(exc).__name__}: {exc}"

    if "Sheet1" not in wb.sheetnames:
        return 0.0, "missing required sheet Sheet1"

    ws = wb["Sheet1"]
    correct = 0
    details = []
    for cell, expected in EXPECTED.items():
        actual = ws[cell].value
        if close_enough(actual, expected):
            correct += 1
        else:
            details.append(f"{cell}: expected {expected}, got {actual!r}")

    score = correct / len(EXPECTED)
    if score == 1.0:
        return score, f"all {len(EXPECTED)} checked cells matched"

    preview = "; ".join(details[:4])
    if len(details) > 4:
        preview += f"; plus {len(details) - 4} more mismatch(es)"
    return score, f"{correct}/{len(EXPECTED)} checked cells matched; {preview}"


def main():
    workbook_path, source = choose_workbook()
    if workbook_path is None:
        write_result(False, 0.0, "no model-produced .xlsx workbook found in /workspace")
        return

    score, reason = evaluate_workbook(workbook_path)
    resolved = score >= RESOLVED_THRESHOLD
    write_result(resolved, score, f"{reason}; workbook={workbook_path.name}; source={source}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        write_result(False, 0.0, f"validator error: {type(exc).__name__}: {exc}")
