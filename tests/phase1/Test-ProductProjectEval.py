from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / "tools" / "evals" / "evaluate_product_project.py"


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)
    print(f"  OK   {message}")


with tempfile.TemporaryDirectory(prefix="mb-product-eval-") as value:
    temporary = Path(value)
    manifest = {
        "evaluationDefaults": {"minimumRectIoU": 0.35},
        "scenarios": [{
            "id": "holdout", "split": "holdout", "scenes": [
                {"order": 1, "representativeTimeRangeMs": [100, 300],
                 "expectedRect": {"x1": 0.45, "y1": 0.45, "x2": 0.55, "y2": 0.55}},
                {"order": 2, "representativeTimeRangeMs": [900, 1100], "expectedRect": None},
            ],
        }],
    }
    project = {"sheets": [{"steps": [
        {"capture": {"videoTimeMs": 600, "targetCandidates": [{
            "id": "coarse-1", "rect": {"x1": 0.2, "y1": 0.2, "x2": 0.8, "y2": 0.8},
        }]}},
        {"capture": {"videoTimeMs": 1100, "targetCandidates": []}},
    ]}]}
    manifest_path, project_path, output_path = (
        temporary / "manifest.json", temporary / "project.json", temporary / "result.json")
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    project_path.write_text(json.dumps(project), encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(TOOL), "--manifest", str(manifest_path), "--project", str(project_path),
         "--scenario", "holdout", "--output", str(output_path)],
        capture_output=True, text=True, encoding="utf-8",
    )
    check(result.returncode == 0, "製品プロジェクト評価器を実行できる")
    score = json.loads(output_path.read_text(encoding="utf-8"))
    check(score["sceneCountPasses"] == 1 and score["representativeTimeHits"] == 2,
          "場面数と300msの走査許容を採点する")
    check(score["top4RectHits"] == 1 and score["strictIoURectHits"] == 0,
          "粗い候補の中心包含と厳密IoUを分けて報告する")
    check(score["noRectFalsePositives"] == 0, "結果画面へ不要な候補がないことを採点する")

print("\nProduct project evaluation checks passed.")

