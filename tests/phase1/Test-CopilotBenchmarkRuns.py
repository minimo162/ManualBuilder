from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / "tools" / "evals" / "evaluate_copilot_benchmark_runs.py"


def write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)
    print(f"  OK   {message}")


with tempfile.TemporaryDirectory(prefix="mb-copilot-runs-") as value:
    temp = Path(value)
    manifest = {
        "fixtureVersion": "test-1", "scenarios": [{
            "id": "practical", "split": "holdout", "expectedSceneCount": 3,
        }],
    }
    base = {
        "representativeTimeHits": 3, "representativeTimeTotal": 3,
        "rectTotal": 2, "top4RectHits": 2,
        "noRectSceneTotal": 1, "noRectFalsePositives": 0,
        "scenarios": [{
            "id": "practical", "rectTotal": 2, "top4RectHits": 2,
            "representativeTimeHits": 3, "representativeTimeTotal": 3,
            "noRectSceneTotal": 1, "noRectFalsePositives": 0,
        }],
    }
    manifest_path, base_path, output_path = temp / "manifest.json", temp / "base.json", temp / "machine.json"
    write(manifest_path, manifest)
    write(base_path, base)
    for number, count, failures in ((1, 3, 0), (2, 2, 1), (3, 3, 0)):
        run = temp / "runs" / f"run-{number:02d}"
        write(run / "status.json", {"state": "completed"})
        write(run / "result.json", {
            "drafts": [{"title": str(index)} for index in range(count)],
            "failures": [{"packet": 2}] * failures,
        })
    result = subprocess.run([
        sys.executable, str(TOOL), "--manifest", str(manifest_path), "--scenario", "practical",
        "--runs-root", str(temp / "runs"), "--base-machine", str(base_path), "--output", str(output_path),
    ], capture_output=True, text=True, encoding="utf-8")
    check(result.returncode == 0, "保存済みrunから機械証拠を生成できる")
    machine = json.loads(output_path.read_text(encoding="utf-8"))
    check(machine["evaluationScope"]["runCount"] == 3, "3回の反復実行を漏らさない")
    check(machine["sceneCountPasses"] == 2 and machine["sceneCountTotal"] == 3,
          "draft不足のrunを場面数合格へ数えない")
    check(machine["scenarios"][0]["expectedSceneCount"] == 9
          and machine["scenarios"][0]["actualSceneCount"] == 8,
          "期待9場面に対する実8場面を実ファイルから集計する")
    check(machine["runs"][1]["failureCount"] == 1,
          "result.jsonのpacket failureを機械証拠へ残す")
    check(machine["rectTotal"] == 6 and machine["top4RectHits"] == 6,
          "同じ入力画像に対する矩形評価を反復回数へ合わせる")

print("\nCopilot benchmark run evidence checks passed.")
