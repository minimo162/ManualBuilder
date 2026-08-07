from __future__ import annotations

import base64
import json
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / "tools" / "evals" / "evaluate_recorder_copilot_runs.py"
PRODUCTION_MANIFEST = ROOT / "samples" / "evals" / "recorder" / "gold" / "manifest.json"
JPEG = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////"
    "2wBDAf//////////////////////////////////////////////////////////////////////////////////////"
    "wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAEf/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIAAwAAABAf/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPxB//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPxB//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxB//9k="
) + b"\0" * 1024


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def write_jsonl(path: Path, values: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(json.dumps(value, ensure_ascii=False) for value in values) + "\n", encoding="utf-8")


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)
    print(f"  OK   {message}")


production = json.loads(PRODUCTION_MANIFEST.read_text(encoding="utf-8"))
production_scenarios = production["scenarios"]
check([item["id"] for item in production_scenarios] == [
    "edge-excel-order-transfer", "excel-multi-operation", "edge-delayed-transition"],
    "実践的なEdge→Excel・Excel複数操作・Edge遅延遷移の3正解を固定する")
check(all(item["captureProvenance"] == "manualbuilder-recorder-real-gui" for item in production_scenarios),
      "各正解が実Recorder GUI captureを要求する")
check((ROOT / production_scenarios[0]["source"]["script"]).is_file()
      and (ROOT / production_scenarios[2]["source"]["fixture"]).is_file(),
      "実機スモークと遅延遷移fixtureの参照先が存在する")
check(sum(len(item["steps"]) for item in production_scenarios) == 16,
      "3シナリオ合計16手順の正解がある")


with tempfile.TemporaryDirectory(prefix="mb-recorder-run-eval-") as value:
    temp = Path(value)
    manifest = {
        "fixtureVersion": "test-real-1", "evaluationDefaults": {
            "minimumRunsPerScenario": 1, "minimumFrameCount": 8,
            "minimumCaptureDurationMs": 3000, "maximumEventDistanceMs": 1000,
            "transitionTitlePatterns": ["読み込み中"],
        },
        "scenarios": [{
            "id": "practical", "forbiddenClaims": ["Excelを開く"], "steps": [
                {"order": 1, "app": "edge", "requiredConcepts": ["顧客コード", "入力"],
                 "exactTokens": ["X-1"], "eventKinds": ["input"], "targetRequired": True,
                 "afterRequired": False},
                {"order": 2, "app": "edge", "requiredConcepts": ["検索"],
                 "exactTokens": [], "eventKinds": ["click"], "targetRequired": True,
                 "afterRequired": True},
            ],
        }],
    }
    manifest_path, output_path = temp / "manifest.json", temp / "machine.json"
    write_json(manifest_path, manifest)
    run = temp / "runs" / "practical" / "run-01"
    frames = []
    for index in range(1, 9):
        name = f"frame-{index:05d}.jpg"
        path = run / "frames" / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(JPEG)
        frames.append({
            "id": f"F{index:05d}", "index": index, "timeMs": (index - 1) * 500,
            "image": name, "windowTitle": "受注検索 - Microsoft Edge",
        })
    events = [
        {"index": 1, "timeMs": 700, "kind": "input", "targetName": "顧客コード",
         "windowTitle": "受注検索 - Microsoft Edge", "clickPoint": {"x": 100, "y": 100}},
        {"index": 2, "timeMs": 1700, "kind": "click", "targetName": "検索",
         "windowTitle": "受注検索 - Microsoft Edge", "rect": {"x1": 0.1, "y1": 0.1, "x2": 0.2, "y2": 0.2}},
    ]
    proposals = [
        {"beforeFrame": "F00002", "afterFrame": "", "eventIds": [1], "targetEventId": 1,
         "title": "顧客コードを入力", "description": "顧客コードへX-1を入力します。"},
        {"beforeFrame": "F00004", "afterFrame": "F00006", "eventIds": [2], "targetEventId": 2,
         "title": "注文を検索", "description": "［検索］をクリックします。"},
    ]
    write_json(run / "status.json", {"jobId": "job-1", "state": "completed"})
    write_json(run / "result.json", {"jobId": "job-1", "proposals": proposals})
    write_jsonl(run / "frames.jsonl", frames)
    write_jsonl(run / "events.jsonl", events)
    (run / "copilot.log").write_text("[INFO] completed\n", encoding="utf-8")

    result = subprocess.run([
        sys.executable, str(TOOL), "--manifest", str(manifest_path),
        "--runs-root", str(temp / "runs"), "--output", str(output_path),
    ], capture_output=True, text=True, encoding="utf-8")
    check(result.returncode == 0, "実Recorder形式のbundleを採点できる")
    machine = json.loads(output_path.read_text(encoding="utf-8"))
    check(machine["evaluationScope"]["fixtureType"] == "real-recorder-copilot",
          "合成映像評価と区別したprovenanceを出力する")
    check(machine["sceneCountPasses"] == 1 and machine["representativeTimeHits"] == 2,
          "手順数・順序・概念・アプリ・イベント時刻を合格にする")
    check(machine["top4RectHits"] == 2 and machine["rectTotal"] == 2,
          "targetEventのclickPoint/rectを赤枠アンカーとして採点する")
    detail = machine["scenarios"][0]["runs"][0]
    check(detail["artifactDigest"] and all(item["imageFilesOk"] for item in detail["scenes"]),
          "参照JPEGを実在確認しbundle digestを残す")

    # JSON回答だけを置いたsynthetic-only runは、Recorderの原本がないため合格できない。
    incomplete = temp / "incomplete" / "practical" / "run-01"
    write_json(incomplete / "result.json", {"jobId": "fake", "proposals": proposals})
    incomplete_output = temp / "incomplete-machine.json"
    subprocess.run([
        sys.executable, str(TOOL), "--manifest", str(manifest_path),
        "--runs-root", str(temp / "incomplete"), "--output", str(incomplete_output),
    ], check=True, capture_output=True, text=True, encoding="utf-8")
    failed = json.loads(incomplete_output.read_text(encoding="utf-8"))
    check(failed["sceneCountPasses"] == 0 and failed["runs"][0]["failureCount"] > 0,
          "回答JSONだけの合成runを合格にしない")

    # 遷移中タイトルのコマを代表画像に選ぶと、他の条件が正しくても失格にする。
    frames[3]["windowTitle"] = "読み込み中 - Microsoft Edge"
    write_jsonl(run / "frames.jsonl", frames)
    transition_output = temp / "transition-machine.json"
    subprocess.run([
        sys.executable, str(TOOL), "--manifest", str(manifest_path),
        "--runs-root", str(temp / "runs"), "--output", str(transition_output),
    ], check=True, capture_output=True, text=True, encoding="utf-8")
    transition = json.loads(transition_output.read_text(encoding="utf-8"))
    check(transition["sceneCountPasses"] == 0,
          "遷移中タイトルのbeforeFrameを選んだrunを不合格にする")

print("\nRecorderCopilot run evaluation checks passed.")
