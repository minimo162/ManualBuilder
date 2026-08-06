from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / "tools" / "evals" / "aggregate_copilot_reviews.py"


def machine(*, false_positives: int = 0) -> dict:
    return {
        "evaluationScope": {"fixtureType": "synthetic-in-sample", "scenarioCount": 1, "isHoldout": False},
        "sceneCountPasses": 1,
        "sceneCountTotal": 1,
        "representativeTimeHits": 2,
        "representativeTimeTotal": 2,
        "rectTotal": 1,
        "top4RectHits": 1,
        "noRectSceneTotal": 1,
        "noRectFalsePositives": false_positives,
        "scenarios": [{"id": "expense-application"}],
    }


def scenario_review(*, passed: bool = True, score: int = 4, critical: bool = False) -> dict:
    return {
        "scenarioId": "expense-application",
        "pass": passed,
        "sceneMetrics": {
            "expected": 2,
            "actual": 2,
            "matched": 2,
            "precision": 1.0,
            "recall": 1.0,
            "orderCorrect": True,
            "rectsPassing": 1,
            "rectsEvaluated": 1,
        },
        "writingScoreOutOf5": score,
        "criticalFactualErrors": ["保存したと誤記"] if critical else [],
        "missingConcepts": [],
        "unsupportedClaims": [],
        "goldIssues": [],
        "findings": [],
        "summary": "評価結果",
    }


def reviewer(reviewer_id: str, **kwargs) -> dict:
    return {
        "schemaVersion": 1,
        "fixtureVersion": "1.0.0",
        "reviewerId": reviewer_id,
        "reviews": [scenario_review(**kwargs)],
    }


def invoke(root: Path, machine_data: dict, a_data: dict, b_data: dict):
    paths = {}
    for name, value in (("machine", machine_data), ("a", a_data), ("b", b_data)):
        path = root / f"{name}.json"
        path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        paths[name] = path
    output = root / "aggregate.json"
    if output.exists():
        output.unlink()
    child_environment = os.environ.copy()
    child_environment["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        [sys.executable, str(TOOL), "--machine", str(paths["machine"]),
         "--reviewer-a", str(paths["a"]), "--reviewer-b", str(paths["b"]),
         "--output", str(output)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=child_environment,
    )
    aggregate = json.loads(output.read_text(encoding="utf-8")) if output.exists() else None
    return result, aggregate


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)
    print(f"  OK   {message}")


with tempfile.TemporaryDirectory(prefix="mb-copilot-eval-") as temporary:
    temp = Path(temporary)

    result, value = invoke(temp, machine(), reviewer("blind-a"), reviewer("blind-b", score=5))
    check(result.returncode == 0, "機械評価と独立評価が合格なら終了コード0")
    check(value["pass"] and not value["arbitrationRequired"], "1点差までは裁定なしで合格")
    check("branch" not in value and "baseline" not in value, "集約結果へ旧結果やブランチを混ぜない")

    result, value = invoke(
        temp,
        machine(),
        reviewer("blind-a", passed=True, score=5, critical=False),
        reviewer("blind-b", passed=False, score=3, critical=True),
    )
    check(result.returncode == 1, "裁定が必要なら終了コード1")
    check(value["status"] == "needs-arbitration" and value["arbitrationRequired"], "不一致を裁定待ちにする")
    reasons = set(value["disagreements"][0]["reasons"])
    check(reasons == {"pass-disagreement", "writing-score-gap", "critical-error-disagreement"},
          "合否・2点差・重大誤認の不一致をすべて記録する")

    result, value = invoke(temp, machine(false_positives=1), reviewer("blind-a"), reviewer("blind-b"))
    check(result.returncode == 1 and value["status"] == "failed", "機械評価の誤検出を合意不合格にする")
    check(not value["arbitrationRequired"], "機械評価だけの失敗をレビュアー裁定へ回さない")
    check(value["machine"]["metrics"]["noRectFalsePositiveRate"] == 1.0,
          "回帰比較用の機械評価値を集約結果へ残す")

    result, value = invoke(
        temp,
        machine(),
        reviewer("blind-a", passed=False, score=3, critical=True),
        reviewer("blind-b", passed=False, score=3, critical=True),
    )
    check(result.returncode == 1 and not value["arbitrationRequired"],
          "両者が重大誤認で一致した不合格は裁定待ちにしない")

    biased = reviewer("blind-a")
    biased["branchName"] = "feature/better-prompt"
    result, value = invoke(temp, machine(), biased, reviewer("blind-b"))
    check(result.returncode == 2 and value is None, "レビュアー入力のブランチ名をschema違反として拒否する")

    duplicate = reviewer("blind-a")
    result, value = invoke(temp, machine(), duplicate, reviewer("blind-a"))
    check(result.returncode == 2 and value is None, "同じreviewerIdによる二重採点を拒否する")

print("\nCopilot評価集約器の検査はすべて成功しました。")
