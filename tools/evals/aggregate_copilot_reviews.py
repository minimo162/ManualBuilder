"""Validate and aggregate blind Copilot quality reviews.

Reviewer files intentionally contain only fixture identity and the generated
artifact review.  Baselines, branch names, commits, and previous results are
rejected so they cannot bias a reviewer before scoring.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any


class SchemaError(ValueError):
    """Raised when an evaluation input does not match the supported schema."""


PROHIBITED_REVIEW_CONTEXT = {
    "branch",
    "branchName",
    "baseBranch",
    "commit",
    "baseline",
    "previousResult",
    "previousResults",
    "oldResult",
    "comparison",
    "expectedImprovement",
}

REVIEW_TOP_LEVEL_KEYS = {"schemaVersion", "fixtureVersion", "reviewerId", "reviews"}
SCENE_METRIC_KEYS = {
    "expected",
    "actual",
    "matched",
    "precision",
    "recall",
    "orderCorrect",
    "rectsPassing",
    "rectsEvaluated",
}


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise SchemaError(f"{path}: JSONを読み取れません: {exc}") from exc
    if not isinstance(value, dict):
        raise SchemaError(f"{path}: 最上位はJSONオブジェクトである必要があります。")
    return value


def require(condition: bool, message: str) -> None:
    if not condition:
        raise SchemaError(message)


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def validate_count(value: Any, name: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{name} は0以上の整数である必要があります。")
    return value


def validate_machine(machine: dict[str, Any]) -> dict[str, Any]:
    require(machine.get("evaluationScope") is not None and isinstance(machine["evaluationScope"], dict),
            "machine.evaluationScope が必要です。")
    scenarios = machine.get("scenarios")
    require(isinstance(scenarios, list) and len(scenarios) > 0, "machine.scenarios が必要です。")

    scenario_ids: list[str] = []
    for index, scenario in enumerate(scenarios):
        require(isinstance(scenario, dict), f"machine.scenarios[{index}] はオブジェクトである必要があります。")
        scenario_id = scenario.get("id")
        require(isinstance(scenario_id, str) and bool(scenario_id.strip()),
                f"machine.scenarios[{index}].id が必要です。")
        require(scenario_id not in scenario_ids, f"machine.scenarios のidが重複しています: {scenario_id}")
        scenario_ids.append(scenario_id)

    count_fields = (
        "sceneCountPasses",
        "sceneCountTotal",
        "representativeTimeHits",
        "representativeTimeTotal",
        "rectTotal",
        "top4RectHits",
        "noRectSceneTotal",
        "noRectFalsePositives",
    )
    counts = {name: validate_count(machine.get(name), f"machine.{name}") for name in count_fields}
    require(counts["sceneCountTotal"] > 0, "machine.sceneCountTotal は1以上である必要があります。")
    require(counts["sceneCountPasses"] <= counts["sceneCountTotal"], "sceneCountPasses が総数を超えています。")
    require(counts["representativeTimeHits"] <= counts["representativeTimeTotal"],
            "representativeTimeHits が総数を超えています。")
    require(counts["top4RectHits"] <= counts["rectTotal"], "top4RectHits が総数を超えています。")
    require(counts["noRectFalsePositives"] <= counts["noRectSceneTotal"],
            "noRectFalsePositives が総数を超えています。")

    checks = {
        "sceneCounts": counts["sceneCountPasses"] == counts["sceneCountTotal"],
        "representativeTimes": counts["representativeTimeHits"] == counts["representativeTimeTotal"],
        "rectCandidateRecallAt4": counts["top4RectHits"] == counts["rectTotal"],
        "noRectFalsePositives": counts["noRectFalsePositives"] == 0,
    }
    metrics = {
        **counts,
        "sceneCountPassRate": counts["sceneCountPasses"] / counts["sceneCountTotal"],
        "representativeTimeHitRate": (
            counts["representativeTimeHits"] / counts["representativeTimeTotal"]
            if counts["representativeTimeTotal"] else None
        ),
        "rectCandidateRecallAt4": counts["top4RectHits"] / counts["rectTotal"] if counts["rectTotal"] else None,
        "noRectFalsePositiveRate": (
            counts["noRectFalsePositives"] / counts["noRectSceneTotal"]
            if counts["noRectSceneTotal"] else None
        ),
    }
    return {"scenarioIds": scenario_ids, "checks": checks, "metrics": metrics, "pass": all(checks.values())}


def validate_string_list(value: Any, name: str) -> list[str]:
    require(isinstance(value, list), f"{name} は配列である必要があります。")
    require(all(isinstance(item, str) for item in value), f"{name} の各要素は文字列である必要があります。")
    return value


def validate_review(review: dict[str, Any], label: str, expected_scenarios: list[str]) -> dict[str, Any]:
    unexpected = set(review) - REVIEW_TOP_LEVEL_KEYS
    prohibited = set(review) & PROHIBITED_REVIEW_CONTEXT
    require(not prohibited,
            f"{label} に盲検評価へ渡せない文脈があります: {', '.join(sorted(prohibited))}")
    require(not unexpected, f"{label} に未対応の最上位項目があります: {', '.join(sorted(unexpected))}")
    require(review.get("schemaVersion") == 1, f"{label}.schemaVersion は1である必要があります。")
    require(isinstance(review.get("fixtureVersion"), str) and bool(review["fixtureVersion"].strip()),
            f"{label}.fixtureVersion が必要です。")
    require(isinstance(review.get("reviewerId"), str) and bool(review["reviewerId"].strip()),
            f"{label}.reviewerId が必要です。")
    reviews = review.get("reviews")
    require(isinstance(reviews, list) and len(reviews) > 0, f"{label}.reviews が必要です。")

    parsed: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(reviews):
        prefix = f"{label}.reviews[{index}]"
        require(isinstance(item, dict), f"{prefix} はオブジェクトである必要があります。")
        scenario_id = item.get("scenarioId")
        require(isinstance(scenario_id, str) and scenario_id in expected_scenarios,
                f"{prefix}.scenarioId が機械評価に存在しません。")
        require(scenario_id not in parsed, f"{label} のscenarioIdが重複しています: {scenario_id}")
        require(isinstance(item.get("pass"), bool), f"{prefix}.pass は真偽値である必要があります。")
        score = item.get("writingScoreOutOf5")
        require(isinstance(score, int) and not isinstance(score, bool) and 1 <= score <= 5,
                f"{prefix}.writingScoreOutOf5 は1から5の整数である必要があります。")

        metrics = item.get("sceneMetrics")
        require(isinstance(metrics, dict), f"{prefix}.sceneMetrics が必要です。")
        require(SCENE_METRIC_KEYS <= set(metrics), f"{prefix}.sceneMetrics に必須項目がありません。")
        for name in ("expected", "actual", "matched", "rectsPassing", "rectsEvaluated"):
            validate_count(metrics.get(name), f"{prefix}.sceneMetrics.{name}")
        require(metrics["matched"] <= metrics["expected"] and metrics["matched"] <= metrics["actual"],
                f"{prefix}.sceneMetrics.matched が場面数を超えています。")
        require(metrics["rectsPassing"] <= metrics["rectsEvaluated"],
                f"{prefix}.sceneMetrics.rectsPassing が評価数を超えています。")
        for name in ("precision", "recall"):
            value = metrics.get(name)
            require(is_number(value) and 0 <= value <= 1, f"{prefix}.sceneMetrics.{name} は0から1である必要があります。")
        require(isinstance(metrics.get("orderCorrect"), bool),
                f"{prefix}.sceneMetrics.orderCorrect は真偽値である必要があります。")

        for name in ("criticalFactualErrors", "missingConcepts", "unsupportedClaims", "goldIssues"):
            validate_string_list(item.get(name), f"{prefix}.{name}")
        findings = item.get("findings")
        require(isinstance(findings, list), f"{prefix}.findings は配列である必要があります。")
        require(all(isinstance(finding, dict)
                    and isinstance(finding.get("message"), str)
                    and finding.get("severity") in {"critical", "major", "minor", "info"}
                    and (finding.get("sceneOrder") is None
                         or (isinstance(finding.get("sceneOrder"), int)
                             and not isinstance(finding.get("sceneOrder"), bool)
                             and finding.get("sceneOrder") >= 1))
                    for finding in findings),
                f"{prefix}.findings の各項目にはseverity、sceneOrder、messageが必要です。")
        require(isinstance(item.get("summary"), str), f"{prefix}.summary は文字列である必要があります。")
        if item["pass"]:
            require(score >= 4, f"{prefix}.pass=true では文章点4以上が必要です。")
            require(len(item["criticalFactualErrors"]) == 0,
                    f"{prefix}.pass=true では重大な事実誤認を記録できません。")
        parsed[scenario_id] = item

    missing = [scenario_id for scenario_id in expected_scenarios if scenario_id not in parsed]
    require(not missing, f"{label} に未採点のシナリオがあります: {', '.join(missing)}")
    return {
        "fixtureVersion": review["fixtureVersion"],
        "reviewerId": review["reviewerId"],
        "reviews": parsed,
    }


def aggregate(machine: dict[str, Any], review_a: dict[str, Any], review_b: dict[str, Any]) -> dict[str, Any]:
    machine_result = validate_machine(machine)
    a = validate_review(review_a, "reviewerA", machine_result["scenarioIds"])
    b = validate_review(review_b, "reviewerB", machine_result["scenarioIds"])
    require(a["reviewerId"] != b["reviewerId"], "reviewerAとreviewerBのreviewerIdは別である必要があります。")
    require(a["fixtureVersion"] == b["fixtureVersion"], "reviewer間でfixtureVersionが一致しません。")

    disagreements: list[dict[str, Any]] = []
    scenario_results: list[dict[str, Any]] = []
    reviewers_pass = True
    for scenario_id in machine_result["scenarioIds"]:
        item_a = a["reviews"][scenario_id]
        item_b = b["reviews"][scenario_id]
        reasons: list[str] = []
        if item_a["pass"] != item_b["pass"]:
            reasons.append("pass-disagreement")
        if abs(item_a["writingScoreOutOf5"] - item_b["writingScoreOutOf5"]) > 1:
            reasons.append("writing-score-gap")
        critical_a = len(item_a["criticalFactualErrors"]) > 0
        critical_b = len(item_b["criticalFactualErrors"]) > 0
        if critical_a != critical_b:
            reasons.append("critical-error-disagreement")
        if reasons:
            disagreements.append({"scenarioId": scenario_id, "reasons": reasons})
        reviewers_pass = reviewers_pass and item_a["pass"] and item_b["pass"]
        scenario_results.append({
            "scenarioId": scenario_id,
            "reviewerA": {
                "pass": item_a["pass"],
                "writingScoreOutOf5": item_a["writingScoreOutOf5"],
                "criticalFactualErrorCount": len(item_a["criticalFactualErrors"]),
            },
            "reviewerB": {
                "pass": item_b["pass"],
                "writingScoreOutOf5": item_b["writingScoreOutOf5"],
                "criticalFactualErrorCount": len(item_b["criticalFactualErrors"]),
            },
        })

    arbitration_required = bool(disagreements)
    passed = machine_result["pass"] and reviewers_pass and not arbitration_required
    return {
        "schemaVersion": 1,
        "fixtureVersion": a["fixtureVersion"],
        "status": "passed" if passed else ("needs-arbitration" if arbitration_required else "failed"),
        "pass": passed,
        "arbitrationRequired": arbitration_required,
        "machine": {
            "pass": machine_result["pass"],
            "checks": machine_result["checks"],
            "metrics": machine_result["metrics"],
        },
        "reviewers": [
            {"reviewerId": a["reviewerId"], "pass": all(item["pass"] for item in a["reviews"].values())},
            {"reviewerId": b["reviewerId"], "pass": all(item["pass"] for item in b["reviews"].values())},
        ],
        "scenarios": scenario_results,
        "disagreements": disagreements,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="ManualBuilderの盲検Copilot評価を検証・集約します。")
    parser.add_argument("--machine", type=Path, required=True, help="場面・矩形の機械評価JSON")
    parser.add_argument("--reviewer-a", type=Path, required=True, help="独立レビュアーAのJSON")
    parser.add_argument("--reviewer-b", type=Path, required=True, help="独立レビュアーBのJSON")
    parser.add_argument("--output", type=Path, required=True, help="集約結果JSON")
    args = parser.parse_args()

    try:
        result = aggregate(load_json(args.machine), load_json(args.reviewer_a), load_json(args.reviewer_b))
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["pass"] else 1
    except SchemaError as exc:
        print(f"入力エラー: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
