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
    fixture_version = machine.get("fixtureVersion")
    require(isinstance(fixture_version, str) and bool(fixture_version.strip()),
            "machine.fixtureVersion が必要です。")
    require(machine.get("evaluationScope") is not None and isinstance(machine["evaluationScope"], dict),
            "machine.evaluationScope が必要です。")
    scenarios = machine.get("scenarios")
    require(isinstance(scenarios, list) and len(scenarios) > 0, "machine.scenarios が必要です。")
    declared_scenario_count = machine["evaluationScope"].get("scenarioCount")
    require(isinstance(declared_scenario_count, int) and not isinstance(declared_scenario_count, bool),
            "machine.evaluationScope.scenarioCount は整数である必要があります。")
    require(declared_scenario_count == len(scenarios),
            "machine.evaluationScope.scenarioCount と machine.scenarios の件数が一致しません。")

    scenario_ids: list[str] = []
    scenario_metrics: dict[str, dict[str, int]] = {}
    for index, scenario in enumerate(scenarios):
        require(isinstance(scenario, dict), f"machine.scenarios[{index}] はオブジェクトである必要があります。")
        scenario_id = scenario.get("id")
        require(isinstance(scenario_id, str) and bool(scenario_id.strip()),
                f"machine.scenarios[{index}].id が必要です。")
        require(scenario_id not in scenario_ids, f"machine.scenarios のidが重複しています: {scenario_id}")
        scenario_ids.append(scenario_id)
        expected = validate_count(scenario.get("expectedSceneCount"),
                                  f"machine.scenarios[{index}].expectedSceneCount")
        actual = validate_count(scenario.get("actualSceneCount"),
                                f"machine.scenarios[{index}].actualSceneCount")
        rect_total = validate_count(scenario.get("rectTotal"), f"machine.scenarios[{index}].rectTotal")
        rect_hits = validate_count(scenario.get("top4RectHits"), f"machine.scenarios[{index}].top4RectHits")
        require(rect_hits <= rect_total,
                f"machine.scenarios[{index}].top4RectHits がrectTotalを超えています。")
        scenario_metrics[scenario_id] = {
            "expected": expected, "actual": actual,
            "rectsPassing": rect_hits, "rectsEvaluated": rect_total,
        }

    runs = machine.get("runs")
    require(isinstance(runs, list) and len(runs) > 0,
            "machine.runs にCopilot実行ごとの証拠が必要です。")
    run_ids: set[str] = set()
    run_expected_by_scenario = {scenario_id: 0 for scenario_id in scenario_ids}
    run_actual_by_scenario = {scenario_id: 0 for scenario_id in scenario_ids}
    completed_runs = draft_count_passes = 0
    packet_failures = 0
    for index, run in enumerate(runs):
        prefix = f"machine.runs[{index}]"
        require(isinstance(run, dict), f"{prefix} はオブジェクトである必要があります。")
        run_id = run.get("id")
        require(isinstance(run_id, str) and bool(run_id.strip()), f"{prefix}.id が必要です。")
        require(run_id not in run_ids, f"machine.runs のidが重複しています: {run_id}")
        run_ids.add(run_id)
        scenario_id = run.get("scenarioId")
        require(scenario_id in scenario_metrics, f"{prefix}.scenarioId がmachine.scenariosに存在しません。")
        require(isinstance(run.get("state"), str), f"{prefix}.state が必要です。")
        expected = validate_count(run.get("expectedDraftCount"), f"{prefix}.expectedDraftCount")
        actual = validate_count(run.get("actualDraftCount"), f"{prefix}.actualDraftCount")
        failures = validate_count(run.get("failureCount"), f"{prefix}.failureCount")
        run_expected_by_scenario[scenario_id] += expected
        run_actual_by_scenario[scenario_id] += actual
        completed_runs += int(run["state"] == "completed")
        draft_count_passes += int(actual == expected)
        packet_failures += failures

    for scenario_id in scenario_ids:
        require(run_expected_by_scenario[scenario_id] == scenario_metrics[scenario_id]["expected"],
                f"{scenario_id} の実行証拠とexpectedSceneCountが一致しません。")
        require(run_actual_by_scenario[scenario_id] == scenario_metrics[scenario_id]["actual"],
                f"{scenario_id} の実行証拠とactualSceneCountが一致しません。")

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
    require(counts["sceneCountTotal"] == len(runs),
            "machine.sceneCountTotal はCopilot実行数と一致する必要があります。")
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
        "completedRuns": completed_runs == len(runs),
        "draftCounts": draft_count_passes == len(runs),
        "packetFailures": packet_failures == 0,
    }
    metrics = {
        **counts,
        "runCount": len(runs),
        "completedRunCount": completed_runs,
        "draftCountPasses": draft_count_passes,
        "packetFailureCount": packet_failures,
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
    return {
        "fixtureVersion": fixture_version,
        "scenarioIds": scenario_ids,
        "scenarioMetrics": scenario_metrics,
        "checks": checks,
        "metrics": metrics,
        "pass": all(checks.values()),
    }


def validate_string_list(value: Any, name: str) -> list[str]:
    require(isinstance(value, list), f"{name} は配列である必要があります。")
    require(all(isinstance(item, str) for item in value), f"{name} の各要素は文字列である必要があります。")
    return value


def validate_review(review: dict[str, Any], label: str,
                    expected_scenarios: list[str], machine_scenarios: dict[str, dict[str, int]]) -> dict[str, Any]:
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
        machine_metrics = machine_scenarios[scenario_id]
        for review_name, machine_name in (
            ("expected", "expected"), ("actual", "actual"),
            ("rectsPassing", "rectsPassing"), ("rectsEvaluated", "rectsEvaluated"),
        ):
            require(metrics[review_name] == machine_metrics[machine_name],
                    f"{prefix}.sceneMetrics.{review_name} が機械評価の実数と一致しません。")
        expected_precision = metrics["matched"] / metrics["actual"] if metrics["actual"] else 0.0
        expected_recall = metrics["matched"] / metrics["expected"] if metrics["expected"] else 0.0
        require(math.isclose(float(metrics["precision"]), expected_precision, abs_tol=1e-6),
                f"{prefix}.sceneMetrics.precision がmatched/actualと一致しません。")
        require(math.isclose(float(metrics["recall"]), expected_recall, abs_tol=1e-6),
                f"{prefix}.sceneMetrics.recall がmatched/expectedと一致しません。")

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
            require(metrics["precision"] == 1 and metrics["recall"] == 1 and metrics["orderCorrect"],
                    f"{prefix}.pass=true では全場面の一致と正しい順序が必要です。")
            require(metrics["rectsPassing"] == metrics["rectsEvaluated"],
                    f"{prefix}.pass=true では全操作矩形が合格する必要があります。")
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
    a = validate_review(review_a, "reviewerA", machine_result["scenarioIds"], machine_result["scenarioMetrics"])
    b = validate_review(review_b, "reviewerB", machine_result["scenarioIds"], machine_result["scenarioMetrics"])
    require(a["reviewerId"] != b["reviewerId"], "reviewerAとreviewerBのreviewerIdは別である必要があります。")
    require(a["fixtureVersion"] == b["fixtureVersion"], "reviewer間でfixtureVersionが一致しません。")
    require(a["fixtureVersion"] == machine_result["fixtureVersion"],
            "reviewerとmachineでfixtureVersionが一致しません。")

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
