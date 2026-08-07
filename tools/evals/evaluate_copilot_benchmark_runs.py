"""Build tamper-resistant machine evidence from saved Copilot benchmark runs.

The blind reviewers judge wording and semantic scene matches.  Counts, worker
state, and packet failures are derived here from the immutable run artifacts so
reviewers cannot accidentally (or optimistically) overwrite those facts.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path}")
    return value


def count_generated(result: dict[str, Any]) -> int:
    for name in ("drafts", "proposals"):
        value = result.get(name)
        if isinstance(value, list):
            return len(value)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, default=Path("samples/evals/gold/manifest.json"))
    parser.add_argument("--scenario", required=True)
    parser.add_argument("--runs-root", type=Path, required=True)
    parser.add_argument("--base-machine", type=Path, required=True,
                        help="One-run scene/time/rectangle machine evaluation")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    manifest = load(args.manifest)
    scenario = next((item for item in manifest.get("scenarios", []) if item.get("id") == args.scenario), None)
    if scenario is None:
        raise SystemExit(f"Scenario not found: {args.scenario}")
    base = load(args.base_machine)
    base_scenario = next((item for item in base.get("scenarios", []) if item.get("id") == args.scenario), None)
    if base_scenario is None:
        raise SystemExit(f"Base machine result does not contain scenario: {args.scenario}")

    run_directories = sorted(path for path in args.runs_root.glob("run-*") if path.is_dir())
    if not run_directories:
        raise SystemExit(f"No run-* directories found: {args.runs_root}")

    expected_per_run = int(scenario["expectedSceneCount"])
    runs: list[dict[str, Any]] = []
    for directory in run_directories:
        status_path, result_path = directory / "status.json", directory / "result.json"
        if not status_path.is_file() or not result_path.is_file():
            raise SystemExit(f"Run evidence is incomplete: {directory}")
        status, result = load(status_path), load(result_path)
        failures = result.get("failures", [])
        if not isinstance(failures, list):
            raise SystemExit(f"result.failures must be an array: {result_path}")
        runs.append({
            "id": directory.name,
            "scenarioId": args.scenario,
            "state": str(status.get("state", "")),
            "expectedDraftCount": expected_per_run,
            "actualDraftCount": count_generated(result),
            "failureCount": len(failures),
        })

    multiplier = len(runs)
    expected_total = expected_per_run * multiplier
    actual_total = sum(item["actualDraftCount"] for item in runs)
    rect_total = int(base_scenario.get("rectTotal", base.get("rectTotal", 0))) * multiplier
    rect_hits = int(base_scenario.get("top4RectHits", base.get("top4RectHits", 0))) * multiplier
    no_rect_total = int(base_scenario.get("noRectSceneTotal", base.get("noRectSceneTotal", 0))) * multiplier
    no_rect_false_positives = int(
        base_scenario.get("noRectFalsePositives", base.get("noRectFalsePositives", 0))) * multiplier
    representative_total = int(
        base_scenario.get("representativeTimeTotal", base.get("representativeTimeTotal", expected_per_run))) * multiplier
    representative_hits = int(
        base_scenario.get("representativeTimeHits", base.get("representativeTimeHits", 0))) * multiplier

    machine = {
        "fixtureVersion": str(manifest.get("fixtureVersion", "")),
        "evaluationScope": {
            "fixtureType": f"copilot-{scenario.get('split', 'development')}",
            "scenarioCount": 1,
            "runCount": multiplier,
            "split": scenario.get("split", "development"),
            "isHoldout": scenario.get("split") == "holdout",
            "source": "saved Copilot benchmark status/result artifacts",
        },
        "sceneCountPasses": sum(item["actualDraftCount"] == item["expectedDraftCount"] for item in runs),
        "sceneCountTotal": multiplier,
        "representativeTimeHits": representative_hits,
        "representativeTimeTotal": representative_total,
        "rectTotal": rect_total,
        "top4RectHits": rect_hits,
        "noRectSceneTotal": no_rect_total,
        "noRectFalsePositives": no_rect_false_positives,
        "scenarios": [{
            "id": args.scenario,
            "expectedSceneCount": expected_total,
            "actualSceneCount": actual_total,
            "rectTotal": rect_total,
            "top4RectHits": rect_hits,
        }],
        "runs": runs,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(machine, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(machine, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
