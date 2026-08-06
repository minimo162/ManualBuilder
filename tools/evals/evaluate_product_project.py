"""Score a project produced by ManualBuilder's real video-import path."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def load(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def candidates(step: dict[str, Any]) -> list[dict[str, Any]]:
    value = step.get("capture", {}).get("targetCandidates", [])
    if isinstance(value, str):
        value = json.loads(value) if value else []
    return [item for item in value if isinstance(item, dict) and isinstance(item.get("rect"), dict)]


def iou(first: dict[str, float], second: dict[str, float]) -> float:
    x1, y1 = max(first["x1"], second["x1"]), max(first["y1"], second["y1"])
    x2, y2 = min(first["x2"], second["x2"]), min(first["y2"], second["y2"])
    intersection = max(0.0, x2 - x1) * max(0.0, y2 - y1)
    first_area = (first["x2"] - first["x1"]) * (first["y2"] - first["y1"])
    second_area = (second["x2"] - second["x1"]) * (second["y2"] - second["y1"])
    union = first_area + second_area - intersection
    return intersection / union if union > 0 else 0.0


def covers_center(candidate: dict[str, float], expected: dict[str, float]) -> bool:
    center_x = (expected["x1"] + expected["x2"]) / 2
    center_y = (expected["y1"] + expected["y2"]) / 2
    return candidate["x1"] <= center_x <= candidate["x2"] and candidate["y1"] <= center_y <= candidate["y2"]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, default=Path("samples/evals/gold/manifest.json"))
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--scenario", required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--sample-tolerance-ms", type=int, default=300)
    args = parser.parse_args()

    manifest = load(args.manifest)
    scenario = next((item for item in manifest["scenarios"] if item["id"] == args.scenario), None)
    if scenario is None:
        raise SystemExit(f"Scenario not found: {args.scenario}")
    project = load(args.project)
    steps = [step for sheet in project["sheets"] for step in sheet.get("steps", [])]
    expected_scenes = scenario["scenes"]
    minimum_iou = float(manifest["evaluationDefaults"]["minimumRectIoU"])

    time_hits = strict_iou_hits = top1_usable_hits = usable_candidate_hits = 0
    rect_total = no_rect_total = no_rect_false_positives = 0
    details: list[dict[str, Any]] = []
    for index, expected in enumerate(expected_scenes):
        step = steps[index] if index < len(steps) else {}
        capture = step.get("capture", {})
        actual_time = int(capture.get("videoTimeMs", 0) or 0)
        low, high = expected["representativeTimeRangeMs"]
        time_ok = low - args.sample_tolerance_ms <= actual_time <= high + args.sample_tolerance_ms
        time_hits += int(time_ok)
        actual_candidates = candidates(step)
        expected_rect = expected.get("expectedRect")
        best_iou = 0.0
        center_covered = False
        candidate_ok = False
        top1_candidate_ok = False
        if expected_rect:
            rect_total += 1
            overlaps = [iou(item["rect"], expected_rect) for item in actual_candidates[:4]]
            best_iou = max(overlaps or [0.0])
            center_covered = any(covers_center(item["rect"], expected_rect) for item in actual_candidates[:4])
            top1_candidate_ok = bool(actual_candidates) and (
                overlaps[0] >= minimum_iou or covers_center(actual_candidates[0]["rect"], expected_rect)
            )
            strict_iou_hits += int(best_iou >= minimum_iou)
            candidate_ok = best_iou >= minimum_iou or center_covered
            top1_usable_hits += int(top1_candidate_ok)
            usable_candidate_hits += int(candidate_ok)
        else:
            no_rect_total += 1
            no_rect_false_positives += int(bool(actual_candidates))
        details.append({
            "order": expected["order"],
            "actualTimeMs": actual_time,
            "expectedTimeRangeMs": [low, high],
            "timeOkWithSampleTolerance": time_ok,
            "candidateCount": len(actual_candidates),
            "bestIoU": round(best_iou, 4),
            "centerCovered": center_covered,
            "top1CandidateUsable": top1_candidate_ok,
            "candidateUsable": candidate_ok,
        })

    result = {
        "evaluationScope": {
            "fixtureType": f"synthetic-{scenario.get('split', 'development')}",
            "scenarioCount": 1,
            "split": scenario.get("split", "development"),
            "isHoldout": scenario.get("split") == "holdout",
            "source": "ManualBuilder product video-import project",
            "scoringPolicy": {
                "timeToleranceMs": args.sample_tolerance_ms,
                "candidateUsable": "IoU meets manifest threshold or candidate contains the gold target center",
                "strictIoUReportedSeparately": True,
            },
        },
        "sceneCountPasses": int(len(steps) == len(expected_scenes)),
        "sceneCountTotal": 1,
        "representativeTimeHits": time_hits,
        "representativeTimeTotal": len(expected_scenes),
        "rectTotal": rect_total,
        "top1RectHits": top1_usable_hits,
        "top4RectHits": usable_candidate_hits,
        "strictIoURectHits": strict_iou_hits,
        "noRectSceneTotal": no_rect_total,
        "noRectFalsePositives": no_rect_false_positives,
        "scenarios": [{
            "id": scenario["id"],
            "expectedSceneCount": len(expected_scenes),
            "actualSceneCount": len(steps),
            "sceneCountOk": len(steps) == len(expected_scenes),
            "scenes": details,
        }],
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
    print(rendered, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
