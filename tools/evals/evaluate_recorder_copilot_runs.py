"""Score real ManualBuilder RecorderCopilot capture bundles against gold steps."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import unicodedata
from pathlib import Path
from typing import Any


JPEG_PREFIXES = (b"\xff\xd8\xff",)
PACKET_FAILURE = re.compile(r"(?:\[WARN\].*\u30d1\u30b1\u30c3\u30c8|\u30d1\u30b1\u30c3\u30c8.*\u56de\u7b54\u3092\u8aad\u307f\u53d6れ)")


def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path}")
    return value


def json_lines(path: Path) -> list[dict[str, Any]]:
    result = []
    for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        value = json.loads(line)
        if not isinstance(value, dict):
            raise ValueError(f"JSON object required: {path}:{number}")
        result.append(value)
    return result


def normalized(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).lower()
    # WindowsのEdgeタイトルには表示されないU+200B等のformat文字が混ざる。
    # 画面表示と同じ文字列として評価し、"Microsoft Edge"の判定を落とさない。
    text = "".join(character for character in text if unicodedata.category(character) != "Cf")
    return " ".join(text.split())


def app_key(title: Any) -> str:
    value = normalized(title)
    if "microsoft edge" in value:
        return "edge"
    if value.endswith("- excel") or "microsoft excel" in value:
        return "excel"
    if value.endswith("- word") or "microsoft word" in value:
        return "word"
    return value


def has_anchor(event: dict[str, Any] | None) -> bool:
    return bool(event and (isinstance(event.get("rect"), dict) or isinstance(event.get("clickPoint"), dict)))


def image_ok(path: Path) -> bool:
    try:
        data = path.read_bytes()
    except OSError:
        return False
    return len(data) >= 1024 and data.startswith(JPEG_PREFIXES)


def score_run(run_dir: Path, scenario: dict[str, Any], defaults: dict[str, Any]) -> dict[str, Any]:
    required = ("status.json", "result.json", "frames.jsonl", "events.jsonl", "copilot.log")
    missing = [name for name in required if not (run_dir / name).is_file()]
    if missing:
        return {"id": f"{scenario['id']}/{run_dir.name}", "scenarioId": scenario["id"], "state": "invalid",
                "expectedDraftCount": len(scenario["steps"]), "actualDraftCount": 0,
                "failureCount": len(missing), "errors": ["missing artifacts: " + ", ".join(missing)],
                "scenes": [], "artifactDigest": ""}

    status, result = load(run_dir / "status.json"), load(run_dir / "result.json")
    frames, events = json_lines(run_dir / "frames.jsonl"), json_lines(run_dir / "events.jsonl")
    proposals = result.get("proposals", [])
    if not isinstance(proposals, list):
        proposals = []
    frame_map = {str(item.get("id", "")).upper(): item for item in frames}
    event_map = {int(item.get("index", 0)): item for item in events if int(item.get("index", 0)) > 0}
    transition_patterns = defaults.get("transitionTitlePatterns", []) + scenario.get("transitionTitlePatterns", [])
    transition_patterns = [normalized(value) for value in transition_patterns]
    max_event_distance = int(defaults.get("maximumEventDistanceMs", 2500))
    errors: list[str] = []

    if len(frames) < int(defaults.get("minimumFrameCount", 8)):
        errors.append("insufficient recorder frames")
    duration = max((int(item.get("timeMs", 0)) for item in frames), default=0)
    if duration < int(defaults.get("minimumCaptureDurationMs", 3000)):
        errors.append("capture duration is too short")
    if not events:
        errors.append("no recorder events")
    if status.get("jobId") and result.get("jobId") and status["jobId"] != result["jobId"]:
        errors.append("status/result jobId mismatch")

    referenced_images: set[str] = set()
    scene_results: list[dict[str, Any]] = []
    rect_total = rect_hits = time_hits = 0
    no_rect_total = no_rect_false_positives = 0
    for index, expected in enumerate(scenario["steps"]):
        proposal = proposals[index] if index < len(proposals) and isinstance(proposals[index], dict) else {}
        before_id = str(proposal.get("beforeFrame", "")).upper()
        after_id = str(proposal.get("afterFrame", "")).upper()
        before, after = frame_map.get(before_id), frame_map.get(after_id) if after_id else None
        text = f"{proposal.get('title', '')} {proposal.get('description', '')}"
        norm_text = normalized(text)
        concepts_ok = all(normalized(term) in norm_text for term in expected.get("requiredConcepts", []))
        exact_ok = all(str(token) in text for token in expected.get("exactTokens", []))
        forbidden_exact_ok = all(str(token) not in text for token in expected.get("forbiddenExactTokens", []))
        forbidden_ok = all(normalized(term) not in norm_text for term in scenario.get("forbiddenClaims", []))
        app_ok = before is not None and app_key(before.get("windowTitle")) == expected.get("app")
        stable_ok = before is not None and not any(pattern in normalized(before.get("windowTitle")) for pattern in transition_patterns)
        if after is not None:
            stable_ok = stable_ok and not any(
                pattern in normalized(after.get("windowTitle")) for pattern in transition_patterns)
            stable_ok = stable_ok and int(after.get("timeMs", 0)) > int(before.get("timeMs", 0))
        after_ok = (after is not None) if expected.get("afterRequired") else True

        event_ids = []
        for raw in proposal.get("eventIds", []):
            try:
                event_id = int(raw)
            except (TypeError, ValueError):
                continue
            if event_id in event_map and event_id not in event_ids:
                event_ids.append(event_id)
        target_id = int(proposal.get("targetEventId", 0) or 0)
        target = event_map.get(target_id)
        kinds = {normalized(event_map[event_id].get("kind")) for event_id in event_ids}
        expected_kinds = {normalized(value) for value in expected.get("eventKinds", [])}
        kind_ok = bool(kinds & expected_kinds) if expected_kinds else True
        anchor_ok = has_anchor(target)
        if expected.get("targetRequired"):
            rect_total += 1
            rect_hits += int(anchor_ok)
        else:
            no_rect_total += 1
            no_rect_false_positives += int(target_id > 0)

        event_times = [int(event_map[event_id].get("timeMs", 0)) for event_id in event_ids]
        event_time_ok = bool(before and event_times and min(
            abs(int(before.get("timeMs", 0)) - value) for value in event_times) <= max_event_distance)
        time_hits += int(event_time_ok)
        image_refs_ok = before is not None
        for frame in (before, after):
            if frame is None:
                continue
            name = str(frame.get("image", ""))
            referenced_images.add(name)
            if Path(name).name != name or not image_ok(run_dir / "frames" / name):
                image_refs_ok = False

        passed = all((concepts_ok, exact_ok, forbidden_exact_ok, forbidden_ok, app_ok,
                      stable_ok, after_ok, kind_ok, event_time_ok, image_refs_ok,
                      anchor_ok if expected.get("targetRequired") else True))
        scene_results.append({
            "order": expected["order"], "pass": passed, "conceptsOk": concepts_ok,
            "exactTokensOk": exact_ok and forbidden_exact_ok, "forbiddenClaimsOk": forbidden_ok,
            "appOk": app_ok, "stableFramesOk": stable_ok, "afterFrameOk": after_ok,
            "eventKindOk": kind_ok, "eventTimeOk": event_time_ok, "anchorOk": anchor_ok,
            "imageFilesOk": image_refs_ok, "beforeFrame": before_id, "afterFrame": after_id,
        })

    packet_failures = len(PACKET_FAILURE.findall((run_dir / "copilot.log").read_text(encoding="utf-8", errors="replace")))
    declared_failures = result.get("failures", [])
    if isinstance(declared_failures, list):
        packet_failures += len(declared_failures)
    digest = hashlib.sha256()
    for path in sorted([run_dir / name for name in required] + [run_dir / "frames" / name for name in referenced_images]):
        if path.is_file():
            digest.update(path.name.encode("utf-8"))
            digest.update(path.read_bytes())
    return {
        "id": f"{scenario['id']}/{run_dir.name}", "scenarioId": scenario["id"], "state": str(status.get("state", "")),
        "expectedDraftCount": len(scenario["steps"]), "actualDraftCount": len(proposals),
        "failureCount": packet_failures + len(errors), "errors": errors,
        "rectTotal": rect_total, "top4RectHits": rect_hits,
        "representativeTimeHits": time_hits, "representativeTimeTotal": len(scenario["steps"]),
        "noRectSceneTotal": no_rect_total, "noRectFalsePositives": no_rect_false_positives,
        "semanticPasses": sum(item["pass"] for item in scene_results),
        "scenes": scene_results, "artifactDigest": digest.hexdigest(),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path,
                        default=Path("samples/evals/recorder/gold/manifest.json"))
    parser.add_argument("--runs-root", type=Path, required=True,
                        help="<scenario-id>/run-*/... capture bundle root")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    manifest = load(args.manifest)
    defaults = manifest.get("evaluationDefaults", {})
    scenarios = manifest.get("scenarios", [])
    all_runs: list[dict[str, Any]] = []
    scenario_results: list[dict[str, Any]] = []
    for scenario in scenarios:
        run_dirs = sorted(path for path in (args.runs_root / scenario["id"]).glob("run-*") if path.is_dir())
        scored = [score_run(path, scenario, defaults) for path in run_dirs]
        if len(scored) < int(defaults.get("minimumRunsPerScenario", 1)):
            scored.append({
                "id": f"{scenario['id']}/missing-run", "scenarioId": scenario["id"], "state": "invalid",
                "expectedDraftCount": len(scenario["steps"]), "actualDraftCount": 0,
                "failureCount": 1, "errors": ["minimum real recorder run count not met"],
                "rectTotal": sum(bool(item.get("targetRequired")) for item in scenario["steps"]),
                "top4RectHits": 0, "representativeTimeHits": 0,
                "representativeTimeTotal": len(scenario["steps"]), "noRectSceneTotal": 0,
                "noRectFalsePositives": 0, "semanticPasses": 0, "scenes": [], "artifactDigest": "",
            })
        all_runs.extend(scored)
        scenario_results.append({
            "id": scenario["id"],
            "expectedSceneCount": sum(item["expectedDraftCount"] for item in scored),
            "actualSceneCount": sum(item["actualDraftCount"] for item in scored),
            "rectTotal": sum(item.get("rectTotal", 0) for item in scored),
            "top4RectHits": sum(item.get("top4RectHits", 0) for item in scored),
            "runCount": len(run_dirs), "runs": scored,
        })
    machine = {
        "fixtureVersion": manifest.get("fixtureVersion", ""),
        "evaluationScope": {
            "fixtureType": "real-recorder-copilot", "scenarioCount": len(scenarios),
            "runCount": len(all_runs), "isHoldout": True,
            "source": "ManualBuilder Recorder frames/events/images/status/log/result bundles",
        },
        "sceneCountPasses": sum(
            item["actualDraftCount"] == item["expectedDraftCount"] and item.get("semanticPasses", 0) == item["expectedDraftCount"]
            for item in all_runs),
        "sceneCountTotal": len(all_runs),
        "representativeTimeHits": sum(item.get("representativeTimeHits", 0) for item in all_runs),
        "representativeTimeTotal": sum(item.get("representativeTimeTotal", 0) for item in all_runs),
        "rectTotal": sum(item.get("rectTotal", 0) for item in all_runs),
        "top4RectHits": sum(item.get("top4RectHits", 0) for item in all_runs),
        "noRectSceneTotal": sum(item.get("noRectSceneTotal", 0) for item in all_runs),
        "noRectFalsePositives": sum(item.get("noRectFalsePositives", 0) for item in all_runs),
        "scenarios": scenario_results,
        "runs": [{key: value for key, value in item.items() if key in {
            "id", "scenarioId", "state", "expectedDraftCount", "actualDraftCount", "failureCount"
        }} for item in all_runs],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(machine, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(machine, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
