"""Evaluate ManualBuilder's coarse video scene/target candidates against gold fixtures.

This mirrors web/assets/js/video-scenes.js closely enough to make candidate coverage
reproducible without opening the product UI. It deliberately scores local candidate
generation separately from Copilot's later candidate selection.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image


DEFAULTS = {
    "cols": 32,
    "rows": 18,
    "sampleIntervalMs": 300,
    "fineIntervalMs": 60,
    "staticThreshold": 0.010,
    "blockChangeThreshold": 0.055,
    "subtleBlockChangeThreshold": 0.030,
    "minCandidateBlocks": 2,
    "minStillMs": 700,
    "settleMs": 250,
    "sceneDistance": 0.030,
    "maxLocalizedRatio": 0.35,
    "maxScenes": 60,
}


def plan_times(duration_ms: int, interval_ms: int) -> list[int]:
    values = [0]
    values.extend(range(interval_ms, duration_ms, interval_ms))
    if duration_ms > 0:
        values.append(duration_ms)
    return values


def signature(frame: np.ndarray) -> np.ndarray:
    cols, rows = DEFAULTS["cols"], DEFAULTS["rows"]
    rgb = Image.fromarray(frame[..., :3]).resize((cols * 4, rows * 4), Image.Resampling.BILINEAR)
    pixels = np.asarray(rgb, dtype=np.float64) / 255.0
    # Keep the product JavaScript's BT.601 coefficients exactly. This evaluator is
    # intentionally a mirror, not a separately tuned detector.
    luminance = pixels[..., 0] * 0.299 + pixels[..., 1] * 0.587 + pixels[..., 2] * 0.114
    return luminance.reshape(rows, 4, cols, 4).mean(axis=(1, 3)).reshape(-1)


def sig_distance(first: np.ndarray, second: np.ndarray) -> float:
    return float(np.abs(first - second).mean())


def connected_clusters(indices: list[int], cols: int, rows: int) -> list[list[int]]:
    members = set(indices)
    seen: set[int] = set()
    clusters: list[list[int]] = []
    for start in indices:
        if start in seen:
            continue
        stack = [start]
        seen.add(start)
        cluster: list[int] = []
        while stack:
            current = stack.pop()
            cluster.append(current)
            col, row = current % cols, current // cols
            neighbours = []
            if col > 0:
                neighbours.append(current - 1)
            if col < cols - 1:
                neighbours.append(current + 1)
            if row > 0:
                neighbours.append(current - cols)
            if row < rows - 1:
                neighbours.append(current + cols)
            for neighbour in neighbours:
                if neighbour in members and neighbour not in seen:
                    seen.add(neighbour)
                    stack.append(neighbour)
        clusters.append(cluster)
    return sorted(clusters, key=len, reverse=True)


def cluster_rect(cluster: list[int], cols: int, rows: int) -> dict[str, float]:
    columns = [index % cols for index in cluster]
    row_values = [index // cols for index in cluster]
    return {
        "x1": max(0.0, (min(columns) - 0.5) / cols),
        "y1": max(0.0, (min(row_values) - 0.5) / rows),
        "x2": min(1.0, (max(columns) + 1.5) / cols),
        "y2": min(1.0, (max(row_values) + 1.5) / rows),
    }


def locate_candidates(before: np.ndarray, after: np.ndarray) -> list[dict]:
    cols, rows = DEFAULTS["cols"], DEFAULTS["rows"]
    def build(threshold: float, subtle: bool = False) -> list[dict]:
        changed = np.flatnonzero(np.abs(before - after) > threshold).tolist()
        # A page transition can be split into several disconnected bands. Looking
        # only at each cluster would turn those bands into false local targets.
        if len(changed) / (cols * rows) > DEFAULTS["maxLocalizedRatio"]:
            return []
        clusters = [cluster for cluster in connected_clusters(changed, cols, rows)
                    if len(cluster) >= DEFAULTS["minCandidateBlocks"]
                    and len(cluster) / (cols * rows) <= DEFAULTS["maxLocalizedRatio"]]
        if not clusters:
            return []
        rects = [cluster_rect(cluster, cols, rows) for cluster in clusters[:3 if subtle else 4]]
        if subtle and len(clusters) >= 2:
            combined = cluster_rect([item for cluster in clusters[:4] for item in cluster], cols, rows)
            area = (combined["x2"] - combined["x1"]) * (combined["y2"] - combined["y1"])
            if area <= DEFAULTS["maxLocalizedRatio"]:
                rects.insert(0, combined)
        return [{
            "id": f"video-diff-{index + 1}", "source": "video-diff",
            "confidence": "low" if subtle else ("medium" if index == 0 else "low"), "rect": rect,
        } for index, rect in enumerate(rects[:4])]

    return build(DEFAULTS["blockChangeThreshold"]) or build(DEFAULTS["subtleBlockChangeThreshold"], True)


def iou(first: dict | None, second: dict | None) -> float:
    if not first or not second:
        return 0.0
    left, top = max(first["x1"], second["x1"]), max(first["y1"], second["y1"])
    right, bottom = min(first["x2"], second["x2"]), min(first["y2"], second["y2"])
    intersection = max(0.0, right - left) * max(0.0, bottom - top)
    area_first = (first["x2"] - first["x1"]) * (first["y2"] - first["y1"])
    area_second = (second["x2"] - second["x1"]) * (second["y2"] - second["y1"])
    union = area_first + area_second - intersection
    return intersection / union if union > 0 else 0.0


def read_video(path: Path) -> tuple[list[np.ndarray], float]:
    reader = imageio.get_reader(path)
    metadata = reader.get_meta_data()
    fps = float(metadata.get("fps") or 10.0)
    frames = [np.asarray(frame) for frame in reader]
    reader.close()
    return frames, fps


def evaluate_scenario(root: Path, scenario: dict, frames_output: Path | None = None) -> dict:
    video_path = (root / "gold" / scenario["videoFile"]).resolve()
    frames, fps = read_video(video_path)
    duration_ms = int(scenario["durationMs"])

    def frame_at(time_ms: int) -> np.ndarray:
        index = min(len(frames) - 1, max(0, int(round(time_ms * fps / 1000.0))))
        return frames[index]

    samples = [{"timeMs": value, "signature": signature(frame_at(value))}
               for value in plan_times(duration_ms, DEFAULTS["sampleIntervalMs"])]
    runs = []
    start = 0
    for index in range(1, len(samples) + 1):
        moved = index < len(samples) and (
            sig_distance(samples[index - 1]["signature"], samples[index]["signature"]) > DEFAULTS["staticThreshold"]
            or sig_distance(samples[start]["signature"], samples[index]["signature"]) > DEFAULTS["staticThreshold"]
        )
        if moved or index == len(samples):
            runs.append({"startIndex": start, "endIndex": index - 1,
                         "startMs": samples[start]["timeMs"], "endMs": samples[index - 1]["timeMs"],
                         "durationMs": samples[index - 1]["timeMs"] - samples[start]["timeMs"]})
            start = index

    selected = []
    previous = None
    for run in runs:
        if run["durationMs"] < DEFAULTS["minStillMs"]:
            continue
        target = run["startMs"] + min(DEFAULTS["settleMs"], run["durationMs"] / 2)
        sample_index = run["endIndex"]
        for index in range(run["startIndex"], run["endIndex"] + 1):
            if samples[index]["timeMs"] >= target:
                sample_index = index
                break
        current = samples[sample_index]["signature"]
        if previous is not None and sig_distance(previous, current) < DEFAULTS["sceneDistance"]:
            continue
        previous = current
        selected.append({"timeMs": samples[sample_index]["timeMs"], "run": run, "signature": current})

    detected = []
    for index, scene in enumerate(selected):
        candidates = []
        if index + 1 < len(selected):
            transition_end_ms = selected[index + 1]["run"]["startMs"]
            operation_start_ms = max(scene["timeMs"], transition_end_ms - 1000)
            baseline = scene["signature"]
            for time_ms in range(operation_start_ms + DEFAULTS["fineIntervalMs"], transition_end_ms, DEFAULTS["fineIntervalMs"]):
                current = signature(frame_at(time_ms))
                found = locate_candidates(baseline, current)
                if found:
                    candidates = found
                    break
        detected.append({"timeMs": scene["timeMs"], "candidates": candidates})

    if frames_output:
        frames_output.mkdir(parents=True, exist_ok=True)
        for index, item in enumerate(detected, start=1):
            file_name = f"{scenario['id']}-{index:02d}.jpg"
            Image.fromarray(frame_at(item["timeMs"])[..., :3]).save(frames_output / file_name, quality=94)
            item["frameFile"] = file_name

    scene_results = []
    rect_total = rect_top1_hits = rect_top4_hits = time_hits = 0
    candidate_scene_total = no_rect_total = no_rect_false_positives = 0
    for index, expected in enumerate(scenario["scenes"]):
        actual = detected[index] if index < len(detected) else {"timeMs": -1, "candidates": []}
        time_range = expected["representativeTimeRangeMs"]
        time_ok = time_range[0] <= actual["timeMs"] <= time_range[1]
        time_hits += int(time_ok)
        expected_rect = expected.get("expectedRect")
        has_candidates = bool(actual["candidates"])
        candidate_scene_total += int(has_candidates)
        top1_iou = iou(actual["candidates"][0]["rect"], expected_rect) if actual["candidates"] and expected_rect else 0.0
        best_iou = max([iou(candidate["rect"], expected_rect) for candidate in actual["candidates"]] or [0.0])
        if expected_rect:
            rect_total += 1
            rect_top1_hits += int(top1_iou >= 0.35)
            rect_top4_hits += int(best_iou >= 0.35)
        else:
            no_rect_total += 1
            no_rect_false_positives += int(has_candidates)
        scene_results.append({
            "order": expected["order"], "expectedTimeRangeMs": time_range, "actualTimeMs": actual["timeMs"],
            "timeOk": time_ok, "candidateCount": len(actual["candidates"]),
            "top1IoU": round(top1_iou, 4), "bestIoU": round(best_iou, 4), "candidates": actual["candidates"],
            "frameFile": actual.get("frameFile", ""),
        })
    # Precision is undefined when the detector emitted no candidates. Reporting 1.0
    # there would overstate quality, so serialize it as null instead.
    candidate_precision = rect_top4_hits / candidate_scene_total if candidate_scene_total else None
    top1_precision = rect_top1_hits / candidate_scene_total if candidate_scene_total else None
    no_rect_false_positive_rate = no_rect_false_positives / no_rect_total if no_rect_total else None
    return {
        "id": scenario["id"], "expectedSceneCount": scenario["expectedSceneCount"],
        "actualSceneCount": len(detected), "sceneCountOk": len(detected) == scenario["expectedSceneCount"],
        "representativeTimeHits": time_hits, "representativeTimeTotal": len(scenario["scenes"]),
        "rectTotal": rect_total, "top1RectHits": rect_top1_hits, "top4RectHits": rect_top4_hits,
        "candidateSceneTotal": candidate_scene_total,
        "candidateSetPrecision": round(candidate_precision, 4) if candidate_precision is not None else None,
        "top1CandidatePrecision": round(top1_precision, 4) if top1_precision is not None else None,
        "noRectSceneTotal": no_rect_total,
        "noRectFalsePositives": no_rect_false_positives,
        "noRectFalsePositiveRate": round(no_rect_false_positive_rate, 4) if no_rect_false_positive_rate is not None else None,
        "scenes": scene_results,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=Path("samples/evals"))
    parser.add_argument("--output", type=Path)
    parser.add_argument("--frames-output", type=Path)
    args = parser.parse_args()
    manifest = json.loads((args.root / "gold" / "manifest.json").read_text(encoding="utf-8"))
    scenarios = [evaluate_scenario(args.root, scenario, args.frames_output) for scenario in manifest["scenarios"]]
    candidate_scene_total = sum(item["candidateSceneTotal"] for item in scenarios)
    top1_rect_hits = sum(item["top1RectHits"] for item in scenarios)
    top4_rect_hits = sum(item["top4RectHits"] for item in scenarios)
    no_rect_total = sum(item["noRectSceneTotal"] for item in scenarios)
    no_rect_false_positives = sum(item["noRectFalsePositives"] for item in scenarios)
    result = {
        "evaluationScope": {
            "fixtureType": "synthetic-in-sample",
            "scenarioCount": len(scenarios),
            "isHoldout": False,
            "overfittingRisk": (
                "These three synthetic fixtures are development fixtures, not a holdout set. "
                "A perfect score does not establish accuracy on real recordings, other codecs, "
                "frame rates, resolutions, applications, animations, or capture conditions."
            ),
        },
        "sceneCountPasses": sum(item["sceneCountOk"] for item in scenarios),
        "sceneCountTotal": len(scenarios),
        "representativeTimeHits": sum(item["representativeTimeHits"] for item in scenarios),
        "representativeTimeTotal": sum(item["representativeTimeTotal"] for item in scenarios),
        "rectTotal": sum(item["rectTotal"] for item in scenarios),
        "top1RectHits": top1_rect_hits,
        "top4RectHits": top4_rect_hits,
        "candidateSceneTotal": candidate_scene_total,
        "candidateSetPrecision": round(top4_rect_hits / candidate_scene_total, 4) if candidate_scene_total else None,
        "top1CandidatePrecision": round(top1_rect_hits / candidate_scene_total, 4) if candidate_scene_total else None,
        "noRectSceneTotal": no_rect_total,
        "noRectFalsePositives": no_rect_false_positives,
        "noRectFalsePositiveRate": round(no_rect_false_positives / no_rect_total, 4) if no_rect_total else None,
        "scenarios": scenarios,
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
