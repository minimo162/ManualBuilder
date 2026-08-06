"""Generate synthetic WebM recordings used by ManualBuilder quality evaluations.

The screens are fictional and contain no customer or employee data. Install the
optional encoder with `python -m pip install imageio-ffmpeg` before running.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

try:
    import imageio_ffmpeg
except ImportError as exc:  # pragma: no cover - developer setup guidance
    raise SystemExit(
        "imageio-ffmpeg is required: python -m pip install imageio-ffmpeg"
    ) from exc


WIDTH = 960
HEIGHT = 540
FPS = 10

NAVY = "#17324d"
BLUE = "#2463a7"
BLUE_DARK = "#184f88"
LINE = "#d7dee7"
TEXT = "#172033"
MUTED = "#64748b"
GREEN = "#157f5b"
WHITE = "#ffffff"


def font_path(bold: bool = False) -> str:
    candidates = (
        [Path(r"C:\Windows\Fonts\BIZ-UDGothicB.ttc"), Path(r"C:\Windows\Fonts\YuGothB.ttc")]
        if bold
        else [Path(r"C:\Windows\Fonts\BIZ-UDGothicR.ttc"), Path(r"C:\Windows\Fonts\YuGothM.ttc")]
    )
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    raise RuntimeError("A Japanese Windows font was not found.")


_font_cache: dict[tuple[int, bool], ImageFont.FreeTypeFont] = {}


def get_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    key = (size, bold)
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(font_path(bold), size=size)
    return _font_cache[key]


def rounded(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str,
            radius: int = 0, outline: str | None = None, width: int = 1) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def label(draw: ImageDraw.ImageDraw, value: str, xy: tuple[int, int], size: int = 20,
          color: str = TEXT, bold: bool = False, anchor: str = "lm") -> None:
    draw.text(xy, value, font=get_font(size, bold), fill=color, anchor=anchor)


def button(draw: ImageDraw.ImageDraw, value: str, x: int, y: int, width: int,
           active: bool = False, color: str = BLUE) -> None:
    rounded(draw, (x, y, x + width, y + 48), BLUE_DARK if active else color, 8)
    label(draw, value, (x + width // 2, y + 24), 18, WHITE, True, "mm")


def field(draw: ImageDraw.ImageDraw, name: str, value: str, x: int, y: int,
          width: int, focused: bool = False) -> None:
    label(draw, name, (x, y - 18), 15, MUTED, True)
    rounded(draw, (x, y, x + width, y + 50), WHITE, 6, BLUE if focused else "#b9c5d3", 2)
    label(draw, value or "選択してください", (x + 16, y + 25), 18,
          TEXT if value else "#8794a5")


def shell(title: str, content_fill: str = "#f3f6f9") -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (WIDTH, HEIGHT), "#f3f6f9")
    draw = ImageDraw.Draw(image)
    draw.rectangle((170, 68, WIDTH, HEIGHT), fill=content_fill)
    draw.rectangle((0, 0, WIDTH, 68), fill=NAVY)
    label(draw, "架空業務ポータル", (30, 34), 22, WHITE, True)
    label(draw, "評価専用・実データではありません", (930, 34), 14, "#c7d6e5", anchor="rm")
    draw.rectangle((0, 68, 170, HEIGHT), fill="#e7edf4")
    label(draw, "ホーム", (28, 112), 17, TEXT, True)
    label(draw, "申請", (28, 158), 17, TEXT, True)
    label(draw, "設定", (28, 204), 17, TEXT, True)
    label(draw, title, (205, 105), 28, TEXT, True)
    return image, draw


def expense_frame(elapsed: int) -> Image.Image:
    if elapsed < 1600:
        image, draw = shell("申請一覧", "#f3f6f9")
        button(draw, "新規申請", 730, 100, 160, 1400 <= elapsed)
        rounded(draw, (205, 176, 890, 236), WHITE, 8, LINE)
        label(draw, "申請番号", (230, 206), 16, MUTED, True)
        label(draw, "種類", (450, 206), 16, MUTED, True)
        label(draw, "状態", (690, 206), 16, MUTED, True)
        rounded(draw, (205, 244, 890, 310), WHITE, 8, LINE)
        label(draw, "A-1042", (230, 277), 17)
        label(draw, "備品購入", (450, 277), 17)
        label(draw, "承認済み", (690, 277), 17, GREEN, True)
        return image
    if elapsed < 3100:
        image, draw = shell("申請の作成", "#d8e4f0")
        field(draw, "申請の種類", "", 220, 185, 520, elapsed >= 2900)
        field(draw, "件名", "", 220, 295, 520)
        button(draw, "保存", 730, 450, 160)
        return image
    if elapsed < 4600:
        image, draw = shell("申請の作成", "#b8cde3")
        field(draw, "申請の種類", "", 220, 185, 520, True)
        rounded(draw, (220, 242, 740, 392), WHITE, 6, "#b9c5d3")
        rounded(draw, (230, 252, 730, 306), "#dbeafe" if elapsed >= 4400 else WHITE, 4)
        label(draw, "出張費", (250, 279), 18, TEXT, True)
        rounded(draw, (230, 316, 730, 370), WHITE, 4)
        label(draw, "備品購入", (250, 343), 18)
        button(draw, "保存", 730, 450, 160)
        return image
    if elapsed < 6100:
        image, draw = shell("申請の作成", "#d2b58a")
        field(draw, "申請の種類", "出張費", 220, 185, 520)
        field(draw, "件名", "東京出張", 220, 295, 520)
        button(draw, "保存", 730, 450, 160, elapsed >= 5900)
        return image
    image, draw = shell("申請を保存しました", "#789b86")
    rounded(draw, (250, 170, 830, 390), WHITE, 12, LINE)
    rounded(draw, (285, 205, 337, 257), "#dff5ea", 26)
    label(draw, "✓", (311, 231), 28, GREEN, True, "mm")
    label(draw, "申請 A-1043 を保存しました", (365, 228), 22, TEXT, True)
    label(draw, "状態: 下書き", (365, 275), 18, MUTED)
    button(draw, "一覧へ戻る", 550, 320, 210, color=GREEN)
    return image


def settings_frame(elapsed: int) -> Image.Image:
    if elapsed < 1600 or elapsed >= 3200:
        image, draw = shell("ホーム")
        rounded(draw, (205, 150, 525, 300), WHITE, 10, LINE)
        label(draw, "未処理の申請", (235, 185), 17, MUTED, True)
        label(draw, "3件", (235, 242), 42, BLUE, True)
        rounded(draw, (550, 150, 890, 300), WHITE, 10, LINE)
        label(draw, "お知らせ", (580, 185), 17, MUTED, True)
        label(draw, "システム更新はありません", (580, 235), 18)
        button(draw, "通知設定", 700, 340, 190, 1400 <= elapsed < 1600)
        return image
    image, draw = shell("通知設定")
    button(draw, "← ホームへ戻る", 205, 92, 190, 3000 <= elapsed < 3200, GREEN)
    rounded(draw, (205, 170, 890, 400), WHITE, 10, LINE)
    label(draw, "メール通知", (240, 215), 20, TEXT, True)
    rounded(draw, (790, 195, 852, 229), GREEN, 17)
    rounded(draw, (824, 199, 850, 225), WHITE, 13)
    label(draw, "申請が承認されたときに通知します。", (240, 260), 17, MUTED)
    label(draw, "毎週のまとめ", (240, 330), 20, TEXT, True)
    rounded(draw, (790, 310, 852, 344), "#aab5c2", 17)
    rounded(draw, (794, 314, 820, 340), WHITE, 13)
    return image


def terms_page(offset: int, agree_active: bool = False) -> Image.Image:
    image, draw = shell("利用規約")
    top = 150 - offset
    rounded(draw, (210, top, 890, top + 760), WHITE, 10, LINE)
    label(draw, "サービス利用規約", (245, top + 42), 24, TEXT, True)
    for index in range(12):
        label(draw, f"{index + 1}. 架空サービスの利用条件を確認します。",
              (245, top + 94 + index * 48), 16)
        draw.rounded_rectangle((245, top + 114 + index * 48,
                                815 - (index % 3) * 55, top + 119 + index * 48),
                               radius=2, fill="#d9e1e9")
    label(draw, "内容を確認したうえで同意してください。", (245, top + 690), 17, MUTED, True)
    button(draw, "同意する", 690, top + 720, 170, agree_active, GREEN)
    draw.rectangle((170, 68, 960, 134), fill="#f3f6f9")
    draw.rectangle((170, 490, 960, 540), fill="#f3f6f9")
    label(draw, "利用規約", (205, 105), 28, TEXT, True)
    rounded(draw, (902, 145, 907, 475), "#e0e6ed", 2)
    thumb_y = 148 + min(270, round((offset / 430) * 270))
    rounded(draw, (901, thumb_y, 908, thumb_y + 60), "#91a0b1", 3)
    return image


def terms_frame(elapsed: int) -> Image.Image:
    if elapsed < 3000:
        offset = 0 if elapsed < 1200 else round(((elapsed - 1200) / 1800) * 430)
        return terms_page(offset)
    if elapsed < 4600:
        return terms_page(430, elapsed >= 4400)
    if elapsed < 6100:
        image = terms_page(430)
        draw = ImageDraw.Draw(image, "RGBA")
        draw.rectangle((0, 68, 960, 540), fill=(15, 23, 42, 107))
        rounded(draw, (280, 170, 760, 405), WHITE, 12)
        label(draw, "利用規約に同意しますか？", (320, 220), 23, TEXT, True)
        label(draw, "同意後にサービスを利用できます。", (320, 270), 17, MUTED)
        button(draw, "キャンセル", 330, 325, 160, color="#718096")
        button(draw, "同意して続行", 520, 325, 190, elapsed >= 5900, GREEN)
        return image.convert("RGB")
    image, draw = shell("利用開始の準備ができました")
    rounded(draw, (260, 175, 830, 385), WHITE, 12, LINE)
    label(draw, "✓", (315, 235), 42, GREEN, True, "mm")
    label(draw, "利用規約への同意が完了しました", (365, 225), 22, TEXT, True)
    label(draw, "ホームからサービスを利用できます。", (365, 275), 17, MUTED)
    button(draw, "ホームを開く", 550, 320, 210, color=GREEN)
    return image


def compact_button(draw: ImageDraw.ImageDraw, value: str, x: int, y: int, width: int,
                   active: bool = False, color: str = BLUE) -> None:
    rounded(draw, (x, y, x + width, y + 46), BLUE_DARK if active else color, 7)
    label(draw, value, (x + width // 2, y + 23), 15, WHITE, True, "mm")


def request_search_page(query: str, status: str, sort: str, rows: list[tuple[str, str, str, str]],
                        active: str = "", notice: str = "") -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image, draw = shell("申請検索")
    field(draw, "検索", query, 205, 150, 250, active == "search")
    compact_button(draw, status, 475, 150, 150, active == "status", "#52677d")
    compact_button(draw, sort, 645, 150, 130, active == "sort", "#52677d")
    compact_button(draw, "クリア", 790, 150, 100, active == "clear", "#718096")
    if notice:
        rounded(draw, (600, 86, 890, 128), "#dff5ea", 7, "#98d5bb")
        label(draw, notice, (620, 107), 15, GREEN, True)

    label(draw, "申請番号", (220, 220), 14, MUTED, True)
    label(draw, "件名", (350, 220), 14, MUTED, True)
    label(draw, "状態", (610, 220), 14, MUTED, True)
    label(draw, "更新日", (700, 220), 14, MUTED, True)
    for index, (request_id, subject, row_status, updated) in enumerate(rows):
        top = 238 + index * 72
        rounded(draw, (205, top, 890, top + 60), WHITE, 7, LINE)
        label(draw, request_id, (220, top + 30), 16, TEXT, True)
        label(draw, subject, (350, top + 30), 16)
        label(draw, row_status, (610, top + 30), 16, GREEN if row_status == "下書き" else MUTED, True)
        label(draw, updated, (700, top + 30), 15, MUTED)
        rounded(draw, (805, top + 11, 875, top + 49),
                BLUE_DARK if active == f"detail-{request_id}" else BLUE, 6)
        label(draw, "詳細", (840, top + 30), 14, WHITE, True, "mm")
    return image, draw


ALL_REQUESTS = [
    ("EV-1006", "研修資料の購入", "下書き", "8月3日"),
    ("EV-1007", "会議室利用", "完了", "8月4日"),
    ("EV-1008", "評価用申請A", "下書き", "8月5日"),
]
FILTERED_REQUESTS = [
    ("EV-1005", "評価用申請B", "下書き", "8月2日"),
    ("EV-1008", "評価用申請A", "下書き", "8月5日"),
]
SORTED_REQUESTS = list(reversed(FILTERED_REQUESTS))


def request_detail_page(expanded: bool = False, active: str = "") -> Image.Image:
    image, draw = shell("申請の詳細")
    compact_button(draw, "← 一覧へ戻る", 205, 92, 170, active == "back", GREEN)
    rounded(draw, (205, 165, 890, 430), WHITE, 10, LINE)
    label(draw, "EV-1008", (235, 205), 18, MUTED, True)
    label(draw, "評価用申請A", (235, 245), 25, TEXT, True)
    label(draw, "状態: 下書き", (235, 290), 17, GREEN, True)
    label(draw, "申請者: テスト利用者", (235, 330), 16, MUTED)
    compact_button(draw, "詳細を表示", 680, 240, 180, active == "expand")
    if expanded:
        rounded(draw, (235, 350, 860, 414), "#eef4fa", 7, "#c6d5e4")
        label(draw, "更新履歴", (255, 372), 15, MUTED, True)
        label(draw, "8月5日  下書きを作成", (255, 398), 16)
    else:
        label(draw, "更新履歴は［詳細を表示］から確認できます。", (235, 382), 16, MUTED)
    return image


def request_search_holdout_frame(elapsed: int) -> Image.Image:
    stage = min(9, elapsed // 1500)
    active = (elapsed % 1500) >= 1300
    if stage == 0:
        return request_search_page("", "状態: すべて", "番号順", ALL_REQUESTS,
                                   "search" if active else "")[0]
    if stage == 1:
        return request_search_page("評価用申請", "状態: すべて", "番号順", FILTERED_REQUESTS,
                                   "status" if active else "")[0]
    if stage == 2:
        image, draw = request_search_page("評価用申請", "状態: すべて", "番号順", FILTERED_REQUESTS)
        rounded(draw, (475, 205, 625, 305), WHITE, 7, "#9aabbd")
        rounded(draw, (485, 258, 615, 296), "#dbeafe" if active else WHITE, 5)
        label(draw, "すべて", (500, 230), 15)
        label(draw, "下書き", (500, 277), 15, TEXT, True)
        return image
    if stage == 3:
        return request_search_page("評価用申請", "状態: 下書き", "番号順", FILTERED_REQUESTS,
                                   "sort" if active else "")[0]
    if stage == 4:
        image, draw = request_search_page("評価用申請", "状態: 下書き", "番号順", FILTERED_REQUESTS)
        rounded(draw, (625, 205, 890, 305), WHITE, 7, "#9aabbd")
        label(draw, "申請番号順", (650, 230), 15)
        rounded(draw, (635, 258, 880, 296), "#dbeafe" if active else WHITE, 5)
        label(draw, "更新日の新しい順", (650, 277), 15, TEXT, True)
        return image
    if stage == 5:
        return request_search_page("評価用申請", "状態: 下書き", "更新日順", SORTED_REQUESTS,
                                   "detail-EV-1008" if active else "")[0]
    if stage == 6:
        return request_detail_page(False, "expand" if active else "")
    if stage == 7:
        return request_detail_page(True, "back" if active else "")
    if stage == 8:
        return request_search_page("評価用申請", "状態: 下書き", "更新日順", SORTED_REQUESTS,
                                   "clear" if active else "")[0]
    return request_search_page("", "状態: すべて", "番号順", ALL_REQUESTS,
                               notice="検索条件をクリアしました")[0]


REQUEST_SEARCH_HOLDOUT = {
    "id": "request-search-workflow",
    "split": "holdout",
    "detectorTuningAllowed": False,
    "videoFile": "../videos/request-search-workflow.webm",
    "durationMs": 15000,
    "summary": "申請を検索・絞り込み・並べ替え、詳細と更新履歴を確認して一覧へ戻り、検索条件をクリアする。",
    "coverage": ["input", "filter", "sort", "detail", "expand", "A-B-A", "result"],
    "expectedSceneCount": 10,
    "scenes": [
        {"order": 1, "state": "申請一覧", "representativeTimeRangeMs": [200, 600],
         "operationKind": "input", "operationTarget": "検索",
         "expectedRect": {"x1": 0.2135, "y1": 0.2778, "x2": 0.4740, "y2": 0.3704},
         "expectedTitle": "申請の検索", "expectedDescription": "［検索］に「評価用申請」と入力します。",
         "requiredConcepts": ["検索", "評価用申請", "入力"], "forbiddenClaims": ["申請を作成", "保存"]},
        {"order": 2, "state": "検索済み一覧", "representativeTimeRangeMs": [1700, 2100],
         "operationKind": "click", "operationTarget": "状態: すべて",
         "expectedRect": {"x1": 0.4948, "y1": 0.2778, "x2": 0.6510, "y2": 0.3630},
         "expectedTitle": "状態フィルターの表示", "expectedDescription": "［状態: すべて］を選択します。",
         "requiredConcepts": ["状態", "すべて", "選択"], "forbiddenClaims": ["下書きを保存", "検索を解除"]},
        {"order": 3, "state": "状態の選択肢", "representativeTimeRangeMs": [3200, 3600],
         "operationKind": "click", "operationTarget": "下書き",
         "expectedRect": {"x1": 0.5052, "y1": 0.4778, "x2": 0.6406, "y2": 0.5481},
         "expectedTitle": "下書きへの絞り込み", "expectedDescription": "状態から［下書き］を選択します。",
         "requiredConcepts": ["下書き", "選択"], "forbiddenClaims": ["完了", "申請を削除"]},
        {"order": 4, "state": "下書きの検索結果", "representativeTimeRangeMs": [4700, 5100],
         "operationKind": "click", "operationTarget": "番号順",
         "expectedRect": {"x1": 0.6719, "y1": 0.2778, "x2": 0.8073, "y2": 0.3630},
         "expectedTitle": "並べ替え項目の表示", "expectedDescription": "［番号順］を選択します。",
         "requiredConcepts": ["番号順", "選択"], "forbiddenClaims": ["申請番号を変更", "保存"]},
        {"order": 5, "state": "並べ替えの選択肢", "representativeTimeRangeMs": [6200, 6600],
         "operationKind": "click", "operationTarget": "更新日の新しい順",
         "expectedRect": {"x1": 0.6615, "y1": 0.4778, "x2": 0.9167, "y2": 0.5481},
         "expectedTitle": "更新日順への並べ替え", "expectedDescription": "［更新日の新しい順］を選択します。",
         "requiredConcepts": ["更新日", "新しい順", "選択"], "forbiddenClaims": ["申請番号順", "更新日を変更"]},
        {"order": 6, "state": "並べ替え済み一覧", "representativeTimeRangeMs": [7700, 8100],
         "operationKind": "click", "operationTarget": "EV-1008の詳細",
         "expectedRect": {"x1": 0.8385, "y1": 0.4611, "x2": 0.9115, "y2": 0.5315},
         "expectedTitle": "申請詳細の表示", "expectedDescription": "申請番号「EV-1008」の［詳細］を選択します。",
         "requiredConcepts": ["EV-1008", "詳細", "選択"], "forbiddenClaims": ["EV-1005", "申請を承認"]},
        {"order": 7, "state": "申請の詳細", "representativeTimeRangeMs": [9200, 9600],
         "operationKind": "click", "operationTarget": "詳細を表示",
         "expectedRect": {"x1": 0.7083, "y1": 0.4444, "x2": 0.8958, "y2": 0.5296},
         "expectedTitle": "更新履歴の表示", "expectedDescription": "［詳細を表示］を選択し、更新履歴を表示します。",
         "requiredConcepts": ["詳細を表示", "更新履歴", "選択"], "forbiddenClaims": ["履歴を削除", "申請を送信"]},
        {"order": 8, "state": "更新履歴を展開した詳細", "representativeTimeRangeMs": [10700, 11100],
         "operationKind": "click", "operationTarget": "一覧へ戻る",
         "expectedRect": {"x1": 0.2135, "y1": 0.1704, "x2": 0.3906, "y2": 0.2556},
         "expectedTitle": "申請一覧へ戻る", "expectedDescription": "更新履歴を確認し、［一覧へ戻る］を選択します。",
         "requiredConcepts": ["更新履歴", "一覧へ戻る", "選択"], "forbiddenClaims": ["申請を保存", "履歴を編集"]},
        {"order": 9, "state": "検索結果への再訪", "representativeTimeRangeMs": [12200, 12600],
         "operationKind": "click", "operationTarget": "クリア",
         "expectedRect": {"x1": 0.8229, "y1": 0.2778, "x2": 0.9271, "y2": 0.3630},
         "expectedTitle": "検索条件のクリア", "expectedDescription": "［クリア］を選択して検索条件を解除します。",
         "requiredConcepts": ["クリア", "検索条件", "解除"], "forbiddenClaims": ["申請を削除", "ログアウト"]},
        {"order": 10, "state": "検索条件の解除結果", "representativeTimeRangeMs": [13700, 14300],
         "operationKind": "result", "operationTarget": "", "expectedRect": None,
         "expectedTitle": "検索条件の解除確認", "expectedDescription": "検索条件がクリアされ、申請一覧が再表示されたことを確認します。",
         "requiredConcepts": ["検索条件", "クリア", "申請一覧", "確認"],
         "forbiddenClaims": ["申請を作成", "申請を保存"]},
    ],
}


SCENARIOS = (
    ("expense-application", 7400, expense_frame),
    ("settings-roundtrip", 4600, settings_frame),
    ("terms-slow-scroll", 7400, terms_frame),
    ("request-search-workflow", 15000, request_search_holdout_frame),
)


def write_video(path: Path, duration_ms: int, frame_factory) -> None:
    writer = imageio_ffmpeg.write_frames(
        str(path),
        (WIDTH, HEIGHT),
        fps=FPS,
        codec="libvpx-vp9",
        pix_fmt_in="rgb24",
        pix_fmt_out="yuv420p",
        bitrate="900k",
        macro_block_size=1,
        ffmpeg_log_level="warning",
        output_params=[
            "-deadline", "good", "-cpu-used", "4", "-row-mt", "1",
            "-fflags", "+bitexact", "-flags:v", "+bitexact", "-map_metadata", "-1",
        ],
    )
    writer.send(None)
    try:
        for frame_index in range(round(duration_ms * FPS / 1000)):
            elapsed = round(frame_index * 1000 / FPS)
            writer.send(frame_factory(elapsed).tobytes())
    finally:
        writer.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    default_output = Path(__file__).resolve().parents[2] / "samples" / "evals" / "videos"
    parser.add_argument("--output", type=Path, default=default_output)
    parser.add_argument(
        "--manifest",
        type=Path,
        default=Path(__file__).resolve().parents[2] / "samples" / "evals" / "gold" / "manifest.json",
    )
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    for scenario_id, duration_ms, frame_factory in SCENARIOS:
        target = args.output / f"{scenario_id}.webm"
        print(f"Generating {target.name}...")
        write_video(target, duration_ms, frame_factory)
        print(f"  {target.stat().st_size:,} bytes")

    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    manifest["fixtureVersion"] = "1.1.0"
    manifest["splits"] = {
        "development": {"detectorTuningAllowed": True},
        "holdout": {"detectorTuningAllowed": False},
    }
    scenarios = []
    for scenario in manifest["scenarios"]:
        if scenario["id"] == REQUEST_SEARCH_HOLDOUT["id"]:
            continue
        scenario["split"] = "development"
        scenario["detectorTuningAllowed"] = True
        scenarios.append(scenario)
    scenarios.append(REQUEST_SEARCH_HOLDOUT.copy())
    for scenario in scenarios:
        video_path = (args.manifest.parent / scenario["videoFile"]).resolve()
        payload = video_path.read_bytes()
        scenario["byteLength"] = len(payload)
        scenario["sha256"] = hashlib.sha256(payload).hexdigest()
    manifest["scenarios"] = scenarios
    args.manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Updated {args.manifest}")


if __name__ == "__main__":
    main()
