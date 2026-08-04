"""Generate synthetic WebM recordings used by ManualBuilder quality evaluations.

The screens are fictional and contain no customer or employee data. Install the
optional encoder with `python -m pip install imageio-ffmpeg` before running.
"""

from __future__ import annotations

import argparse
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


SCENARIOS = (
    ("expense-application", 7400, expense_frame),
    ("settings-roundtrip", 4600, settings_frame),
    ("terms-slow-scroll", 7400, terms_frame),
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
        output_params=["-deadline", "good", "-cpu-used", "4", "-row-mt", "1"],
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
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    for scenario_id, duration_ms, frame_factory in SCENARIOS:
        target = args.output / f"{scenario_id}.webm"
        print(f"Generating {target.name}...")
        write_video(target, duration_ms, frame_factory)
        print(f"  {target.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
