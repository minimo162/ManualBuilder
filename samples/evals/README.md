# 録画・文章品質の評価用データ

完全に架空の業務画面を使った、ManualBuilder専用の正解付き録画です。
実在する会社、利用者、申請、URL、金額、アカウントは含みません。

## 収録内容

- `expense-application.webm`: 一覧から申請を作成して保存する基本操作
- `settings-roundtrip.webm`: ホームから設定へ移動し、元の画面へ戻る操作
- `terms-slow-scroll.webm`: ゆっくりしたスクロールと確認ダイアログ
- `request-search-workflow.webm`: 検索、絞り込み、並べ替え、詳細確認、一覧復帰、条件クリアを行う10場面のholdout

正解は [`gold/manifest.json`](gold/manifest.json) にあります。場面数、時刻範囲、操作対象、
正規化矩形、期待する手順文、必須概念、書いてはいけない内容を記録しています。
動画の意図しない差し替えを検出できるよう、ファイル長とSHA-256も固定しています。

独立したサブエージェントで採点するときは、[`EVALUATION-PROMPT.md`](EVALUATION-PROMPT.md)を使います。

## 評価時の扱い

- 文章は完全一致で採点しません。`requiredConcepts`を満たし、`forbiddenClaims`を含まず、
  映像上の事実と矛盾しないことを評価します。
- `expectedRect`は操作対象そのものの矩形です。場面分割器が付ける余白を考慮し、IoU 0.35以上を合格目安とします。
- 場面時刻は動画圧縮と300ms走査の差を考慮し、`representativeTimeRangeMs`内なら正解です。
- 最終結果画面も、確認手順として残すことを正解にしています。

## 場面・操作候補の再評価

製品と同じ閾値で、場面時刻と操作候補を再評価できます。

```powershell
python tools/evals/evaluate_video_candidates.py --split development --output out/video-candidates.json
```

`top1RectHits` / `top4RectHits` は正解枠を候補へ含められた割合、
`candidateSetPrecision` は候補を出した場面のうち正解枠を含んだ割合、
`noRectFalsePositiveRate` は本来枠を出さない場面で候補を出した割合です。
既定では調整可能なdevelopment 3本だけを評価します。未調整の確認は
`--split holdout`を明示し、結果を見た後で検出器を調整した場合は新しいholdoutへ交換します。
満点でも実録画や別アプリでの精度を保証しません。

製品のJavaScriptをブラウザーで直接確認するときは、リポジトリをローカルHTTPで配信し、
`tools/evals/video-pipeline-harness.html`を開いて録画を選択します。

製品UIで取り込んだプロジェクトは、300ms走査の許容と粗い候補の中心包含を含む
候補再現率を次で採点できます。厳密なIoU合格数も別に残ります。

```powershell
python tools/evals/evaluate_product_project.py --project <project.json> `
  --scenario request-search-workflow --output out/machine.json
```


## 再生成

Python、Pillow、`imageio-ffmpeg`を用意し、リポジトリ直下で次を実行します。

```powershell
python tools/evals/generate_video_fixtures.py
```

生成条件は 960×540px、10fps、WebM/VP9です。
