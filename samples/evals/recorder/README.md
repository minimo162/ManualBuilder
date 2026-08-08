# RecorderCopilot 実機評価

`gold/manifest.json` は、現在の主経路である「実際の画面操作をRecorderで記録し、
コンタクトシートをM365 Copilotへ渡して手順候補を得る」処理の正解データです。
従来の合成WebM評価とは別のゲートであり、回答JSONだけでは合格しません。

## 実行bundle

評価ルートを次の配置にします。

```text
<runs-root>/
  edge-excel-order-transfer/run-01/
  excel-multi-operation/run-01/
  edge-delayed-transition/run-01/
```

各 `run-*` には、同じRecorder/Copilot jobから保存した次のファイルが必要です。

```text
status.json
result.json
frames.jsonl
events.jsonl
copilot.log
frames/frame-00001.jpg ...
```

- `result.json` は `proposals` を持つRecorderCopilot workerの結果です。
- `status.json` と `result.json` の `jobId` は一致させます。
- `frames.jsonl` の `image` は `frames/` 直下のJPEGだけを参照します。
- 原本を後から差し替えた評価と区別できるよう、評価結果にはbundle全体のSHA-256 digestが残ります。
- 最低1回を必須、各シナリオ3回を推奨します。合格判定は全run一致です。

## 実機シナリオの起動

各シナリオは同じ安全ハーネスから選択します。ハーネスが作成したExcel COM
インスタンスとEdgeウィンドウだけを終了し、起動前から存在したExcel/Edgeには触れません。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA `
  -File tests/manual/Invoke-RecorderCopilotScenario.ps1 `
  -Scenario edge-excel-order-transfer `
  -WorkbookPath <temporary-order-workbook.xlsx> `
  -ReadyPath <ready.json> -StartPath <go.signal>
```

`-Scenario` は次の3つです。

- `edge-excel-order-transfer`（`-WorkbookPath` 必須）
- `excel-multi-operation`（`-WorkbookPath` 必須）
- `edge-delayed-transition`

`ReadyPath` が作成された後にManualBuilderで録画を開始し、`StartPath` の空ファイルを
作成すると操作が始まります。操作終了後にManualBuilderで録画を終了し、Copilotの
候補が表示されるまで待ちます。

候補表示後、取り込みまたは破棄を行う前にbundleを非破壊コピーします。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA `
  -File tests/manual/Save-RecorderCopilotBenchmarkBundle.ps1 `
  -Scenario edge-excel-order-transfer -RunId run-01 `
  -DestinationRoot <runs-root> -RecordingJobDirectory <record-job-directory>
```

保存先がすでに存在する場合は上書きしません。録画jobが複数残っている場合は、誤った
証跡を選ばないよう `-RecordingJobDirectory` の明示が必須です。コピー後も録画親jobと
AI子jobは削除されません。

## 採点

```powershell
python tools/evals/evaluate_recorder_copilot_runs.py `
  --runs-root <runs-root> --output <machine.json>
```

採点器は次を検査します。

- 期待手順数、順序、必須概念、画面に現れた値と英字の大文字小文字
- before/afterフレームのアプリ、時系列、JPEG原本の存在
- 操作イベントの種類と時刻、`targetEventId` の矩形またはクリック座標
- 読み込み中タイトルのフレーム、架空のアプリ切替、記録開始・停止の混入
- worker未完了、packet failure、status/resultのjob不一致

機械評価が不合格なら、最終結果も不合格です。
