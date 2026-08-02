# src

Phase 1のWindows PowerShell 5.1バックエンドです。

現在のファイル:

- `Start-ManualBuilderLauncher.ps1`: 共有版の確認、ローカルキャッシュ更新、起動中判定、フォールバック
- `ManualBuilder.Launcher.psm1`: バージョン判定、ファイル照合、原子的なキャッシュ切替、1世代保持
- `Start-ManualBuilder.ps1`: 起動、Mutex、localhost HTTPサーバー、セキュリティ検証、ルーティング
- `ManualBuilder.Storage.psm1`: ユーザー別データ保存先と旧プロジェクトの非破壊移行
- `ManualBuilder.Project.psm1`: プロジェクトモデル、入力制限、JSON読込、原子的保存、直前バックアップ
- `ManualBuilder.Capture.psm1`: 画像実体検証、保存先監視、重複除去、画像付き手順の追加
- `ManualBuilder.Excel.psm1`: 注釈合成、シート名変換、Excel COM生成、所有PID確認、自己検査
- `Export-ManualBuilderExcel.ps1`: localhostサーバーから分離したExcel出力ワーカー
- `ManualBuilder.Word.psm1`: Word COM生成、表紙・目次・手順文書、所有PID確認、原子的保存、自己検査
- `Export-ManualBuilderWord.ps1`: localhostサーバーから分離したWord副出力ワーカー
- `ManualBuilder.Web.psm1`: htmxへ返す編集画面のHTML断片

スクリーンショット監視、Excel主出力、Word副出力は製品側へ統合済みです。Office出力は専用プロセスで実行し、既存のExcel／Wordが開いている場合は安全側へ停止します。

PowerShellファイルはWindows PowerShell 5.1で日本語を安全に扱えるよう、UTF-8 BOM + CRLFで保存してください。
