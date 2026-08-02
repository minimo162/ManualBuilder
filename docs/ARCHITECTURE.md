# アーキテクチャ方針

## 境界

```text
Shared distribution (read only)
    └─ versioned app package / update launcher
            │ verified copy when version changes
            ▼
Per-user app cache (%LOCALAPPDATA%\ManualBuilder\app)
    │
Browser UI (web)
    │  localhost HTTP / htmx
    ▼
PowerShell application (src)
    ├─ per-user data (%LOCALAPPDATA%\ManualBuilder\data)
    │   ├─ project JSON and images
    │   ├─ runtime.json
    │   └─ Office export job snapshots
    ├─ screenshot watcher / clipboard import
    ├─ Excel export job / dedicated PowerShell worker (primary)
    └─ Word export job / dedicated PowerShell worker (secondary)
```

ブラウザーは編集UIだけを担当し、ファイル監視、ローカル保存、Office COMはPowerShell側で行う。外部サーバーやCDNを実行時依存にせず、業務データをPC外へ送信しない。

アプリコードの配布元は部内共有フォルダーに置くが、実行前に `%LOCALAPPDATA%\ManualBuilder\app`
へコピーする。通常起動では共有側の小さなバージョン情報だけを読み、版が同じならローカル版を
そのまま起動する。書込み可能な正本データは `%LOCALAPPDATA%\ManualBuilder\data` を使用し、Windowsユーザーごとに分離する。
旧版のアプリ配下データは、ローカル正本が存在しない初回に限り、検証後にコピーする。

## データモデル

```text
Project
├─ metadata
└─ sheets[]
   ├─ id, name, order, summary
   └─ steps[]
      └─ id, order, title, description, note, imageId, annotations[]
```

注釈は `rect`、`arrow`、`number`、`blackout` と、画像左上を原点とする0〜1の
比率座標で保持する。元画像は変更せず、ブラウザーではSVGとして重ね、Office出力時に
派生画像へ合成する。これにより表示サイズが変わっても位置を維持し、後から再編集できる。

アプリのプロジェクトデータを正本とする。ExcelとWordは同じデータから作る出力物であり、出力後の変更をアプリへ戻す往復編集はMVP対象外とする。

## Excel出力ジョブ

「Excelで作成」を押した時点の `project.json` と画像をジョブ専用フォルダーへ複製し、
`Export-ManualBuilderExcel.ps1` を別のWindows PowerShellプロセスで起動する。編集用localhost
サーバーはCOM処理を行わないため、作成中もブラウザーの編集操作と進捗ポーリングを継続できる。

ワーカーはステータスJSONを原子的に更新し、中止要求ファイルを手順ごとに確認する。完成xlsxは
一時名で保存して自己検査後に完成名へ移動し、既定では「ドキュメント\ManualBuilder」へ置く。
注釈はジョブ内の派生PNGにだけ合成し、プロジェクトの元画像は変更しない。

## Word出力ジョブ

Wordも開始時点のプロジェクトと画像をジョブ専用フォルダーへ複製し、
`Export-ManualBuilderWord.ps1`を別プロセスで起動する。Wordが起動中の場合はCOMを
生成する前に停止する。空のベースラインから作ったWINWORDだけをHwndとPIDで識別し、
編集済み画像を使って表紙、自動目次、シート見出し、手順、説明、補足、ページ番号を生成する。
完成docxは一時名で保存し、手順数・画像数・表紙・順序の自己検査後に完成名へ移動する。

## 安全境界

- HTTP待受はlocalhostだけに限定する
- Host、Origin、Content-Type、要求サイズ、画像実体を検証する
- プロジェクトIDや画像IDから任意パスを組み立てさせない
- 注釈種類、ID、件数、座標範囲、番号をPowerShell側でも再検証する
- 保存は同一ボリューム上の一時ファイルを完成名へ置換する
- Office COMは所有PIDを確認してから設定変更し、所有不明のプロセスを終了しない
- Excel作成は専用プロセスに分離し、HwndとPIDで所有を証明したExcelだけを終了対象にする
- 既存Excelへ接続した場合は設定変更・ブック作成・Quitを行わず安全停止する
- Word起動中はCOMを生成せず、既存Wordへ設定変更・文書作成・Quitを行わない
- 出力ファイルを開くAPIは、アプリが決定した出力フォルダー配下だけを許可する
- UIのHTMLへ入力値を返すときは必ずエスケープする
- 共有アプリ配下の旧データは移行時にも削除せず、既存のローカル正本を上書きしない
- 移行は一時フォルダーへコピーし、JSON読込と全ファイルのSHA-256照合後に確定する
- アプリ更新も一時フォルダーへコピーし、全ファイルのSHA-256照合後にフォルダー名を切り替える
- 更新前のローカル実行版を1世代保持し、共有版の欠損・更新失敗時に現在版を壊さない
- アプリ起動中はローカル実行コードを差し替えない

## Phase 0との関係

`tests/phase0` はWindows固有機能とOffice COMの成立性を確認した固定スナップショットである。製品実装から直接呼び出す最終構造ではなく、Phase 1へ移植する安全条件と回帰試験の基準として扱う。
