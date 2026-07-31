# アーキテクチャ方針

## 境界

```text
Browser UI (web)
    │  localhost HTTP / htmx
    ▼
PowerShell application (src)
    ├─ project JSON and images (local only)
    ├─ screenshot watcher / clipboard import
    ├─ Excel exporter (primary)
    └─ Word exporter (secondary)
```

ブラウザーは編集UIだけを担当し、ファイル監視、ローカル保存、Office COMはPowerShell側で行う。外部サーバーやCDNを実行時依存にせず、業務データをPC外へ送信しない。

## データモデル

```text
Project
├─ metadata
└─ sheets[]
   ├─ id, name, order, summary
   └─ steps[]
      └─ id, order, title, description, note, imageId
```

アプリのプロジェクトデータを正本とする。ExcelとWordは同じデータから作る出力物であり、出力後の変更をアプリへ戻す往復編集はMVP対象外とする。

## 安全境界

- HTTP待受はlocalhostだけに限定する
- Host、Origin、Content-Type、要求サイズ、画像実体を検証する
- プロジェクトIDや画像IDから任意パスを組み立てさせない
- 保存は同一ボリューム上の一時ファイルを完成名へ置換する
- Office COMは所有PIDを確認してから設定変更し、所有不明のプロセスを終了しない
- UIのHTMLへ入力値を返すときは必ずエスケープする

## Phase 0との関係

`tests/phase0` はWindows固有機能とOffice COMの成立性を確認した固定スナップショットである。製品実装から直接呼び出す最終構造ではなく、Phase 1へ移植する安全条件と回帰試験の基準として扱う。
