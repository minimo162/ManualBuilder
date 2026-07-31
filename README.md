# ManualBuilder

スクリーンショットと説明文から、部内向けの操作マニュアルを効率よく作るWindowsアプリです。

主出力はExcel、必要な場合の副出力はWordとします。印刷用の厳密な組版より、PC上で画像と説明を同時に読みやすく、出力後も部内で追記・差し替えしやすいことを優先します。

## 現在の状態

- Phase 0（技術・安全性検証）完了
- Excel COMによる主出力: 単発、10回連続、キャンセル、異常分岐を実機確認済み
- 既存の未保存Excelブックへ影響しないことを実機確認済み
- Word COMによる副出力の安全性と基本レイアウトを確認済み
- Phase 1（製品MVP）は未実装

Phase 0の確定スナップショットは `tests/phase0` に保存しています。Windowsで `tests\phase0\run.cmd` を実行すると検証メニューが開きます。

## 製品方針

- HTML + htmxで、撮影・入力・並べ替え・プレビューを行う
- Windows PowerShell 5.1で、localhostサーバー、画像監視、保存、Office出力を担う
- ユーザーが作った「シート」を、同じ順序のExcelワークシートへ出力する
- 1手順を「画像・タイトル・説明・補足」の固定カードとして扱う
- Excelでは画像を左、説明を右に配置し、PCの横長画面で見やすくする
- Wordは同じプロジェクトデータから生成できる副出力とする
- ユーザーが開いている未保存のExcel・Wordを変更または終了しない

詳しい要件は [docs/REQUIREMENTS.md](docs/REQUIREMENTS.md)、UI/UX方針は [docs/UIUX-DIRECTION.md](docs/UIUX-DIRECTION.md) を参照してください。

## リポジトリ構成

```text
ManualBuilder/
├─ .github/              GitHub用テンプレート
├─ docs/                 要件、UI/UX、検証結果、ロードマップ
├─ samples/              架空・匿名化したサンプルだけを置く場所
├─ src/                  Phase 1のPowerShellバックエンド
├─ tests/phase0/         Phase 0確定版 v0.4.16（改変しない基準点）
└─ web/                  Phase 1のHTML・htmx・CSS・JavaScript
```

## Phase 1の開始位置

実装順は [docs/PHASE1-ROADMAP.md](docs/PHASE1-ROADMAP.md) にまとめています。最初はプロジェクトJSON、自動保存、静的編集画面をつなぎ、その後に撮影監視とExcel出力を統合します。

## GitHubへ登録する前に

このリポジトリには、実際の業務画面、顧客情報、個人情報、認証情報、生成されたExcel・Word・PDFを登録しないでください。`.gitignore` で一般的な生成物は除外していますが、初回コミット前に必ず `git status` で対象ファイルを確認します。

登録方法は [PUSH-GUIDE.md](PUSH-GUIDE.md) を参照してください。部内用のため、GitHubリポジトリは **Private** で作成してください。

## ライセンス

部内利用の非公開プロジェクトを想定しているため、オープンソースライセンスは付与していません。
