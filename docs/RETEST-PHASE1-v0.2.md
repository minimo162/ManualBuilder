# Phase 1 v0.2 再テスト

## 適用前

1. 起動中のManualBuilderを画面の `…` → `ManualBuilderを終了` から閉じる。
2. リポジトリで `git status` を実行し、未コミット変更がないことを確認する。
3. ZIP内の `ManualBuilder` の内容を既存リポジトリへ上書きする。

ZIPに `data` は含まれないため、既存のプロジェクトJSONは上書きしない。

## 自動テスト

リポジトリ直下で次を実行する。

```text
tests\phase1\run-tests.cmd
```

期待結果:

- `Static tests passed.`
- `Project store tests passed.`
- `Server integration tests passed.`
- `All Phase 1 foundation tests passed.`
- モジュールの「承認されていない動詞」警告が出ない

## 目視テスト

1. `run.cmd` を実行する。
2. 空状態に「手順を追加」が1つだけ表示されることを確認する。
3. 手順を追加し、タイトル・説明・補足が自動保存されることを確認する。
4. 手順、シート、上部バーの `…` メニューがマウスとキーボードで開くことを確認する。
5. `Esc` で開いているメニューを閉じられることを確認する。
6. ManualBuilderを起動したまま、もう一度 `run.cmd` を実行する。
7. 「既存のブラウザータブへ戻ってください」と表示され、新しいタブが増えないことを確認する。
8. 1280px幅およびWindows表示倍率125%で、保存状態・監視状態・Excelボタンが確認できることを確認する。

Excel／Wordボタンは出力コードの製品統合前なので、今回も無効が正常である。

## GitHubへ反映する場合

実機確認後に差分を確認し、問題がなければコミットする。

```powershell
git status
git diff --stat
git add .
git commit -m "feat: refine Phase 1 editor UI"
git push
```
