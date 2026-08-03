# Phase 1 v0.32.6 再確認手順

v0.32.6は、ManualBuilderのブラウザータブだけを閉じた後に、起動中の画面へ戻れなくなる問題を修正した版です。

## 自動検査

Windowsで次を実行します。

```text
tests\phase1\run-tests.cmd
```

最後に `All Phase 1 foundation tests passed.` と表示されることを確認します。

## 目視確認

1. `run.cmd` からManualBuilderを起動する
2. ManualBuilderのブラウザータブだけを閉じる
3. もう一度 `run.cmd` を実行すると、起動中のManualBuilderが新しいブラウザータブで開く
4. 警告ダイアログだけが表示されて終了しないことを確認する
5. 共有HTMLの「編集する」から同じ状態で起動した場合は、既存画面が開き、画面からManualBuilderを終了して再実行するよう案内される
6. 画面左下の版数が `v0.32.6` になっている

## 修正内容

- `runtime.json` に記録した起動中URLを再実行時に既定のブラウザーへ渡す
- ローカルキャッシュ版から直接再実行した場合も同じ復帰動作にする
- 外部URLを開かないよう、復帰先を `http://localhost:<port>/` に限定する
- 配布版数をv0.32.6へ更新する
