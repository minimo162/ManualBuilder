# Phase 1 v0.15.0 ユーザー別保存先 再確認

## 目的

共有フォルダーに置いたManualBuilderを複数人が起動しても、プロジェクト、画像、実行時情報、
Office出力の一時ファイルが各ユーザーのPC内で分離されることを確認する。

## 事前確認

v0.14.1以前の作業データがアプリ配下の `data\projects\default` にある場合は、削除せずそのまま残す。
新しいZIPを上書き展開するときも `data` を消さない。

## 自動テスト

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"
tests\phase1\run-tests.cmd
```

次が表示され、最後に `All Phase 1 foundation tests passed.` となることを確認する。

- 既定保存先をLocalApplicationData配下にする
- 既存のdefaultプロジェクトを初回だけ移行する
- 移行後も元のproject.jsonと元画像を削除しない
- 2回目の起動では移行を繰り返さない
- 既存のローカルデータを旧データで上書きしない

## 実機確認

1. `run.cmd` を起動する。
2. PowerShell画面の `ユーザーデータ` が次の場所になっていることを確認する。

```text
%LOCALAPPDATA%\ManualBuilder\data
```

3. ブラウザー左下に `v0.15.0` と表示されることを確認する。
4. 旧データがあった場合、以前のシート、手順、画像が表示されることを確認する。
5. 次の両方を確認する。

```powershell
Test-Path "$env:LOCALAPPDATA\ManualBuilder\data\projects\default\project.json"
Test-Path ".\data\projects\default\project.json"
```

ローカル側は `True` になる。旧データがあった場合はアプリ側も `True` のままで、移行後も削除されない。

6. タイトルまたは説明を編集し、終了・再起動後も変更が残ることを確認する。
7. アプリ配下の旧 `project.json` が編集内容で上書きされていないことを確認する。

## 判定

- 通常の読書きは各ユーザーの `%LOCALAPPDATA%` だけで行われる
- 旧データは初回だけコピーされ、元データを残す
- 既存ローカルデータがある場合は移行を再実行しない
- 完成Excel／Wordはユーザーの `ドキュメント\ManualBuilder` へ出力される

この4点を満たせば、v0.15.0のユーザー別保存先分離は合格とする。
