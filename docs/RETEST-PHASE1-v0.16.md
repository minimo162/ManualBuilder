# Phase 1 v0.16.3 再確認手順

## 目的

1台のPCで複数のマニュアルを安全に管理できることを確認します。既存の `default` データは削除せず、最初のマニュアルとして一覧へ表示します。

## 自動テスト

ExcelとWordを閉じる必要はありません。PowerShellで次を実行します。

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"
tests\phase1\run-tests.cmd
```

次の2行を含め、最後に `All Phase 1 foundation tests passed.` と表示されることを確認します。

- `Project catalog tests passed.`
- `Project library server tests passed.`

## 画面確認

1. `run.cmd` を起動する。
2. 最初に「マニュアルを選ぶ」画面が表示される。
3. 従来のマニュアルがカードとして残っている。
4. 名前を入力して「＋ 新規作成」を押すと、空の編集画面が開く。
5. 左上付近の「一覧」を押すとマニュアル一覧へ戻る。
6. 検索欄へ名前の一部を入力し、該当カードだけが残る。
7. カードの `…` から「複製」を押し、名前に「- コピー」が付いた別カードが増える。
8. 複製したカードを開き、元と別の内容として編集できる。
9. 一覧へ戻り、カードの `…` から「アーカイブ」を押す。
10. カードが通常一覧から消え、下部の「アーカイブ」内へ移る。
11. 「復元」を押すと通常一覧へ戻る。
12. アプリを終了して再起動し、すべてのカードと内容が残っている。

## 保存先確認

```powershell
Get-ChildItem "$env:LOCALAPPDATA\ManualBuilder\data\projects" -Directory
Get-ChildItem "$env:LOCALAPPDATA\ManualBuilder\data\projects-archive" -Directory -ErrorAction SilentlyContinue
Get-Content "$env:LOCALAPPDATA\ManualBuilder\data\settings.json"
```

通常のマニュアルは `projects`、アーカイブしたマニュアルは `projects-archive` に分かれます。`settings.json` は前回開いたカードの目印だけを保持します。

## 安全性

- アーカイブは削除ではなくフォルダー移動です。
- 複製は新しいプロジェクトIDを発行し、元データを変更しません。
- 読み込めないプロジェクトは自動削除せず、一覧にエラーとして残します。
- Office出力中はプロジェクトの切替、複製、アーカイブ、復元を拒否します。
