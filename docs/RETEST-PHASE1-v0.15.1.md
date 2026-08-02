# Phase 1 v0.15.1 ローカルキャッシュ起動 再確認

## 目的

共有フォルダーのManualBuilderを配布元として使い、実際のサーバー、画面、Office出力処理が
各PCのローカルキャッシュから起動することを確認する。

## 事前条件

- ManualBuilderを終了してから更新する
- `%LOCALAPPDATA%\ManualBuilder\data` は削除しない
- 旧アプリ配下の `data` も移行確認が終わるまで残す

## 自動テスト

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"
tests\phase1\run-tests.cmd
```

`Local application cache tests passed.` と、最後の
`All Phase 1 foundation tests passed.` を確認する。

## 初回起動

共有配布版またはダウンロードフォルダーの `run.cmd` を起動する。コンソールへ次が表示される。

```text
[OK] ローカル実行版を更新しました: v0.15.1
[INFO] ローカル版から起動します: C:\Users\<ユーザー名>\AppData\Local\ManualBuilder\app\src\Start-ManualBuilder.ps1
```

次の項目を確認する。

```powershell
Test-Path "$env:LOCALAPPDATA\ManualBuilder\app\src\Start-ManualBuilder.ps1"
Test-Path "$env:LOCALAPPDATA\ManualBuilder\app\.install.json"
Test-Path "$env:LOCALAPPDATA\ManualBuilder\ManualBuilder.cmd"
```

すべて `True` になる。

## 2回目の起動

ManualBuilderを終了して共有側の `run.cmd` をもう一度起動し、次が表示されることを確認する。

```text
[INFO] ローカル実行版は最新です: v0.15.1
```

共有側の全ファイルを再コピーせず、ローカル版から起動する。

## ローカル起動

ManualBuilderを終了し、次を実行する。

```powershell
& "$env:LOCALAPPDATA\ManualBuilder\ManualBuilder.cmd"
```

共有側のランチャーを経由せず、同じプロジェクトが開くことを確認する。

## 起動中の再実行

ManualBuilderを起動したまま共有側の `run.cmd` を実行し、既存タブへ戻る案内だけが出ることを
確認する。このときローカル実行コードは更新されない。

## 判定

- 初回だけ検証済みローカルキャッシュを作る
- 同じ版では再コピーしない
- 正本データは従来どおり `%LOCALAPPDATA%\ManualBuilder\data` に残る
- ローカル起動ファイルからも起動できる
- 起動中のコードを差し替えない

この5点を満たせばv0.15.1は合格とする。
