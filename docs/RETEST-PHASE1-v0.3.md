# Phase 1 v0.3 画像取込み再テスト

## 自動テスト

ManualBuilderを終了してから、リポジトリ直下で実行する。

```text
tests\phase1\run-tests.cmd
```

次の4つが成功することを確認する。

- `Static tests passed.`
- `Project store tests passed.`
- `Capture store tests passed.`
- `Server integration tests passed.`

## 画像取込み

1. `run.cmd` で起動する。
2. ヘッダーが `監視中・このタブに追加` になることを確認する。
3. `Win + Shift + S` で1枚撮影し、2秒程度で選択中シート末尾へ追加されることを確認する。
4. `Ctrl + V` で画像を貼り付け、追加後に説明欄へフォーカスすることを確認する。
5. PNGまたはJPEGを複数選択してドロップし、ファイル順に追加されることを確認する。
6. `画像を選択` からPNG、JPEG、BMPを追加できることを確認する。
7. 同じ画像をもう一度貼り付け、手順が増えず重複通知が出ることを確認する。
8. タイトル・説明・補足を入力し、再起動後も画像と文章が復元されることを確認する。

画像実体は次へ保存される。

```text
data\projects\default\images
```

## 監視停止とバックフィル防止

1. ManualBuilderのタブを閉じるが、PowerShell画面は閉じない。
2. 30秒以上待ってからスクリーンショットを2枚撮る。
3. `http://localhost:8765/` を開き直す。
4. タブを閉じていた間の2枚が後から追加されないことを確認する。
5. 開き直した後に新しく撮影した画像だけが追加されることを確認する。

## 表示確認

- 1920×1080・100%: 標準レイアウトとして画像と説明を同時に確認できる
- 2560×1440・100%: 編集領域が横へ伸びすぎず、最大1800px程度で中央に収まる
- 実効1280×720: 主要操作、保存状態、監視状態が隠れない

## GitHubへ反映する場合

```powershell
git status
git diff --stat
git add .
git commit -m "feat: integrate screenshot capture workflow"
git push
```
