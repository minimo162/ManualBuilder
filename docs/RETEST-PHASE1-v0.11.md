# Phase 1 v0.11 Word副出力 再検証

## 1. 適用確認

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"
Test-Path .\src\ManualBuilder.Word.psm1
Test-Path .\src\Export-ManualBuilderWord.ps1
```

どちらも`True`なら適用済み。

## 2. Officeを起動しない基盤テスト

```powershell
tests\phase1\run-tests.cmd
```

`All Phase 1 foundation tests passed.`を確認する。

## 3. Word COM出力テスト

Wordをすべて閉じてから実行する。

```powershell
tests\phase1\run-word-test.cmd
```

次を確認する。

- `Word COM export test passed.`
- docx内部に表紙、目次、2つのシート見出し、2枚の画像がある
- テスト終了後にWINWORDプロセスが残らない

## 4. 製品UI

1. `run.cmd`で起動する
2. シートを2件以上、各シートへ手順を追加する
3. 画像へ切り抜き、赤枠、番号を設定する
4. 右上の`…`から`Wordで作成`を選ぶ
5. 進捗が更新され、完成後に`Wordを開く`と`保存先を開く`が表示される
6. 完成docxでシート順・手順順・説明・補足・画像編集結果を確認する

## 5. Word起動中の安全停止

1. Wordで新規文書を開き、`未保存Wordを残すテスト`と入力する（保存しない）
2. ManualBuilderから`Wordで作成`を選ぶ
3. COMを生成せず、Wordを閉じて再実行またはExcel作成の案内になることを確認する
4. 未保存文書、入力文字、Wordウィンドウがそのまま残ることを確認する
5. ダイアログの`Excelで作成`からExcel出力へ移れることを確認する

## 6. 目視確認

- 表紙タイトルと作成日が読める
- 目次にシートと手順がページ番号付きで並ぶ
- ナビゲーションウィンドウにシートと手順が階層表示される
- 補足の黄色い枠が次の段落へ漏れない
- 画像が縦横比を保ち、切り抜きと注釈が反映される
- フッター中央にページ番号がある
- 見出しと本文がBIZ UDPゴシック（またはフォールバック字体）である

確認後、生成docxと画面キャプチャを共有する。レイアウトは実出力を見て調整する。
