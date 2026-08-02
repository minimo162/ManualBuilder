# Phase 1 v0.12 Wordレイアウト 再検証

## 1. 適用確認

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"
Select-String .\src\ManualBuilder.Word.psm1 -Pattern "Get-MbWordContentSheets"
Select-String .\src\ManualBuilder.Word.psm1 -Pattern "450.0 /"
```

両方が見つかればv0.12の主要修正は適用済み。

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
- 空シート除外、画像450pt配置、フッター中央揃えのテストが`OK`
- テスト終了後にWINWORDプロセスが残らない

## 4. 製品UIと実出力

1. 手順があるシートと、手順が0件の空シートを用意する
2. 横長画像、縦長画像、小さく切り抜いた画像を含める
3. 右上の`…`から`Wordで作成`を選ぶ
4. 完成docxを開く
5. 空シートが目次と本文に出ないことを確認する
6. 画像が中央に配置され、以前より大きく読めることを確認する
7. フッターのページ番号が中央にあることを確認する

## 5. Word起動中の安全停止

1. Wordで未保存文書を開いたままManualBuilderから`Wordで作成`を選ぶ
2. 見出しが「Wordが開いているため、作成を開始しませんでした」になることを確認する
3. 影響がない旨と「Wordを閉じて再実行するか、Excelで作成」の対処が重複せず表示されることを確認する
4. 未保存文書とManualBuilderの入力内容がそのまま残ることを確認する

