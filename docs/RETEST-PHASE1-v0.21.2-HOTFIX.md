# ManualBuilder Phase 1 v0.21.2 テスト修正版

`Test-ExcelUtilities.ps1` が非公開の内部関数を通常コマンドとして呼び出していたため、
長文レイアウトの試験開始時に停止していました。

製品のExcel出力処理には影響しません。内部関数を公開せず、既存の進捗JSON試験と同じく
モジュールスコープ内で呼び出すように試験だけを修正しました。

```powershell
cd "$env:USERPROFILE\Downloads\ManualBuilder"
tests\phase1\run-tests.cmd
```

すべて通過後、Excelを閉じて次を実行します。

```powershell
tests\phase1\run-excel-test.cmd
```
