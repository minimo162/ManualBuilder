# Recorder Companion

ManualBuilderの操作記録中に表示する、WPF + Microsoft Edge WebView2製の記録モニターです。
WinForms版との併存や実行時フォールバックは行いません。Smart App Controlに拒否される
未署名EXEも生成せず、Microsoft署名済みのWindows PowerShell内でWPFホストを実行します。

## 構成

- `Invoke-RecorderCompanion.ps1`: WPFホストを同一プロセス内へ読み込む正式な起動口
- `ManualBuilder.RecorderCompanion.cs`: 常に手前、フォーカス復帰、記録指示、終了処理を担当するWPFホスト
- `web/`: WebView2内に表示するローカル限定UI
- `vendor/WebView2/`: Microsoft署名済みWebView2 SDK 1.0.4078.44の実行依存ファイル

WebView2は外部サイトへ遷移できない設定で使用し、表示内容と記録データはローカルファイルだけから読み込みます。

## 起動方式

アプリ本体が `Invoke-RecorderCompanion.ps1` をSTAのWindows PowerShellで起動します。
ホストのC#ソースは `Add-Type` により同一プロセス内へ読み込むため、未署名バイナリは
ディスクへ生成されません。

WebView2 SDKのライセンスは `vendor/WebView2/LICENSE.txt` にあります。
