# 修正指示書 — ManualBuilder v0.21.8

作成日: 2026-08-02
対象: `src/`（PowerShell 5.1 バックエンド）、`web/`（htmx + バニラJS フロント）
レビュー範囲: HTTPサーバー／Web、Excel/Word COM出力、データ永続化・画像取込み、Webフロント、**アプリ画面／Excel・Word出力のビジュアルデザイン**

## 第2回レビュー（実物を見ての検証・2026-08-02）

架空のサンプルマニュアル（備品管理システム／2シート5手順・注釈4種・画像なし手順を含む）を作り、
**実際のアプリ画面とExcel／Word出力を生成して確認**した。静的な読み合わせでは出なかった問題が出た。

### 🔴 実バグ: 注釈が1件のとき編集画面から消え、上書きで失われる

**対象**: `src/ManualBuilder.Web.psm1:99`

```powershell
# 修正前
$annotationsJson = if ($annotations.Count -eq 0) { '[]' } else { $annotations | ConvertTo-Json -Compress -Depth 5 }
```

PowerShellはパイプで単一要素の配列を「要素そのもの」へ展開するため、注釈がちょうど**1件**のときだけ
`[{...}]` ではなく `{...}` が出力される。画面側 `readCardAnnotations` は
`Array.isArray(parsed) ? parsed : []` で判定するため、**注釈が0件として扱われる**。

実害（サンプルで再現・修正後に解消を確認）:

| 影響箇所 | 症状 |
|---|---|
| `app.js:301` 注釈オーバーレイ | 編集画面に注釈が表示されない（バッジは「注釈 1」と出るので画面内で矛盾） |
| `app.js:924` 画像編集の初期値 | **空の状態から開始するため、保存すると既存の注釈が失われる**（データ損失） |
| `app.js:1514/1538` 拡大表示・編集判定 | 注釈が無いものとして扱われる |
| Excel／Word出力 | PowerShell側は `@()` で包むため**注釈が出る** → 画面と出力が食い違う |

`project.json` 自体は正しく配列で保存されているため、データは壊れていない。HTML生成だけの問題。

**修正**: `ConvertTo-Json -InputObject @($annotations)` でパイプを介さず配列を維持し、
念のため先頭が `[` でない場合は包み直す。回帰テスト `tests/phase1/Test-WebRender.ps1` を追加し、
注釈0/1/2/3件すべてで配列になることと件数が欠けないことを検証（`run-tests.cmd` へ登録済み）。

### 🔴 実バグ: Word出力後にWINWORDが残留する（#5の再評価）

初回レビューでは「Hwnd限定の強制終了は安全のため緩和しない」と判断したが、実機で計測したところ
**Wordの`Application.Hwnd`は常に0**で、所有モードは必ず「空ベースライン＋PID差分」に落ちていた。
その結果、`Quit(0)`で終了しきれなかったプロセスが**毎回残留**していた（実測で再現）。

| 検証項目 | 結果 |
|---|---|
| `Application.Hwnd`（文書追加の前後とも） | **0**（Wordは公開していない） |
| `ActiveWindow.Hwnd`（文書追加後） | 取得でき、**所有PIDと一致** |
| `Quit(0)`＋参照解放＋GC後のWINWORD | **1個残留** |

**修正**: 強制終了の条件を緩めるのではなく、**所有証明の手段を増やした**。
`Get-MbWordProcessId` が `ActiveWindow.Hwnd` を優先して見るようにし、文書を追加した直後に
「Hwnd由来のPIDがPID差分で特定した所有PIDと一致するか」を確認して、一致したときだけ
所有モードを `Hwnd` へ引き上げる。これにより既存の（設計docどおりの）安全網が正しく働く。
一致しなければ従来どおり強制終了しない。

検証: 修正後は `ownershipMode: Hwnd` となり、**出力後のWINWORD残留は0個**になった。

### デザイン改善（実物を見て判断）

| # | 対象 | 変更前の見え方 | 変更 |
|---|---|---|---|
| E1 | Excel 画像なし手順 | 左半分が「画像なし」だけの巨大な空きグレー領域 | **文章をカード全幅**へ。編集画面の「画像なし手順の空白を縮小する」と方針を揃えた |
| E2 | Excel 文章側の余白 | 説明が短いと白い帯がグレーの中に浮き、未完成に見えた | 文章側を**下端まで白の一枚面**に統一 |
| E3 | Excel 説明・補足 | 文字が区切り線に張り付いていた | `IndentLevel = 1` で一段字下げ |
| A1 | アプリ 一覧画面 | 日本語UIの中で見出しラベルだけ英語（`YOUR MANUALS`） | **「作成したマニュアル」**へ。字間も欧文用の0.12emから0.04emへ |

Word出力はPDFの画素描画ができなかったため、COMで**構造と書式値を実測**して確認した。
余白56.7pt(=20mm)、表題28pt／見出1 16pt／見出2 13pt／本文10.5pt、画像450pt幅（1920×1080→450×253pt）、
フッター中央揃え、目次にシートと手順、いずれも設計どおり。見出し色 `rgb(58,91,160)` と
本文色 `rgb(24,32,51)` がExcelと一致していることも実ファイルで確認できた（D4の統一が効いている）。

### 検証方法（再現手順）

作業用スクリプトは本リポジトリ外（セッションのscratchpad）に置いた。要点のみ記す。

- 架空サンプル生成 → アプリ起動 → `msedge --headless --screenshot` で画面を撮影
- Excelは出力後に `Range.CopyPicture` + `Chart.Export` でシートを**画面表示のままPNG化**（PDFだと印刷分割が入る）
- Wordは出力後にCOMで段落スタイル・フォント・色・画像寸法・フッターを列挙

> 注意: PowerShell 5.1はBOM無しUTF-8をANSIとして読むため、日本語を含む`.ps1`は**BOM付きで保存**すること。

---

## 実装状況（2026-08-02 反映済み）

`tests\phase1` の全10テストがパスすることを確認済み。

| 項目 | 状態 | 備考 |
|---|---|---|
| #1 サーバー停止 | ✅ 実装 | catch内送信とHeartbeat/WatcherFlushをtry/catchで保護 |
| #2 #3 `\`サニタイズ | ✅ 実装 | ファイル名・シート名の文字クラスに `\\` を追加 |
| #4 PNG後始末 | ✅ 実装 | Word同様 `$generatedImages` を追跡しfinallyで削除 |
| #6 `.bak`フォールバック | ✅ 実装 | 検証に通ったバックアップのみ採用し本体へ書き戻す |
| #7 `updatedAt` | ✅ 実装 | `TryParse` で不正値を最古扱いにし一覧を守る |
| #8 注釈JSONガード | ✅ 実装 | `Test-MbProject`/`Set-MbStepAnnotations` の両方に形式チェック |
| #9 loadリスナー | ✅ 実装 | `onload` に一本化しリスナー蓄積を解消 |
| #10 ナビ再構築 | ✅ 実装 | 200msデバウンス |
| #18 revision先行加算 | ✅ 実装 | 書込み失敗時にメモリ側を巻き戻す |
| D2 コントラスト | ✅ 実装 | `--text-3` を `#6b7280`（約4.6:1）へ |
| D3 常設影 | ✅ 実装 | `.project-card` / `.project-create` の影を削除、角丸9px→8px |
| D4 テーマ色統一 | ✅ 実装 | Wordの本文/見出し色をExcel側（24,32,51 / 58,91,160）へ統一 |
| D14 D15 | ✅ 実装 | 未使用の色定義を削除、シート概要をAA達成する本文色へ |
| **フォント統一** | ✅ 実装 | BIZ UDPゴシックへ統一。後述 |
| **D5 ウェイト正規化** | ✅ 実装 | フォント統一に伴い400/700へ（36箇所中15箇所を変更） |
| **D6 サイドバー幅** | 📝 doc修正 | コードの280pxが正。`UIUX-DIRECTION.md` を更新 |
| **D17 タブ色** | 📝 doc修正 | 濃淡での区別が正。`EXCEL-DESIGN-v0.19.md` を更新 |
| **D1 Excel画像幅** | ⛔ **誤検知・差し戻し** | 後述 |
| **#5 強制終了の緩和** | ⛔ **意図的に不採用** | 後述 |
| D7（角丸・アイコン寸法） | ⏸ 一部のみ | 角丸9px→8pxは解消。アイコン30/32/34pxとフォーカスリング3系統は未着手 |

### フォント統一（BIZ UDPゴシック）

実機で `BIZ UDPGothic` がインストール済み（登録名は英語表記のみ）で、Excel・Word・アプリの
いずれも解決順の先頭にあることを確認したうえで、次を統一した。

- **フォントスタックの並びを3箇所で一致**させた。CSSの `--font` を
  `Resolve-MbExcelBodyFont` / `Resolve-MbWordBodyFont` と同じ順序
  （BIZ UDP系 → BIZ UD系 → Meiryo → Yu Gothic UI → MS Pゴシック）に揃えた。
- **`font-weight` を 400 / 700 の2値へ正規化**（500→400、600・650・750・800→700）。
  基準フォントは2ウェイトしか持たないため、細かい階調は実機で再現されない。
  ブラウザーで実測し、使用中のウェイトが 400 と 700 だけになったことを確認済み。
- **番号注釈のフォントも統一**。従来はアプリのSVG・Excel/Word出力とも `Arial` 固定だった。
  ただし実測したところ BIZ UDPゴシックの「99」は ink幅41.0px で、直径48pxの円に対して窮屈
  （Arialは30.0px）。そのため**番号の文字サイズを 27 → 24 に調整**し、2桁でも余白が残るようにした。
  この値は `web/assets/js/app.js` のSVGと `New-MbAnnotatedImage`（GDI+）で**必ず同じにすること**。
  プレビューと出力の見た目が食い違う原因になる。

> 番号バッジだけはArialの方が数字が締まって見えるという判断もあり得る。その場合は
> app.js の `ANNOTATION_NUMBER_FONT` と `New-MbAnnotatedImage -NumberFontName` の
> 既定値を戻し、サイズを27へ戻せばよい（2箇所を必ず同時に戻す）。

### D1（Excel画像幅）は誤検知だった

レビューでは「設計docの画像8列・60%に対し実装が50:50」と指摘したが、実装を変更したところ
`tests\phase1\Test-Static.ps1` の3件が失敗した。同テストは
**「Excelの画像と説明を約半幅ずつへ再配分する」**として `A:G` / `H:L` と列幅 `9 / 15 / 20` を
仕様として固定しており、README も「画像は約半幅で表示し」と記載している。
つまり**半幅構成が v0.19以降の意図的な決定**で、`EXCEL-DESIGN-v0.19.md` の
「画像は左8列、説明は右4列」「約800×450px」の記述が古かった。

→ コード変更は全て差し戻し、`docs/EXCEL-DESIGN-v0.19.md` を実測値
（画像A:G・説明H:Lで約半分ずつ、1920×1080は約700×395px相当）へ更新した。

### #5（Stop-ProcessをHwnd限定から緩和）は採用しない → 第2回レビューで別解を実装

> 追記: 第2回レビューで、Wordの`Application.Hwnd`が常に0であること、その結果プロセスが
> 毎回残留することを実測で確認した。**強制終了の条件は緩めず**、`ActiveWindow.Hwnd`で
> 所有を証明できるようにして既存の安全網を働かせる形で解決した（前掲）。以下は初回時点の判断。

`WORD-OUTPUT-DESIGN-v0.12.md:37` は「`Stop-Process`はHwndで所有を証明したPIDにだけ許可する」と
明記している。PID差分モードは「起動直後に利用者がWord/Excelを開いた」場合に他人のプロセスを
指し得るため、緩和すると**利用者の未保存文書を強制終了する**恐れがある。これは本アプリの
最優先方針（未保存のExcel・Wordを変更または終了しない）に反するため、緩和せず
意図を明示するコメントを追加するに留めた。プロセス残留が実際に観測された場合は、
強制終了ではなく所有確認の強化（Hwnd取得の改善）で対応すべき。

### 古い記述を修正したdoc

D1と同じく「コードが正しく、docが古い」ケースを実装と突き合わせて洗い出し、docを修正した。

| doc | 古い記述 | 実装（正） |
|---|---|---|
| `UIUX-DIRECTION.md` | 左サイドバー**224px** | **280px**。v0.18で手順アウトラインを左へ集約した際に拡張 |
| `UIUX-DIRECTION.md` | フォントは4段の並び | BIZ UDPゴシック基準に統一し、Excel/Wordと同じ解決順＋ウェイト400/700を明記 |
| `EXCEL-DESIGN-v0.19.md` | 画像は左**8列**・約800×450px | 画像A:G・説明H:Lで**約半分ずつ**、約700×395px相当 |
| `EXCEL-DESIGN-v0.19.md` | 目次と手順シートのタブへ**同じ**テーマ色 | 目次は濃紺、手順シートは同系の青。**濃淡で役割を区別** |
| `WORD-OUTPUT-DESIGN-v0.12.md` | 本文は「濃い青灰色」 | Excelと同一色に統一（見出し58,91,160／本文24,32,51） |

なお、`UIUX-DIRECTION.md` の「上部バー54px」「編集領域最大1800px」「一覧行38px以上」
「通常ボタン36px」「本文15〜16px」「影は一時的に前面へ出る要素だけ」は実装と一致しており、
docは正しいままである（影は今回のD3修正で一致した）。

### D7 の残り

角丸の外れ値（9px）は8pxへ寄せて解消したが、次は未着手。影響範囲が広くテストで保護されていないため、
まとめて実施するのが安全。

- アイコンボタンの寸法が 30 / 32 / 34px と混在（方針は32px基準）
- フォーカスリングが `box-shadow` 2系統と `outline` 1系統に分裂
- ダイアログの角丸10pxが「6〜8px基準」から外れている

## サマリ

全体としてセキュリティ設計（ID正規表現検証、Zip Slip対策、マジックバイト検証、HTMLエスケープ、CSRFトークン、原子的置換）は堅実で、**致命的なデータ破壊・RCE・パストラバーサルは検出されなかった**。ただしサーバー全体を停止させる高リスクのバグが1件、出力を最終段で失敗させるバグが1件あり、最優先で対応が必要。デザイン面でも、方針に反する常設影や、主役であるはずのExcel画像が縮んでいる等の要修正がある。以下、コード不具合を優先度順にまとめ、後半に**デザインレビュー**を追記した。

| # | 優先 | 対象 | 要約 |
|---|------|------|------|
| 1 | 🔴 高 | Start-ManualBuilder.ps1:1496-1511 | ダウンロード中断等でcatch内送信が再スロー→`exit 1`でサーバー停止 |
| 2 | 🔴 高 | Excel.psm1:118 | ファイル名の`\`未サニタイズ→全処理後のMoveで出力失敗 |
| 3 | 🟠 中 | Excel.psm1:143,166 | シート名サニタイズ/検証で`\`欠落 |
| 4 | 🟠 中 | Excel.psm1:1200-1236 | 注釈描画PNGの後始末漏れ（毎回蓄積） |
| 5 | 🟠 中 | Excel.psm1:1222 / Word.psm1:507 | 強制終了がHwnd限定→WINWORD.EXE残留 |
| 6 | 🟠 中 | Project.psm1:259-275 | 破損時に`.bak`へフォールバックしない |
| 7 | 🟠 中 | Workspace.psm1:97 | 不正な`updatedAt`1件でカタログ全体が閲覧不能 |
| 8 | 🟠 中 | Project.psm1:471-496 | 不正な注釈JSONでStrictModeの生エラー漏れ |
| 9 | 🟠 中 | app.js:1037-1038 | 画像差し替えで`load`リスナー蓄積・二重描画 |
| 10 | 🟠 中 | app.js:1624-1627 | 打鍵ごとにナビ全DOM再構築＋注釈再描画 |
| 11〜21 | 🟡 低 | 各所 | 後述（堅牢性・多層防御の改善） |

---

## 🔴 高優先

### 1. レスポンス送信中の例外でサーバーが停止する

**対象**: `src/Start-ManualBuilder.ps1:1496-1511`（併せて `1487-1488`）

**現象**:
`Invoke-MbRoute` がレスポンスのヘッダー／本文を書き始めた後に例外が起きると、内側catch（1497/1500行）が `Write-MbResponse` を呼び直す。送信済みレスポンスに対して `Write-MbResponse` は `StatusCode`／`Headers.Add`／`ContentLength64` を設定するため `InvalidOperationException` を再スロー。この2次例外はcatch節内（try保護外）で発生するため内側tryでは捕捉されず、外側catch（1509行）→ `exit 1`（1511行）に達し**サーバープロセスごと停止**する。

**再現シナリオ**:
大きな画像／ZIPのダウンロード（`Write-MbFile`／`Write-MbDownload` の `CopyTo`）中にブラウザーが接続を中断 → `HttpListenerException` → 内側catchが `Write-MbResponse` → 再スロー → `exit 1`。1回のキャンセルで全ユーザー作業（未保存セッション含む）が失われる。待機ループ内の `Update-MbCaptureHeartbeatState`／`Invoke-MbWatcherFlush`（1487-1488行）もtry保護外で同じ経路をたどる。

**修正方針**:
- 内側catchの `Write-MbResponse` 呼び出しを `try{}catch{}` で囲む（これだけでプロセス停止は防げる）。
- できれば「まだ何も送信していない」場合のみエラー応答を書き込むフラグ管理を追加。
- 待機ループ内の Heartbeat／WatcherFlush も `try{}catch{}` で囲む。

```powershell
} catch [System.UnauthorizedAccessException] {
    if ($context) { try { Write-MbResponse $context $_.Exception.Message 403 'text/plain; charset=utf-8' } catch { } }
    Write-MbLog $_.Exception.Message 'WARN'
} catch {
    if ($context) { try { Write-MbResponse $context '処理中にエラーが発生しました。' 500 'text/plain; charset=utf-8' } catch { } }
    Write-MbLog $_.Exception.Message 'ERROR'
}
```

```powershell
while (-not $contextTask.AsyncWaitHandle.WaitOne(200)) {
    try { Update-MbCaptureHeartbeatState } catch { Write-MbLog $_.Exception.Message 'WARN' }
    try { Invoke-MbWatcherFlush } catch { Write-MbLog $_.Exception.Message 'WARN' }
    if (-not $script:Running) { break }
}
```

---

### 2. Excelファイル名のサニタイズがバックスラッシュを除去しない

**対象**: `src/ManualBuilder.Excel.psm1:118`（`Get-MbSafeExcelFileName`）

**現象**:
```powershell
$safe = $Name -replace '[\/:*?"<>|]', '_'
```
文字クラス `[\/...]` の `\/` は「エスケープされた `/`」であり、リテラルの `\` を含まない。`/` は置換されるが `\` はすり抜ける。Word側（`ManualBuilder.Word.psm1:87` の `'[\\/:*?"<>|]'`）は `\` を含んでおり**非対称**。

**再現シナリオ**:
`Project.title` が `売上\集計` の場合、返り値は `売上\集計_....xlsx`。全描画処理を終えた最終段 `Excel.psm1:1172` の `[IO.File]::Move($temporaryPath, $outputPath)` が、存在しないサブフォルダ `売上\` へのMoveとなり `DirectoryNotFoundException` で**出力全体が失敗**する（重い処理を全部終えた最後に落ちる、最悪のUX）。

**修正方針**: 先頭に `\\` を追加してWordと統一。
```powershell
$safe = $Name -replace '[\\/:*?"<>|]', '_'
```

---

## 🟠 中優先

### 3. Excelシート名のサニタイズ／検証でバックスラッシュが欠落

**対象**: `src/ManualBuilder.Excel.psm1:143`（サニタイズ）、`166`（検証）

Excelがワークシート名に禁止する文字は `: \ / ? * [ ]` の7種。どちらの文字クラス `[:\/\?\*\[\]]` にも `\` が無い。シート名が `A\B` だとサニタイズを抜け、`Test-MbExcelWorksheetName` も `$true` を返す。結果 `worksheet.Name = ...`（923行）でExcelが例外→catchフォールバックで `手順シート N` に置換され、**ユーザー指定名が失われる**。

**修正方針**: 両所の文字クラスに `\\` を追加。
```powershell
# 143行
$safe = $safe -replace '[:\\/\?\*\[\]]', '・'
# 166行
if ($Name -match '[:\\/\?\*\[\]]') { return $false }
```

---

### 4. Excel注釈描画PNGの後始末漏れ

**対象**: `src/ManualBuilder.Excel.psm1:781, 1031, 1200-1236`（Word実装との非対称）

Excelは `rendered-images\<step.id>.png` を書き出すが、Wordのような追跡・削除が無い。Word側は `$generatedImages`（`Word.psm1:294,402`）に登録しfinally（`517-519`）で削除している。Excelのfinally（`1200-1236`）は一時xlsxのみ削除。**注釈付き手順を含むExcel出力のたびにPNGが恒久蓄積**（成功・失敗・キャンセルのいずれでも残留）。

**修正方針**: Word同様、生成した派生画像パスをリストに蓄積し、finallyで削除する。

---

### 5. 強制終了フォールバックがHwndモード限定でプロセス残留

**対象**: `src/ManualBuilder.Excel.psm1:1222` / `src/ManualBuilder.Word.psm1:507`

```powershell
if ($ownershipProven -and $ownPid -gt 0 -and $ownershipMode -eq 'Hwnd') { ... Stop-Process ... }
```
所有権解決には `Hwnd` と「空ベースライン＋PID差分」の2モードがあるが、強制終了の安全網は `Hwnd` のみ対象。特にWordは `Word.Application` が `Hwnd` を持たない可能性が高く、常にPID差分モードになりやすい。`Quit()` が完全終了に失敗すると **WINWORD.EXE が残留**する（本アプリの「プロセスを残さない」設計に反する）。

**修正方針**: 所有が「証明済み」なら（モードを問わず）締めのタイムアウト＋`Stop-Process` を適用する。
```powershell
if ($ownershipProven -and $ownPid -gt 0) { ... Stop-Process -Id $ownPid ... }
```

---

### 6. project.json破損時に`.bak`へフォールバックしない

**対象**: `src/ManualBuilder.Project.psm1:259-275`（`Get-MbProject`）

`Save-MbProject`（`240,247`）は `[IO.File]::Replace` で `project.json.bak` を維持するのに、`Get-MbProject` は `ConvertFrom-Json` 失敗時に例外を投げるだけで `.bak` を参照しない。書込み途中の電源断等で本体が破損すると、**直前の正常コピーがあるのに復旧不能**になり、バックアップの意味が失われている。

**修正方針**: `Get-MbProject` のcatchで `"$Path.bak"` が存在すれば読込＋`Test-MbProject` 検証を試み、成功時はそれを返して本体を再生成する。

---

### 7. 不正な`updatedAt`1件でカタログ全体が閲覧不能

**対象**: `src/ManualBuilder.Workspace.psm1:97`（`Get-MbProjectCatalog`）

`Sort-Object @{ Expression = { [DateTime]$_.updatedAt } }` を使うが、`Test-MbProject` は `updatedAt` の形式を検証していない（`Repair-MbProject` も欠損補完のみ）。手編集・レガシー移行（`Storage.psm1:101-103` は無検証コピー）でパース不能な `updatedAt` を持つプロジェクトが1件でもあると、`[DateTime]` キャストがソート中に例外→**一覧全体が表示不能**になり、正常なマニュアルにも到達できなくなる。

**修正方針**: `[DateTime]::TryParse` を使い、失敗時はフォールバック値（`[DateTime]::MinValue` かディレクトリの `LastWriteTimeUtc`）を使う。
```powershell
$expr = {
    $d = [DateTime]::MinValue
    [void][DateTime]::TryParse($_.updatedAt, [ref]$d)
    $d
}
... | Sort-Object -Property @{ Expression = $expr; Descending = $true }
```

---

### 8. 不正な注釈JSONでStrictModeの生エラーが漏れる

**対象**: `src/ManualBuilder.Project.psm1:471-496`（`Set-MbStepAnnotations`）、併せて `Test-MbProject:197`

try/catchが `ConvertFrom-Json`（462-466行）しか保護しておらず、その後の要素走査（`[string]$annotation.id`（472）、`[int]$annotation.label`（485）等）は保護外。`Set-StrictMode -Version 2.0` 下で、要素がオブジェクトでない（`[1,2,3]`）や `id` 欠如（`[{}]`）だと「Property 'id' cannot be found」等の**生の英語例外**がそのまま伝播する（意図した「注釈データを読み込めません。」でない）。`485` は `PSObject.Properties.Name -contains 'label'` でガードしているが `472` は非対称。`Test-MbProject:197` の `[int]$annotation.label` も存在チェック無し。

**修正方針**: 要素走査全体をtry/catchで囲み既存の親切なメッセージに統一。プロパティアクセス前に `PSObject.Properties.Name -contains` でガードを揃える。

---

### 9. 画像差し替え時に`load`リスナーが蓄積・二重描画

**対象**: `web/assets/js/app.js:1037-1038`（`updateReplacedImageCard`）

`renderCardAnnotations`（300-303行）は `dataset.annotationLoadBound` フラグで `addEventListener('load', …)` を一度だけ登録する設計。しかし差し替え時に
```js
image.removeAttribute('data-annotation-load-bound');
image.onload = () => renderCardAnnotations(card);
```
としており、`removeAttribute` は**フラグを消すだけで既存の `addEventListener` リスナーを解除しない**（img要素を使い回すため残存）。差し替え／「元の画像へ戻す」をN回行うと旧リスナーがN個累積し、以降のロードごとに `renderCardAnnotations`（SVG全再描画）が多重に走る。

**修正方針**: `image.onload` 方式に一本化し `addEventListener` 側を廃止するか、`removeEventListener` で明示解除する。または `image.replaceWith(image.cloneNode())` で旧リスナーごと差し替える。

---

### 10. 打鍵ごとに手順ナビを全DOM再構築＋注釈再描画

**対象**: `web/assets/js/app.js:1624-1627 → 328-388`

タイトル／説明の `input` ごとに `rebuildStepNavigation()` が発火し、`stepCards()` 全件をループしてナビDOMを全再生成（`replaceChildren`）、さらに `setActiveStep` 経由で `renderCardAnnotations(active)` まで走る。手順数が多いマニュアルでは1文字ごとに全再生成＋SVG再描画が同期実行され、**入力レイテンシ・ジャンク**が顕在化する（IME変換中は特に重い）。

**修正方針**: `input` ハンドラをデバウンス（150〜250ms）する、または変更カードに対応するナビ項目のテキストのみ部分更新する。テキスト編集時（注釈不変）は `renderCardAnnotations` を呼ばない。

---

## 🟡 低優先（堅牢性・多層防御）

| # | 対象 | 内容 | 方針 |
|---|------|------|------|
| 11 | Excel.psm1:1012 / Word.psm1:396 | `$step.imageId` を未ガードアクセス（StrictMode 2.0で旧フォーマット時に例外）。`crop` は `PSObject.Properties.Name -contains` でガード済で非対称 | `imageId` も同様にガード統一 |
| 12 | Excel.psm1:806,812-813 | 既存Excel稼働時に所有権未解決だと、生成した不可視Excelが `Quit` も強制終了もされず稀に残留 | 生成済みかつブック未追加なら明示 `Quit()` |
| 13 | app.js:794 | `Math.min(99, max+1)` により番号注釈が「99」で重複 | 上限到達時は追加拒否＋トースト、または空き番号探索 |
| 14 | Start-ManualBuilder.ps1:117,780,1038,1362 | セッショントークンをURLクエリで送出（履歴・ログに残存）。`Referrer-Policy: no-referrer` で一部緩和済 | 画像も `X-Manual-Token` ヘッダー認証へ統一 |
| 15 | Start-ManualBuilder.ps1:770-773 | Hostヘッダー欠如時に検証スキップ（DNSリバインディング防御の穴） | 期待値に厳密一致しない場合は空も含め拒否 |
| 16 | Start-ManualBuilder.ps1:786-791 | Origin欠如のPOSTを許可（トークンで緩和済） | 可能ならOrigin欠如も拒否し二重化 |
| 17 | Start-ManualBuilder.ps1:861 | `LocalPath` の二重 `UnescapeDataString`（バイパスは不成立だが不要） | `$request.Url.LocalPath` をそのまま使う |
| 18 | Project.psm1:231 | `revision` を書込み前にインクリメント。Replace失敗でメモリとディスクが不整合（カウンタが飛ぶ） | 書込み成功後にインクリメント |
| 19 | Capture.psm1:107-128 | 画像をJSON永続化前にディスクへ書くため、保存失敗時に孤児ファイルが残る | 起動時／保存時に `Remove-MbUnreferencedImages` を明示実行 |
| 20 | app.js:1440-1448 / サーバー | Excel/Word出力の同時起動を相互排他していない | 出力中は両ボタン無効化、またはサーバー側で排他確認 |
| 21 | app.js:1974-1980 | シート切替時に未保存テキストのフラッシュ無し（テンプレートのデバウンス次第で最後の入力喪失） | 切替前に保留中の保存を明示フラッシュ |

---

# デザインレビュー（追記）

コードによる不具合とは別に、**アプリ画面／Excel・Word出力のビジュアルデザイン**を、実行中アプリの実DOM・完全なCSS・出力生成コードを設計ドキュメント（`UIUX-DIRECTION.md`／`EXCEL-DESIGN-v0.19.md`／`WORD-OUTPUT-DESIGN-v0.12.md`）と突き合わせてレビューした。全体の完成度は高いが、**明文化された方針に反する箇所**と**Excelの主役であるはずの画像が小さい問題**を優先で直したい。

> 補足: Browserペインが非表示のためピクセル単位のスクリーンショットは取得できず、実DOM構造＋CSS実値＋生成コードの静的解析による評価。レイアウトの目視確認が必要なら、Browserペインを表示のうえ再取得する。

| # | 優先 | 対象 | 要約 |
|---|------|------|------|
| D1 | 🔴 高 | Excel.psm1:595-598,931-933 | 画像/説明の幅が実質50:50（設計は画像8列60%・800px相当）。主役の画像が小さい |
| D2 | 🔴 高 | app.css:8 | 二次テキスト色 `#8a909b` が白地で約3.2:1、WCAG AA未達（多数の小サイズ文字に使用） |
| D3 | 🟠 中 | app.css:2038,2120,2125 | カードに常設の影＝方針「常設影・グラデ禁止」に違反 |
| D4 | 🟠 中 | Word.psm1:191-192 / Excel.psm1:526-530 | Excel↔Wordでアクセント青・本文色のRGBが不一致（別テーマに見える） |
| D5 | 🟠 中 | app.css 全体 | `font-weight` 500/650/750/800 が日本語フォントで再現されず階層が崩れる |
| D6 | 🟠 中 | app.css:400 | 左サイドバー280px＝方針224pxと乖離、狭幅で早期に縦積み |
| D7 | 🟠 中 | app.css 各所 | 角丸6/7/8/9/10px混在、アイコンボタン22/30/32/34px混在、フォーカスリング3系統 |
| D8〜 | 🟡 低 | 各所 | トースト色依存、補足黄のトークン分岐、出力の色値重複 など |

## 🔴 デザイン高優先

### D1. Excel出力で画像が主役になっていない（実質50:50）

**対象**: `src/ManualBuilder.Excel.psm1:595-598`（範囲定義）、`931-933`（列幅）

**現象**: 列幅は A=9 / B〜G=15×6=90（画像領域 `A:G` 合計=**99**）、H〜L=20×5=**100**（説明領域 `H:L`）。つまり画像側99・説明側100で**むしろ説明側がわずかに広い**。設計方針と二重に食い違う:
- `EXCEL-DESIGN-v0.19.md:27`「画像は左**8列**、説明は右**4列**」 → 実装は画像 A〜G(**7列**)／説明 H〜L(**5列**)で1列ずれ。
- `EXCEL-LAYOUT-PLAN-v0.7.md:33-34`「画像側は約**60%**、文章側は約40%」 → 実装は約50:50。
- `EXCEL-DESIGN-v0.19.md:28`「1920×1080画像を約**800×450px**相当まで表示」 → 使用可能画像幅（`imageArea.Width-18`, 692行）は約700px相当で約12%不足。

Excelは「PCの横長画面で画像を大きく読む」ことが主眼（`README.md`／方針§3）なのに、その主役が縮んでいる。

**修正方針**: 画像側を A〜H(8列)、説明側 I〜L(4列)へ寄せる、または列幅を画像側優位に再配分（例: 画像側合計≈145／説明側≈95）し、16:9画像が約800px幅で表示され、かつ `imageArea.Width > textPanel.Width` を必ず満たすようにする。範囲定義（595-598行）とヘッダー/番号/カード枠（592-599行の `L` 基準）も新しい列境界に合わせて更新する。

### D2. 二次テキスト色がコントラスト基準（WCAG AA）に未達

**対象**: `web/assets/css/app.css:8` — `--text-3: #8a909b`

**現象**: `#8a909b`(rgb 138,144,155) は白（`--surface #fff`）上で約**3.2:1**。18px未満の本文・ラベルは4.5:1が必要。`--text-3` は `.step-total`(12px)、`.field__optional`(11px)、`.sidebar__version`(10px)、`.image-placeholder`(12px)、`.count-badge` など**小サイズ文字に多用**されており、方針§6 UX-10（色覚・可読性）にも触れる。

**修正方針**: 本文用途の値を約4.6:1を満たす `#6b7280` 前後へ引き上げ、10px以下での情報色使用を撤廃。
```css
--text-3: #6b7280;
```

## 🟠 デザイン中優先

### D3. カードに常設の影（方針違反）

**対象**: `web/assets/css/app.css:2038`（`.project-create`）、`2120,2125`（`.project-card` の常設影・hover影）

方針§3「**影はメニューやトーストなど一時的に前面へ出る要素だけに使う。カード常設の影とグラデーションは使わない**」に明確に反する。一覧カードが常時 `box-shadow` を持つ。

**修正方針**: 常設影を削除し境界線で区切る。hoverは弱い `border-color` 変化に留める。
```css
.project-card { box-shadow: none; }
.project-card:hover { border-color: var(--border-strong); box-shadow: none; }
.project-create { box-shadow: none; }
```

### D4. Excel↔Wordでテーマ色が不一致

**対象**: `src/ManualBuilder.Word.psm1:191-192` vs `src/ManualBuilder.Excel.psm1:526-530`

同一 `project.json` から作る2出力なのに青系・本文色のRGBが異なり、並べると別テーマに見える。

| 役割 | Excel | Word |
|---|---|---|
| アクセント青 | 58,91,160 | 46,94,140 |
| 本文文字色 | 24,32,51 | 32,55,72 |

各docには個別適合だが、クロス出力の統一が取れていない。

**修正方針**: テーマ色を1組の定数に統一（例: アクセント=58,91,160／濃紺=38,57,104／本文=24,32,51）し、両モジュールで共有する。

### D5. `font-weight` の階調が日本語フォントで再現されない

**対象**: `web/assets/css/app.css` 全体（`.button` 650、`.eyebrow`/`.step-number` 800、`.excel-export-dialog__mark` 750 など、500/600/650/700/750/800が混在）

主フォント `BIZ UDPGothic`／`Meiryo` は Regular/Bold の2ウェイトのみ。600以上は全てBold、500はRegularに丸められ、**設計意図の細かな太さ階層が実機で再現されない**。

**修正方針**: 太さは `400 / 700` の2値に統一し、階層はサイズと色で表現する。

### D6. 左サイドバー幅が方針(224px)と乖離

**対象**: `web/assets/css/app.css:400` — `grid-template-columns: 280px ...`

方針§3「左サイドバー224px」に対し280px。編集領域を56px圧迫し、狭幅（1280px・125%/150%）で早期に縦積みへ追い込む。

**修正方針**: `224px` に寄せる。方針側を実測値に更新して整合を取る選択でも可（どちらが正か判断の上で統一）。
```css
.app-layout { grid-template-columns: 224px minmax(0, 1fr); }
```

### D7. 仕上げの不統一（角丸・アイコンボタン・フォーカスリング）

**対象**: `web/assets/css/app.css` 各所
- 角丸: `6/7/8/9/10px` 混在（`.project-card` 9px:2118、`.annotation-editor` 10px:1504、`.excel-export-dialog` 10px:1735 等）。方針§3「6〜8px基準」。→ 面=8px／部品=6px／一時前面=8px に統一、9・10pxを廃止。
- アイコンボタン: `.icon-button` 32px は方針どおりだが `.drag-handle`/`.step-action-button` 30px、`.step-nav__drag` 22px幅、`.excel-export-dialog__close` 34px とばらつく。→ 基準32pxに統一、密度が要る箇所のみ明示的に28pxサブサイズ。
- フォーカスリング: グローバル `box-shadow ...rgba(108,139,211,.28)`(60行)／`.brand--home` `outline 2px`(166行)／入力欄 `rgba(58,91,160,.16)`(232行)の3系統。→ 全要素 `outline: 2px solid var(--focus); outline-offset: 2px;` に統一。

## 🟡 デザイン低優先

| # | 対象 | 内容 | 方針 |
|---|------|------|------|
| D8 | app.css:1465 | トースト種別が背景色のみ＋色`#30343b`がトークン外 | 種別アイコン(✓/！)＋文言を併記、色をトークン化 |
| D9 | app.css:1373-1378 | 補足の黄(`#fffdf5`/`#ead9a8`)が `--warning-soft` と別系統、ラベルが低コントラスト | `--warning-soft`/`--warning` へ寄せ、ラベルは `--text-2` |
| D10 | app.css:819 | 状態ドット`#d99b19`がトークン外、「説明未入力」と「画像なし」が同色 | `--warning`系にトークン化、2状態を塗り/枠線で差別化 |
| D11 | app.css 全体 | padding/gapに 2,3,5,7,9,11,13,26,28px 等の非スケール値が頻出 | 4px基準スケール＋`--space-*` トークン化 |
| D12 | app.css:1378 | 補足textareaが14px＝本文下限15〜16px(方針§3)を下回る | 15pxに戻すか副次テキストとして明確化 |
| D13 | app.css:2455 | 極狭幅で監視状態テキストが32pxに切詰め(UX-04に触れる) | アイコン＋短縮ラベルを残す |
| D14 | Excel.psm1:528 vs 790 | `colorAccentSoft` が2箇所で不一致(239,243,251 / 238,243,252)＋528行はデッドコード | 528行を削除し790行へ一本化 |
| D15 | Excel.psm1:981-982 | シート概要が `colorMuted`×`accentSoft`地で約4.4:1(AA未達) | 文字色を一段濃く、または地を白に |
| D16 | Excel.psm1:663,676 / Word.psm1:246 | 補足の文字色が3系統(茶122,83,0／紺／茶92,72,20) | 1色(例 茶92,72,20)に統一 |
| D17 | Excel.psm1:902,930 | 目次タブ/手順タブが別色＝設計「同じテーマ色」と不一致 | 同色にするか設計docを「濃淡で区別」に更新 |
| D18(要確認) | app.js カード生成部 | `.step-card__actions`/`.step-delete-button`/`.drag-handle`(app.css:966-1017)が残存。方針(EDITOR-UX-v0.18)は管理操作を左ナビへ集約 | 実描画されていれば操作重複。描画有無を確認し不要なら撤去 |

## デザインで良かった点（回帰防止のため記録）

- **アプリ**: 色/フォントのトークン化、フォントスタックは方針どおり（BIZ UDPGothic→Meiryo→Yu Gothic UI）、上部バー54px・編集領域max1800px・一覧行min38pxが方針一致。状態ドットに`title`/`aria-label`併記、保存状態は記号＋文言、`prefers-reduced-motion`対応、`:focus-visible`常備。
- **Excel/Word**: `STEP 01`書式、罫線をカード外周・領域境界に限定＋グリッド線非表示、画像淡グレー/説明白/補足淡黄の領域配色、濃紺地×白文字のコントラスト約6.5:1、Word余白20mm・表紙28pt/見出し16・13pt/本文10.5pt・フッター中央ページ番号・1920×1080画像を450pt幅、画像拡大上限(Excel1.5倍/Word2倍)はいずれも設計適合。

---

## レビューで問題なしと確認した点（回帰防止のため記録）

- **XSS**: HTML出力のユーザーテキスト（title/description/note/シート名/注釈）は全て `ConvertTo-MbHtml`（`WebUtility.HtmlEncode`）を通す。JS側もユーザー入力は全て `textContent`、`innerHTML` は静的リテラルのみ。`insertAdjacentHTML` 不使用。
- **パストラバーサル**: ID類は `^(default|project-[a-f0-9]{32})$`、画像は `^image-[a-f0-9]{32}\.(png|jpg|bmp)$`、シートは `^sheet-[a-f0-9]{32}$` で厳格検証。
- **Zip Slip**: エントリ名の `\` 混入拒否、`[IO.Path]::GetFileName` 使用、`CreateNew` 書込み、サイズ・件数上限あり。
- **画像検証**: マジックバイト＋System.Drawing実体デコード＋総画素上限＋SHA-256 dedup。
- **CSRF**: 状態変更系は `X-Manual-Token` カスタムヘッダー必須。htmxは `selfRequestsOnly=true`。
- **原子性**: 一時ファイル→`[IO.File]::Replace`（同一ディレクトリ）→`.bak`退避。
- **$null比較**: 全箇所 `$null -ne $x` の左置き順で正しい。日付/数値は明示フォーマットでカルチャ非依存。
- **COM解放**: `try/finally + Release-Mb*ComObject`、`DisplayAlerts=$false`、既存インスタンス誤接続防止（`MB_CONNECTED_TO_EXISTING_*`）は堅実。

## 推奨対応順

1. **#1（サーバー停止）** — 最優先。catch内送信のtry/catch保護だけでも即入れる。
2. **#2, #3（`\`サニタイズ）** — 小さく安全な修正。#5（プロセス残留）と併せてOffice出力の安定化。
3. **D1（Excel画像幅）, D2（コントラスト）, D3（常設影）** — 出力/画面の見た目に直結し、方針違反も明確。小〜中規模の修正。
4. **#4（PNG蓄積）, #6, #7, #8（永続化の堅牢化）**
5. **#9, #10（フロントの性能・リーク）, D4〜D7（テーマ色・トークン統一）**
6. **#11〜21・D8〜D18** — 余力に応じて。#14〜17はlocalhost単一ユーザー前提のため実害は限定的。デザイン低優先はトークン整理としてまとめて実施すると効率的。

> デザイン方針とコードの数値が食い違う項目（D1の列数・D6のサイドバー幅・D17のタブ色）は、**コードを方針に合わせるか／方針docを実装に合わせるか**の判断が必要。方針を正とするのが基本だが、実装側が最新の意図であれば docs を更新して整合を取ること。
