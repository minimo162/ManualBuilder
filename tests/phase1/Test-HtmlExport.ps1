# Phase 1 HTML exporter test.
# HTML出力はCOMを使わないため、生成結果をここで完全に確かめられる。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Html.psm1') -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-HtmlTest-' + [guid]::NewGuid().ToString('N'))

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)

    # --- 出力フォルダー名 ---
    Assert-Mb ((Get-MbSafeHtmlFolderName -Name '営業手順' -Directory $testRoot) -eq '営業手順') '日本語のフォルダー名をそのまま使う'
    Assert-Mb ((Get-MbSafeHtmlFolderName -Name 'a/b:c*d?' -Directory $testRoot) -eq 'a_b_c_d_') '使えない文字を置き換える'
    Assert-Mb ((Get-MbSafeHtmlFolderName -Name 'CON' -Directory $testRoot) -eq '_CON') '予約された名前を避ける'
    Assert-Mb ((Get-MbSafeHtmlFolderName -Name '   ' -Directory $testRoot) -eq 'manual') '空の名前でも出力できる'
    # 日付も連番も付けない。マニュアル1つ＝フォルダー1つにして、共有フォルダー側も同じ場所を更新できるようにする。
    [void](New-Item -ItemType Directory -Path (Join-Path $testRoot '重複') -Force)
    Assert-Mb ((Get-MbSafeHtmlFolderName -Name '重複' -Directory $testRoot) -eq '重複') '同名フォルダーがあっても同じ名前を返す'
    Assert-Mb ((Get-MbHtmlExportFolderName -Project ([pscustomobject]@{ title = '営業手順' }) -OutputDirectory $testRoot) -eq '営業手順') 'マニュアル名からフォルダー名を決める'
    Assert-Mb ((Get-MbHtmlExportFolderName -Project ([pscustomobject]@{ title = '営業手順' }) -OutputDirectory $testRoot) -notmatch '\d{8}') 'フォルダー名に日付を付けない'

    # --- 出力対象のシート ---
    $project = New-MbProject
    $project.title = '営業システム操作マニュアル'
    $sheetId = [string]$project.sheets[0].id
    $project.sheets[0].name = 'ログイン'
    $emptySheet = Add-MbSheet -Project $project
    $emptySheet.name = '手順のない章'
    Assert-Mb ((@(Get-MbHtmlContentSheets -Project $project)).Count -eq 0) '手順が無ければ出力対象にならない'

    $step1 = Add-MbStep -Project $project -SheetId $sheetId
    $step1.title = 'ログイン画面を開く'
    $step1.description = "ブラウザーを起動します。`nアドレスを入力します。"
    $step1.note = '社内ネットワークからのみ接続できます'
    $step2 = Add-MbStep -Project $project -SheetId $sheetId
    $step2.title = ''
    $step3 = Add-MbStep -Project $project -SheetId $sheetId
    $step3.title = '<script>alert(1)</script> & "危険" な名前'
    $step3.description = 'タグ <b> や & を含む説明'

    $sheets = @(Get-MbHtmlContentSheets -Project $project)
    Assert-Mb ($sheets.Count -eq 1) '手順の無いシートは出力しない'

    # --- HTML生成 ---
    $imageNames = @{ ([string]$step1.id) = 'images/step-0001.png' }
    $videoNames = @{ ([string]$step1.id) = '_source/videos/video-0123456789abcdef0123456789abcdef.mp4' }
    $html = ConvertTo-MbManualHtml -Project $project -ImageNames $imageNames -VideoNames $videoNames -GeneratedAt '2026年8月3日 10:00'

    Assert-Mb ($html.StartsWith('<!doctype html>')) 'HTML文書として始まる'
    Assert-Mb ($html -match '<meta charset="utf-8">') '文字コードを宣言する'
    Assert-Mb ($html -match '<meta name="viewport"') 'スマートフォンでも読める指定を入れる'
    Assert-Mb ($html -match '<title>営業システム操作マニュアル</title>') 'タイトルを出す'

    # 共有フォルダー上ではゾーン判定でスクリプトが止まることがあるため、JavaScriptを使わない。
    Assert-Mb ($html -notmatch '(?i)<script') 'JavaScriptを一切使わない'
    Assert-Mb ($html -notmatch '(?i)\son\w+\s*=') 'onclickなどのイベント属性も使わない'
    Assert-Mb ($html -notmatch '(?i)<link[^>]+href') '外部ファイルを読み込まない'
    Assert-Mb ($html -match '<style>') 'スタイルを埋め込む'

    Assert-Mb ($html -match '<img src="images/step-0001\.png"') '画像を相対パスで参照する'
    # 動画は元データ（_source）の1本だけを参照する。同じ動画をフォルダー内に二重に持たない。
    Assert-Mb ($html -match '<video src="_source/videos/video-0123456789abcdef0123456789abcdef\.mp4"') '動画は元データの1本を参照する'
    Assert-Mb ($html -match 'poster="images/step-0001\.png"') '動画の再生前は焼き込み済み画像を出す'
    Assert-Mb ($html -match 'controls') '動画に再生操作を付ける'
    Assert-Mb ($html -match 'この手順に画像はありません') '画像の無い手順もその旨を出す'
    Assert-Mb ($html -match 'class="mb-step mb-step--has-video"') '動画つきの手順に印を付ける'
    Assert-Mb ($html -match '\.mb-step--has-video \.mb-step__still \{ display: none; \}') '画面では静止画と動画を二重に出さない'
    Assert-Mb ($html -match '\.mb-step--has-video \.mb-step__still \{ display: block; \}') '印刷では動画の代わりに静止画を出す'

    Assert-Mb ($html -match 'id="sheet-1"') 'シートに移動先を付ける'
    Assert-Mb ($html -match 'href="#sheet-1-step-1"') '目次から手順へ移動できる'
    Assert-Mb ($html -match 'id="sheet-1-step-3"') '手順ごとに移動先を付ける'
    Assert-Mb ($html -match '手順名未入力') '手順名が空でも見出しを出す'

    # --- エスケープ ---
    Assert-Mb ($html -notmatch '<script>alert\(1\)</script>') '手順名のタグをそのまま出力しない'
    Assert-Mb ($html -match '&lt;script&gt;alert\(1\)&lt;/script&gt;') '手順名のタグをエスケープする'
    Assert-Mb ($html -match '&amp;') 'アンパサンドをエスケープする'
    Assert-Mb ($html -match '&quot;危険&quot;') '引用符をエスケープする'
    Assert-Mb ($html -match 'タグ &lt;b&gt; や &amp; を含む説明') '説明文もエスケープする'

    # --- 印刷 ---
    Assert-Mb ($html -match '@media print') '印刷用の指定を持つ'
    Assert-Mb ($html -match 'page-break-inside: avoid') '手順が印刷で分断されないようにする'

    # --- 改行の扱い ---
    Assert-Mb ($html -match 'white-space: pre-wrap') '説明文の改行を保つ'

    # --- 出力そのもの（画像はGDI+が要るため、画像なしの手順で確かめる） ---
    $exportProject = New-MbProject
    $exportProject.title = '出力テスト'
    $exportSheetId = [string]$exportProject.sheets[0].id
    $exportStep = Add-MbStep -Project $exportProject -SheetId $exportSheetId
    $exportStep.title = '画像なしの手順'
    $exportStep.description = '説明文'
    $projectPath = Join-Path $testRoot 'project.json'
    $exportProject = Save-MbProject -Project $exportProject -Path $projectPath

    $outputRoot = Join-Path $testRoot 'output'
    $result = Invoke-MbHtmlExport -Project $exportProject -ProjectPath $projectPath -OutputDirectory $outputRoot
    Assert-Mb (Test-Path -LiteralPath $result.IndexPath -PathType Leaf) 'index.htmlを出力する'
    Assert-Mb ([int]$result.StepCount -eq 1) '手順数を返す'
    Assert-Mb ([int]$result.ImageCount -eq 0) '画像の無い手順を数えない'
    Assert-Mb ([long]$result.TotalBytes -gt 0) '出力の合計サイズを返す'
    Assert-Mb ((Split-Path -Leaf $result.OutputPath) -eq [string]$result.FolderName) 'フォルダー名を返す'

    $writtenBytes = [IO.File]::ReadAllBytes($result.IndexPath)
    Assert-Mb ($writtenBytes[0] -eq 0xEF -and $writtenBytes[1] -eq 0xBB -and $writtenBytes[2] -eq 0xBF) 'UTF-8 BOM付きで書き出す'
    $written = [IO.File]::ReadAllText($result.IndexPath, [Text.Encoding]::UTF8)
    Assert-Mb ($written -match '出力テスト') '出力したHTMLにタイトルが入る'
    Assert-Mb ($written -match '画像なしの手順') '出力したHTMLに手順が入る'

    # 未完成フォルダーを残さない
    Assert-Mb (@(Get-ChildItem -LiteralPath $outputRoot -Directory -Force | Where-Object { $_.Name -like '.mb-html-*' }).Count -eq 0) '作業用フォルダーを残さない'

    # --- 同梱する元データ（_source） ---
    $sourcePath = Join-Path $result.OutputPath '_source'
    Assert-Mb (Test-Path -LiteralPath $sourcePath -PathType Container) '出力フォルダーへ元データを同梱する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $sourcePath 'project.json') -PathType Leaf) '元データに project.json を入れる'
    Assert-Mb ([string]$result.SourcePath -eq $sourcePath) '元データの場所を返す'
    $sourceItem = Get-Item -LiteralPath $sourcePath -Force
    $onWindows = [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT
    Assert-Mb ((($sourceItem.Attributes -band [IO.FileAttributes]::Hidden) -ne 0) -or -not $onWindows) '元データは隠しフォルダーにする'

    # --- 一緒に配る.cmd ---
    $openCmd = Join-Path $result.OutputPath 'マニュアルを開く.cmd'
    $editCmd = Join-Path $result.OutputPath '編集する.cmd'
    Assert-Mb (Test-Path -LiteralPath $openCmd -PathType Leaf) '「マニュアルを開く.cmd」を一緒に出す'
    Assert-Mb (Test-Path -LiteralPath $editCmd -PathType Leaf) '「編集する.cmd」を一緒に出す'
    $openText = Get-MbHtmlOpenCommandText
    $editText = Get-MbHtmlEditCommandText
    # 共有フォルダー上の.htmlはInternet Explorerモードへ回されることがあるため、Edgeを明示して開く。
    Assert-Mb ($openText -match 'msedge\.exe') 'Edgeを明示して開く'
    Assert-Mb ($openText -match 'start "" "%MB_INDEX%"') 'Edgeが無ければ既定のブラウザーで開く'
    Assert-Mb ($editText -match '-ImportFrom "%~dp0_source"') '編集時は同梱した元データを取り込む'
    Assert-Mb ($editText -match '-PublishTo "%~dp0\."') '編集時は反映先も一緒に渡す'
    # "%~dp0" は末尾が \ のため、そのまま引数にすると引用符が壊れる。
    Assert-Mb ($editText -notmatch '"%~dp0"') '末尾が区切り文字のままの引数を渡さない'
    Assert-Mb ($editText -match 'Start-ManualBuilderLauncher\.ps1') 'ローカル実行版のランチャーを起動する'
    Assert-Mb (($openText -replace "`r`n", '') -notmatch "`n") 'cmdファイルはCRLFで書く'

    # --- 同じ名前のフォルダーは、確認したときだけ作り直す ---
    $blocked = $false
    try { [void](Invoke-MbHtmlExport -Project $exportProject -ProjectPath $projectPath -OutputDirectory $outputRoot) }
    catch { $blocked = ([string]$_.Exception.Message -match 'すでにあります') }
    Assert-Mb $blocked '同じ名前のフォルダーがあれば黙って上書きしない'

    $markerPath = Join-Path $result.OutputPath 'marker.txt'
    [IO.File]::WriteAllText($markerPath, 'old')
    $again = Invoke-MbHtmlExport -Project $exportProject -ProjectPath $projectPath -OutputDirectory $outputRoot -Overwrite
    Assert-Mb ([string]$again.FolderName -eq [string]$result.FolderName) '作り直しても同じフォルダー名になる'
    Assert-Mb ([bool]$again.Replaced) '作り直したことを返す'
    Assert-Mb (-not (Test-Path -LiteralPath $markerPath)) '前のフォルダーの中身を残さない'
    Assert-Mb (Test-Path -LiteralPath $again.IndexPath -PathType Leaf) '作り直したindex.htmlがある'
    Assert-Mb (@(Get-ChildItem -LiteralPath $outputRoot -Directory -Force | Where-Object { $_.Name -like '.mb-old-*' }).Count -eq 0) '差し替え用の退避フォルダーを残さない'

    # --- 共有フォルダーへの反映 ---
    $shareRoot = Join-Path $testRoot 'share'
    [void](New-Item -ItemType Directory -Path $shareRoot -Force)
    $shareTarget = Join-Path $shareRoot '出力テスト'
    $published = Copy-MbHtmlManualFolder -SourceFolder $again.OutputPath -DestinationFolder $shareTarget
    Assert-Mb (Test-Path -LiteralPath (Join-Path $shareTarget 'index.html') -PathType Leaf) '反映先へindex.htmlをコピーする'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $shareTarget '_source\project.json') -PathType Leaf) '反映先へ元データもコピーする'
    Assert-Mb (-not [bool]$published.Replaced) '初回は置き換えではない'
    Assert-Mb ([int]$published.FileCount -gt 0) '反映したファイル数を返す'

    $staleMarker = Join-Path $shareTarget 'stale.txt'
    [IO.File]::WriteAllText($staleMarker, 'stale')
    $republished = Copy-MbHtmlManualFolder -SourceFolder $again.OutputPath -DestinationFolder $shareTarget
    Assert-Mb ([bool]$republished.Replaced) '2回目は置き換えになる'
    Assert-Mb (-not (Test-Path -LiteralPath $staleMarker)) '反映先に古いファイルを残さない'
    Assert-Mb (@(Get-ChildItem -LiteralPath $shareRoot -Directory -Force | Where-Object { $_.Name -like '.mb-publish-*' -or $_.Name -like '.mb-old-*' }).Count -eq 0) 'コピー途中のフォルダーを残さない'

    $selfRejected = $false
    try { [void](Copy-MbHtmlManualFolder -SourceFolder $again.OutputPath -DestinationFolder $again.OutputPath) }
    catch { $selfRejected = ([string]$_.Exception.Message -match '同じフォルダー') }
    Assert-Mb $selfRejected '反映元と反映先が同じなら断る'

    $nestedRejected = $false
    try { [void](Copy-MbHtmlManualFolder -SourceFolder $again.OutputPath -DestinationFolder (Join-Path $again.OutputPath 'nested')) }
    catch { $nestedRejected = ([string]$_.Exception.Message -match '中へは指定できません') }
    Assert-Mb $nestedRejected '反映先を反映元の中へは指定できない'

    $missingRejected = $false
    try { [void](Copy-MbHtmlManualFolder -SourceFolder $again.OutputPath -DestinationFolder (Join-Path (Join-Path $testRoot 'no-such-share') '出力テスト')) }
    catch { $missingRejected = ([string]$_.Exception.Message -match '共有フォルダーが見つかりません') }
    Assert-Mb $missingRejected '共有フォルダーへ届かないときは分かるように断る'

    # --- 手順が無ければ出力しない ---
    $emptyProject = New-MbProject
    $emptyPath = Join-Path $testRoot 'empty.json'
    $emptyProject = Save-MbProject -Project $emptyProject -Path $emptyPath
    $rejected = $false
    try { [void](Invoke-MbHtmlExport -Project $emptyProject -ProjectPath $emptyPath -OutputDirectory $outputRoot) }
    catch { $rejected = ([string]$_.Exception.Message -match 'HTMLへ出力する手順がありません') }
    Assert-Mb $rejected '手順が1件も無ければ出力しない'

    Write-Host ''
    Write-Host 'HTML export tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
