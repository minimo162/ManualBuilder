# ManualBuilder HTML exporter.
# Excel・Wordと違いCOMを使わない。文字列を組み立ててファイルへ書くだけなので、
# Officeの有無に依存せず、既存のブック・文書へ影響することもない。
# 出力したHTMLはJavaScriptを使わない。共有フォルダー上のファイルはゾーン判定で
# スクリプトが制限されることがあるため、HTML標準の機能だけで成立させる。

Set-StrictMode -Version 2.0

# 注釈の焼き込み（New-MbAnnotatedImage）と安全な移動（Move-MbDirectorySafely）を借りる。
# サーバーが先に読み込むため今までは動いていたが、暗黙の依存はテストから使えないため明示する。
# -Force は付けない。入れ子の再読込で呼び出し元のコマンドが消えるため。
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Excel.psm1')
# 画像と動画の実ファイルの場所を解決するために使う。
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1')

# 画面表示に使う画像の基準。Excel（760px幅）より大きめにして、拡大表示にも耐えるようにする。
$script:MbHtmlImageWidth = 1100
$script:MbHtmlImageHeight = 1100

function ConvertTo-MbHtmlEscaped {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-MbHtmlContentSheets {
    param([Parameter(Mandatory = $true)][object]$Project)
    return @($Project.sheets | Where-Object { @($_.steps).Count -gt 0 })
}

function Get-MbSafeHtmlFolderName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Directory
    )
    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = ($safe.ToCharArray() | ForEach-Object { if ([int]$_ -lt 32) { '_' } else { $_ } }) -join ''
    $safe = $safe.TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'manual' }
    if ($safe -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $safe = '_' + $safe }
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    while ((Join-Path $Directory $safe).Length -gt 200 -and $safe.Length -gt 8) {
        $safe = $safe.Substring(0, $safe.Length - 8)
    }
    # 日付や連番は付けない。同じマニュアルは毎回同じフォルダー名にして、
    # 共有フォルダー側も同じ場所を上書きできるようにする。
    return $safe
}

function Get-MbHtmlExportFolderName {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$OutputDirectory
    )
    return Get-MbSafeHtmlFolderName -Name ([string]$Project.title) -Directory $OutputDirectory
}

function Get-MbHtmlStyle {
    # 外部ファイルを読まない自己完結のスタイル。印刷時は目次を隠し、手順が分断されないようにする。
    return @'
:root { color-scheme: light; }
* { box-sizing: border-box; }
body {
  margin: 0;
  color: #182033;
  background: #f4f6f9;
  font-family: "BIZ UDPGothic", "BIZ UDPゴシック", Meiryo, "Yu Gothic UI", "MS Pゴシック", sans-serif;
  font-size: 15px;
  line-height: 1.7;
}
a { color: #3a5ba0; }
.mb-header {
  padding: 22px 28px;
  color: #ffffff;
  background: #26355a;
}
.mb-header h1 { margin: 0; font-size: 24px; }
.mb-header p { margin: 6px 0 0; color: #c6cee2; font-size: 12px; }
.mb-layout { display: flex; align-items: flex-start; gap: 24px; padding: 24px; }
.mb-nav {
  position: sticky;
  top: 24px;
  width: 260px;
  flex: 0 0 260px;
  max-height: calc(100vh - 48px);
  overflow: auto;
  padding: 16px;
  background: #ffffff;
  border: 1px solid #dfe4ec;
  border-radius: 8px;
}
.mb-nav__title { margin: 0 0 8px; font-size: 12px; font-weight: 700; color: #5a6373; }
.mb-nav ol { margin: 0; padding-left: 18px; }
.mb-nav > ol { padding-left: 16px; }
.mb-nav li { margin: 4px 0; font-size: 13px; }
.mb-nav ol ol { padding-left: 14px; }
.mb-nav ol ol li { font-size: 12px; color: #5a6373; }
.mb-main { flex: 1; min-width: 0; }
.mb-sheet { margin-bottom: 32px; }
.mb-sheet > h2 {
  margin: 0 0 4px;
  padding-bottom: 8px;
  font-size: 20px;
  border-bottom: 2px solid #3a5ba0;
}
.mb-sheet__summary { margin: 8px 0 16px; color: #5a6373; font-size: 13px; }
.mb-step {
  margin: 0 0 16px;
  padding: 18px 20px;
  background: #ffffff;
  border: 1px solid #dfe4ec;
  border-radius: 8px;
}
.mb-step > h3 { display: flex; align-items: center; gap: 10px; margin: 0 0 12px; font-size: 17px; }
.mb-step__num {
  width: 28px;
  height: 28px;
  flex: 0 0 auto;
  display: inline-grid;
  place-items: center;
  color: #ffffff;
  background: #3a5ba0;
  border-radius: 50%;
  font-size: 14px;
}
.mb-step__body { display: flex; align-items: flex-start; gap: 20px; flex-wrap: wrap; }
.mb-step__visual { flex: 1 1 460px; min-width: 0; }
.mb-step__text { flex: 1 1 300px; min-width: 0; }
.mb-step__visual figure { margin: 0; }
/* 動画つきの手順は、画面では動画（再生前に焼き込み済み画像が出る）だけを見せる。
   同じ画像が二重に並ばないようにし、印刷時は逆に静止画だけを残す。 */
.mb-step--has-video .mb-step__still { display: none; }
.mb-step__visual img {
  max-width: 100%;
  height: auto;
  display: block;
  border: 1px solid #d2d6dc;
  border-radius: 4px;
}
.mb-step__desc { margin: 0; white-space: pre-wrap; }
.mb-step__note {
  margin: 12px 0 0;
  padding: 10px 12px;
  color: #4a5262;
  background: #f4f6f9;
  border-left: 3px solid #c3cad8;
  border-radius: 0 4px 4px 0;
  font-size: 13px;
  white-space: pre-wrap;
}
.mb-step__note strong { display: block; margin-bottom: 2px; font-size: 12px; color: #5a6373; }
.mb-step__novisual { margin: 0; color: #8a919e; font-size: 13px; }
.mb-video { margin-top: 12px; }
.mb-video video { width: 100%; max-width: 720px; display: block; background: #182033; border-radius: 6px; }
.mb-video__caption { margin: 6px 0 0; color: #5a6373; font-size: 12px; }
.mb-footer { padding: 16px 28px 32px; color: #8a919e; font-size: 12px; }
@media (max-width: 900px) {
  .mb-layout { display: block; padding: 16px; }
  .mb-nav { position: static; width: auto; max-height: none; margin-bottom: 16px; }
}
@media print {
  body { background: #ffffff; font-size: 11pt; }
  .mb-layout { display: block; padding: 0; }
  .mb-nav, .mb-video { display: none; }
  .mb-step--has-video .mb-step__still { display: block; }
  .mb-header { color: #182033; background: #ffffff; border-bottom: 2px solid #26355a; }
  .mb-header p { color: #5a6373; }
  .mb-step { page-break-inside: avoid; border: 1px solid #c3cad8; }
  .mb-sheet > h2 { page-break-after: avoid; }
}
'@
}

function ConvertTo-MbManualHtml {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [hashtable]$ImageNames = @{},
        [hashtable]$VideoNames = @{},
        [string]$GeneratedAt = ''
    )

    if ([string]::IsNullOrWhiteSpace($GeneratedAt)) { $GeneratedAt = (Get-Date).ToString('yyyy年M月d日 HH:mm') }
    $title = ConvertTo-MbHtmlEscaped $Project.title
    $sheets = @(Get-MbHtmlContentSheets -Project $Project)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine('<!doctype html>')
    [void]$sb.AppendLine('<html lang="ja">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine('<title>' + $title + '</title>')
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine((Get-MbHtmlStyle))
    [void]$sb.AppendLine('</style>')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')
    [void]$sb.AppendLine('<header class="mb-header"><h1>' + $title + '</h1><p>操作マニュアル ・ ' + (ConvertTo-MbHtmlEscaped $GeneratedAt) + ' 作成 ・ ManualBuilder</p></header>')
    [void]$sb.AppendLine('<div class="mb-layout">')

    # 目次。JavaScriptを使わずアンカーだけで移動する。
    [void]$sb.AppendLine('<nav class="mb-nav" aria-label="目次"><p class="mb-nav__title">目次</p><ol>')
    for ($sheetIndex = 0; $sheetIndex -lt $sheets.Count; $sheetIndex++) {
        $sheetAnchor = 'sheet-' + ($sheetIndex + 1)
        [void]$sb.Append('<li><a href="#' + $sheetAnchor + '">' + (ConvertTo-MbHtmlEscaped $sheets[$sheetIndex].name) + '</a>')
        $steps = @($sheets[$sheetIndex].steps)
        if ($steps.Count -gt 0) {
            [void]$sb.Append('<ol>')
            for ($stepIndex = 0; $stepIndex -lt $steps.Count; $stepIndex++) {
                $stepAnchor = $sheetAnchor + '-step-' + ($stepIndex + 1)
                $stepTitle = [string]$steps[$stepIndex].title
                if ([string]::IsNullOrWhiteSpace($stepTitle)) { $stepTitle = '手順名未入力' }
                [void]$sb.Append('<li><a href="#' + $stepAnchor + '">' + (ConvertTo-MbHtmlEscaped $stepTitle) + '</a></li>')
            }
            [void]$sb.Append('</ol>')
        }
        [void]$sb.AppendLine('</li>')
    }
    [void]$sb.AppendLine('</ol></nav>')

    [void]$sb.AppendLine('<main class="mb-main">')
    for ($sheetIndex = 0; $sheetIndex -lt $sheets.Count; $sheetIndex++) {
        $sheet = $sheets[$sheetIndex]
        $sheetAnchor = 'sheet-' + ($sheetIndex + 1)
        [void]$sb.AppendLine('<section class="mb-sheet" id="' + $sheetAnchor + '">')
        [void]$sb.AppendLine('<h2>' + (ConvertTo-MbHtmlEscaped $sheet.name) + '</h2>')
        if ($sheet.PSObject.Properties.Name -contains 'summary' -and -not [string]::IsNullOrWhiteSpace([string]$sheet.summary)) {
            [void]$sb.AppendLine('<p class="mb-sheet__summary">' + (ConvertTo-MbHtmlEscaped $sheet.summary) + '</p>')
        }

        $steps = @($sheet.steps)
        for ($stepIndex = 0; $stepIndex -lt $steps.Count; $stepIndex++) {
            $step = $steps[$stepIndex]
            $stepId = [string]$step.id
            $stepAnchor = $sheetAnchor + '-step-' + ($stepIndex + 1)
            $stepTitle = [string]$step.title
            if ([string]::IsNullOrWhiteSpace($stepTitle)) { $stepTitle = '手順名未入力' }
            $stepNumber = $stepIndex + 1

            $stepClass = if ($VideoNames.ContainsKey($stepId)) { 'mb-step mb-step--has-video' } else { 'mb-step' }
            [void]$sb.AppendLine('<article class="' + $stepClass + '" id="' + $stepAnchor + '">')
            [void]$sb.AppendLine('<h3><span class="mb-step__num" aria-hidden="true">' + $stepNumber + '</span>' + (ConvertTo-MbHtmlEscaped $stepTitle) + '</h3>')
            [void]$sb.AppendLine('<div class="mb-step__body">')

            $imageName = if ($ImageNames.ContainsKey($stepId)) { [string]$ImageNames[$stepId] } else { '' }
            $videoName = if ($VideoNames.ContainsKey($stepId)) { [string]$VideoNames[$stepId] } else { '' }
            [void]$sb.AppendLine('<div class="mb-step__visual">')
            if ($imageName) {
                $altText = ConvertTo-MbHtmlEscaped ("手順 $stepNumber の画面: " + $stepTitle)
                [void]$sb.AppendLine('<figure class="mb-step__still"><img src="' + (ConvertTo-MbHtmlEscaped $imageName) + '" alt="' + $altText + '" loading="lazy"></figure>')
            } else {
                [void]$sb.AppendLine('<p class="mb-step__novisual">この手順に画像はありません。</p>')
            }
            if ($videoName) {
                $posterAttribute = if ($imageName) { ' poster="' + (ConvertTo-MbHtmlEscaped $imageName) + '"' } else { '' }
                [void]$sb.AppendLine('<div class="mb-video"><video src="' + (ConvertTo-MbHtmlEscaped $videoName) + '"' + $posterAttribute + ' controls preload="metadata"></video><p class="mb-video__caption">▶ 再生ボタンで操作の動画を確認できます</p></div>')
            }
            [void]$sb.AppendLine('</div>')

            [void]$sb.AppendLine('<div class="mb-step__text">')
            if (-not [string]::IsNullOrWhiteSpace([string]$step.description)) {
                [void]$sb.AppendLine('<p class="mb-step__desc">' + (ConvertTo-MbHtmlEscaped $step.description) + '</p>')
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$step.note)) {
                [void]$sb.AppendLine('<p class="mb-step__note"><strong>補足</strong>' + (ConvertTo-MbHtmlEscaped $step.note) + '</p>')
            }
            [void]$sb.AppendLine('</div>')

            [void]$sb.AppendLine('</div></article>')
        }
        [void]$sb.AppendLine('</section>')
    }
    [void]$sb.AppendLine('</main>')
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<footer class="mb-footer">ManualBuilderで作成しました。検索はCtrl+F、印刷とPDF化はCtrl+Pが使えます。</footer>')
    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')
    return $sb.ToString()
}

function Get-MbHtmlCommandEncoding {
    # cmdは既定でANSIコードページ（日本語版は932）として読む。UTF-8で書くと日本語のメッセージが化ける。
    try { [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance) } catch { }
    try { return [Text.Encoding]::GetEncoding(932) } catch { return (New-Object Text.UTF8Encoding($false)) }
}

function Get-MbHtmlOpenCommandText {
    # 共有フォルダー上の.htmlは、組織のポリシーでInternet Explorerモードへ回されることがある。
    # Edgeを明示的に指定して開くことで、そこを避ける。見つからなければ既定のブラウザーに任せる。
    return @'
@echo off
setlocal
set "MB_INDEX=%~dp0index.html"
if not exist "%MB_INDEX%" goto :missing
set "MB_EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if exist "%MB_EDGE%" goto :edge
set "MB_EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
if exist "%MB_EDGE%" goto :edge
start "" "%MB_INDEX%"
exit /b 0
:edge
start "" "%MB_EDGE%" "%MB_INDEX%"
exit /b 0
:missing
echo index.html が見つかりません。フォルダーごとコピーしてください。
pause
exit /b 1
'@ -replace "`r?`n", "`r`n"
}

function Get-MbHtmlEditCommandText {
    # 元データ（_source）を取り込んだ状態でManualBuilderを起動する。
    # 配布先（このフォルダー）も一緒に渡し、次回は「共有フォルダーへ反映」だけで更新できるようにする。
    return @'
@echo off
setlocal
set "MB_LAUNCHER=%LOCALAPPDATA%\ManualBuilder\app\src\Start-ManualBuilderLauncher.ps1"
if exist "%MB_LAUNCHER%" goto :run
set "MB_LAUNCHER=%LOCALAPPDATA%\ManualBuilder\app.previous\src\Start-ManualBuilderLauncher.ps1"
if exist "%MB_LAUNCHER%" goto :run
echo このPCにはManualBuilderがまだ入っていません。
echo 共有フォルダーの run.cmd から一度ManualBuilderを起動してから、もう一度実行してください。
pause
exit /b 1
:run
if not exist "%~dp0_source\project.json" goto :nosource
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%MB_LAUNCHER%" -ImportFrom "%~dp0_source" -PublishTo "%~dp0."
if errorlevel 1 pause
exit /b %ERRORLEVEL%
:nosource
echo 元データ（_source フォルダー）が見つかりません。フォルダーごとコピーしてください。
pause
exit /b 1
'@ -replace "`r?`n", "`r`n"
}

function New-MbHtmlSourceFolder {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$StagingPath
    )
    # 出力フォルダーの中へ元データを同梱する。マニュアル1つ＝フォルダー1つになり、
    # 「見る人が消したら元も消える」ので、容量の管理が普通のフォルダーと同じ感覚になる。
    $sourceDirectory = Join-Path $StagingPath '_source'
    [void](New-Item -ItemType Directory -Path $sourceDirectory -Force)
    [IO.File]::Copy($ProjectPath, (Join-Path $sourceDirectory 'project.json'), $true)

    $projectDirectory = Split-Path -Parent $ProjectPath
    foreach ($pair in @(
        [pscustomobject]@{ Name = 'images'; Items = @($Project.images) },
        [pscustomobject]@{ Name = 'videos'; Items = @($Project.videos) }
    )) {
        if ($pair.Items.Count -eq 0) { continue }
        $destination = Join-Path $sourceDirectory $pair.Name
        [void](New-Item -ItemType Directory -Path $destination -Force)
        foreach ($item in $pair.Items) {
            $fileName = [string]$item.fileName
            $sourceFile = Join-Path (Join-Path $projectDirectory $pair.Name) $fileName
            if (-not (Test-Path -LiteralPath $sourceFile -PathType Leaf)) { throw "元データのファイルが見つかりません: $fileName" }
            [IO.File]::Copy($sourceFile, (Join-Path $destination $fileName), $true)
        }
    }

    # 隠し属性にして、読む人の目に触れないようにする（読み取りは妨げないのでHTMLから動画を参照できる）。
    try {
        $entry = Get-Item -LiteralPath $sourceDirectory -Force
        $entry.Attributes = $entry.Attributes -bor [IO.FileAttributes]::Hidden
    } catch { }
    return $sourceDirectory
}

function Invoke-MbHtmlExport {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [AllowEmptyString()][string]$NumberFontName = '',
        [switch]$Overwrite
    )

    $sheets = @(Get-MbHtmlContentSheets -Project $Project)
    if ($sheets.Count -eq 0) { throw 'HTMLへ出力する手順がありません。' }
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }

    $folderName = Get-MbHtmlExportFolderName -Project $Project -OutputDirectory $OutputDirectory
    $stagingPath = Join-Path $OutputDirectory ('.mb-html-' + [guid]::NewGuid().ToString('N'))
    $outputPath = Join-Path $OutputDirectory $folderName
    $replacedPath = Join-Path $OutputDirectory ('.mb-old-' + [guid]::NewGuid().ToString('N'))
    $replacedMoved = $false
    $moveCompleted = $false

    try {
        [void](New-Item -ItemType Directory -Path $stagingPath -Force)
        $imageDirectory = Join-Path $stagingPath 'images'

        $imageNames = @{}
        $videoNames = @{}
        $stepCount = 0
        $imageCount = 0
        $videoCount = 0

        foreach ($sheet in $sheets) {
            foreach ($step in @($sheet.steps)) {
                $stepCount++
                $stepId = [string]$step.id

                if (-not [string]::IsNullOrWhiteSpace([string]$step.imageId)) {
                    $sourcePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$step.imageId)
                    if (-not $sourcePath -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "画像が見つかりません: $($step.imageId)" }
                    if (-not (Test-Path -LiteralPath $imageDirectory)) { [void](New-Item -ItemType Directory -Path $imageDirectory -Force) }
                    $imageCount++
                    $destinationPath = Join-Path $imageDirectory ("step-{0:d4}.png" -f $stepCount)
                    # 注釈と切り抜きは画像へ焼き込む。Excel・Wordと同じ見た目になり、
                    # 黒塗りがHTMLのソースから読み取られることもない。
                    $renderedPath = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations @($step.annotations) -Crop $step.crop `
                        -DestinationPath $destinationPath `
                        -TargetDisplayWidth $script:MbHtmlImageWidth -TargetDisplayHeight $script:MbHtmlImageHeight `
                        -MaximumDisplayScale 1.5 -NumberFontName $NumberFontName
                    if ([string]$renderedPath -ne $destinationPath) {
                        # 注釈も切り抜きも無い手順では焼き込みが不要なため、元画像のパスがそのまま返る。
                        # 戻り値を捨てると出力フォルダーへ画像が入らず、リンク切れになる。
                        $extension = [IO.Path]::GetExtension([string]$renderedPath).ToLowerInvariant()
                        if ($extension -notin @('.png', '.jpg', '.jpeg', '.bmp')) { $extension = '.png' }
                        $destinationPath = (Join-Path $imageDirectory ("step-{0:d4}" -f $stepCount)) + $extension
                        [IO.File]::Copy([string]$renderedPath, $destinationPath, $true)
                    }
                    if (-not (Test-Path -LiteralPath $destinationPath -PathType Leaf)) { throw "画像を出力できませんでした: $($step.imageId)" }
                    $imageNames[$stepId] = 'images/' + [IO.Path]::GetFileName($destinationPath)
                }

                if ($step.PSObject.Properties.Name -contains 'videoId' -and -not [string]::IsNullOrWhiteSpace([string]$step.videoId)) {
                    $videoSourcePath = Get-MbVideoFilePath -Project $Project -ProjectPath $ProjectPath -VideoId ([string]$step.videoId)
                    if (-not $videoSourcePath -or -not (Test-Path -LiteralPath $videoSourcePath -PathType Leaf)) { throw "動画が見つかりません: $($step.videoId)" }
                    $extension = [IO.Path]::GetExtension($videoSourcePath).ToLowerInvariant()
                    if ($extension -notin @('.mp4', '.webm')) { throw "対応しない動画形式です: $extension" }
                    $videoCount++
                    # 動画は _source の中の1本だけを参照する。同じ動画をフォルダー内に二重に持たない。
                    $videoNames[$stepId] = '_source/videos/' + [IO.Path]::GetFileName($videoSourcePath)
                }
            }
        }

        [void](New-MbHtmlSourceFolder -Project $Project -ProjectPath $ProjectPath -StagingPath $stagingPath)

        $html = ConvertTo-MbManualHtml -Project $Project -ImageNames $imageNames -VideoNames $videoNames
        # BOM付きにする。file:// で開いたときに文字化けしないよう、charset指定に加えて念のため付ける。
        [IO.File]::WriteAllText((Join-Path $stagingPath 'index.html'), $html, (New-Object Text.UTF8Encoding($true)))
        $commandEncoding = Get-MbHtmlCommandEncoding
        [IO.File]::WriteAllText((Join-Path $stagingPath 'マニュアルを開く.cmd'), (Get-MbHtmlOpenCommandText), $commandEncoding)
        [IO.File]::WriteAllText((Join-Path $stagingPath '編集する.cmd'), (Get-MbHtmlEditCommandText), $commandEncoding)

        $replaced = $false
        if (Test-Path -LiteralPath $outputPath) {
            if (-not $Overwrite) { throw '同じ名前の出力フォルダーがすでにあります。' }
            # 先に新しい方を作り終えてから差し替える。差し替えに失敗しても、前のフォルダーを戻せるようにする。
            Move-MbDirectorySafely -SourcePath $outputPath -DestinationPath $replacedPath
            $replacedMoved = $true
            $replaced = $true
        }
        try {
            Move-MbDirectorySafely -SourcePath $stagingPath -DestinationPath $outputPath
            $moveCompleted = $true
        } catch {
            if ($replacedMoved) {
                try {
                    Move-MbDirectorySafely -SourcePath $replacedPath -DestinationPath $outputPath
                    $replacedMoved = $false
                } catch {
                    # 戻せなかった場合、退避先が前のフォルダーの唯一の実体になる。
                    # 消してしまわないよう場所を伝える。
                    throw ('出力フォルダーを差し替えられませんでした。前の内容は「' + (Split-Path -Leaf $replacedPath) +
                        '」という名前で保存先に残っています。名前を戻してから、もう一度実行してください。')
                }
            }
            throw
        }

        $totalBytes = [long]0
        foreach ($file in @(Get-ChildItem -LiteralPath $outputPath -Recurse -File -Force)) { $totalBytes += [long]$file.Length }
        return [pscustomobject]@{
            OutputPath  = $outputPath
            FolderName  = $folderName
            IndexPath   = Join-Path $outputPath 'index.html'
            SourcePath  = Join-Path $outputPath '_source'
            StepCount   = $stepCount
            ImageCount  = $imageCount
            VideoCount  = $videoCount
            TotalBytes  = $totalBytes
            Replaced    = $replaced
        }
    } finally {
        # 途中で失敗した場合、未完成のフォルダーを残さない。
        if (Test-Path -LiteralPath $stagingPath) {
            Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        # 差し替えを終えたときだけ、退避しておいた前のフォルダーを消す。
        # 差し替えに失敗して戻せていない場合は、前の内容がここにしか無いので消さない。
        if ($moveCompleted -and $replacedMoved -and (Test-Path -LiteralPath $replacedPath)) {
            Remove-Item -LiteralPath $replacedPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Copy-MbHtmlManualFolder {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$DestinationFolder
    )
    # 共有フォルダーへの反映。コピー途中のフォルダーを他の人に見せないよう、
    # 別名で全部コピーしてから、最後に名前の差し替えだけで切り替える。
    $source = [IO.Path]::GetFullPath($SourceFolder).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $destination = [IO.Path]::GetFullPath($DestinationFolder).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if (-not (Test-Path -LiteralPath $source -PathType Container)) { throw '反映するマニュアルのフォルダーが見つかりません。' }
    if ($source.Equals($destination, [StringComparison]::OrdinalIgnoreCase)) { throw '反映元と反映先が同じフォルダーです。' }
    if ($destination.StartsWith($source + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw '反映先を反映元の中へは指定できません。'
    }

    $parent = Split-Path -Parent $destination
    if ([string]::IsNullOrWhiteSpace($parent)) { throw '反映先の場所を確認できません。' }
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw '反映先の共有フォルダーが見つかりません。ネットワークの接続を確認してください。' }

    $stagingPath = Join-Path $parent ('.mb-publish-' + [guid]::NewGuid().ToString('N'))
    $replacedPath = Join-Path $parent ('.mb-old-' + [guid]::NewGuid().ToString('N'))
    $replacedMoved = $false
    $moveCompleted = $false
    try {
        Copy-Item -LiteralPath $source -Destination $stagingPath -Recurse -Force -ErrorAction Stop
        $replaced = $false
        if (Test-Path -LiteralPath $destination) {
            Move-MbDirectorySafely -SourcePath $destination -DestinationPath $replacedPath
            $replacedMoved = $true
            $replaced = $true
        }
        try {
            Move-MbDirectorySafely -SourcePath $stagingPath -DestinationPath $destination
            $moveCompleted = $true
        } catch {
            if ($replacedMoved) {
                try {
                    Move-MbDirectorySafely -SourcePath $replacedPath -DestinationPath $destination
                    $replacedMoved = $false
                } catch {
                    throw ('共有フォルダーを差し替えられませんでした。前の内容は「' + (Split-Path -Leaf $replacedPath) +
                        '」という名前で共有フォルダーに残っています。名前を戻してから、もう一度実行してください。')
                }
            }
            throw
        }
        $fileCount = 0
        $totalBytes = [long]0
        foreach ($file in @(Get-ChildItem -LiteralPath $destination -Recurse -File -Force)) {
            $fileCount++
            $totalBytes += [long]$file.Length
        }
        return [pscustomobject]@{
            DestinationPath = $destination
            FolderName      = Split-Path -Leaf $destination
            Replaced        = $replaced
            FileCount       = $fileCount
            TotalBytes      = $totalBytes
        }
    } finally {
        if (Test-Path -LiteralPath $stagingPath) {
            Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        # 差し替えを終えたときだけ、退避しておいた前のフォルダーを消す。
        if ($moveCompleted -and $replacedMoved -and (Test-Path -LiteralPath $replacedPath)) {
            Remove-Item -LiteralPath $replacedPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-MbManualHtml',
    'Get-MbSafeHtmlFolderName',
    'Get-MbHtmlExportFolderName',
    'Get-MbHtmlContentSheets',
    'Get-MbHtmlOpenCommandText',
    'Get-MbHtmlEditCommandText',
    'Invoke-MbHtmlExport',
    'Copy-MbHtmlManualFolder'
)
