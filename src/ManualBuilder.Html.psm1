# ManualBuilder HTML exporter.
# Excel・Word・PowerPointと違いCOMを使わない。文字列を組み立ててファイルへ書くだけなので、
# Officeの有無に依存せず、既存のブック・文書へ影響することもない。
# 出力したHTMLはJavaScriptを使わない。共有フォルダー上のファイルはゾーン判定で
# スクリプトが制限されることがあるため、HTML標準の機能だけで成立させる。

Set-StrictMode -Version 2.0

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
    $candidate = $safe
    $suffix = 2
    while (Test-Path -LiteralPath (Join-Path $Directory $candidate)) {
        $candidate = "${safe}_$suffix"
        $suffix++
    }
    return $candidate
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

function Invoke-MbHtmlExport {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [AllowEmptyString()][string]$NumberFontName = ''
    )

    $sheets = @(Get-MbHtmlContentSheets -Project $Project)
    if ($sheets.Count -eq 0) { throw 'HTMLへ出力する手順がありません。' }
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }

    $folderName = Get-MbSafeHtmlFolderName -Name (([string]$Project.title) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss')) -Directory $OutputDirectory
    $stagingPath = Join-Path $OutputDirectory ('.mb-html-' + [guid]::NewGuid().ToString('N'))
    $outputPath = Join-Path $OutputDirectory $folderName

    try {
        [void](New-Item -ItemType Directory -Path $stagingPath -Force)
        $imageDirectory = Join-Path $stagingPath 'images'
        $videoDirectory = Join-Path $stagingPath 'videos'

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
                    $fileName = "step-{0:d4}.png" -f $stepCount
                    # 注釈と切り抜きは画像へ焼き込む。Excel・Word・PowerPointと同じ見た目になり、
                    # 黒塗りがHTMLのソースから読み取られることもない。
                    [void](New-MbAnnotatedImage -SourcePath $sourcePath -Annotations @($step.annotations) -Crop $step.crop `
                        -DestinationPath (Join-Path $imageDirectory $fileName) `
                        -TargetDisplayWidth $script:MbHtmlImageWidth -TargetDisplayHeight $script:MbHtmlImageHeight `
                        -MaximumDisplayScale 1.5 -NumberFontName $NumberFontName)
                    $imageNames[$stepId] = 'images/' + $fileName
                }

                if ($step.PSObject.Properties.Name -contains 'videoId' -and -not [string]::IsNullOrWhiteSpace([string]$step.videoId)) {
                    $videoSourcePath = Get-MbVideoFilePath -Project $Project -ProjectPath $ProjectPath -VideoId ([string]$step.videoId)
                    if (-not $videoSourcePath -or -not (Test-Path -LiteralPath $videoSourcePath -PathType Leaf)) { throw "動画が見つかりません: $($step.videoId)" }
                    if (-not (Test-Path -LiteralPath $videoDirectory)) { [void](New-Item -ItemType Directory -Path $videoDirectory -Force) }
                    $videoCount++
                    $extension = [IO.Path]::GetExtension($videoSourcePath).ToLowerInvariant()
                    if ($extension -notin @('.mp4', '.webm')) { throw "対応しない動画形式です: $extension" }
                    $videoFileName = ("step-{0:d4}" -f $stepCount) + $extension
                    [IO.File]::Copy($videoSourcePath, (Join-Path $videoDirectory $videoFileName), $true)
                    $videoNames[$stepId] = 'videos/' + $videoFileName
                }
            }
        }

        $html = ConvertTo-MbManualHtml -Project $Project -ImageNames $imageNames -VideoNames $videoNames
        # BOM付きにする。file:// で開いたときに文字化けしないよう、charset指定に加えて念のため付ける。
        [IO.File]::WriteAllText((Join-Path $stagingPath 'index.html'), $html, (New-Object Text.UTF8Encoding($true)))

        if (Test-Path -LiteralPath $outputPath) { throw '同じ名前の出力フォルダーがすでにあります。' }
        [IO.Directory]::Move($stagingPath, $outputPath)

        $totalBytes = [long]0
        foreach ($file in @(Get-ChildItem -LiteralPath $outputPath -Recurse -File)) { $totalBytes += [long]$file.Length }
        return [pscustomobject]@{
            OutputPath = $outputPath
            FolderName = $folderName
            IndexPath  = Join-Path $outputPath 'index.html'
            StepCount  = $stepCount
            ImageCount = $imageCount
            VideoCount = $videoCount
            TotalBytes = $totalBytes
        }
    } finally {
        # 途中で失敗した場合、未完成のフォルダーを残さない。
        if (Test-Path -LiteralPath $stagingPath) {
            Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-MbManualHtml',
    'Get-MbSafeHtmlFolderName',
    'Get-MbHtmlContentSheets',
    'Invoke-MbHtmlExport'
)
