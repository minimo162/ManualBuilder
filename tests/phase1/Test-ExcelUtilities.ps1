# Phase 1 Excel naming and annotation composition tests (Excel COM is not started).

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Excel.psm1') -Force
$excelModule = Get-Module -Name 'ManualBuilder.Excel'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-ExcelUtilityTest-' + [guid]::NewGuid().ToString('N'))

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Get-MbRedAnnotationBounds {
    param([Parameter(Mandatory = $true)][string]$Path)
    $image = [Drawing.Bitmap]::FromFile($Path)
    try {
        $minX = $image.Width
        $minY = $image.Height
        $maxX = -1
        $maxY = -1
        for ($y = 0; $y -lt $image.Height; $y++) {
            for ($x = 0; $x -lt $image.Width; $x++) {
                $pixel = $image.GetPixel($x, $y)
                if ($pixel.R -gt 170 -and $pixel.G -lt 110 -and $pixel.B -lt 100) {
                    $minX = [Math]::Min($minX, $x)
                    $minY = [Math]::Min($minY, $y)
                    $maxX = [Math]::Max($maxX, $x)
                    $maxY = [Math]::Max($maxY, $y)
                }
            }
        }
        if ($maxX -lt 0) { throw "赤い注釈が見つかりません: $Path" }
        return [pscustomobject]@{ Width = ($maxX - $minX + 1); Height = ($maxY - $minY + 1) }
    } finally {
        $image.Dispose()
    }
}

function Get-MbRedVerticalRunAtX {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$X
    )
    $image = [Drawing.Bitmap]::FromFile($Path)
    try {
        $longest = 0
        $current = 0
        for ($y = 0; $y -lt $image.Height; $y++) {
            $pixel = $image.GetPixel($X, $y)
            if ($pixel.R -gt 170 -and $pixel.G -lt 110 -and $pixel.B -lt 100) {
                $current++
                $longest = [Math]::Max($longest, $current)
            } else {
                $current = 0
            }
        }
        return [int]$longest
    } finally {
        $image.Dispose()
    }
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    $usedNames = @{}
    $first = Get-MbSafeExcelWorksheetName -RequestedName '経費/申請:国内' -UsedNames $usedNames
    $second = Get-MbSafeExcelWorksheetName -RequestedName '経費/申請:国内' -UsedNames $usedNames
    $long = Get-MbSafeExcelWorksheetName -RequestedName ('長いシート名' * 10) -UsedNames $usedNames
    Assert-Mb ($first -eq '経費・申請・国内') 'Excel禁止文字を安全なシート名へ変換する'
    Assert-Mb ($second -eq '経費・申請・国内 (2)') '重複シート名へ連番を付ける'
    Assert-Mb ($long.Length -le 31 -and (Test-MbExcelWorksheetName -Name $long)) 'シート名を31文字以内へ収める'

    $fileName = Get-MbSafeExcelFileName -Name '経費/申請:*?' -Directory $testRoot
    Assert-Mb ($fileName -eq '経費_申請___.xlsx') 'Excelファイル名の禁止文字を置換する'
    Assert-Mb ((Get-MbSafeExcelFileName -Name 'CON' -Directory $testRoot) -eq '_CON.xlsx') 'Windows予約名を安全なファイル名へ変換する'

    # --- 動画つきのフォルダー出力 ---
    Assert-Mb ((Get-MbSafeExcelFolderName -Name '経費/申請:*?' -Directory $testRoot) -eq '経費_申請___') '出力フォルダー名の禁止文字を置換する'
    Assert-Mb ((Get-MbSafeExcelFolderName -Name 'CON' -Directory $testRoot) -eq '_CON') 'Windows予約名を安全なフォルダー名へ変換する'
    [void](New-Item -ItemType Directory -Path (Join-Path $testRoot '重複マニュアル') -Force)
    Assert-Mb ((Get-MbSafeExcelFolderName -Name '重複マニュアル' -Directory $testRoot) -eq '重複マニュアル_2') '同名フォルダーがあれば連番を付ける'

    $videoProjectPath = Join-Path $testRoot 'video-project\project.json'
    $videoDirectory = Join-Path $testRoot 'video-project\videos'
    [void](New-Item -ItemType Directory -Path $videoDirectory -Force)
    $videoAName = 'video-' + ('a' * 32) + '.mp4'
    $videoBName = 'video-' + ('b' * 32) + '.webm'
    [IO.File]::WriteAllBytes((Join-Path $videoDirectory $videoAName), ([byte[]](1, 2, 3)))
    [IO.File]::WriteAllBytes((Join-Path $videoDirectory $videoBName), ([byte[]](4, 5, 6)))
    $videoProject = [pscustomobject]@{
        videos = @(
            [pscustomobject]@{ id = 'video-a'; fileName = $videoAName },
            [pscustomobject]@{ id = 'video-b'; fileName = $videoBName }
        )
        sheets = @(
            [pscustomobject]@{ steps = @(
                [pscustomobject]@{ id = 'step-1'; videoId = 'video-a' },
                [pscustomobject]@{ id = 'step-2'; videoId = '' },
                [pscustomobject]@{ id = 'step-3'; videoId = 'video-a' }
            ) },
            [pscustomobject]@{ steps = @(
                [pscustomobject]@{ id = 'step-4'; videoId = 'video-b' }
            ) }
        )
    }
    $plan = Get-MbExcelVideoPlan -Project $videoProject -ProjectPath $videoProjectPath
    Assert-Mb ([int]$plan.Count -eq 2) '同じ動画を複数の手順へ付けてもファイルは1本にする'
    Assert-Mb ($plan.StepLinks['step-1'] -eq '動画\動画001.mp4') '動画つきの手順へ相対パスのリンクを作る'
    Assert-Mb ($plan.StepLinks['step-3'] -eq '動画\動画001.mp4') '同じ動画の手順は同じファイルを指す'
    Assert-Mb ($plan.StepLinks['step-4'] -eq '動画\動画002.webm') '拡張子は元の動画に合わせる'
    Assert-Mb (-not $plan.StepLinks.ContainsKey('step-2')) '動画の無い手順にはリンクを作らない'
    # 相対パスでなければ、フォルダーごとコピーしたときにリンクが切れる。
    Assert-Mb (@($plan.StepLinks.Values | Where-Object { $_ -match '^[A-Za-z]:\\|^\\\\' }).Count -eq 0) 'リンクへ絶対パスを使わない'

    $noVideoProject = [pscustomobject]@{
        videos = @()
        sheets = @([pscustomobject]@{ steps = @([pscustomobject]@{ id = 'step-1'; videoId = '' }) })
    }
    Assert-Mb ([int](Get-MbExcelVideoPlan -Project $noVideoProject -ProjectPath $videoProjectPath).Count -eq 0) '動画が無ければフォルダー出力にしない'

    $missingVideoProject = [pscustomobject]@{
        videos = @([pscustomobject]@{ id = 'video-c'; fileName = 'video-' + ('c' * 32) + '.mp4' })
        sheets = @([pscustomobject]@{ steps = @([pscustomobject]@{ id = 'step-1'; videoId = 'video-c' }) })
    }
    $missingRejected = $false
    try { [void](Get-MbExcelVideoPlan -Project $missingVideoProject -ProjectPath $videoProjectPath) }
    catch { $missingRejected = ([string]$_.Exception.Message -match '動画ファイルが見つかりません') }
    Assert-Mb $missingRejected '動画ファイルが無ければ出力しない'

    $wideLayout = Get-MbExcelStepCardLayout -Description '短い説明' -ImageWidth 1920 -ImageHeight 500
    $screenLayout = Get-MbExcelStepCardLayout -Description '短い説明' -ImageWidth 1920 -ImageHeight 1080
    $portraitLayout = Get-MbExcelStepCardLayout -Description '短い説明' -ImageWidth 1080 -ImageHeight 1920
    $extremePortraitLayout = Get-MbExcelStepCardLayout -Description '短い説明' -ImageWidth 500 -ImageHeight 2000
    $longDescriptionText = '長い説明です。' * 300
    $longDescriptionLayout = Get-MbExcelStepCardLayout -Description $longDescriptionText -ImageWidth 1920 -ImageHeight 1080
    $longDescriptionLines = & $excelModule { param($Value) Get-MbExcelTextLineEstimate -Text $Value -CharactersPerLine 68 } $longDescriptionText
    $noteText = '補足事項です。' * 20
    $noteLayout = Get-MbExcelStepCardLayout -Description '短い説明' -Note $noteText -ImageWidth 1920 -ImageHeight 1080
    $noteLines = & $excelModule { param($Value) Get-MbExcelTextLineEstimate -Text $Value -CharactersPerLine 72 } $noteText
    Assert-Mb ($wideLayout.ImageRows -eq 7 -and $wideLayout.ContentRows -eq 7) '横長画像のカード高さを実表示寸法へ縮める'
    Assert-Mb ($screenLayout.ImageRows -eq 14 -and $screenLayout.ContentRows -eq 14) '16対9画像を実表示寸法に必要な高さへ収める'
    $officeScreenLayout = Get-MbExcelStepCardLayout -Description '短い説明' -ImageWidth 1920 -ImageHeight 1200
    Assert-Mb ($officeScreenLayout.ImageRows -eq 15 -and $officeScreenLayout.ContentRows -eq 15) '16対10画像を実表示寸法に必要な高さへ収める'
    Assert-Mb ($portraitLayout.ImageRows -eq 26 -and $portraitLayout.ContentRows -eq 26) '縦長画像のカード高さを上限まで広げる'
    Assert-Mb ($extremePortraitLayout.ImageRows -eq 22 -and $extremePortraitLayout.ContentRows -eq 22) '極端な縦長画像のカード高さを22行へ抑える'
    Assert-Mb ($longDescriptionLayout.ContentRows -gt 14) '長い説明に必要な本文行を確保する'
    Assert-Mb ($longDescriptionLayout.DescriptionBodyRows -lt ($longDescriptionLines + 1)) '長文を26pt行へ詰め直して過剰な下余白を抑える'
    Assert-Mb (($longDescriptionLayout.DescriptionBodyRows * 26) -ge (($longDescriptionLines * 16) + 4)) '長文の表示に必要な物理高さを確保する'
    $previousDescriptionRows = [int][Math]::Ceiling((($longDescriptionLines * 17.0) + 8.0) / 26.0)
    Assert-Mb ($longDescriptionLayout.DescriptionBodyRows -le $previousDescriptionRows) '長文カードの旧安全余裕を詰める'
    Assert-Mb ($noteLayout.HasNote -and $noteLayout.NoteBodyRows -ge 2) '補足の長さに応じた行数を確保する'
    Assert-Mb (($noteLayout.NoteBodyRows * 26) -ge (($noteLines * 15) + 4)) '補足の表示に必要な物理高さを確保する'

    $mixedWidthText = ('商品MAZDA CX-5 2026 / ' * 20)
    $mixedWidthLines = & $excelModule { param($Value) Get-MbExcelTextLineEstimate -Text $Value -CharactersPerLine 68 } $mixedWidthText
    $fullWidthOnlyLines = & $excelModule { param($Value) Get-MbExcelTextLineEstimate -Text $Value -CharactersPerLine 68 } ('商' * $mixedWidthText.Length)
    Assert-Mb ($mixedWidthLines -lt $fullWidthOnlyLines) '半角英数字を実表示幅に近い文字数として見積もる'
    Assert-Mb ($screenLayout.NextRowOffset -eq ($screenLayout.ContentRows + 2)) '次のカード開始行を可変高さから計算する'

    $pageBreakRows = @(Get-MbExcelPageBreakRows -Cards @(
        [pscustomobject]@{ StartRow = 3; Height = 398.0 },
        [pscustomobject]@{ StartRow = 19; Height = 404.0 },
        [pscustomobject]@{ StartRow = 35; Height = 242.0 },
        [pscustomobject]@{ StartRow = 45; Height = 500.0 }
    ))
    Assert-Mb (($pageBreakRows -join ',') -eq '19,45') '画像カードを分割せず収まる手順だけ同じ印刷ページへまとめる'

    $normalRenderTarget = Get-MbExcelAnnotationRenderTarget -ImageWidth 1920 -ImageHeight 1080 -Crop $null
    $wideRenderTarget = Get-MbExcelAnnotationRenderTarget -ImageWidth 1920 -ImageHeight 500 -Crop $null
    $portraitRenderTarget = Get-MbExcelAnnotationRenderTarget -ImageWidth 500 -ImageHeight 2000 -Crop $null
    $croppedWideRenderTarget = Get-MbExcelAnnotationRenderTarget -ImageWidth 1920 -ImageHeight 1080 `
        -Crop ([pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 0.2 })
    Assert-Mb ($normalRenderTarget.Width -eq 760 -and $normalRenderTarget.Height -eq 880) '通常画像の注釈表示基準を維持する'
    Assert-Mb ($wideRenderTarget.Width -eq 646 -and $wideRenderTarget.Height -eq 880) '極端な横長画像の注釈表示基準をExcel配置へ合わせる'
    Assert-Mb ($portraitRenderTarget.Width -eq 760 -and $portraitRenderTarget.Height -eq 620) '極端な縦長画像の注釈表示基準をExcel配置へ合わせる'
    Assert-Mb ($croppedWideRenderTarget.Width -eq 646) '切り抜き後の比率で注釈表示基準を決める'

    $statusPath = Join-Path $testRoot 'status.json'
    $status = [pscustomobject]@{ state = 'queued'; message = '開始'; updatedAt = '' }
    & $excelModule { param($Path, $Value) Write-MbExcelStatus -StatusPath $Path -Status $Value } $statusPath $status
    $status.state = 'running'
    $status.message = '更新'
    & $excelModule { param($Path, $Value) Write-MbExcelStatus -StatusPath $Path -Status $Value } $statusPath $status
    $savedStatus = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$savedStatus.state -eq 'running') '同じExcel進捗JSONをWindows互換方式で連続更新する'
    $statusTemporaryFiles = @(Get-ChildItem -LiteralPath $testRoot -Filter '.status-*.tmp' -File -ErrorAction SilentlyContinue)
    Assert-Mb ($statusTemporaryFiles.Count -eq 0) '進捗JSON更新後に一時・バックアップファイルを残さない'

    Add-Type -AssemblyName System.Drawing
    $sourcePath = Join-Path $testRoot 'source.png'
    $renderedPath = Join-Path $testRoot 'rendered.png'
    $croppedPath = Join-Path $testRoot 'cropped.png'
    $bitmap = New-Object Drawing.Bitmap 320, 200
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::White)
        $bitmap.Save($sourcePath, [Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
    $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    $annotations = @(
        [pscustomobject]@{ type = 'rect'; x1 = 0.1; y1 = 0.1; x2 = 0.5; y2 = 0.5; label = 0 },
        [pscustomobject]@{ type = 'arrow'; x1 = 0.2; y1 = 0.8; x2 = 0.55; y2 = 0.55; label = 0 },
        [pscustomobject]@{ type = 'number'; x1 = 0.25; y1 = 0.25; x2 = 0.25; y2 = 0.25; label = 1 },
        [pscustomobject]@{ type = 'blackout'; x1 = 0.7; y1 = 0.6; x2 = 0.9; y2 = 0.8; label = 0 }
    )
    $resultPath = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations $annotations -DestinationPath $renderedPath
    Assert-Mb ($resultPath -eq $renderedPath -and (Test-Path -LiteralPath $renderedPath -PathType Leaf)) '4種類の注釈を出力画像へ合成する'
    Assert-Mb ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -eq $sourceHash) '注釈合成後も元画像を変更しない'
    $nullAnnotationResult = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations $null -DestinationPath (Join-Path $testRoot 'unused-null-annotations.png')
    Assert-Mb ($nullAnnotationResult -eq $sourcePath) '操作後注釈がnullでも元画像をそのまま利用する'

    $rendered = [Drawing.Bitmap]::FromFile($renderedPath)
    try {
        $blackPixel = $rendered.GetPixel(256, 140)
        Assert-Mb ($blackPixel.R -lt 40 -and $blackPixel.G -lt 50 -and $blackPixel.B -lt 60) '黒塗り注釈を画像へ反映する'
        Assert-Mb ($rendered.Width -eq 320 -and $rendered.Height -eq 200) '注釈合成後も画像寸法を維持する'
    } finally {
        $rendered.Dispose()
    }

    $crop = [pscustomobject]@{ x = 0.25; y = 0.25; width = 0.5; height = 0.5 }
    $cropResult = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations @() -Crop $crop -DestinationPath $croppedPath
    Assert-Mb ($cropResult -eq $croppedPath -and (Test-Path -LiteralPath $croppedPath -PathType Leaf)) '切り抜いた出力用画像を作成する'
    $cropped = [Drawing.Bitmap]::FromFile($croppedPath)
    try {
        Assert-Mb ($cropped.Width -eq 160 -and $cropped.Height -eq 100) '正規化した切り抜き範囲を画像寸法へ反映する'
    } finally {
        $cropped.Dispose()
    }
    Assert-Mb ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -eq $sourceHash) '切り抜き後も元画像を変更しない'

    $scaleSourcePath = Join-Path $testRoot 'annotation-scale-source.png'
    $fullAnnotationPath = Join-Path $testRoot 'annotation-full.png'
    $wideAnnotationPath = Join-Path $testRoot 'annotation-wide-crop.png'
    $portraitAnnotationPath = Join-Path $testRoot 'annotation-portrait-crop.png'
    $rectAnnotationPath = Join-Path $testRoot 'annotation-rect.png'
    $arrowAnnotationPath = Join-Path $testRoot 'annotation-arrow.png'
    $scaleBitmap = New-Object Drawing.Bitmap 1200, 800
    $scaleGraphics = [Drawing.Graphics]::FromImage($scaleBitmap)
    try {
        $scaleGraphics.Clear([Drawing.Color]::White)
        $scaleBitmap.Save($scaleSourcePath, [Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $scaleGraphics.Dispose()
        $scaleBitmap.Dispose()
    }
    $numberAnnotation = @([pscustomobject]@{ type = 'number'; x1 = 0.5; y1 = 0.5; x2 = 0.5; y2 = 0.5; label = 2 })
    [void](New-MbAnnotatedImage -SourcePath $scaleSourcePath -Annotations $numberAnnotation `
        -DestinationPath $fullAnnotationPath -TargetDisplayWidth 760 -TargetDisplayHeight 880)
    $wideCrop = [pscustomobject]@{ x = 0.25; y = 0.45; width = 0.5; height = 0.1 }
    [void](New-MbAnnotatedImage -SourcePath $scaleSourcePath -Annotations $numberAnnotation -Crop $wideCrop `
        -DestinationPath $wideAnnotationPath -TargetDisplayWidth 646 -TargetDisplayHeight 880 -MaximumDisplayScale 1.5)
    $portraitCrop = [pscustomobject]@{ x = 0.45; y = 0.1; width = 0.1; height = 0.8 }
    [void](New-MbAnnotatedImage -SourcePath $scaleSourcePath -Annotations $numberAnnotation -Crop $portraitCrop `
        -DestinationPath $portraitAnnotationPath -TargetDisplayWidth 760 -TargetDisplayHeight 620 -MaximumDisplayScale 1.5)
    $rectAnnotation = @([pscustomobject]@{ type = 'rect'; x1 = 0.2; y1 = 0.2; x2 = 0.8; y2 = 0.7; label = 0 })
    [void](New-MbAnnotatedImage -SourcePath $scaleSourcePath -Annotations $rectAnnotation `
        -DestinationPath $rectAnnotationPath -TargetDisplayWidth 760 -TargetDisplayHeight 880)
    $arrowAnnotation = @([pscustomobject]@{ type = 'arrow'; x1 = 0.2; y1 = 0.5; x2 = 0.8; y2 = 0.5; label = 0 })
    [void](New-MbAnnotatedImage -SourcePath $scaleSourcePath -Annotations $arrowAnnotation `
        -DestinationPath $arrowAnnotationPath -TargetDisplayWidth 760 -TargetDisplayHeight 880)
    $fullBounds = Get-MbRedAnnotationBounds -Path $fullAnnotationPath
    $wideBounds = Get-MbRedAnnotationBounds -Path $wideAnnotationPath
    $portraitBounds = Get-MbRedAnnotationBounds -Path $portraitAnnotationPath
    $fullDisplayScale = [Math]::Min(2.0, [Math]::Min(760.0 / 1200.0, 880.0 / 800.0))
    $wideDisplayScale = [Math]::Min(1.5, [Math]::Min(646.0 / 600.0, 880.0 / 80.0))
    $portraitDisplayScale = [Math]::Min(1.5, [Math]::Min(760.0 / 120.0, 620.0 / 640.0))
    $fullDisplayedDiameter = [double]([Math]::Max($fullBounds.Width, $fullBounds.Height) * $fullDisplayScale)
    $wideDisplayedDiameter = [double]([Math]::Max($wideBounds.Width, $wideBounds.Height) * $wideDisplayScale)
    $portraitDisplayedDiameter = [double]([Math]::Max($portraitBounds.Width, $portraitBounds.Height) * $portraitDisplayScale)
    $fullAnnotationUnit = 0.62
    $wideAnnotationUnit = 0.62
    $portraitAnnotationUnit = 0.62
    Assert-Mb ([Math]::Abs($fullDisplayedDiameter - (48 * $fullAnnotationUnit)) -le 3.0) '全画面画像の番号注釈を共通表示寸法へ合わせる'
    Assert-Mb ([Math]::Abs($wideDisplayedDiameter - (48 * $wideAnnotationUnit)) -le 3.0) '横長画像の番号注釈を共通表示寸法へ合わせる'
    Assert-Mb ($wideDisplayScale -le 1.5) 'Excelの拡大上限1.5倍でも注釈寸法を同期する'
    Assert-Mb ($wideDisplayScale -lt (760.0 / 600.0)) '細長い画像を通常幅の85%へ抑えて注釈倍率を同期する'
    Assert-Mb ([Math]::Abs($portraitDisplayedDiameter - (48 * $portraitAnnotationUnit)) -le 3.0) '縦長画像の番号注釈を共通表示寸法へ合わせる'
    Assert-Mb ($portraitDisplayScale -lt (880.0 / 640.0)) '極端な縦長画像を通常高の85%へ抑えて注釈倍率を同期する'
    Assert-Mb ([Math]::Abs($wideDisplayedDiameter - $fullDisplayedDiameter) -le 3.0 -and [Math]::Abs($portraitDisplayedDiameter - $fullDisplayedDiameter) -le 3.0) '画像比率が異なっても注釈の視認サイズを揃える'
    $rectDisplayedThickness = [double]((Get-MbRedVerticalRunAtX -Path $rectAnnotationPath -X 600) * $fullDisplayScale)
    $arrowDisplayedThickness = [double]((Get-MbRedVerticalRunAtX -Path $arrowAnnotationPath -X 480) * $fullDisplayScale)
    Assert-Mb ([Math]::Abs($rectDisplayedThickness - (7 * $fullAnnotationUnit)) -le 1.5) '赤枠を共通表示寸法の太さへ合わせる'
    Assert-Mb ([Math]::Abs($arrowDisplayedThickness - (8 * $fullAnnotationUnit)) -le 1.5) '赤矢印を共通表示寸法の太さへ合わせる'

    $comparisonPath = Join-Path $testRoot 'before-after.png'
    [void](New-MbBeforeAfterImage -BeforePath $sourcePath -AfterPath $croppedPath -DestinationPath $comparisonPath)
    Assert-Mb (Test-Path -LiteralPath $comparisonPath -PathType Leaf) 'Excel用に操作前と操作後を1枚へまとめる'
    $comparison = [Drawing.Image]::FromFile($comparisonPath)
    try {
        Assert-Mb ($comparison.Height -gt $comparison.Width) '比較画像で操作前と操作後を縦に並べる'
    } finally { $comparison.Dispose() }

    $horizontalComparisonPath = Join-Path $testRoot 'before-after-horizontal.png'
    [void](New-MbBeforeAfterImage -BeforePath $sourcePath -AfterPath $croppedPath -DestinationPath $horizontalComparisonPath -Orientation horizontal -Order after-before)
    $horizontalComparison = [Drawing.Image]::FromFile($horizontalComparisonPath)
    try {
        Assert-Mb ($horizontalComparison.Width -gt $horizontalComparison.Height) '比較画像で操作前と操作後を左右に並べる'
    } finally { $horizontalComparison.Dispose() }

    Write-Host ''
    Write-Host 'Excel utility tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
