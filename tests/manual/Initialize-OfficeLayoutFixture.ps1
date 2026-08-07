[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProjectPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
Add-Type -AssemblyName System.Drawing

function New-MbOfficeFixturePng {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][Drawing.Color]$Color,
        [ValidateRange(120, 2400)][int]$Width = 1200,
        [ValidateRange(120, 2400)][int]$Height = 675
    )

    $bitmap = New-Object Drawing.Bitmap $Width, $Height
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $fontSize = [single][Math]::Max(20, [Math]::Min(54, [Math]::Min($Width, $Height) / 8))
    $font = New-Object Drawing.Font 'Yu Gothic UI', $fontSize, ([Drawing.FontStyle]::Bold), ([Drawing.GraphicsUnit]::Pixel)
    $smallFont = New-Object Drawing.Font 'Yu Gothic UI', ([single][Math]::Max(14, $fontSize * 0.42)), ([Drawing.FontStyle]::Regular), ([Drawing.GraphicsUnit]::Pixel)
    $stream = New-Object IO.MemoryStream
    try {
        $graphics.Clear($Color)
        $graphics.DrawString($Text, $font, [Drawing.Brushes]::White, ([single]($Width * 0.05)), ([single]($Height * 0.43)))
        $graphics.DrawString(("{0} x {1}" -f $Width, $Height), $smallFont, [Drawing.Brushes]::White,
            ([single]($Width * 0.05)), ([single]($Height * 0.82)))
        $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    } finally {
        $stream.Dispose()
        $smallFont.Dispose()
        $font.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function New-MbFixtureAnnotationId {
    return 'annotation-' + [guid]::NewGuid().ToString('N')
}

function ConvertTo-MbFixtureAnnotationJson {
    param([Parameter(Mandatory = $true)][object[]]$Annotations)
    return ConvertTo-Json -InputObject @($Annotations) -Compress -Depth 4
}

$fullProjectPath = [IO.Path]::GetFullPath($ProjectPath)
if (Test-Path -LiteralPath $fullProjectPath) {
    throw "既存のプロジェクトは上書きしません。新しいパスを指定してください: $fullProjectPath"
}

$projectDirectory = Split-Path -Parent $fullProjectPath
[void](New-Item -ItemType Directory -Path $projectDirectory -Force)
$project = Get-MbProject -Path $fullProjectPath
Set-MbProjectTitle -Project $project -Title 'Office出力 4パターン検証'
$sheet = @($project.sheets)[0]
Rename-MbSheet -Project $project -SheetId $sheet.id -Name '画像配置4パターン'

# A: 操作前だけ。操作前画像へ赤枠と切り抜きを設定する。
$a = Add-MbImageStep -Project $project -ProjectPath $fullProjectPath -SheetId $sheet.id `
    -Bytes (New-MbOfficeFixturePng -Text 'A BEFORE ONLY' -Color ([Drawing.Color]::FromArgb(40, 90, 160))) -Source file
Update-MbStep -Project $project -StepId $a.Step.id -Title 'A 操作前だけ' `
    -Description '操作前画像の赤枠と切り抜きを確認します。' -Note 'ExcelとWordで同じ範囲が見えること。'
$aAnnotations = @([pscustomobject]@{
    id = New-MbFixtureAnnotationId; type = 'rect'; x1 = 0.15; y1 = 0.18; x2 = 0.55; y2 = 0.50; label = 0
})
Set-MbStepImageEdits -Project $project -StepId $a.Step.id `
    -AnnotationsJson (ConvertTo-MbFixtureAnnotationJson $aAnnotations) `
    -CropJson '{"x":0.05,"y":0.05,"width":0.9,"height":0.9}'
[void](Set-MbStepImageLayout -Project $project -StepId $a.Step.id -Layout before -Order before-after)

# B: 操作後だけ。操作後画像へ番号と切り抜きを設定する。
$b = Add-MbImageStep -Project $project -ProjectPath $fullProjectPath -SheetId $sheet.id `
    -Bytes (New-MbOfficeFixturePng -Text 'B BEFORE HIDDEN' -Color ([Drawing.Color]::FromArgb(90, 100, 110))) -Source file
Update-MbStep -Project $project -StepId $b.Step.id -Title 'B 操作後だけ' `
    -Description '操作後画像だけを表示し、番号と切り抜きを確認します。' -Note ''
[void](Set-MbStepResultImage -Project $project -ProjectPath $fullProjectPath -StepId $b.Step.id `
    -Bytes (New-MbOfficeFixturePng -Text 'B AFTER ONLY' -Color ([Drawing.Color]::FromArgb(40, 150, 90))) -Source file)
$bResultAnnotations = @([pscustomobject]@{
    id = New-MbFixtureAnnotationId; type = 'number'; x1 = 0.72; y1 = 0.25; x2 = 0.72; y2 = 0.25; label = 1
})
Set-MbStepImageEdits -Project $project -StepId $b.Step.id -Target result `
    -AnnotationsJson (ConvertTo-MbFixtureAnnotationJson $bResultAnnotations) `
    -CropJson '{"x":0.1,"y":0.1,"width":0.8,"height":0.8}'
[void](Set-MbStepImageLayout -Project $project -StepId $b.Step.id -Layout after -Order before-after)

# C: 左右比較。操作後を先にし、操作前と操作後の双方へ注釈を設定する。
$c = Add-MbImageStep -Project $project -ProjectPath $fullProjectPath -SheetId $sheet.id `
    -Bytes (New-MbOfficeFixturePng -Text 'C BEFORE SECOND' -Color ([Drawing.Color]::FromArgb(110, 70, 160))) -Source file
Update-MbStep -Project $project -StepId $c.Step.id -Title 'C 左右比較・操作後が先' `
    -Description '左右比較で操作後、操作前の順に並ぶことを確認します。' -Note '双方の注釈を確認します。'
[void](Set-MbStepResultImage -Project $project -ProjectPath $fullProjectPath -StepId $c.Step.id `
    -Bytes (New-MbOfficeFixturePng -Text 'C AFTER FIRST' -Color ([Drawing.Color]::FromArgb(220, 125, 45))) -Source file)
$cBeforeAnnotations = @([pscustomobject]@{
    id = New-MbFixtureAnnotationId; type = 'rect'; x1 = 0.20; y1 = 0.20; x2 = 0.50; y2 = 0.50; label = 0
})
$cResultAnnotations = @([pscustomobject]@{
    id = New-MbFixtureAnnotationId; type = 'number'; x1 = 0.75; y1 = 0.30; x2 = 0.75; y2 = 0.30; label = 2
})
Set-MbStepImageEdits -Project $project -StepId $c.Step.id `
    -AnnotationsJson (ConvertTo-MbFixtureAnnotationJson $cBeforeAnnotations) -CropJson ''
Set-MbStepImageEdits -Project $project -StepId $c.Step.id -Target result `
    -AnnotationsJson (ConvertTo-MbFixtureAnnotationJson $cResultAnnotations) -CropJson ''
[void](Set-MbStepImageLayout -Project $project -StepId $c.Step.id -Layout side-by-side -Order after-before)

# D: 上下比較。極端な横長と縦長を操作前、操作後の順に並べる。
$d = Add-MbImageStep -Project $project -ProjectPath $fullProjectPath -SheetId $sheet.id `
    -Bytes (New-MbOfficeFixturePng -Text 'D BEFORE WIDE' -Color ([Drawing.Color]::FromArgb(25, 130, 170)) -Width 1800 -Height 360) -Source file
Update-MbStep -Project $project -StepId $d.Step.id -Title 'D 上下比較・極端な縦横比' `
    -Description '極端な横長の操作前と縦長の操作後を上下に並べます。' -Note '過拡大や欠けがないこと。'
[void](Set-MbStepResultImage -Project $project -ProjectPath $fullProjectPath -StepId $d.Step.id `
    -Bytes (New-MbOfficeFixturePng -Text 'D AFTER PORTRAIT' -Color ([Drawing.Color]::FromArgb(190, 55, 90)) -Width 360 -Height 1800) -Source file)
$dBeforeAnnotations = @([pscustomobject]@{
    id = New-MbFixtureAnnotationId; type = 'arrow'; x1 = 0.12; y1 = 0.65; x2 = 0.82; y2 = 0.65; label = 0
})
$dResultAnnotations = @([pscustomobject]@{
    id = New-MbFixtureAnnotationId; type = 'rect'; x1 = 0.20; y1 = 0.25; x2 = 0.72; y2 = 0.70; label = 0
})
Set-MbStepImageEdits -Project $project -StepId $d.Step.id `
    -AnnotationsJson (ConvertTo-MbFixtureAnnotationJson $dBeforeAnnotations) -CropJson ''
Set-MbStepImageEdits -Project $project -StepId $d.Step.id -Target result `
    -AnnotationsJson (ConvertTo-MbFixtureAnnotationJson $dResultAnnotations) -CropJson ''
[void](Set-MbStepImageLayout -Project $project -StepId $d.Step.id -Layout stacked -Order before-after)

[void](Save-MbProject -Project $project -Path $fullProjectPath)
Write-Output $fullProjectPath
