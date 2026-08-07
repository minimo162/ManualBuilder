# Office出力の4配置fixtureと、Excel/Wordが共有する画像前処理を検証する。
# Excel/Word COMは起動しない。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$fixtureScript = Join-Path $repoRoot 'tests\manual\Initialize-OfficeLayoutFixture.ps1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-OfficeLayoutFixtureTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'
$renderRoot = Join-Path $testRoot 'rendered'
Add-Type -AssemblyName System.Drawing

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Get-MbProcessIds {
    param([Parameter(Mandatory = $true)][string]$Name)
    return @(
        Get-Process -Name $Name -ErrorAction SilentlyContinue |
            Sort-Object -Property Id |
            ForEach-Object { [int]$_.Id }
    )
}

function Assert-MbPixelNear {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$X,
        [Parameter(Mandatory = $true)][int]$Y,
        [Parameter(Mandatory = $true)][Drawing.Color]$Expected,
        [int]$Tolerance = 18,
        [Parameter(Mandatory = $true)][string]$Message
    )
    $bitmap = [Drawing.Bitmap]::FromFile($Path)
    try {
        $safeX = [Math]::Max(0, [Math]::Min($bitmap.Width - 1, $X))
        $safeY = [Math]::Max(0, [Math]::Min($bitmap.Height - 1, $Y))
        $actual = $bitmap.GetPixel($safeX, $safeY)
        $matches = ([Math]::Abs([int]$actual.R - [int]$Expected.R) -le $Tolerance -and
            [Math]::Abs([int]$actual.G - [int]$Expected.G) -le $Tolerance -and
            [Math]::Abs([int]$actual.B - [int]$Expected.B) -le $Tolerance)
        Assert-Mb $matches ($Message + " (actual=$($actual.R),$($actual.G),$($actual.B))")
    } finally {
        $bitmap.Dispose()
    }
}

function Test-MbImageHasRedAnnotation {
    param([Parameter(Mandatory = $true)][string]$Path)
    $bitmap = [Drawing.Bitmap]::FromFile($Path)
    try {
        $step = [Math]::Max(1, [int]([Math]::Min($bitmap.Width, $bitmap.Height) / 220))
        for ($y = 0; $y -lt $bitmap.Height; $y += $step) {
            for ($x = 0; $x -lt $bitmap.Width; $x += $step) {
                $pixel = $bitmap.GetPixel($x, $y)
                if ($pixel.R -ge 180 -and $pixel.G -le 105 -and $pixel.B -le 105) { return $true }
            }
        }
        return $false
    } finally {
        $bitmap.Dispose()
    }
}

$excelBefore = Get-MbProcessIds -Name EXCEL
$wordBefore = Get-MbProcessIds -Name WINWORD

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    $createdPath = & $fixtureScript -ProjectPath $projectPath
    Assert-Mb ([IO.Path]::GetFullPath([string]$createdPath) -eq [IO.Path]::GetFullPath($projectPath)) 'Office配置fixtureを指定パスへ作成する'
    $existingProjectRejected = $false
    try { [void](& $fixtureScript -ProjectPath $projectPath) } catch { $existingProjectRejected = $true }
    Assert-Mb $existingProjectRejected 'fixture生成で既存プロジェクトを上書きしない'

    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Excel.psm1') -Force
    $project = Get-MbProject -Path $projectPath
    $steps = @($project.sheets[0].steps)

    Assert-Mb ($steps.Count -eq 4) '最小4手順だけをfixtureへ作成する'
    Assert-Mb ((@($steps | ForEach-Object { [string]$_.title }) -join '|') -eq
        'A 操作前だけ|B 操作後だけ|C 左右比較・操作後が先|D 上下比較・極端な縦横比') 'AからDの順序を固定する'
    Assert-Mb (@($project.images).Count -eq 7) '操作前4枚と操作後3枚を別画像として保持する'

    $a, $b, $c, $d = $steps
    Assert-Mb ([string]$a.imageLayout -eq 'before' -and [string]$a.imageOrder -eq 'before-after' -and
        @($a.annotations).Count -eq 1 -and [double]$a.crop.width -eq 0.9) 'Aへ操作前赤枠と切り抜きを設定する'
    Assert-Mb ([string]$b.imageLayout -eq 'after' -and -not [string]::IsNullOrWhiteSpace([string]$b.resultImageId) -and
        @($b.resultAnnotations).Count -eq 1 -and [string]$b.resultAnnotations[0].type -eq 'number' -and
        [double]$b.resultCrop.width -eq 0.8) 'Bへ操作後だけ・番号・切り抜きを設定する'
    Assert-Mb ([string]$c.imageLayout -eq 'side-by-side' -and [string]$c.imageOrder -eq 'after-before' -and
        @($c.annotations).Count -eq 1 -and @($c.resultAnnotations).Count -eq 1) 'Cへ左右・操作後先・前後双方の注釈を設定する'
    Assert-Mb ([string]$d.imageLayout -eq 'stacked' -and [string]$d.imageOrder -eq 'before-after') 'Dへ上下・操作前先を設定する'
    $dBeforeMeta = @($project.images | Where-Object { $_.id -eq $d.imageId })[0]
    $dAfterMeta = @($project.images | Where-Object { $_.id -eq $d.resultImageId })[0]
    Assert-Mb ([int]$dBeforeMeta.width -eq 1800 -and [int]$dBeforeMeta.height -eq 360 -and
        [int]$dAfterMeta.width -eq 360 -and [int]$dAfterMeta.height -eq 1800) 'Dへ極端な横長と縦長を設定する'

    [void](New-Item -ItemType Directory -Path $renderRoot -Force)
    function Render-MbFixtureSide {
        param(
            [Parameter(Mandatory = $true)][object]$Step,
            [ValidateSet('before', 'result')][string]$Target,
            [Parameter(Mandatory = $true)][string]$DestinationPath
        )
        $isResult = $Target -eq 'result'
        $imageId = if ($isResult) { [string]$Step.resultImageId } else { [string]$Step.imageId }
        $annotations = if ($isResult) { @($Step.resultAnnotations) } else { @($Step.annotations) }
        $crop = if ($isResult) { $Step.resultCrop } else { $Step.crop }
        $meta = @($project.images | Where-Object { $_.id -eq $imageId })[0]
        $sourcePath = Get-MbImageFilePath -Project $project -ProjectPath $projectPath -ImageId $imageId
        $targetSize = Get-MbExcelAnnotationRenderTarget -ImageWidth ([int]$meta.width) -ImageHeight ([int]$meta.height) -Crop $crop
        return New-MbAnnotatedImage -SourcePath $sourcePath -Annotations $annotations -Crop $crop `
            -DestinationPath $DestinationPath -TargetDisplayWidth $targetSize.Width -TargetDisplayHeight $targetSize.Height
    }

    $aOutput = Render-MbFixtureSide -Step $a -Target before -DestinationPath (Join-Path $renderRoot 'A-before-only.png')
    $bOutput = Render-MbFixtureSide -Step $b -Target result -DestinationPath (Join-Path $renderRoot 'B-after-only.png')
    $cBefore = Render-MbFixtureSide -Step $c -Target before -DestinationPath (Join-Path $renderRoot 'C-before.png')
    $cAfter = Render-MbFixtureSide -Step $c -Target result -DestinationPath (Join-Path $renderRoot 'C-after.png')
    $cOutput = New-MbBeforeAfterImage -BeforePath $cBefore -AfterPath $cAfter `
        -DestinationPath (Join-Path $renderRoot 'C-side-by-side.png') -Orientation horizontal -Order after-before
    $dBefore = Render-MbFixtureSide -Step $d -Target before -DestinationPath (Join-Path $renderRoot 'D-before.png')
    $dAfter = Render-MbFixtureSide -Step $d -Target result -DestinationPath (Join-Path $renderRoot 'D-after.png')
    $dOutput = New-MbBeforeAfterImage -BeforePath $dBefore -AfterPath $dAfter `
        -DestinationPath (Join-Path $renderRoot 'D-stacked.png') -Orientation vertical -Order before-after

    $officeImages = @($aOutput, $bOutput, $cOutput, $dOutput)
    Assert-Mb (@($officeImages | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }).Count -eq 4) 'Officeへ渡す最終画像を4手順分作成する'
    Assert-Mb ((Test-MbImageHasRedAnnotation $aOutput) -and (Test-MbImageHasRedAnnotation $bOutput) -and
        (Test-MbImageHasRedAnnotation $cOutput) -and (Test-MbImageHasRedAnnotation $dOutput)) '4手順の最終画像へ注釈を焼き込む'

    $aBitmap = [Drawing.Image]::FromFile($aOutput)
    try { Assert-Mb ($aBitmap.Width -lt 1200 -and $aBitmap.Height -lt 675) 'Aの切り抜きを最終画像寸法へ反映する' }
    finally { $aBitmap.Dispose() }
    $bBitmap = [Drawing.Image]::FromFile($bOutput)
    try { Assert-Mb ($bBitmap.Width -eq 960 -and $bBitmap.Height -eq 540) 'Bは隠れた操作前でなく切り抜いた操作後だけを使う' }
    finally { $bBitmap.Dispose() }

    Assert-MbPixelNear -Path $bOutput -X 20 -Y 20 -Expected ([Drawing.Color]::FromArgb(40, 150, 90)) `
        -Message 'Bの最終画像を操作後の色で作る'
    Assert-MbPixelNear -Path $cOutput -X 60 -Y 90 -Expected ([Drawing.Color]::FromArgb(220, 125, 45)) `
        -Message 'Cの左側へ操作後を置く'
    $cBitmap = [Drawing.Image]::FromFile($cOutput)
    try { $cSecondX = [int]($cBitmap.Width / 2) + 60 } finally { $cBitmap.Dispose() }
    Assert-MbPixelNear -Path $cOutput -X $cSecondX -Y 90 -Expected ([Drawing.Color]::FromArgb(110, 70, 160)) `
        -Message 'Cの右側へ操作前を置く'
    Assert-MbPixelNear -Path $dOutput -X 70 -Y 70 -Expected ([Drawing.Color]::FromArgb(25, 130, 170)) `
        -Message 'Dの上側へ横長の操作前を置く'
    $dBitmap = [Drawing.Image]::FromFile($dOutput)
    try { $dCenterX = [int]($dBitmap.Width / 2) } finally { $dBitmap.Dispose() }
    Assert-MbPixelNear -Path $dOutput -X $dCenterX -Y 400 -Expected ([Drawing.Color]::FromArgb(190, 55, 90)) `
        -Message 'Dの下側へ縦長の操作後を置く'

    Assert-Mb (((Get-MbProcessIds -Name EXCEL) -join ',') -eq ($excelBefore -join ',')) '検証中にExcelプロセスを起動・終了しない'
    Assert-Mb (((Get-MbProcessIds -Name WINWORD) -join ',') -eq ($wordBefore -join ',')) '検証中にWordプロセスを起動・終了しない'
    Write-Host ''
    Write-Host 'Office layout fixture tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
