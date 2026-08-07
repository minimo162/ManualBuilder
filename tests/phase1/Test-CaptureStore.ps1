# Phase 1 image validation, storage, and duplicate detection test.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
$testRoot = Join-Path $env:TEMP ('ManualBuilder-CaptureTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function New-MbTestPng {
    param([Drawing.Color]$Color = [Drawing.Color]::CornflowerBlue)
    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object Drawing.Bitmap 8, 6
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $stream = New-Object IO.MemoryStream
    try {
        $graphics.Clear($Color)
        $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
        $stream.Dispose()
    }
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    $project = Get-MbProject -Path $projectPath
    $sheetId = [string]$project.selectedSheetId
    $bytes = New-MbTestPng

    $result = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $sheetId -Bytes $bytes -Source paste
    Assert-Mb ($result.Status -eq 'added') 'PNG画像を手順として追加できる'
    Assert-Mb ($result.Image.width -eq 8 -and $result.Image.height -eq 6) '画像寸法を検証して記録できる'
    Assert-Mb (@($project.images).Count -eq 1) '画像メタデータが1件追加される'
    Assert-Mb (@($project.sheets[0].steps).Count -eq 1) '画像付き手順が1件追加される'
    Assert-Mb ($project.sheets[0].steps[0].imageId -eq $result.Image.id) '手順が画像IDを参照する'

    $annotationJson = '[{"id":"annotation-00000000000000000000000000000001","type":"rect","x1":0.1,"y1":0.2,"x2":0.7,"y2":0.8,"label":0},{"id":"annotation-00000000000000000000000000000002","type":"number","x1":0.3,"y1":0.4,"x2":0.3,"y2":0.4,"label":1}]'
    Set-MbStepImageEdits -Project $project -StepId $project.sheets[0].steps[0].id -AnnotationsJson $annotationJson -CropJson '{"x":0.1,"y":0.1,"width":0.8,"height":0.8}'
    Assert-Mb (@($project.sheets[0].steps[0].annotations).Count -eq 2) '画像付き手順へ注釈を保存できる'
    Assert-Mb ([double]$project.sheets[0].steps[0].crop.width -eq 0.8) '画像付き手順へ切り抜き範囲を保存できる'

    $invalidAnnotationRejected = $false
    try {
        Set-MbStepAnnotations -Project $project -StepId $project.sheets[0].steps[0].id -AnnotationsJson '[{"id":"annotation-00000000000000000000000000000003","type":"arrow","x1":-1,"y1":0,"x2":1,"y2":1,"label":0}]'
    } catch { $invalidAnnotationRejected = $true }
    Assert-Mb $invalidAnnotationRejected '範囲外の注釈座標を拒否する'

    $invalidCropRejected = $false
    try {
        Set-MbStepImageEdits -Project $project -StepId $project.sheets[0].steps[0].id -AnnotationsJson $annotationJson -CropJson '{"x":0.9,"y":0,"width":0.2,"height":1}'
    } catch { $invalidCropRejected = $true }
    Assert-Mb $invalidCropRejected '画像外の切り抜き範囲を拒否する'

    $step = $project.sheets[0].steps[0]
    $step.title = '維持するタイトル'
    $step.description = '維持する説明'
    $originalImageId = [string]$step.imageId
    $replacementBytes = New-MbTestPng -Color ([Drawing.Color]::OrangeRed)
    $replacement = Set-MbStepImage -Project $project -ProjectPath $projectPath -StepId $step.id -Bytes $replacementBytes -Source file
    Assert-Mb ($replacement.Status -eq 'replaced' -and $replacement.Image.id -ne $originalImageId) '既存手順の画像を差し替えできる'
    Assert-Mb (@($project.sheets[0].steps).Count -eq 1) '画像差し替えで手順を増やさない'
    Assert-Mb ($step.title -eq '維持するタイトル' -and $step.description -eq '維持する説明') '画像差し替えで文章を維持する'
    Assert-Mb (@($step.annotations).Count -eq 0 -and [double]$step.crop.width -eq 1.0) '差し替え画像の注釈と切り抜きを初期化する'
    Assert-Mb ([string]$replacement.Previous.imageId -eq $originalImageId -and @($replacement.Previous.annotations).Count -eq 2) '元画像と編集内容を復元用に返す'

    $replacementImageId = [string]$step.imageId
    $restored = Restore-MbStepImage -Project $project -StepId $step.id -Previous $replacement.Previous
    Assert-Mb ([string]$step.imageId -eq $originalImageId) '元の画像へ戻せる'
    Assert-Mb (@($step.annotations).Count -eq 2 -and [double]$step.crop.width -eq 0.8) '元の注釈と切り抜きを復元できる'
    $replacementPath = Remove-MbUnusedImage -Project $project -ProjectPath $projectPath -ImageId $replacementImageId
    if ($replacementPath -and (Test-Path -LiteralPath $replacementPath)) { Remove-Item -LiteralPath $replacementPath -Force }
    Assert-Mb (@($project.images).Count -eq 1) '復元後の未参照画像を整理できる'

    $resultBytes = New-MbTestPng -Color ([Drawing.Color]::SeaGreen)
    $resultImage = Set-MbStepResultImage -Project $project -ProjectPath $projectPath -StepId $step.id -Bytes $resultBytes -Source file
    Assert-Mb ($resultImage.Status -eq 'set' -and -not [string]::IsNullOrWhiteSpace([string]$step.resultImageId)) '手順へ操作後画像を追加できる'
    Assert-Mb ([string]$step.imageLayout -eq 'side-by-side') '操作後画像の追加時は左右比較を初期値にする'
    Set-MbStepImageEdits -Project $project -StepId $step.id -Target result -AnnotationsJson $annotationJson -CropJson '{"x":0.2,"y":0.2,"width":0.7,"height":0.7}'
    Assert-Mb (@($step.resultAnnotations).Count -eq 2 -and [double]$step.resultCrop.width -eq 0.7) '操作後画像の注釈と切り抜きを操作前と分けて保存する'
    [void](Set-MbStepImageLayout -Project $project -StepId $step.id -Layout stacked -Order after-before)
    Assert-Mb ([string]$step.imageLayout -eq 'stacked' -and [string]$step.imageOrder -eq 'after-before') '比較画像の配置と前後順を保存できる'
    $removedResult = Remove-MbStepResultImage -Project $project -ProjectPath $projectPath -StepId $step.id
    Assert-Mb ([string]::IsNullOrWhiteSpace([string]$step.resultImageId) -and [string]$step.imageLayout -eq 'before') '操作後画像を外すと操作前だけへ戻る'
    Assert-Mb (@($step.resultAnnotations).Count -eq 0 -and [double]$step.resultCrop.width -eq 1.0) '操作後画像を外すと専用の編集状態も初期化する'
    if ($removedResult.RemovedPath -and (Test-Path -LiteralPath $removedResult.RemovedPath)) { Remove-Item -LiteralPath $removedResult.RemovedPath -Force }

    $imagePath = Get-MbImageFilePath -Project $project -ProjectPath $projectPath -ImageId $result.Image.id
    Assert-Mb (Test-Path -LiteralPath $imagePath -PathType Leaf) '画像実体をプロジェクト内へ保存する'
    $project = Save-MbProject -Project $project -Path $projectPath

    $duplicate = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $sheetId -Bytes $bytes -Source drop
    Assert-Mb ($duplicate.Status -eq 'duplicate') '同じ画像をSHA-256で重複判定する'
    Assert-Mb (@($project.images).Count -eq 1 -and @($project.sheets[0].steps).Count -eq 1) '重複画像では手順を増やさない'

    $reused = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $sheetId -Bytes $bytes `
        -Source recorder -AllowDuplicateStep
    Assert-Mb ($reused.Status -eq 'added') '操作記録では同じ画面でも別の手順として追加できる'
    Assert-Mb (@($project.images).Count -eq 1 -and @($project.sheets[0].steps).Count -eq 2) '同じ画面の手順は画像実体を共有する'
    Assert-Mb ([string]$project.sheets[0].steps[1].imageId -eq [string]$result.Image.id) '共有した画像IDを新しい手順から参照する'

    $invalidRejected = $false
    try {
        [void](Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $sheetId -Bytes ([byte[]](1, 2, 3, 4)) -Source file)
    } catch { $invalidRejected = $true }
    Assert-Mb $invalidRejected '画像でないバイト列を拒否する'

    $loaded = Get-MbProject -Path $projectPath
    Assert-Mb (@($loaded.images).Count -eq 1 -and $loaded.sheets[0].steps[0].imageId) '画像参照を再読込できる'
    Assert-Mb (@($loaded.sheets[0].steps[0].annotations).Count -eq 2) '注釈データを再読込できる'
    Assert-Mb ([double]$loaded.sheets[0].steps[0].crop.x -eq 0.1) '切り抜きデータを再読込できる'
    $temporaryFiles = @(Get-ChildItem -LiteralPath (Join-Path $testRoot 'images') -Filter '.image-*.tmp' -File -ErrorAction SilentlyContinue)
    Assert-Mb ($temporaryFiles.Count -eq 0) '画像保存後に一時ファイルが残らない'

    Write-Host ''
    Write-Host 'Capture store tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
