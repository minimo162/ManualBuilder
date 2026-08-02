# Phase 1 project model and atomic save test.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
$testRoot = Join-Path $env:TEMP ('ManualBuilder-ProjectTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    $project = Get-MbProject -Path $projectPath
    Assert-Mb (@($project.sheets).Count -eq 1) '新規プロジェクトにシートが1件ある'
    Assert-Mb (@($project.images).Count -eq 0) '新規プロジェクトの画像一覧が空である'
    Assert-Mb ($project.title -eq '新しいマニュアル') '新規タイトルが設定される'

    Set-MbProjectTitle -Project $project -Title '経費精算 操作マニュアル'
    $sheet2 = Add-MbSheet -Project $project
    Rename-MbSheet -Project $project -SheetId $sheet2.id -Name '申請'
    $step = Add-MbStep -Project $project -SheetId $sheet2.id
    Update-MbStep -Project $project -StepId $step.id -Title '申請画面を開く' -Description 'メニューから申請を選択します。' -Note '事前にログインが必要です。'
    Assert-Mb (@($step.annotations).Count -eq 0) '新しい手順の注釈一覧が空である'
    Assert-Mb ([double]$step.crop.width -eq 1 -and [double]$step.crop.height -eq 1) '新しい手順の切り抜き範囲が画像全体である'
    $project = Save-MbProject -Project $project -Path $projectPath

    Assert-Mb (Test-Path -LiteralPath $projectPath) 'project.jsonが作成される'
    Assert-Mb (@($project.sheets).Count -eq 2) 'シートを追加できる'
    Assert-Mb (@((Get-MbSelectedSheet -Project $project).steps).Count -eq 1) '選択シートへ手順を追加できる'

    $loaded = Get-MbProject -Path $projectPath
    $loadedStep = @((Get-MbSelectedSheet -Project $loaded).steps)[0]
    Assert-Mb ($loaded.title -eq '経費精算 操作マニュアル') 'タイトルを再読込できる'
    Assert-Mb ($loadedStep.description -eq 'メニューから申請を選択します。') '説明を再読込できる'
    Assert-Mb (@($loadedStep.annotations).Count -eq 0) '空の注釈一覧を再読込できる'
    Assert-Mb ([double]$loadedStep.crop.x -eq 0 -and [double]$loadedStep.crop.width -eq 1) '切り抜き範囲を再読込できる'
    $loadedStep.PSObject.Properties.Remove('crop')
    [void](Save-MbProject -Project $loaded -Path $projectPath)
    Assert-Mb ([double]$loadedStep.crop.width -eq 1 -and [double]$loadedStep.crop.height -eq 1) '既存プロジェクトへ画像全体の切り抜き範囲を補完する'

    $temporaryStep = Add-MbStep -Project $loaded -SheetId $sheet2.id
    Update-MbStep -Project $loaded -StepId $temporaryStep.id -Title '並べ替え対象' -Description '' -Note ''
    Set-MbStepOrder -Project $loaded -SheetId $sheet2.id -StepIds @($temporaryStep.id, $loadedStep.id)
    Assert-Mb ([string]@((Get-MbSelectedSheet -Project $loaded).steps)[0].id -eq $temporaryStep.id) '手順を並べ替えできる'
    Remove-MbStep -Project $loaded -StepId $temporaryStep.id
    Assert-Mb (@((Get-MbSelectedSheet -Project $loaded).steps).Count -eq 1) '手順を削除できる'

    $temporarySheet = Add-MbSheet -Project $loaded
    Set-MbSheetOrder -Project $loaded -SheetIds @($temporarySheet.id, $sheet2.id, $loaded.sheets[0].id)
    Assert-Mb ([string]$loaded.sheets[0].id -eq $temporarySheet.id) 'シートを並べ替えできる'
    [void](Move-MbStepToSheet -Project $loaded -StepId $loadedStep.id -TargetSheetId $temporarySheet.id)
    Assert-Mb (@($temporarySheet.steps).Count -eq 1 -and [string]$temporarySheet.steps[0].id -eq $loadedStep.id) '手順を別シートへ移動できる'
    Assert-Mb ([string]$temporarySheet.steps[0].description -eq 'メニューから申請を選択します。') 'シート移動後も手順内容を維持する'
    [void](Move-MbStepToSheet -Project $loaded -StepId $loadedStep.id -TargetSheetId $sheet2.id)
    Remove-MbSheet -Project $loaded -SheetId $temporarySheet.id
    Select-MbSheet -Project $loaded -SheetId $sheet2.id
    Assert-Mb (@($loaded.sheets).Count -eq 2) 'シートを削除できる'

    Set-MbProjectTitle -Project $loaded -Title '経費精算 操作マニュアル 改訂'
    [void](Save-MbProject -Project $loaded -Path $projectPath)
    Assert-Mb (Test-Path -LiteralPath "$projectPath.bak") '2回目の保存で直前バックアップが残る'

    $tempFiles = @(Get-ChildItem -LiteralPath $testRoot -Filter '.project-*.tmp' -File -ErrorAction SilentlyContinue)
    Assert-Mb ($tempFiles.Count -eq 0) '保存後に一時ファイルが残らない'
    Write-Host ''
    Write-Host 'Project store tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
