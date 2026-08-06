# Copilot評価用プロジェクトが、候補順を正解として漏らさず操作前後を保持することを検査する。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$builderPath = Join-Path $repoRoot 'tools\evals\New-MbCopilotEvalProject.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-CopilotEval-' + [guid]::NewGuid().ToString('N'))
$framesDirectory = Join-Path $testRoot 'frames'
$evaluationPath = Join-Path $testRoot 'evaluation.json'
$projectPath = Join-Path $testRoot 'project\project.json'
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

function New-MbEvalJpeg {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][Drawing.Color]$Color)
    $bitmap = New-Object Drawing.Bitmap 80, 50
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear($Color)
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Jpeg)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

try {
    Add-Type -AssemblyName System.Drawing
    [void](New-Item -ItemType Directory -Path $framesDirectory -Force)
    New-MbEvalJpeg -Path (Join-Path $framesDirectory 'before.jpg') -Color ([Drawing.Color]::SteelBlue)
    New-MbEvalJpeg -Path (Join-Path $framesDirectory 'after.jpg') -Color ([Drawing.Color]::SeaGreen)

    $evaluation = [pscustomobject]@{
        scenarios = @([pscustomobject]@{
            id = 'fair-candidate-test'
            projectTitle = '架空の申請システム'
            sheetName = '申請操作'
            scenes = @(
                [pscustomobject]@{
                    order = 1; actualTimeMs = 1250; frameFile = 'before.jpg'; resultFrameFile = 'after.jpg'
                    candidates = @(
                        [pscustomobject]@{ id = 'video-diff-1'; source = 'video-diff'; confidence = 'low'; label = ''; targetType = ''; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.3; y2 = 0.3 } },
                        [pscustomobject]@{ id = 'video-diff-2'; source = 'video-diff'; confidence = 'medium'; label = ''; targetType = ''; rect = [pscustomobject]@{ x1 = 0.6; y1 = 0.6; x2 = 0.9; y2 = 0.9 } }
                    )
                },
                [pscustomobject]@{ order = 2; actualTimeMs = 2500; frameFile = 'after.jpg'; candidates = @() }
            )
        })
    }
    [IO.File]::WriteAllText($evaluationPath, ($evaluation | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))

    [void](& $builderPath -EvaluationJson $evaluationPath -FramesDirectory $framesDirectory `
        -ProjectPath $projectPath -ScenarioId 'fair-candidate-test')
    $project = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $steps = @($project.sheets[0].steps)

    Add-Result ([string]$project.title -eq '架空の申請システム' -and [string]$project.sheets[0].name -eq '申請操作') `
        '入力した一般コンテキストだけをプロジェクトへ保持する'
    Add-Result ($steps.Count -eq 2) '評価場面を順番どおり手順へ変換する'
    Add-Result (@($steps[0].capture.targetCandidates).Count -eq 2) '全候補をCopilotの比較対象として保持する'
    Add-Result ([string]$steps[0].capture.targetCandidates[0].id -eq 'video-diff-1' -and
        [string]$steps[0].capture.targetCandidates[1].id -eq 'video-diff-2') '候補IDと順序を改変しない'
    Add-Result ([string]::IsNullOrWhiteSpace([string]$steps[0].capture.targetCandidateId) -and
        [string]::IsNullOrWhiteSpace([string]$steps[0].capture.targetConfidence)) '先頭候補を選択済みの正解として漏らさない'
    Add-Result (@($steps[0].annotations).Count -eq 0) '候補を確定注釈として保存しない'
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$steps[0].resultImageId) -and
        [string]$steps[0].resultImageId -ne [string]$steps[0].imageId) '操作後フレームを別画像として保持する'
    Add-Result ([string]$steps[0].imageLayout -eq 'side-by-side') '操作前後を比較できる配置にする'
    Add-Result ([string]::IsNullOrWhiteSpace([string]$steps[1].resultImageId)) '操作後フレームがない場面へ画像を捏造しない'
    Add-Result ([int]$steps[0].capture.videoTimeMs -eq 1250 -and [int]$steps[1].capture.videoTimeMs -eq 2500) `
        '抽出した録画内時刻を保持する'
} catch {
    Add-Result $false ('Copilot評価用プロジェクトの検査中に例外: ' + $_.Exception.Message)
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($errors.Count -gt 0) {
    Write-Host "`n$($errors.Count) checks failed." -ForegroundColor Red
    exit 1
}

Write-Host "`nCopilot evaluation project checks passed." -ForegroundColor Green
exit 0
