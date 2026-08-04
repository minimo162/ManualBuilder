[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$EvaluationJson,
    [Parameter(Mandatory = $true)][string]$FramesDirectory,
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [string]$ScenarioId = 'expense-application'
)

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force

$evaluation = [IO.File]::ReadAllText([IO.Path]::GetFullPath($EvaluationJson), [Text.Encoding]::UTF8) | ConvertFrom-Json
$scenario = @($evaluation.scenarios | Where-Object { [string]$_.id -eq $ScenarioId }) | Select-Object -First 1
if ($null -eq $scenario) { throw "Scenario not found: $ScenarioId" }

$project = New-MbProject
$project.title = 'Fictional business portal (Copilot evaluation)'
$project.sheets[0].name = 'Expense application'
$projectDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))
if (-not (Test-Path -LiteralPath $projectDirectory)) { [void](New-Item -ItemType Directory -Path $projectDirectory -Force) }

foreach ($scene in @($scenario.scenes)) {
    $framePath = Join-Path ([IO.Path]::GetFullPath($FramesDirectory)) ([string]$scene.frameFile)
    if (-not (Test-Path -LiteralPath $framePath -PathType Leaf)) { throw "Frame not found: $framePath" }
    $added = Add-MbImageStep -Project $project -ProjectPath $ProjectPath -SheetId $project.sheets[0].id `
        -Bytes ([IO.File]::ReadAllBytes($framePath)) -Source 'video' -AllowDuplicateStep
    if ([string]$added.Status -ne 'added') { throw "Could not add frame: $framePath" }
    $candidates = @($scene.candidates)
    $selected = if ($candidates.Count -gt 0) { $candidates[0] } else { $null }
    if ($null -ne $selected) {
        [void](Set-MbStepAnnotations -Project $project -StepId ([string]$added.Step.id) -AnnotationsJson `
            (ConvertTo-Json -InputObject @([pscustomobject]@{
                id = 'annotation-' + [guid]::NewGuid().ToString('N'); type = 'rect'; label = 0
                x1 = $selected.rect.x1; y1 = $selected.rect.y1; x2 = $selected.rect.x2; y2 = $selected.rect.y2
            }) -Depth 6))
    }
    [void](Set-MbStepCapture -Project $project -StepId ([string]$added.Step.id) -Kind 'video-scene' `
        -VideoTimeMs ([int]$scene.actualTimeMs) -TargetSource 'video-diff' `
        -TargetConfidence $(if ($null -ne $selected) { [string]$selected.confidence } else { '' }) `
        -TargetCandidateId $(if ($null -ne $selected) { [string]$selected.id } else { '' }) `
        -TargetCandidatesJson $(if ($candidates.Count -gt 0) { ConvertTo-Json -InputObject $candidates -Depth 8 -Compress } else { '' }))
}

[void](Save-MbProject -Project $project -Path $ProjectPath)
Write-Output $ProjectPath
