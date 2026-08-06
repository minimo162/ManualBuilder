[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$EvaluationJson,
    [Parameter(Mandatory = $true)][string]$FramesDirectory,
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [string]$ScenarioId = 'expense-application'
)

# Evaluation JSON input (one entry in scenarios):
#   projectTitle / sheetName are optional, user-supplied context only.
#   scenes[].frameFile is the operation-before image.
#   scenes[].resultFrameFile is an optional operation-after image.
#   scenes[].candidates contains every local target candidate. None of them is
#   selected in advance; Copilot must choose a candidate (or none) itself.

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force

$evaluation = [IO.File]::ReadAllText([IO.Path]::GetFullPath($EvaluationJson), [Text.Encoding]::UTF8) | ConvertFrom-Json
$scenario = @($evaluation.scenarios | Where-Object { [string]$_.id -eq $ScenarioId }) | Select-Object -First 1
if ($null -eq $scenario) { throw "Scenario not found: $ScenarioId" }

$project = New-MbProject
$project.title = if ($scenario.PSObject.Properties.Name -contains 'projectTitle' -and
    -not [string]::IsNullOrWhiteSpace([string]$scenario.projectTitle)) {
    [string]$scenario.projectTitle
} else { '録画から作成した手順' }
$project.sheets[0].name = if ($scenario.PSObject.Properties.Name -contains 'sheetName' -and
    -not [string]::IsNullOrWhiteSpace([string]$scenario.sheetName)) {
    [string]$scenario.sheetName
} else { '操作手順' }
$projectDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))
if (-not (Test-Path -LiteralPath $projectDirectory)) { [void](New-Item -ItemType Directory -Path $projectDirectory -Force) }
$framesRoot = [IO.Path]::GetFullPath($FramesDirectory)

foreach ($scene in @($scenario.scenes)) {
    $framePath = Join-Path $framesRoot ([string]$scene.frameFile)
    if (-not (Test-Path -LiteralPath $framePath -PathType Leaf)) { throw "Frame not found: $framePath" }
    $added = Add-MbImageStep -Project $project -ProjectPath $ProjectPath -SheetId $project.sheets[0].id `
        -Bytes ([IO.File]::ReadAllBytes($framePath)) -Source 'video' -AllowDuplicateStep
    if ([string]$added.Status -ne 'added') { throw "Could not add frame: $framePath" }

    if ($scene.PSObject.Properties.Name -contains 'resultFrameFile' -and
        -not [string]::IsNullOrWhiteSpace([string]$scene.resultFrameFile)) {
        $resultFramePath = Join-Path $framesRoot ([string]$scene.resultFrameFile)
        if (-not (Test-Path -LiteralPath $resultFramePath -PathType Leaf)) {
            throw "Result frame not found: $resultFramePath"
        }
        [void](Set-MbStepResultImage -Project $project -ProjectPath $ProjectPath `
            -StepId ([string]$added.Step.id) -Bytes ([IO.File]::ReadAllBytes($resultFramePath)) -Source 'file')
    }

    $candidates = @(if ($scene.PSObject.Properties.Name -contains 'candidates') { @($scene.candidates) })
    [void](Set-MbStepCapture -Project $project -StepId ([string]$added.Step.id) -Kind 'video-scene' `
        -VideoTimeMs ([int]$scene.actualTimeMs) -TargetSource 'video-diff' `
        -TargetConfidence '' -TargetCandidateId '' `
        -TargetCandidatesJson $(if ($candidates.Count -gt 0) { ConvertTo-Json -InputObject $candidates -Depth 8 -Compress } else { '' }))
    # Set-MbStepCapture normally falls back to the first candidate when the current
    # selection is empty. In an evaluation that would disclose candidate order as
    # a preselected answer, so leave every candidate explicitly unselected.
    $added.Step.capture.targetCandidateId = ''
    $added.Step.capture.targetConfidence = ''
}

[void](Save-MbProject -Project $project -Path $ProjectPath)
Write-Output $ProjectPath
