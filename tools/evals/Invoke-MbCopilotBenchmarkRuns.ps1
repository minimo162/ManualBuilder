# Runs the real M365 Copilot draft worker repeatedly against an immutable
# evaluation project and preserves each raw result for blind review.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [ValidateRange(1, 10)][int]$Runs = 3,
    [ValidateRange(60, 1800)][int]$TimeoutSeconds = 600,
    [AllowEmptyString()][string]$ProfileRoot = '',
    [AllowEmptyString()][string]$ConfigPath = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$sourceRoot = Join-Path $repositoryRoot 'src'
$resolvedProjectPath = [IO.Path]::GetFullPath($ProjectPath)
$resolvedOutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if (-not (Test-Path -LiteralPath $resolvedProjectPath -PathType Leaf)) {
    throw "評価プロジェクトが見つかりません: $resolvedProjectPath"
}
if ([string]::IsNullOrWhiteSpace($ProfileRoot)) {
    $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
    $ProfileRoot = Join-Path $localAppData 'ManualBuilder\data\copilot-edge-profile'
}
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
    $ConfigPath = Join-Path $localAppData 'ManualBuilder\data\copilot.json'
}

Import-Module (Join-Path $sourceRoot 'ManualBuilder.CopilotServer.psm1') -Force

if (-not (Test-Path -LiteralPath $resolvedOutputRoot)) {
    [void](New-Item -ItemType Directory -Path $resolvedOutputRoot -Force)
}
$jobsRoot = Join-Path $resolvedOutputRoot '_jobs'
[void](New-Item -ItemType Directory -Path $jobsRoot -Force)
Initialize-MbCopilotServer -JobsRoot $jobsRoot -ScriptRoot $sourceRoot `
    -ProfileRoot ([IO.Path]::GetFullPath($ProfileRoot)) -ConfigPath $ConfigPath

$sourceProjectDirectory = Split-Path -Parent $resolvedProjectPath
$sourceProjectFileName = Split-Path -Leaf $resolvedProjectPath
$summaries = New-Object System.Collections.Generic.List[object]
$utf8NoBom = New-Object Text.UTF8Encoding($false)

for ($runNumber = 1; $runNumber -le $Runs; $runNumber++) {
    $runName = 'run-{0:d2}' -f $runNumber
    $runDirectory = Join-Path $resolvedOutputRoot $runName
    if (Test-Path -LiteralPath $runDirectory) {
        throw "出力先がすでに存在します。既存結果を上書きしません: $runDirectory"
    }
    [void](New-Item -ItemType Directory -Path $runDirectory)
    $inputDirectory = Join-Path $runDirectory 'input'
    [void](New-Item -ItemType Directory -Path $inputDirectory)
    foreach ($entry in @(Get-ChildItem -LiteralPath $sourceProjectDirectory -Force)) {
        Copy-Item -LiteralPath $entry.FullName -Destination $inputDirectory -Recurse -Force
    }
    $runProjectPath = Join-Path $inputDirectory $sourceProjectFileName

    Write-Host ("[{0}/{1}] 実M365 Copilotへ下書きを依頼します" -f $runNumber, $Runs)
    $startedAt = [DateTime]::UtcNow
    $status = Start-MbCopilotDraftJob -ProjectPath $runProjectPath
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $status = Read-MbCopilotDraftStatus
        Write-Host ("  {0,3}% {1}" -f [int]$status.percent, [string]$status.message)
        if ([string]$status.state -in @('completed', 'failed', 'cancelled')) { break }
    } while ((Get-Date) -lt $deadline)

    if ([string]$status.state -notin @('completed', 'failed', 'cancelled')) {
        [void](Request-MbCopilotDraftCancel)
        $status = Read-MbCopilotDraftStatus
        $status | Add-Member -NotePropertyName benchmarkTimeout -NotePropertyValue $true -Force
    }
    $result = Get-MbCopilotDraftResult
    [IO.File]::WriteAllText((Join-Path $runDirectory 'status.json'),
        ($status | ConvertTo-Json -Depth 12), $utf8NoBom)
    [IO.File]::WriteAllText((Join-Path $runDirectory 'result.json'),
        ($result | ConvertTo-Json -Depth 30), $utf8NoBom)

    $resultPath = [string]$status.resultPath
    if (-not [string]::IsNullOrWhiteSpace($resultPath)) {
        $jobDirectory = Split-Path -Parent $resultPath
        $logPath = Join-Path $jobDirectory 'copilot.log'
        if (Test-Path -LiteralPath $logPath -PathType Leaf) {
            Copy-Item -LiteralPath $logPath -Destination (Join-Path $runDirectory 'copilot.log') -Force
        }
    }
    $draftCount = @($result.drafts).Count
    $failureCount = @($result.failures).Count
    $summaries.Add([pscustomobject]@{
        run = $runNumber
        state = [string]$status.state
        errorCode = [string]$status.errorCode
        draftCount = $draftCount
        failureCount = $failureCount
        elapsedSeconds = [Math]::Round(([DateTime]::UtcNow - $startedAt).TotalSeconds, 1)
        result = (Join-Path $runName 'result.json')
    })
    Remove-MbCopilotDraftJob
}

$summary = [pscustomobject]@{
    schemaVersion = 1
    generatedAt = [DateTime]::UtcNow.ToString('o')
    projectPath = $resolvedProjectPath
    requestedRuns = $Runs
    completedRuns = @($summaries | Where-Object { $_.state -eq 'completed' }).Count
    runs = @($summaries | ForEach-Object { $_ })
}
[IO.File]::WriteAllText((Join-Path $resolvedOutputRoot 'summary.json'),
    ($summary | ConvertTo-Json -Depth 12), $utf8NoBom)
$summary | ConvertTo-Json -Depth 12
