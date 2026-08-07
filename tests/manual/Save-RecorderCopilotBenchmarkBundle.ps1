[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('edge-excel-order-transfer', 'excel-multi-operation', 'edge-delayed-transition')]
    [string]$Scenario,
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^run-\d{2,3}$')]
    [string]$RunId,
    [Parameter(Mandatory = $true)][string]$DestinationRoot,
    [string]$RecordingJobsRoot = '',
    [string]$RecordingJobDirectory = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if ([string]::IsNullOrWhiteSpace($RecordingJobsRoot)) {
    $RecordingJobsRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) `
        'ManualBuilder\data\recording-jobs'
}
$jobsRoot = [IO.Path]::GetFullPath($RecordingJobsRoot)
if (-not (Test-Path -LiteralPath $jobsRoot -PathType Container)) {
    throw "Recording jobs root was not found: $jobsRoot"
}

if ([string]::IsNullOrWhiteSpace($RecordingJobDirectory)) {
    $candidates = @(Get-ChildItem -LiteralPath $jobsRoot -Directory |
        Where-Object { $_.Name -match '^record-[0-9a-f]{32}$' } |
        Sort-Object LastWriteTime -Descending)
    $eligible = New-Object System.Collections.ArrayList
    foreach ($candidate in $candidates) {
        try {
            $candidateStatus = [IO.File]::ReadAllText((Join-Path $candidate.FullName 'status.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
            if ([string]$candidateStatus.state -ne 'completed') { continue }
            $candidateAi = @(Get-ChildItem -LiteralPath $candidate.FullName -Directory |
                Where-Object { $_.Name -match '^recording-ai-[0-9a-f]{32}$' })
            if ($candidateAi.Count -gt 0) { [void]$eligible.Add($candidate) }
        } catch { continue }
    }
    if ($eligible.Count -lt 1) { throw 'No completed recording job with RecorderCopilot results was found.' }
    if ($eligible.Count -gt 1) {
        throw 'Multiple completed recording jobs have Copilot results. Specify -RecordingJobDirectory explicitly.'
    }
    $recordingJob = $eligible[0]
} else {
    $recordingPath = [IO.Path]::GetFullPath($RecordingJobDirectory)
    $rootPrefix = $jobsRoot.TrimEnd('\') + '\'
    if (-not $recordingPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'RecordingJobDirectory must be inside RecordingJobsRoot.'
    }
    $recordingJob = Get-Item -LiteralPath $recordingPath -ErrorAction Stop
    if (-not $recordingJob.PSIsContainer -or $recordingJob.Name -notmatch '^record-[0-9a-f]{32}$') {
        throw 'RecordingJobDirectory is not a ManualBuilder recording job.'
    }
}

$recordingStatusPath = Join-Path $recordingJob.FullName 'status.json'
foreach ($path in @(
    $recordingStatusPath,
    (Join-Path $recordingJob.FullName 'frames.jsonl'),
    (Join-Path $recordingJob.FullName 'events.jsonl')
)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Recording artifact was not found: $path" }
}
$recordingStatus = [IO.File]::ReadAllText($recordingStatusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
if ([string]$recordingStatus.state -ne 'completed') {
    throw "Recording job is not completed: $([string]$recordingStatus.state)"
}

$aiJobs = @(Get-ChildItem -LiteralPath $recordingJob.FullName -Directory |
    Where-Object { $_.Name -match '^recording-ai-[0-9a-f]{32}$' } |
    Sort-Object LastWriteTime -Descending)
if ($aiJobs.Count -lt 1) { throw 'No RecorderCopilot child job was found.' }
$aiJob = $aiJobs[0]
$aiStatusPath = Join-Path $aiJob.FullName 'status.json'
$resultPath = Join-Path $aiJob.FullName 'result.json'
$logPath = Join-Path $aiJob.FullName 'copilot.log'
foreach ($path in @($aiStatusPath, $resultPath, $logPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Copilot artifact was not found: $path" }
}
$aiStatus = [IO.File]::ReadAllText($aiStatusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$result = [IO.File]::ReadAllText($resultPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
if ([string]$aiStatus.state -ne 'completed') { throw "RecorderCopilot job is not completed: $([string]$aiStatus.state)" }
if ([string]$aiStatus.jobId -ne [string]$result.jobId) { throw 'RecorderCopilot status/result jobId mismatch.' }
if (@($result.proposals).Count -lt 1) { throw 'RecorderCopilot produced no proposals.' }

$sourceFramesDirectory = Join-Path $recordingJob.FullName 'frames'
$frameFiles = @(Get-ChildItem -LiteralPath $sourceFramesDirectory -File -ErrorAction Stop |
    Where-Object { $_.Name -match '^frame-\d{5}\.jpg$' } |
    Sort-Object Name)
if ($frameFiles.Count -lt 1) { throw 'No recorder frame JPEG files were found.' }

$destinationBase = [IO.Path]::GetFullPath($DestinationRoot)
$destination = [IO.Path]::GetFullPath((Join-Path $destinationBase (Join-Path $Scenario $RunId)))
$destinationPrefix = $destinationBase.TrimEnd('\') + '\'
if (-not $destination.StartsWith($destinationPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The resolved destination escaped DestinationRoot.'
}
if (Test-Path -LiteralPath $destination) {
    throw "The benchmark run already exists; existing evidence will not be overwritten: $destination"
}

$created = $false
try {
    [void](New-Item -ItemType Directory -Path (Join-Path $destination 'frames') -Force)
    $created = $true
    Copy-Item -LiteralPath $aiStatusPath -Destination (Join-Path $destination 'status.json')
    Copy-Item -LiteralPath $resultPath -Destination (Join-Path $destination 'result.json')
    Copy-Item -LiteralPath $logPath -Destination (Join-Path $destination 'copilot.log')
    Copy-Item -LiteralPath (Join-Path $recordingJob.FullName 'frames.jsonl') -Destination $destination
    Copy-Item -LiteralPath (Join-Path $recordingJob.FullName 'events.jsonl') -Destination $destination
    foreach ($frame in $frameFiles) {
        Copy-Item -LiteralPath $frame.FullName -Destination (Join-Path $destination 'frames')
    }

    foreach ($name in @('status.json', 'result.json', 'copilot.log', 'frames.jsonl', 'events.jsonl')) {
        if (-not (Test-Path -LiteralPath (Join-Path $destination $name) -PathType Leaf)) {
            throw "Copied bundle is incomplete: $name"
        }
    }
    $copiedFrames = @(Get-ChildItem -LiteralPath (Join-Path $destination 'frames') -File)
    if ($copiedFrames.Count -ne $frameFiles.Count) { throw 'Not all recorder frames were copied.' }
} catch {
    if ($created -and (Test-Path -LiteralPath $destination)) {
        Remove-Item -LiteralPath $destination -Recurse -Force -ErrorAction SilentlyContinue
    }
    throw
}

[pscustomobject]@{
    scenario = $Scenario
    runId = $RunId
    destination = $destination
    recordingJobId = [string]$recordingStatus.jobId
    copilotJobId = [string]$aiStatus.jobId
    proposalCount = @($result.proposals).Count
    frameCount = $frameFiles.Count
    sourcePreserved = (Test-Path -LiteralPath $recordingJob.FullName -PathType Container)
} | ConvertTo-Json -Depth 4
