[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('edge-excel-order-transfer', 'excel-multi-operation', 'edge-delayed-transition')]
    [string]$Scenario,
    [Parameter(Mandatory = $true)][string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$output = [IO.Path]::GetFullPath($OutputDirectory)
$repoPrefix = $repoRoot.TrimEnd('\') + '\'
if (-not $output.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputDirectory must be inside the ManualBuilder workspace.'
}
if (Test-Path -LiteralPath $output) { throw "OutputDirectory already exists: $output" }

$jobsRoot = Join-Path $output 'recording-jobs'
$readyPath = Join-Path $output 'ready.json'
$startPath = Join-Path $output 'go.signal'
$donePath = Join-Path $output 'done.signal'
$releasePath = Join-Path $output 'release.signal'
$stdoutPath = Join-Path $output 'scenario-output.txt'
$stderrPath = Join-Path $output 'scenario-error.txt'
$workbookPath = Join-Path $output 'scenario.xlsx'
[void](New-Item -ItemType Directory -Path $jobsRoot -Force)

Import-Module (Join-Path $repoRoot 'src\ManualBuilder.RecorderServer.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.RecorderCopilot.psm1') -Force
Initialize-MbRecorderServer -JobsRoot $jobsRoot -ScriptRoot (Join-Path $repoRoot 'src')

$powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$scenarioScript = Join-Path $PSScriptRoot 'Invoke-RecorderCopilotScenario.ps1'
$quote = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
$arguments = @(
    '-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $scenarioScript),
    '-Scenario', $Scenario,
    '-ReadyPath', (& $quote $readyPath),
    '-StartPath', (& $quote $startPath),
    '-DonePath', (& $quote $donePath),
    '-ReleasePath', (& $quote $releasePath)
)
if ($Scenario -in @('excel-multi-operation', 'edge-excel-order-transfer')) {
    $arguments += @('-WorkbookPath', (& $quote $workbookPath))
}

$scenarioProcess = $null
try {
    $scenarioProcess = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $readyDeadline = (Get-Date).AddSeconds(25)
    while (-not (Test-Path -LiteralPath $readyPath -PathType Leaf)) {
        if ($scenarioProcess.HasExited) {
            $errorText = if (Test-Path -LiteralPath $stderrPath) { [IO.File]::ReadAllText($stderrPath) } else { '' }
            throw "Scenario stopped before it was ready. $errorText"
        }
        if ((Get-Date) -ge $readyDeadline) { throw 'Scenario did not become ready within 25 seconds.' }
        Start-Sleep -Milliseconds 200
        $scenarioProcess.Refresh()
    }

    $started = Start-MbRecordingJob -IgnoreTitlePatterns @('ManualBuilder', 'Codex', 'ChatGPT', 'Claude')
    if ([string]$started.state -ne 'recording') { throw "Recorder did not start: $([string]$started.state)" }
    Start-Sleep -Milliseconds 700
    [IO.File]::WriteAllText($startPath, 'go', [Text.UTF8Encoding]::new($false))

    $doneDeadline = (Get-Date).AddSeconds(90)
    while (-not (Test-Path -LiteralPath $donePath -PathType Leaf)) {
        if ($scenarioProcess.HasExited) {
            $errorText = if (Test-Path -LiteralPath $stderrPath) { [IO.File]::ReadAllText($stderrPath) } else { '' }
            throw "Scenario stopped before its actions were complete. $errorText"
        }
        if ((Get-Date) -ge $doneDeadline) { throw 'Scenario actions did not finish within 90 seconds.' }
        Start-Sleep -Milliseconds 200
        $scenarioProcess.Refresh()
    }
    Start-Sleep -Milliseconds 900
    $stopped = Stop-MbRecordingJob
    if ([string]$stopped.state -ne 'completed') { throw "Recorder did not complete: $([string]$stopped.state)" }
    [IO.File]::WriteAllText($releasePath, 'release', [Text.UTF8Encoding]::new($false))

    if (-not $scenarioProcess.WaitForExit(10000)) { throw 'Scenario did not close within 10 seconds after recording stopped.' }
    # リダイレクトした標準出力を最後まで排出し、ExitCodeを確定させる。
    $scenarioProcess.WaitForExit()
    $scenarioProcess.Refresh()
    $scenarioOutput = if (Test-Path -LiteralPath $stdoutPath) { [IO.File]::ReadAllText($stdoutPath) } else { '' }
    $scenarioExitCode = $scenarioProcess.ExitCode
    if (($null -ne $scenarioExitCode -and [int]$scenarioExitCode -ne 0) -or
        $scenarioOutput -notmatch '"completed"\s*:\s*true') {
        $errorText = if (Test-Path -LiteralPath $stderrPath) { [IO.File]::ReadAllText($stderrPath) } else { '' }
        throw "Scenario failed with exit code $($scenarioProcess.ExitCode). $errorText"
    }
    $jobDirectory = @(Get-ChildItem -LiteralPath $jobsRoot -Directory |
        Where-Object { $_.Name -match '^record-[0-9a-f]{32}$' })
    if ($jobDirectory.Count -ne 1) { throw "Expected one recording job, found $($jobDirectory.Count)." }

    $framesPath = Join-Path $jobDirectory[0].FullName 'frames.jsonl'
    $eventsPath = Join-Path $jobDirectory[0].FullName 'events.jsonl'
    $framesDirectory = Join-Path $jobDirectory[0].FullName 'frames'
    $frames = @(Read-MbRecorderJsonLines -Path $framesPath)
    $events = @(Read-MbRecorderJsonLines -Path $eventsPath)
    $frames = @(Add-MbRecorderFrameVisualMetrics -Frames $frames -FramesDirectory $framesDirectory)
    $localCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $frames -Events $events -MaximumFrames 30)
    $reviewFrames = @(Select-MbRecorderCandidateFrames -Frames $frames -Candidates $localCandidates)
    if ($localCandidates.Count -lt 1 -or $reviewFrames.Count -lt 2) {
        throw 'Recorder did not produce enough local candidates for review.'
    }

    $contactDirectory = Join-Path $output 'contact-selected'
    $sheets = @(New-MbRecorderContactSheets -Frames $reviewFrames -FramesDirectory $framesDirectory `
        -OutputDirectory $contactDirectory)
    [IO.File]::WriteAllText((Join-Path $output 'local-proposals.json'),
        (($localCandidates | ConvertTo-Json -Depth 10)), [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $output 'review-frames.json'),
        (($reviewFrames | ConvertTo-Json -Depth 8)), [Text.UTF8Encoding]::new($false))

    [pscustomobject]@{
        scenario = $Scenario
        outputDirectory = $output
        eventCount = $events.Count
        timelineFrameCount = $frames.Count
        proposalCount = $localCandidates.Count
        reviewFrameCount = $reviewFrames.Count
        contactSheetCount = $sheets.Count
        status = [string]$stopped.state
    } | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $scenarioProcess) {
        try {
            $scenarioProcess.Refresh()
            if (-not $scenarioProcess.HasExited) { $scenarioProcess.Kill() }
        } catch { }
        $scenarioProcess.Dispose()
    }
    try { [void](Stop-MbRecordingJob) } catch { }
}
