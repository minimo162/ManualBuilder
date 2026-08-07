[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$scenarioScript = Join-Path $repoRoot 'tests\manual\Invoke-RecorderCopilotScenario.ps1'
$existingScript = Join-Path $repoRoot 'tests\manual\Invoke-RealisticRecordingSmoke.ps1'
$bundleScript = Join-Path $repoRoot 'tests\manual\Save-RecorderCopilotBenchmarkBundle.ps1'
$localBenchmarkScript = Join-Path $repoRoot 'tests\manual\Invoke-RecorderLocalBenchmark.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-RecorderTools-' + [guid]::NewGuid().ToString('N'))
$jobsRoot = Join-Path $testRoot 'recording-jobs'
$recordingJob = Join-Path $jobsRoot ('record-' + ('1' * 32))
$aiJob = Join-Path $recordingJob ('recording-ai-' + ('2' * 32))
$destinationRoot = Join-Path $testRoot 'benchmark'
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

function Write-TestUtf8 {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Text)
    $directory = Split-Path -Parent $Path
    if ($directory) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    [IO.File]::WriteAllText($Path, $Text, [Text.UTF8Encoding]::new($false))
}

try {
    foreach ($path in @($scenarioScript, $existingScript, $bundleScript, $localBenchmarkScript)) {
        $bytes = [IO.File]::ReadAllBytes($path)
        Add-Result ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) `
            ('Windows PowerShell用UTF-8 BOM: ' + [IO.Path]::GetFileName($path))
        $text = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
        Add-Result (-not [regex]::IsMatch($text, '(?<!\r)\n')) `
            ('Windows PowerShell用CRLF: ' + [IO.Path]::GetFileName($path))
    }

    $scenarioText = [IO.File]::ReadAllText($scenarioScript, [Text.Encoding]::UTF8)
    Add-Result ($scenarioText -match "ValidateSet\('edge-excel-order-transfer', 'excel-multi-operation', 'edge-delayed-transition'\)") `
        '3つの実践シナリオだけを選択できる'
    Add-Result ($scenarioText -match '\$existingHandles -notcontains' -and
        $scenarioText -match '\$ownedEdgeHandle' -and $scenarioText -match 'PostMessage\(\$ownedEdgeHandle') `
        '起動前のEdgeと所有Edgeを区別し、所有ウィンドウだけを閉じる'
    Add-Result ($scenarioText -match '\$excel = New-Object -ComObject Excel.Application' -and
        $scenarioText -match '\$excel\.Quit\(\)' -and $scenarioText -notmatch 'Stop-Process') `
        '専用Excel COMインスタンスだけを終了する'
    Add-Result ($scenarioText -match '-AllowExistingExcel') `
        '既存ExcelがあってもEdge→Excelシナリオを拒否しない'
    $localBenchmarkText = [IO.File]::ReadAllText($localBenchmarkScript, [Text.Encoding]::UTF8)
    Add-Result ($localBenchmarkText -match 'Start-MbRecordingJob' -and
        $localBenchmarkText -match 'Select-MbRecorderCandidateFrames' -and
        $localBenchmarkText -match 'New-MbRecorderContactSheets' -and
        $localBenchmarkText -match '-WindowStyle Hidden' -and
        $localBenchmarkText -match 'done\.signal' -and $localBenchmarkText -match 'release\.signal') `
        '実機ベンチマークが前面を奪わず、操作完了から代表コマ一覧まで一括生成する'

    [void](New-Item -ItemType Directory -Path (Join-Path $recordingJob 'frames') -Force)
    [void](New-Item -ItemType Directory -Path $aiJob -Force)
    Write-TestUtf8 -Path (Join-Path $recordingJob 'status.json') -Text `
        (([pscustomobject]@{ jobId = 'record-' + ('1' * 32); state = 'completed' } | ConvertTo-Json -Compress))
    Write-TestUtf8 -Path (Join-Path $recordingJob 'frames.jsonl') -Text `
        (([pscustomobject]@{ id = 'F00001'; index = 1; timeMs = 0; image = 'frame-00001.jpg'; windowTitle = 'Test - Microsoft Edge' } | ConvertTo-Json -Compress) + "`n")
    Write-TestUtf8 -Path (Join-Path $recordingJob 'events.jsonl') -Text `
        (([pscustomobject]@{ index = 1; timeMs = 100; kind = 'click'; targetName = 'Test' } | ConvertTo-Json -Compress) + "`n")
    [IO.File]::WriteAllBytes((Join-Path $recordingJob 'frames\frame-00001.jpg'), ([byte[]](0xFF, 0xD8, 0xFF, 0xD9)))
    $aiId = 'recording-ai-' + ('2' * 32)
    Write-TestUtf8 -Path (Join-Path $aiJob 'status.json') -Text `
        (([pscustomobject]@{ jobId = $aiId; state = 'completed' } | ConvertTo-Json -Compress))
    Write-TestUtf8 -Path (Join-Path $aiJob 'result.json') -Text `
        (([pscustomobject]@{ jobId = $aiId; proposals = @([pscustomobject]@{ beforeFrame = 'F00001'; title = 'Test' }) } | ConvertTo-Json -Depth 5 -Compress))
    Write-TestUtf8 -Path (Join-Path $aiJob 'copilot.log') -Text "[INFO] completed`n"

    $json = & $bundleScript -Scenario 'edge-delayed-transition' -RunId 'run-01' `
        -DestinationRoot $destinationRoot -RecordingJobsRoot $jobsRoot -RecordingJobDirectory $recordingJob
    $saved = ($json -join "`n") | ConvertFrom-Json
    $destination = Join-Path $destinationRoot 'edge-delayed-transition\run-01'
    Add-Result ([string]$saved.destination -eq [IO.Path]::GetFullPath($destination) -and [bool]$saved.sourcePreserved) `
        '明示した録画jobを指定シナリオ/runへ非破壊コピーする'
    Add-Result ((Test-Path -LiteralPath (Join-Path $destination 'status.json')) -and
        (Test-Path -LiteralPath (Join-Path $destination 'result.json')) -and
        (Test-Path -LiteralPath (Join-Path $destination 'frames\frame-00001.jpg'))) `
        '評価器が必要とするAI状態・結果・録画フレームを保存する'
    Add-Result ((Test-Path -LiteralPath $recordingJob -PathType Container) -and
        (Test-Path -LiteralPath $aiJob -PathType Container)) `
        'コピー後も録画親jobとAI子jobを削除しない'

    $overwriteRejected = $false
    try {
        & $bundleScript -Scenario 'edge-delayed-transition' -RunId 'run-01' `
            -DestinationRoot $destinationRoot -RecordingJobsRoot $jobsRoot -RecordingJobDirectory $recordingJob | Out-Null
    } catch { $overwriteRejected = $_.Exception.Message -match 'will not be overwritten' }
    Add-Result $overwriteRejected '既存の評価証跡を上書きしない'

    $outside = Join-Path $testRoot ('record-' + ('3' * 32))
    [void](New-Item -ItemType Directory -Path $outside -Force)
    $outsideRejected = $false
    try {
        & $bundleScript -Scenario 'edge-delayed-transition' -RunId 'run-02' `
            -DestinationRoot $destinationRoot -RecordingJobsRoot $jobsRoot -RecordingJobDirectory $outside | Out-Null
    } catch { $outsideRejected = $_.Exception.Message -match 'inside RecordingJobsRoot' }
    Add-Result $outsideRejected '録画root外のディレクトリを証跡として受け付けない'
} catch {
    Add-Result $false ('manual tool test raised an exception: ' + $_.Exception.Message)
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($errors.Count -gt 0) {
    Write-Host "`n$($errors.Count) checks failed." -ForegroundColor Red
    exit 1
}
Write-Host "`nRecorderCopilot manual tool checks passed." -ForegroundColor Green
exit 0
