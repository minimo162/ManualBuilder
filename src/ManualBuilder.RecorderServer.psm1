# 操作記録のジョブ管理と、記録した操作の取り込み。
#
# 記録そのものは別プロセス（Invoke-ManualBuilderRecorder.ps1）で行う。
# サーバーは1本のスレッドでHTTPを捌いているため、60Hzのポーリングを同居させると
# 画面の操作が止まる。Excel出力などと同じく、状態は status.json、停止は
# stop.requested のファイルで受け渡す。

Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Recorder.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.LocalDraft.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.RecorderCopilot.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Ocr.psm1')

$script:MbRecordingJobsRoot = ''
$script:MbRecordingScriptRoot = ''
$script:MbRecordingJob = $null
$script:MbRecorderOcrSnapshotCache = @{}

function New-MbRecorderProcessIdentity {
    param([Parameter(Mandatory = $true)]$Process)
    try {
        return [pscustomobject]@{
            Id = [int]$Process.Id
            ProcessName = [string]$Process.ProcessName
            StartTimeUtcTicks = [int64]$Process.StartTime.ToUniversalTime().Ticks
            ExecutablePath = [string]$Process.Path
        }
    } catch { return $null }
}

function Test-MbRecorderProcessIdentity {
    param([AllowNull()]$Identity)
    if ($null -eq $Identity -or [int]$Identity.Id -le 0 -or
        [string]::IsNullOrWhiteSpace([string]$Identity.ProcessName) -or
        [int64]$Identity.StartTimeUtcTicks -le 0 -or
        [string]::IsNullOrWhiteSpace([string]$Identity.ExecutablePath)) { return $false }
    $process = $null
    try {
        $process = Get-Process -Id ([int]$Identity.Id) -ErrorAction Stop
        return ([string]$process.ProcessName -eq [string]$Identity.ProcessName -and
            [int64]$process.StartTime.ToUniversalTime().Ticks -eq [int64]$Identity.StartTimeUtcTicks -and
            [string]::Equals([string]$process.Path, [string]$Identity.ExecutablePath, [StringComparison]::OrdinalIgnoreCase))
    } catch { return $false }
    finally { if ($null -ne $process) { try { $process.Dispose() } catch { } } }
}

function Stop-MbRecorderOwnedProcess {
    param([AllowNull()]$Identity)
    if ($null -eq $Identity -or [int]$Identity.Id -le 0) { return $false }
    $process = $null
    try {
        $process = Get-Process -Id ([int]$Identity.Id) -ErrorAction Stop
        # 所有確認とKillを同じProcessオブジェクト上で行い、その間のPID再利用余地を作らない。
        if ([string]$process.ProcessName -ne [string]$Identity.ProcessName -or
            [int64]$process.StartTime.ToUniversalTime().Ticks -ne [int64]$Identity.StartTimeUtcTicks -or
            -not [string]::Equals([string]$process.Path, [string]$Identity.ExecutablePath, [StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
        $process.Kill()
        return $true
    } catch { return $false }
    finally { if ($null -ne $process) { try { $process.Dispose() } catch { } } }
}

function Initialize-MbRecorderServer {
    param(
        [Parameter(Mandatory = $true)][string]$JobsRoot,
        [Parameter(Mandatory = $true)][string]$ScriptRoot
    )
    $script:MbRecordingJobsRoot = $JobsRoot
    $script:MbRecordingScriptRoot = $ScriptRoot
}

function Remove-MbOrphanedRecordingJobs {
    # 記録ジョブは Remove-MbRecordingJob でしか消えず、それはメモリ上のジョブが
    # ある場合にしか動かない。取り込む前にアプリが落ちる・再起動されると、
    # 500msごとの全画面フレームを含むフォルダーがそのまま残り続ける。
    # 起動時点で動いている記録は存在しないため、残っているものはすべて取り残し。
    if ([string]::IsNullOrWhiteSpace($script:MbRecordingJobsRoot)) { return 0 }
    if (-not (Test-Path -LiteralPath $script:MbRecordingJobsRoot -PathType Container)) { return 0 }
    $removed = 0
    foreach ($directory in @(Get-ChildItem -LiteralPath $script:MbRecordingJobsRoot -Directory -ErrorAction SilentlyContinue)) {
        # 想定した命名のものだけを消す。利用者が置いた別のフォルダーには触れない。
        if ($directory.Name -notmatch '^record-[a-f0-9]{32}$') { continue }
        try {
            Remove-Item -LiteralPath $directory.FullName -Recurse -Force -ErrorAction Stop
            $removed++
        } catch { }
    }
    return $removed
}

function Get-MbRecordingIdleStatus {
    return [pscustomobject]@{
        jobId = ''; state = 'idle'; count = 0; message = ''; lastTarget = ''; updatedAt = ''
        lastImage = ''; lastResultImage = ''; lastEventIndex = 0; controllerAvailable = $false
    }
}

function Get-MbRecordingLatestPreview {
    if ($null -eq $script:MbRecordingJob) {
        return [pscustomobject]@{ image = ''; resultImage = ''; index = 0 }
    }

    $eventsPath = [string]$script:MbRecordingJob.EventsPath
    $eventsDirectory = [string]$script:MbRecordingJob.EventsDirectory
    if (-not (Test-Path -LiteralPath $eventsPath -PathType Leaf)) {
        return [pscustomobject]@{ image = ''; resultImage = ''; index = 0 }
    }

    try {
        $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
        $stream = [IO.File]::Open($eventsPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
        try {
            $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true, 1024, $false)
            try { $raw = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $stream.Dispose() }

        $lines = @($raw -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            try { $event = $lines[$i] | ConvertFrom-Json } catch { continue }
            if ($null -eq $event -or -not ($event.PSObject.Properties.Name -contains 'image')) { continue }
            $image = [IO.Path]::GetFileName([string]$event.image)
            if ($image -notmatch '^event-\d{3}\.jpg$' -or
                -not (Test-Path -LiteralPath (Join-Path $eventsDirectory $image) -PathType Leaf)) { continue }
            $index = if ($event.PSObject.Properties.Name -contains 'index') { [int]$event.index } else { 0 }
            $resultImage = if ($index -gt 0) { 'event-{0:d3}-result.jpg' -f $index } else { '' }
            if ([string]::IsNullOrWhiteSpace($resultImage) -or
                -not (Test-Path -LiteralPath (Join-Path $eventsDirectory $resultImage) -PathType Leaf)) {
                $resultImage = ''
            }
            return [pscustomobject]@{ image = $image; resultImage = $resultImage; index = $index }
        }
    } catch { }
    return [pscustomobject]@{ image = ''; resultImage = ''; index = 0 }
}

function Read-MbRecordingStatus {
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $statusPath = [string]$script:MbRecordingJob.StatusPath
    $status = $null
    # 原子的な差し替えのごく短い隙間を idle と誤表示しないよう、この要求内で
    # 完成済みJSONを数回読み直す。状態の推測や旧データへの退避は行わない。
    for ($readAttempt = 0; $readAttempt -lt 5 -and $null -eq $status; $readAttempt++) {
        try {
            # 記録ワーカーが完成済みのstatus.jsonを差し替えられるよう、
            # 読み取り中も書き込みと削除（原子的な置換）を共有する。
            $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
            $stream = [IO.File]::Open($statusPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
            try {
                $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true, 1024, $false)
                try {
                    $raw = $reader.ReadToEnd()
                } finally {
                    $reader.Dispose()
                }
            } finally {
                $stream.Dispose()
            }
            $status = $raw | ConvertFrom-Json
        } catch {
            if ($readAttempt -lt 4) { Start-Sleep -Milliseconds 15 }
        }
    }
    if ($null -eq $status) { return (Get-MbRecordingIdleStatus) }
    if ($null -eq $status) { return (Get-MbRecordingIdleStatus) }

    $preview = Get-MbRecordingLatestPreview
    $status | Add-Member -NotePropertyName 'lastImage' -NotePropertyValue ([string]$preview.image) -Force
    $status | Add-Member -NotePropertyName 'lastResultImage' -NotePropertyValue ([string]$preview.resultImage) -Force
    $status | Add-Member -NotePropertyName 'lastEventIndex' -NotePropertyValue ([int]$preview.index) -Force
    $controllerAvailable = $false
    if ($script:MbRecordingJob.PSObject.Properties.Name -contains 'ControllerProcessIdentity' -and
        $null -ne $script:MbRecordingJob.ControllerProcessIdentity) {
        try { $controllerAvailable = Test-MbRecorderProcessIdentity -Identity $script:MbRecordingJob.ControllerProcessIdentity } catch { }
    }
    $status | Add-Member -NotePropertyName 'controllerAvailable' -NotePropertyValue ([bool]$controllerAvailable) -Force

    # 記録プロセスが落ちたまま recording / paused が残らないようにする。
    if ([string]$status.state -in @('starting', 'recording', 'paused')) {
        $alive = $false
        try {
            $alive = if ($script:MbRecordingJob.PSObject.Properties.Name -contains 'ProcessIdentity') {
                Test-MbRecorderProcessIdentity -Identity $script:MbRecordingJob.ProcessIdentity
            } else { $false }
        } catch { $alive = $false }
        if (-not $alive) {
            $status.state = 'failed'
            $status.message = '記録が途中で終わりました。もう一度実行してください。'
        }
        if ($script:MbRecordingJob.PSObject.Properties.Name -contains 'UiaWorkerProcessId') {
            $uiaWorkerAlive = $false
            $uiaWorkerProcessId = [int]$script:MbRecordingJob.UiaWorkerProcessId
            if ($uiaWorkerProcessId -gt 0) {
                try {
                    $uiaWorkerAlive = if ($script:MbRecordingJob.PSObject.Properties.Name -contains 'UiaWorkerProcessIdentity') {
                        Test-MbRecorderProcessIdentity -Identity $script:MbRecordingJob.UiaWorkerProcessIdentity
                    } else { $false }
                } catch { $uiaWorkerAlive = $false }
            }
            if (-not $uiaWorkerAlive) {
                $uiaWarning = 'クリック前のWindows対象監視を開始できなかったため、クリック後の対象検出で記録を続けています。'
                if ($status.PSObject.Properties.Name -contains 'warning' -and
                    -not [string]::IsNullOrWhiteSpace([string]$status.warning)) {
                    $uiaWarning = [string]$status.warning + ' ' + $uiaWarning
                }
                $status | Add-Member -NotePropertyName 'warning' -NotePropertyValue $uiaWarning -Force
            }
        }
    }
    return $status
}

function Start-MbRecordingJob {
    param(
        [string[]]$IgnoreTitlePatterns = @('ManualBuilder', 'ManualBuilder Recorder'),
        [ValidateRange(200, 3000)][int]$ResultCaptureDelayMs = 700
    )

    $current = Read-MbRecordingStatus
    if ([string]$current.state -in @('starting', 'recording', 'paused')) { return $current }

    $capability = Get-MbRecorderCapability
    if (-not $capability.available) { throw ([string]$capability.reason) }

    # 記録モニターはWPF + WebView2版だけを使用する。欠落時に旧UIへ退避せず、
    # 記録を開始する前に明確なエラーとして止める。
    $controllerRoot = Join-Path $script:MbRecordingScriptRoot 'RecorderCompanion'
    $controllerPath = Join-Path $controllerRoot 'Invoke-RecorderCompanion.ps1'
    $controllerSourcePath = Join-Path $controllerRoot 'ManualBuilder.RecorderCompanion.cs'
    $controllerVendorRoot = Join-Path $controllerRoot 'vendor\WebView2'
    $controllerWebRoot = Join-Path $controllerRoot 'web'
    foreach ($requiredPath in @(
        $controllerPath,
        $controllerSourcePath,
        (Join-Path $controllerVendorRoot 'Microsoft.Web.WebView2.Core.dll'),
        (Join-Path $controllerVendorRoot 'Microsoft.Web.WebView2.Wpf.dll'),
        (Join-Path $controllerVendorRoot 'WebView2Loader.dll'),
        (Join-Path $controllerWebRoot 'index.html'),
        (Join-Path $controllerWebRoot 'styles.css'),
        (Join-Path $controllerWebRoot 'app.js')
    )) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw "記録モニターが見つかりません。ManualBuilderを再配置してください: $requiredPath"
        }
    }

    # 前回の記録が残っていれば片付けてから始める。
    Remove-MbRecordingJob
    $script:MbRecorderOcrSnapshotCache = @{}

    $jobId = 'record-' + [guid]::NewGuid().ToString('N')
    $jobDirectory = Join-Path $script:MbRecordingJobsRoot $jobId
    $eventsDirectory = Join-Path $jobDirectory 'events'
    [void](New-Item -ItemType Directory -Path $eventsDirectory -Force)
    $evidenceDirectory = Join-Path $jobDirectory 'evidence'
    [void](New-Item -ItemType Directory -Path $evidenceDirectory -Force)
    $framesDirectory = Join-Path $jobDirectory 'frames'
    [void](New-Item -ItemType Directory -Path $framesDirectory -Force)
    $statusPath = Join-Path $jobDirectory 'status.json'
    $eventsPath = Join-Path $jobDirectory 'events.jsonl'
    $ledgerPath = Join-Path $jobDirectory 'evidence-ledger.jsonl'
    $framesPath = Join-Path $jobDirectory 'frames.jsonl'
    $stopPath = Join-Path $jobDirectory 'stop.requested'
    $pausePath = Join-Path $jobDirectory 'pause.requested'
    $undoPath = Join-Path $jobDirectory 'undo.requested'
    $manualResultPath = Join-Path $jobDirectory 'result.requested'
    $uiaTargetPath = Join-Path $jobDirectory 'uia-target.json'
    $uiaLogPath = Join-Path $jobDirectory 'uia-monitor.log'

    $queued = [pscustomobject]@{
        jobId = $jobId; state = 'starting'; count = 0
        message = '記録の準備をしています'
        lastTarget = ''; updatedAt = [DateTime]::UtcNow.ToString('o')
    }
    [IO.File]::WriteAllText($statusPath, ($queued | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))

    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { throw 'Windows PowerShell 5.1が見つかりません。' }
    $workerPath = Join-Path $script:MbRecordingScriptRoot 'Invoke-ManualBuilderRecorder.ps1'
    $quote = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
    # エクスプローラーや標準ダイアログはクリック直後に対象が消えることがあるため、
    # クリック前のカーソル下を別プロセスで保持する。起動に失敗しても従来のクリック後検索は使える。
    $uiaWorkerProcessId = 0; $uiaWorkerProcessIdentity = $null
    try {
        $uiaWorkerPath = Join-Path $script:MbRecordingScriptRoot 'Invoke-ManualBuilderUiaRecorder.ps1'
        $uiaArguments = @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $uiaWorkerPath),
            '-CachePath', (& $quote $uiaTargetPath),
            '-StopPath', (& $quote $stopPath),
            '-LogPath', (& $quote $uiaLogPath)
        )
        if (@($IgnoreTitlePatterns).Count -gt 0) {
            $uiaArguments += @('-IgnoreTitlePatterns', (& $quote ((@($IgnoreTitlePatterns) -join ','))))
        }
        $uiaWorker = Start-Process -FilePath $powerShellPath -ArgumentList $uiaArguments -WindowStyle Hidden -PassThru
        $uiaWorkerProcessId = [int]$uiaWorker.Id
        $uiaWorkerProcessIdentity = New-MbRecorderProcessIdentity -Process $uiaWorker
        $uiaWorker.Dispose()
    } catch {
        $uiaWorkerProcessId = 0; $uiaWorkerProcessIdentity = $null
    }

    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $workerPath),
        '-EventsDirectory', (& $quote $eventsDirectory),
        '-EventsPath', (& $quote $eventsPath),
        '-EvidenceDirectory', (& $quote $evidenceDirectory),
        '-LedgerPath', (& $quote $ledgerPath),
        '-FramesDirectory', (& $quote $framesDirectory),
        '-FramesPath', (& $quote $framesPath),
        '-StatusPath', (& $quote $statusPath),
        '-StopPath', (& $quote $stopPath),
        '-PausePath', (& $quote $pausePath),
        '-UndoPath', (& $quote $undoPath),
        '-ManualResultPath', (& $quote $manualResultPath),
        '-JobId', (& $quote $jobId),
        '-UiaTargetPath', (& $quote $uiaTargetPath),
        '-ResultCaptureDelayMs', ([string]$ResultCaptureDelayMs)
    )
    if (@($IgnoreTitlePatterns).Count -gt 0) {
        $arguments += '-IgnoreTitlePatterns'
        $arguments += (@($IgnoreTitlePatterns) | ForEach-Object { & $quote $_ }) -join ','
    }

    try {
        $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
        $processId = [int]$worker.Id
        $processIdentity = New-MbRecorderProcessIdentity -Process $worker
        if ($null -eq $processIdentity) { throw '記録プロセスの所有情報を確認できません。' }
        $worker.Dispose()
    } catch {
        try { [IO.File]::WriteAllText($stopPath, 'stop', (New-Object Text.UTF8Encoding($false))) } catch { }
        [void](Stop-MbRecorderOwnedProcess -Identity $uiaWorkerProcessIdentity)
        try { Remove-Item -LiteralPath $jobDirectory -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        throw
    }

    # 対象アプリを操作したまま使えるWebView2記録モニター。UIが準備できなければ
    # 記録だけを裏で続けず、開始処理全体を取り消す。
    $controllerProcessId = 0; $controllerProcessIdentity = $null
    try {
        $controllerArguments = @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA',
            '-File', (& $quote $controllerPath),
            '-StatusPath', (& $quote $statusPath),
            '-EventsDirectory', (& $quote $eventsDirectory),
            '-PausePath', (& $quote $pausePath),
            '-UndoPath', (& $quote $undoPath),
            '-ResultPath', (& $quote $manualResultPath),
            '-StopPath', (& $quote $stopPath),
            '-JobId', (& $quote $jobId),
            '-WebRoot', (& $quote $controllerWebRoot)
        )
        $controller = Start-Process -FilePath $powerShellPath -ArgumentList $controllerArguments `
            -WorkingDirectory $controllerVendorRoot -WindowStyle Hidden -PassThru
        $controllerProcessId = [int]$controller.Id
        $controllerProcessIdentity = New-MbRecorderProcessIdentity -Process $controller
        if ($null -eq $controllerProcessIdentity) { throw '記録モニターの所有情報を確認できません。' }
        $controller.Dispose()
        $controllerReadyPath = $statusPath + '.companion.ready'
        # WebView2 Runtimeの初回初期化は、PC起動直後やウイルス対策ソフトの検査中に
        # 数秒を超えることがある。表示準備を省略せず最大15秒まで待つ。
        $controllerReadyTimeoutMs = 15000
        $controllerReadyIntervalMs = 50
        $controllerReadyAttempts = [Math]::Ceiling($controllerReadyTimeoutMs / $controllerReadyIntervalMs)
        for ($attempt = 0; $attempt -lt $controllerReadyAttempts -and -not (Test-Path -LiteralPath $controllerReadyPath -PathType Leaf); $attempt++) {
            if (-not (Test-MbRecorderProcessIdentity -Identity $controllerProcessIdentity)) { break }
            Start-Sleep -Milliseconds $controllerReadyIntervalMs
        }
        if (-not (Test-Path -LiteralPath $controllerReadyPath -PathType Leaf)) {
            throw '記録レシートを開けませんでした。ManualBuilderを再起動してから、もう一度お試しください。'
        }
    } catch {
        try { [IO.File]::WriteAllText($stopPath, 'stop', (New-Object Text.UTF8Encoding($false))) } catch { }
        [void](Stop-MbRecorderOwnedProcess -Identity $controllerProcessIdentity)
        [void](Stop-MbRecorderOwnedProcess -Identity $processIdentity)
        [void](Stop-MbRecorderOwnedProcess -Identity $uiaWorkerProcessIdentity)
        try { Remove-Item -LiteralPath $jobDirectory -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        throw
    }

    $script:MbRecordingJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $processId; JobDirectory = $jobDirectory
        ProcessIdentity = $processIdentity
        EventsDirectory = $eventsDirectory; EventsPath = $eventsPath
        EvidenceDirectory = $evidenceDirectory; LedgerPath = $ledgerPath
        FramesDirectory = $framesDirectory; FramesPath = $framesPath
        StatusPath = $statusPath; StopPath = $stopPath; PausePath = $pausePath; UndoPath = $undoPath
        ManualResultPath = $manualResultPath; StartedAt = Get-Date
        UiaTargetPath = $uiaTargetPath; UiaLogPath = $uiaLogPath
        UiaWorkerProcessId = $uiaWorkerProcessId
        UiaWorkerProcessIdentity = $uiaWorkerProcessIdentity
        ControllerProcessId = $controllerProcessId
        ControllerProcessIdentity = $controllerProcessIdentity
    }
    # 開始APIが返った直後から最初の操作を拾えるよう、ワーカーがフックを準備して
    # recording を書くまで短時間だけ待つ。遅い環境では starting のまま返し、UIが巡回する。
    $status = Read-MbRecordingStatus
    for ($attempt = 0; $attempt -lt 60 -and [string]$status.state -eq 'starting'; $attempt++) {
        Start-Sleep -Milliseconds 50
        $status = Read-MbRecordingStatus
    }
    return $status
}

function Stop-MbRecordingJob {
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $status = Read-MbRecordingStatus
    if ([string]$status.state -in @('starting', 'recording', 'paused')) {
        [IO.File]::WriteAllText([string]$script:MbRecordingJob.StopPath, 'stop', (New-Object Text.UTF8Encoding($false)))
        # 記録プロセスが停止を見て後始末を終えるまで少しだけ待つ。
        for ($i = 0; $i -lt 40; $i++) {
            Start-Sleep -Milliseconds 100
            $status = Read-MbRecordingStatus
            if ([string]$status.state -notin @('starting', 'recording', 'paused')) { break }
        }
    }
    return $status
}

function Set-MbRecordingPaused {
    param([bool]$Paused)
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $status = Read-MbRecordingStatus
    if ([string]$status.state -notin @('recording', 'paused')) { return $status }
    $pausePath = [string]$script:MbRecordingJob.PausePath
    if ($Paused) {
        [IO.File]::WriteAllText($pausePath, 'pause', (New-Object Text.UTF8Encoding($false)))
    } else {
        Remove-Item -LiteralPath $pausePath -Force -ErrorAction SilentlyContinue
    }
    return (Read-MbRecordingStatus)
}

function Undo-MbLastRecordingEvent {
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $status = Read-MbRecordingStatus
    if ([string]$status.state -notin @('recording', 'paused')) { return $status }
    $previousCount = [int]$status.count
    $requestId = [guid]::NewGuid().ToString('N')
    [IO.File]::WriteAllText([string]$script:MbRecordingJob.UndoPath, $requestId, (New-Object Text.UTF8Encoding($false)))
    # UIの件数を先に減らさない。ワーカーがログと画像を削除してstatusを更新するまで待ち、
    # 確定した件数とプレビューを返す。HTTPサーバーは1本なので連打も直列化される。
    for ($attempt = 0; $attempt -lt 200; $attempt++) {
        Start-Sleep -Milliseconds 50
        $status = Read-MbRecordingStatus
        if ([string]$status.state -notin @('recording', 'paused') -or
            ([string]$status.undoRequestId -eq $requestId -and [int]$status.count -le $previousCount)) { break }
    }
    return $status
}

# 記録した操作を読み出す。取り込む前に一覧を見せて選んでもらうために使う。
function Merge-MbRecordedEditInteractions {
    param(
        [AllowEmptyCollection()][object[]]$Events = @(),
        [int]$MaxGapMs = 5000
    )

    function Test-SamePosition {
        param($First, $Second, [double]$Tolerance = 0.045)
        if ($null -eq $First -or $null -eq $Second) { return $false }
        if ($First.PSObject.Properties.Name -contains 'rect' -and $null -ne $First.rect -and
            $Second.PSObject.Properties.Name -contains 'rect' -and $null -ne $Second.rect) {
            $overlapWidth = [Math]::Min([double]$First.rect.x2, [double]$Second.rect.x2) -
                [Math]::Max([double]$First.rect.x1, [double]$Second.rect.x1)
            $overlapHeight = [Math]::Min([double]$First.rect.y2, [double]$Second.rect.y2) -
                [Math]::Max([double]$First.rect.y1, [double]$Second.rect.y1)
            if ($overlapWidth -gt 0.0 -and $overlapHeight -gt 0.0) { return $true }
        }
        if ($First.PSObject.Properties.Name -contains 'clickPoint' -and $null -ne $First.clickPoint -and
            $Second.PSObject.Properties.Name -contains 'clickPoint' -and $null -ne $Second.clickPoint) {
            return [Math]::Abs([double]$First.clickPoint.x - [double]$Second.clickPoint.x) -le $Tolerance -and
                [Math]::Abs([double]$First.clickPoint.y - [double]$Second.clickPoint.y) -le $Tolerance
        }
        return $false
    }

    function Test-HasPosition {
        param($Event)
        return $null -ne $Event -and ((
            $Event.PSObject.Properties.Name -contains 'rect' -and $null -ne $Event.rect) -or (
            $Event.PSObject.Properties.Name -contains 'clickPoint' -and $null -ne $Event.clickPoint))
    }

    function Test-SelfInteraction {
        param($Event)
        if ($null -eq $Event) { return $false }
        $name = ([string]$Event.targetName).Trim()
        $type = [string]$Event.targetType
        # 記録開始ボタンと、停止するためにManualBuilderタブへ戻るクリックは
        # 作業手順ではない。ブラウザのタイトルが前タブのままでも対象名から除外する。
        if ($type -eq 'ControlType.TabItem' -and $name -match '^ManualBuilder(?:\s|$)') { return $true }
        return $name -in @('操作を記録して手順書を作る', '操作の記録を開始', '記録を停止', 'ManualBuilderへ戻る')
    }

    function Get-EvidenceScore {
        param($Event)
        $score = 0
        if (-not [string]::IsNullOrWhiteSpace([string]$Event.targetName)) { $score += 4 }
        if ([string]$Event.targetSource -notin @('', 'click-point')) { $score += 2 }
        if ([string]$Event.confidence -eq 'high') { $score += 2 }
        elseif ([string]$Event.confidence -eq 'medium') { $score += 1 }
        return $score
    }

    # 1回のタップ・ダブルクリックがUIAと押下履歴の両方から複数件に見えることがある。
    # 同じウィンドウの近接点で短時間に続くクリックは1操作にし、最も根拠の強い対象と
    # 最後の操作後画像を残す。メニュー項目など別位置のクリックは統合しない。
    $collapsed = New-Object System.Collections.ArrayList
    foreach ($event in @($Events)) {
        if (Test-SelfInteraction -Event $event) { continue }
        $previous = if ($collapsed.Count -gt 0) { $collapsed[$collapsed.Count - 1] } else { $null }
        $clickKinds = @('click', 'right-click')
        $canCollapse = $null -ne $previous -and [string]$previous.kind -in $clickKinds -and
            [string]$event.kind -eq [string]$previous.kind -and
            [string]$event.windowTitle -eq [string]$previous.windowTitle -and
            ([int]$event.timeMs - [int]$previous.timeMs) -ge 0 -and
            ([int]$event.timeMs - [int]$previous.timeMs) -le 450 -and
            (Test-SamePosition -First $previous -Second $event)
        if ($canCollapse) {
            if ((Get-EvidenceScore -Event $event) -gt (Get-EvidenceScore -Event $previous)) {
                foreach ($propertyName in @('targetName', 'targetType', 'targetSource', 'confidence', 'rect',
                        'targetCandidates', 'targetCandidateId', 'clickPoint', 'captureRegion')) {
                    if ($event.PSObject.Properties.Name -contains $propertyName) {
                        $previous | Add-Member -NotePropertyName $propertyName -NotePropertyValue $event.$propertyName -Force
                    }
                }
            }
            if ($event.PSObject.Properties.Name -contains 'resultImage' -and
                -not [string]::IsNullOrWhiteSpace([string]$event.resultImage)) {
                $previous | Add-Member -NotePropertyName resultImage -NotePropertyValue ([string]$event.resultImage) -Force
            }
            # 次の入力との間隔は、クラスタの最初ではなく最後のクリックから判定する。
            $previous.timeMs = [int]$event.timeMs
            continue
        }
        [void]$collapsed.Add($event)
    }

    # セルや入力欄をクリックしてそのまま入力した場合は、入力済み画面の1手順だけを残す。
    # トリプルクリック相当の連続操作も、同じ欄への入力ならまとめて除く。
    $result = New-Object System.Collections.ArrayList
    foreach ($current in @($collapsed)) {
        if ([string]$current.kind -eq 'input') {
            while ($result.Count -gt 0) {
                $candidate = $result[$result.Count - 1]
                $sameWindow = [string]$candidate.windowTitle -eq [string]$current.windowTitle
                $samePosition = Test-SamePosition -First $candidate -Second $current -Tolerance 0.02
                $isEditClick = [string]$candidate.kind -eq 'click' -and (
                    [string]$candidate.targetType -in @('ControlType.Edit', 'ControlType.DataItem') -or
                    ([string]$candidate.targetType -eq 'ControlType.ClickPoint' -and $samePosition))
                $sameNamedField = -not [string]::IsNullOrWhiteSpace([string]$candidate.targetName) -and
                    [string]$candidate.targetName -eq [string]$current.targetName
                $sameUnknownField = [string]::IsNullOrWhiteSpace([string]$candidate.targetName) -and
                    [string]::IsNullOrWhiteSpace([string]$current.targetName) -and $samePosition
                $sameField = $sameWindow -and (($sameNamedField -and ($samePosition -or
                        (-not (Test-HasPosition -Event $candidate) -and -not (Test-HasPosition -Event $current)))) -or
                    $sameUnknownField)
                $gapMs = [int]$current.timeMs - [int]$candidate.timeMs
                if (-not ($isEditClick -and $sameField -and $gapMs -ge 0 -and $gapMs -le $MaxGapMs)) { break }
                $result.RemoveAt($result.Count - 1)
            }
        }
        [void]$result.Add($current)
    }
    return @($result)
}

function ConvertTo-MbRecorderOcrLabel {
    param([AllowEmptyString()][string]$Value = '')

    $label = [regex]::Replace([string]$Value, '\s+', ' ').Trim()
    # Windows OCRは大きな日本語ボタンを「詳 細 を 表 示」のように1文字ずつ
    # 分けることがある。英単語間の空白は維持し、日本語同士だけをつなぐ。
    $japanese = '一-龯々ぁ-んァ-ヶー'
    $label = [regex]::Replace($label, "(?<=[$japanese])\s+(?=[$japanese])", '')
    # 横棒を漢数字の「一」と読み違える既知ケース（一覧へ戻る）を限定補正する。
    $label = [regex]::Replace($label, '^[\-‐‑‒–—―−]\s*覧(?=へ|に|を|で|$)', '一覧')
    if ($label.Length -gt 80) { $label = $label.Substring(0, 79).TrimEnd() + '…' }
    return $label
}

function Test-MbRecorderUnreliableOcrLabel {
    param(
        [AllowEmptyString()][string]$Label = '',
        [AllowEmptyString()][string]$ActionKind = '',
        [AllowEmptyString()][string]$Source = ''
    )
    if ($Source -notmatch '(?i)OCR') { return $false }
    $value = ([string]$Label).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { return $true }
    # 入力欄内OCRは、項目名ではなく入力済みの値を読むことが多い。
    if ($ActionKind -eq 'input') { return $true }
    # Excelエラーや1文字だけの断片を、ボタン名・セル名として断定しない。
    if ($value -match '^#[A-Z0-9/]+[!?]$' -or $value.Length -le 1) { return $true }
    return $false
}

function Get-MbRecorderOcrLabelConfidence {
    param([AllowEmptyString()][string]$Label = '')

    if ([string]::IsNullOrWhiteSpace($Label)) { return 'low' }
    $hasJapanese = $Label -match '[一-龯々ぁ-んァ-ヶー]'
    $latinTokens = @([regex]::Matches($Label, '[A-Za-z]+') | ForEach-Object { [string]$_.Value })
    if (-not $hasJapanese -or $latinTokens.Count -eq 0) { return 'medium' }
    # 日本語OCRが英字を「自 i ロ n」のような日英1文字の交互列へ崩す場合は、
    # 文章を確定扱いにせず要確認へ残す。CSV、PDFなど複数文字の実用語は維持する。
    if (@($latinTokens | Where-Object { $_.Length -ge 2 }).Count -eq 0) { return 'low' }
    return 'medium'
}

function Get-MbRecorderOcrSnapshotCached {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $file = Get-Item -LiteralPath $Path
    $key = $file.FullName + '|' + [string]$file.Length + '|' + [string]$file.LastWriteTimeUtc.Ticks
    if ($script:MbRecorderOcrSnapshotCache.ContainsKey($key)) {
        return $script:MbRecorderOcrSnapshotCache[$key]
    }
    $snapshot = Get-MbOcrSnapshot -Path $file.FullName
    # 1回の記録は最大300件。別ジョブを繰り返してもメモリを増やし続けない。
    if ($script:MbRecorderOcrSnapshotCache.Count -ge 320) { $script:MbRecorderOcrSnapshotCache = @{} }
    $script:MbRecorderOcrSnapshotCache[$key] = $snapshot
    return $snapshot
}

function Add-MbRecorderLocalOcrEvidence {
    param(
        [AllowEmptyCollection()][object[]]$Events = @(),
        [Parameter(Mandatory = $true)][string]$EventsDirectory
    )

    foreach ($record in @($Events)) {
        if ($null -eq $record) { continue }
        $name = if ($record.PSObject.Properties.Name -contains 'targetName') { [string]$record.targetName } else { '' }
        $source = if ($record.PSObject.Properties.Name -contains 'targetSource') { [string]$record.targetSource } else { '' }
        $type = if ($record.PSObject.Properties.Name -contains 'targetType') { [string]$record.targetType } else { '' }
        $isClickPoint = $source -eq 'click-point' -or $type -eq 'ControlType.ClickPoint'
        if (-not [string]::IsNullOrWhiteSpace($name) -or -not $isClickPoint -or
            $record.PSObject.Properties.Name -notcontains 'rect' -or $null -eq $record.rect -or
            -not (Test-MbNormalizedRect -Rect $record.rect)) { continue }
        $imageName = if ($record.PSObject.Properties.Name -contains 'image') { [string]$record.image } else { '' }
        if ($imageName -notmatch '^event-\d{3}\.jpg$') { continue }

        $snapshot = Get-MbRecorderOcrSnapshotCached -Path (Join-Path $EventsDirectory $imageName)
        if ($null -eq $snapshot -or -not [bool]$snapshot.available) { continue }
        $resolved = Resolve-MbOperationRect -Rect $record.rect -Snapshot $snapshot -NearestLimit 0.055
        $label = ConvertTo-MbRecorderOcrLabel -Value ([string]$resolved.label)
        if ([string]::IsNullOrWhiteSpace($label) -or [string]$resolved.matched -notin @('inside', 'nearest')) { continue }
        $ocrConfidence = Get-MbRecorderOcrLabelConfidence -Label $label

        $ocrCandidate = [pscustomobject]@{
            id = 'ocr-click-1'
            source = 'click-point+OCR'
            confidence = $ocrConfidence
            label = $label
            targetType = 'ControlType.OcrText'
            rect = $resolved.rect
        }
        $existing = if ($record.PSObject.Properties.Name -contains 'targetCandidates') { @($record.targetCandidates) } else { @() }
        $remaining = @($existing | Where-Object { $null -ne $_ -and [string]$_.id -ne 'ocr-click-1' } | Select-Object -First 3)
        $record | Add-Member -NotePropertyName targetName -NotePropertyValue $label -Force
        $record | Add-Member -NotePropertyName targetSource -NotePropertyValue 'click-point+OCR' -Force
        $record | Add-Member -NotePropertyName confidence -NotePropertyValue $ocrConfidence -Force
        $record | Add-Member -NotePropertyName targetType -NotePropertyValue 'ControlType.OcrText' -Force
        $record | Add-Member -NotePropertyName rect -NotePropertyValue $resolved.rect -Force
        $record | Add-Member -NotePropertyName targetCandidateId -NotePropertyValue 'ocr-click-1' -Force
        $record | Add-Member -NotePropertyName targetCandidates -NotePropertyValue (@($ocrCandidate) + @($remaining)) -Force
    }
    return @($Events)
}

function Repair-MbRecorderUnlabeledInputAnchors {
    param(
        [AllowEmptyCollection()][object[]]$Events = @(),
        [int]$MaximumGapMs = 5000
    )

    $ordered = @($Events | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    for ($index = 1; $index -lt $ordered.Count; $index++) {
        $current = $ordered[$index]
        $previous = $ordered[$index - 1]
        if ([string]$current.kind -ne 'input' -or
            -not [string]::IsNullOrWhiteSpace([string]$current.targetName) -or
            [string]$previous.kind -notin @('click', 'double-click') -or
            [string]::IsNullOrWhiteSpace([string]$previous.targetName) -or
            [string]$current.windowTitle -ne [string]$previous.windowTitle) { continue }
        $gap = [int]$current.timeMs - [int]$previous.timeMs
        if ($gap -lt 0 -or $gap -gt $MaximumGapMs) { continue }
        $previousType = [string]$previous.targetType
        if ($previousType -notin @('ControlType.Edit', 'ControlType.DataItem', 'ControlType.OcrText')) { continue }

        foreach ($property in @('targetName', 'targetSource', 'confidence', 'rect', 'targetCandidateId', 'targetCandidates', 'clickPoint')) {
            if ($previous.PSObject.Properties.Name -contains $property) {
                $current | Add-Member -NotePropertyName $property -NotePropertyValue $previous.$property -Force
            }
        }
        # 文字入力が実際に続いたため、OCR文字は単なる画面ラベルではなく入力欄のラベルである。
        if ($previousType -eq 'ControlType.OcrText') {
            $previous.targetType = 'ControlType.Edit'
            $current | Add-Member -NotePropertyName targetType -NotePropertyValue 'ControlType.Edit' -Force
            foreach ($event in @($previous, $current)) {
                if ($event.PSObject.Properties.Name -contains 'targetCandidates') {
                    foreach ($candidate in @($event.targetCandidates | Where-Object { [string]$_.id -eq 'ocr-click-1' })) {
                        $candidate.targetType = 'ControlType.Edit'
                    }
                }
            }
        } else {
            $current | Add-Member -NotePropertyName targetType -NotePropertyValue $previousType -Force
        }
    }
    return @($ordered)
}

function Get-MbRecordedRawEvents {
    if ($null -eq $script:MbRecordingJob) { return @() }
    $eventsPath = [string]$script:MbRecordingJob.EventsPath
    if (-not (Test-Path -LiteralPath $eventsPath -PathType Leaf)) { return @() }

    $events = New-Object System.Collections.ArrayList
    foreach ($line in [IO.File]::ReadAllLines($eventsPath, [Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $record = $null
        try { $record = $line | ConvertFrom-Json } catch { continue }
        if ($null -eq $record) { continue }
        $resultFileName = ('event-{0:d3}-result.jpg' -f [int]$record.index)
        $resultPath = Join-Path ([string]$script:MbRecordingJob.EventsDirectory) $resultFileName
        if (Test-Path -LiteralPath $resultPath -PathType Leaf) {
            $record | Add-Member -NotePropertyName resultImage -NotePropertyValue $resultFileName -Force
        }
        [void]$events.Add($record)
    }
    # UIAを取得できなかったクリックは、外部送信しないWindows OCRと実クリック位置を
    # 組み合わせる。ページ全体の文字から推測せず、クリック枠の中か直近の語だけを採用する。
    $events = @(Add-MbRecorderLocalOcrEvidence -Events @($events) `
        -EventsDirectory ([string]$script:MbRecordingJob.EventsDirectory))
    # 直後に入力が続いた場合、そのクリックが入力欄だったことは操作列から確定できる。
    $events = @(Repair-MbRecorderUnlabeledInputAnchors -Events @($events))
    # AI workerだけで補助アンカーを復元しても、取り込み時に原本eventsを
    # 読み直すと赤枠が消える。同じ保守的なExcel補間を表示・取り込み側にも適用する。
    return @(Repair-MbRecorderExcelInputEventAnchors -Events @($events))
}

function Get-MbRecordedEvents {
    return @(Merge-MbRecordedEditInteractions -Events @(Get-MbRecordedRawEvents))
}

function Get-MbRecordedFrames {
    if ($null -eq $script:MbRecordingJob) { return @() }
    $framesPath = [string]$script:MbRecordingJob.FramesPath
    if (-not (Test-Path -LiteralPath $framesPath -PathType Leaf)) { return @() }
    $frames = New-Object System.Collections.ArrayList
    foreach ($line in [IO.File]::ReadAllLines($framesPath, [Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $frame = $line | ConvertFrom-Json
            if ($null -ne $frame -and [string]$frame.id -match '^F\d{5}$' -and
                [string]$frame.image -match '^frame-\d{5}\.jpg$') { [void]$frames.Add($frame) }
        } catch { }
    }
    return @($frames)
}

function Get-MbRecorderTransformationReason {
    param(
        [AllowEmptyCollection()][object[]]$Events = @(),
        [AllowEmptyString()][string]$ActionKind = '',
        [bool]$HasAfterImage = $false
    )
    $kinds = @($Events | ForEach-Object { [string]$_.kind } | Where-Object { $_ })
    if ($ActionKind -eq 'visual-change' -and $Events.Count -eq 0) {
        return '操作イベントが無い区間ですが、画面変化を証拠候補として残しました。'
    }
    $basis = if ($Events.Count -le 1) {
        '1件の操作証拠から1手順を作りました。'
    } elseif ($kinds -contains 'input' -and @($kinds | Where-Object { $_ -in @('click', 'right-click') }).Count -gt 0) {
        '入力欄を選んだ操作と、その直後の入力を1手順にまとめました。'
    } else {
        ("{0}件の連続した操作証拠を、同じ目的の1手順としてまとめました。" -f $Events.Count)
    }
    if ($HasAfterImage) { return $basis + ' 操作後に安定した画面も結果候補として保持しています。' }
    return $basis
}

# UI Automation が返した名前と矩形を、クリック座標で独立に照合できた場合だけ
# 「そのまま使える」候補へ昇格する。DOM/OCR/推定UIAや入力操作は対象外にし、
# 自動採用を増やすために安全条件を緩めない。
function Test-MbRecorderTrustedLocalTarget {
    param(
        [AllowNull()]$Event,
        [AllowEmptyString()][string]$ActionKind = ''
    )
    if ($null -eq $Event -or $ActionKind -in @('input', 'visual-change')) { return $false }
    $source = if ($Event.PSObject.Properties.Name -contains 'targetSource') { [string]$Event.targetSource } else { '' }
    $confidence = if ($Event.PSObject.Properties.Name -contains 'confidence') { [string]$Event.confidence } else { '' }
    $label = if ($Event.PSObject.Properties.Name -contains 'targetName') { [string]$Event.targetName } else { '' }
    if ($source -notin @('UIA', 'UIA-CACHE') -or $confidence -notin @('medium', 'high')) { return $false }
    if ([string]::IsNullOrWhiteSpace($label) -or
        (Test-MbRecorderUnreliableOcrLabel -Label $label -ActionKind $ActionKind -Source $source)) {
        return $false
    }
    $allowedTypes = @(
        'ControlType.Button', 'ControlType.SplitButton', 'ControlType.Hyperlink', 'ControlType.MenuItem',
        'ControlType.ListItem', 'ControlType.ComboBox', 'ControlType.TabItem', 'ControlType.DataItem',
        'ControlType.CheckBox', 'ControlType.RadioButton', 'ControlType.TreeItem'
    )
    $targetType = if ($Event.PSObject.Properties.Name -contains 'targetType') { [string]$Event.targetType } else { '' }
    if ($targetType -notin $allowedTypes -or
        $Event.PSObject.Properties.Name -notcontains 'rect' -or
        $Event.PSObject.Properties.Name -notcontains 'clickPoint' -or
        $null -eq $Event.rect -or $null -eq $Event.clickPoint) { return $false }
    try {
        $x1 = [double]$Event.rect.x1; $y1 = [double]$Event.rect.y1
        $x2 = [double]$Event.rect.x2; $y2 = [double]$Event.rect.y2
        $x = [double]$Event.clickPoint.x; $y = [double]$Event.clickPoint.y
        $width = $x2 - $x1; $height = $y2 - $y1
        if ($x1 -lt 0 -or $y1 -lt 0 -or $x2 -gt 1 -or $y2 -gt 1 -or
            $width -le 0 -or $height -le 0 -or ($width * $height) -gt 0.08) { return $false }
        return $x -ge $x1 -and $x -le $x2 -and $y -ge $y1 -and $y -le $y2
    } catch { return $false }
}

function Get-MbRecordedLocalProposals {
    $events = @(Get-MbRecordedRawEvents)
    $frames = @(Get-MbRecordedFrames)
    if ($null -ne $script:MbRecordingJob -and $frames.Count -gt 0) {
        $frames = @(Add-MbRecorderFrameVisualMetrics -Frames $frames -FramesDirectory ([string]$script:MbRecordingJob.FramesDirectory))
    }
    $candidates = @(New-MbRecorderLocalFrameCandidates -Frames $frames -Events $events -MaximumFrames 30)
    if ($candidates.Count -lt 1) { return @() }
    $eventMap = @{}
    foreach ($event in $events) { $eventMap[[int]$event.index] = $event }
    $result = New-Object System.Collections.ArrayList
    foreach ($candidate in $candidates) {
        $groupEvents = @($candidate.eventIds | ForEach-Object {
            $id = [int]$_
            if ($eventMap.ContainsKey($id)) { $eventMap[$id] }
        } | Where-Object { $null -ne $_ })
        $draftEvent = @($groupEvents | Where-Object {
            [string]$_.kind -eq 'input' -and -not [string]::IsNullOrWhiteSpace([string]$_.targetName)
        } | Select-Object -Last 1)
        if (@($draftEvent).Count -gt 0) { $draftEvent = @($draftEvent)[0] }
        elseif ($eventMap.ContainsKey([int]$candidate.targetEventId)) { $draftEvent = $eventMap[[int]$candidate.targetEventId] }
        elseif ($groupEvents.Count -gt 0) { $draftEvent = $groupEvents[0] }
        else { $draftEvent = $null }

        $targetName = if ($null -ne $draftEvent -and $draftEvent.PSObject.Properties.Name -contains 'targetName') { [string]$draftEvent.targetName } else { '' }
        $targetType = if ($null -ne $draftEvent -and $draftEvent.PSObject.Properties.Name -contains 'targetType') { [string]$draftEvent.targetType } else { '' }
        $windowTitle = if ($null -ne $draftEvent -and $draftEvent.PSObject.Properties.Name -contains 'windowTitle') { [string]$draftEvent.windowTitle } else { '' }
        $targetSource = if ($null -ne $draftEvent -and $draftEvent.PSObject.Properties.Name -contains 'targetSource') { [string]$draftEvent.targetSource } else { '' }
        $targetConfidence = if ($null -ne $draftEvent -and $draftEvent.PSObject.Properties.Name -contains 'confidence') { [string]$draftEvent.confidence } else { '' }
        $isVisualChange = [string]$candidate.actionKind -eq 'visual-change'
        # 赤枠アンカーを選べなかった入力では、イベントに残った古いセル名を文章へ
        # 使わない。OCRで入力値を項目名として読んだ場合も同様に安全な文言へ退避する。
        $hasTrustedDraftAnchor = [int]$candidate.targetEventId -gt 0
        if (([string]$candidate.actionKind -eq 'input' -and -not $hasTrustedDraftAnchor) -or
            (Test-MbRecorderUnreliableOcrLabel -Label $targetName -ActionKind ([string]$candidate.actionKind) -Source $targetSource)) {
            $targetName = ''
            $targetType = ''
            $targetSource = if ($hasTrustedDraftAnchor) { $targetSource } else { '' }
            $targetConfidence = 'low'
        }
        if (Test-MbRecorderTrustedLocalTarget -Event $draftEvent -ActionKind ([string]$candidate.actionKind)) {
            $targetConfidence = 'high'
        }
        $draft = if ($isVisualChange) {
            [pscustomobject]@{
                title = '画面の変化を確認'
                description = '操作前後を比較し、必要な操作内容を確認します。'
                reviewRequired = $true
                reviewReason = '操作イベントが欠けた区間を画面変化から補いました。文章と赤枠を確認してください。'
            }
        } else {
            Get-MbLocalStepDraft -ActionKind ([string]$candidate.actionKind) -TargetName $targetName `
                -TargetType $targetType -WindowTitle $windowTitle -TargetSource $targetSource `
                -TargetConfidence $targetConfidence
        }
        $evidenceIds = @($groupEvents | ForEach-Object {
            if ($_.PSObject.Properties.Name -contains 'evidenceId' -and
                [string]$_.evidenceId -match '^evidence-[a-f0-9]{32}$') { [string]$_.evidenceId }
        } | Where-Object { $_ } | Select-Object -Unique)
        $transformationReason = Get-MbRecorderTransformationReason -Events $groupEvents `
            -ActionKind ([string]$candidate.actionKind) -HasAfterImage (-not [string]::IsNullOrWhiteSpace([string]$candidate.afterImage))
        [void]$result.Add([pscustomobject]@{
            id = [string]$candidate.id
            beforeFrame = [string]$candidate.beforeFrame
            afterFrame = [string]$candidate.afterFrame
            eventIds = @($candidate.eventIds)
            evidenceIds = @($evidenceIds)
            sourceOperationCount = $groupEvents.Count
            transformationReason = $transformationReason
            targetEventId = [int]$candidate.targetEventId
            actionKind = [string]$candidate.actionKind
            title = [string]$draft.title
            description = [string]$draft.description
            reviewRequired = [bool]$draft.reviewRequired
            confidence = $(if ([bool]$draft.reviewRequired) { 'low' } else { 'high' })
            reason = $(if ([bool]$draft.reviewRequired) { [string]$draft.reviewReason } else { 'このPCで記録した操作前後から作成しました。' })
            timeMs = [int]$candidate.timeMs
            beforeImage = [string]$candidate.beforeImage
            afterImage = [string]$candidate.afterImage
            source = 'local'
        })
    }
    # ローカル候補でも、AI経路と同じ安全補正を通す。読込中の操作後画像や
    # Excelの選択範囲を、確認画面へ出す前に可能な範囲で修復する。
    $repaired = @(Repair-MbRecorderTransientAfterFrames -Frames $frames -Events $events -Proposals @($result))
    $repaired = @(Expand-MbRecorderExcelRangeSelectionProposals -Events $events -Proposals $repaired)
    return @($repaired)
}

function Get-MbRecordedEventImagePath {
    param([Parameter(Mandatory = $true)][string]$FileName)
    if ($null -eq $script:MbRecordingJob) { return '' }
    if ($FileName -notmatch '^event-\d{3}(?:-result)?\.jpg$') { return '' }
    $path = Join-Path ([string]$script:MbRecordingJob.EventsDirectory) $FileName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    return $path
}

function New-MbRecorderAnnotationId {
    return 'annotation-' + [guid]::NewGuid().ToString('N')
}

function Test-MbRecordedSelectionAnchorContext {
    param(
        [AllowNull()]$Event,
        [Parameter(Mandatory = $true)]$BeforeFrame,
        [AllowNull()]$AfterFrame
    )
    if ($null -eq $Event -or $Event.PSObject.Properties.Name -notcontains 'rect' -or
        $null -eq $Event.rect -or -not (Test-MbNormalizedRect -Rect $Event.rect)) { return $false }
    $beforeApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$BeforeFrame.windowTitle)
    $eventApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$Event.windowTitle)
    if ([string]::IsNullOrWhiteSpace($beforeApp) -or $beforeApp -ne $eventApp) { return $false }
    $beforeTime = [int]$BeforeFrame.timeMs; $eventTime = [int]$Event.timeMs
    if ([Math]::Abs($eventTime - $beforeTime) -gt 2500) { return $false }
    if ($null -ne $AfterFrame) {
        $afterApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$AfterFrame.windowTitle)
        $afterTime = [int]$AfterFrame.timeMs
        if ([string]::IsNullOrWhiteSpace($afterApp) -or $afterApp -ne $beforeApp -or
            $afterTime -le $beforeTime -or $eventTime -gt ($afterTime + 1000)) { return $false }
    }
    return $true
}

# 記録した操作を手順にする。
#
# 録画からの取り込みと違い、赤枠はUI Automationの矩形、またはUIA非対応画面の
# 小さなクリック位置枠である。文字認識や全画面矩形への推定は行わず、受け取った矩形を使う。
function Import-MbRecordedEvents {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [AllowEmptyString()][string]$SelectionJson = ''
    )

    $events = @(Get-MbRecordedEvents)
    if ($events.Count -eq 0) { return [pscustomobject]@{ added = 0; skipped = 0; generated = 0; needsReview = 0 } }

    # 選択が渡されていれば、その番号だけを取り込む。
    $wanted = $null
    if (-not [string]::IsNullOrWhiteSpace($SelectionJson)) {
        $parsed = $null
        try { $parsed = $SelectionJson | ConvertFrom-Json } catch { throw '取り込む操作の指定を読み取れません。' }
        if ($null -ne $parsed) {
            $items = @($parsed)
            if ($parsed.PSObject.Properties.Name -contains 'accept') { $items = @($parsed.accept) }
            $wanted = New-Object 'System.Collections.Generic.HashSet[int]'
            foreach ($item in $items) {
                $value = 0
                try { $value = [int]$item } catch { $value = 0 }
                if ($value -gt 0) { [void]$wanted.Add($value) }
            }
        }
    }

    $evidenceSession = Import-MbRecordedEvidenceSession -Project $Project -ProjectPath $ProjectPath
    $sessionId = if ($null -ne $evidenceSession) { [string]$evidenceSession.id } else { '' }
    $acceptedIndexes = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($record in $events) {
        $recordIndex = 0
        try { $recordIndex = [int]$record.index } catch { $recordIndex = 0 }
        if ($recordIndex -gt 0 -and ($null -eq $wanted -or $wanted.Contains($recordIndex))) {
            [void]$acceptedIndexes.Add($recordIndex)
        }
    }
    Save-MbRecordedRawTransformationDecisions -ProjectPath $ProjectPath -SessionId $sessionId `
        -Events $events -AcceptedIndexes $acceptedIndexes

    $added = 0
    $skipped = 0
    $generated = 0
    $needsReview = 0
    foreach ($record in $events) {
        $index = 0
        try { $index = [int]$record.index } catch { $index = 0 }
        if ($null -ne $wanted -and -not $wanted.Contains($index)) { continue }

        $imagePath = Get-MbRecordedEventImagePath -FileName ([string]$record.image)
        if ([string]::IsNullOrWhiteSpace($imagePath)) { $skipped++; continue }

        $recordTargetType = if ($record.PSObject.Properties.Name -contains 'targetType') { [string]$record.targetType } else { '' }
        $resultImagePath = ''
        if ($record.PSObject.Properties.Name -contains 'resultImage' -and
            -not [string]::IsNullOrWhiteSpace([string]$record.resultImage)) {
            $resultImagePath = Get-MbRecordedEventImagePath -FileName ([string]$record.resultImage)
        }
        # Excelのセル選択は、操作前後を2枚並べても情報が増えない。起動中の灰色画面を
        # 残すより、セルとシートが表示された操作後画面を1枚だけ使う。
        $preferResultAsPrimary = $recordTargetType -eq 'ControlType.DataItem' -and
            -not [string]::IsNullOrWhiteSpace($resultImagePath)
        if ($preferResultAsPrimary) { $imagePath = $resultImagePath }

        $bytes = [IO.File]::ReadAllBytes($imagePath)
        # 同じ静止画でも「入力する」「送信を押す」のように別の操作が続くことがある。
        # 画像ファイルは共有しつつ、選ばれた操作はそれぞれ別の手順として残す。
        $result = Add-MbImageStep -Project $Project -ProjectPath $ProjectPath -SheetId $SheetId -Bytes $bytes `
            -Source 'recorder' -AllowDuplicateStep
        if ($result.Status -ne 'added') {
            # 将来別の状態が増えても、不完全な手順は作らない。
            $skipped++
            continue
        }
        $stepId = [string]$result.Step.id

        # クリック後の安定画面は、操作箇所を示す画像とは別の正式な画像資産にする。
        # 取得に失敗した古い録画も、そのまま取り込める。
        if (-not $preferResultAsPrimary -and -not [string]::IsNullOrWhiteSpace($resultImagePath)) {
                try {
                    $resultAsset = Add-MbImageAsset -Project $Project -ProjectPath $ProjectPath `
                        -Bytes ([IO.File]::ReadAllBytes($resultImagePath)) -Source 'recorder'
                    $result.Step.resultImageId = [string]$resultAsset.Image.id
                    # 結果画像は失わず保存するが、初稿では案内画像1枚を使う。
                    $result.Step.imageLayout = 'before'
                    $result.Step.imageOrder = 'before-after'
                    $result.Step.updatedAt = [DateTime]::UtcNow.ToString('o')
                } catch {
                    # 結果画像だけ壊れていても、クリック前画像の手順は取り込む。
                }
        }

        $rect = $null
        if ($record.PSObject.Properties.Name -contains 'rect') { $rect = $record.rect }
        $annotation = @()
        if ($null -ne $rect -and (Test-MbNormalizedRect -Rect $rect)) {
            $annotation = @([pscustomobject]@{
                id    = New-MbRecorderAnnotationId
                type  = 'rect'
                x1    = [Math]::Round([double]$rect.x1, 6)
                y1    = [Math]::Round([double]$rect.y1, 6)
                x2    = [Math]::Round([double]$rect.x2, 6)
                y2    = [Math]::Round([double]$rect.y2, 6)
                label = 0
            })
        }

        $targetType = $recordTargetType
        # 記録時は画面全体を残す。対象だけへ自動で寄ると、誤検出時に周辺の文脈まで失われるため、
        # 切り抜きは利用者が編集画面で明示的に行う。
        if (@($annotation).Count -gt 0) {
            [void](Set-MbStepAnnotations -Project $Project -StepId $stepId -AnnotationsJson (ConvertTo-Json -InputObject $annotation -Depth 5))
        }

        $kind = if ([string]$record.kind -eq 'input') { 'recorded-input' } else { 'recorded-click' }
        $targetName = [string]$record.targetName
        $suffix = ''
        # 入力の手順は「どの欄に入れたか」を示す。入力した文字は記録していない。
        if ($kind -eq 'recorded-input' -and -not [string]::IsNullOrWhiteSpace($targetName)) {
            $suffix = '（入力）'
        } elseif ([string]$record.kind -eq 'right-click' -and -not [string]::IsNullOrWhiteSpace($targetName)) {
            $suffix = '（右クリック）'
        }
        $targetName = ConvertTo-MbRecorderTargetName -Value $targetName -Suffix $suffix
        $targetSource = if ($record.PSObject.Properties.Name -contains 'targetSource') { [string]$record.targetSource } else { '' }
        $targetConfidence = if ($record.PSObject.Properties.Name -contains 'confidence' -and
            [string]$record.confidence -in @('high', 'medium', 'low')) { [string]$record.confidence } else { '' }
        $candidateId = if ($record.PSObject.Properties.Name -contains 'targetCandidateId') { [string]$record.targetCandidateId } else { '' }
        $recordedCandidates = New-Object System.Collections.ArrayList
        if ($record.PSObject.Properties.Name -contains 'targetCandidates') {
            foreach ($candidate in @($record.targetCandidates | Select-Object -First 4)) {
                if ($null -eq $candidate -or $candidate.PSObject.Properties.Name -notcontains 'id' -or
                    $candidate.PSObject.Properties.Name -notcontains 'rect') { continue }
                $candidateLabel = if ($candidate.PSObject.Properties.Name -contains 'label') { [string]$candidate.label } else { '' }
                $candidateLabel = ConvertTo-MbRecorderTargetName -Value $candidateLabel -Suffix $(if ([string]::IsNullOrWhiteSpace($candidateLabel)) { '' } else { $suffix })
                [void]$recordedCandidates.Add([pscustomobject]@{
                    id = [string]$candidate.id
                    source = $(if ($candidate.PSObject.Properties.Name -contains 'source') { [string]$candidate.source } else { '' })
                    confidence = $(if ($candidate.PSObject.Properties.Name -contains 'confidence') { [string]$candidate.confidence } else { '' })
                    label = $candidateLabel
                    targetType = $(if ($candidate.PSObject.Properties.Name -contains 'targetType') { [string]$candidate.targetType } else { '' })
                    rect = $candidate.rect
                })
            }
        }
        if ($recordedCandidates.Count -eq 0 -and $null -ne $rect -and (Test-MbNormalizedRect -Rect $rect)) {
            $safeSource = if ([string]::IsNullOrWhiteSpace($targetSource)) { 'observed' } else { $targetSource.ToLowerInvariant() -replace '[^a-z0-9]+', '-' }
            $candidateId = ($safeSource.Trim('-') + '-1')
            if ($candidateId -eq '-1') { $candidateId = 'observed-1' }
            [void]$recordedCandidates.Add([pscustomobject]@{
                id = $candidateId; source = $targetSource; confidence = $targetConfidence
                label = $targetName; targetType = $targetType; rect = $rect
            })
        }
        $candidateIds = @($recordedCandidates | ForEach-Object { if ($null -ne $_ -and $_.PSObject.Properties.Name -contains 'id') { [string]$_.id } })
        if ($candidateIds -notcontains $candidateId) {
            $candidateId = if ($candidateIds.Count -gt 0) { [string]$candidateIds[0] } else { '' }
        }
        $candidatesJson = if ($recordedCandidates.Count -gt 0) {
            ConvertTo-Json -InputObject @($recordedCandidates) -Depth 8 -Compress
        } else { '' }
        $clickPointJson = if ($record.PSObject.Properties.Name -contains 'clickPoint' -and $null -ne $record.clickPoint) {
            ConvertTo-Json -InputObject ([pscustomobject]@{
                x = [double]$record.clickPoint.x; y = [double]$record.clickPoint.y
            }) -Compress
        } else { '' }
        $recordEvidenceIds = @()
        if ($record.PSObject.Properties.Name -contains 'evidenceId' -and
            [string]$record.evidenceId -match '^evidence-[a-f0-9]{32}$') {
            $recordEvidenceIds = @([string]$record.evidenceId)
        }
        $recordEvidenceIdsJson = if ($recordEvidenceIds.Count -gt 0) {
            ConvertTo-Json -InputObject ([object[]]$recordEvidenceIds) -Compress
        } else { '' }
        $recordTransformationReason = Get-MbRecorderTransformationReason -Events @($record) `
            -ActionKind ([string]$record.kind) -HasAfterImage (-not [string]::IsNullOrWhiteSpace($resultImagePath))
        [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind $kind `
            -VideoTimeMs ([int]$record.timeMs) -ClickLabel $targetName `
            -WindowTitle ([string]$record.windowTitle) -Narration '' -TargetType $targetType `
            -TargetSource $targetSource -TargetConfidence $targetConfidence `
            -TargetCandidateId $candidateId -TargetCandidatesJson $candidatesJson -ClickPointJson $clickPointJson `
            -SourceSessionId $sessionId -EvidenceIdsJson $recordEvidenceIdsJson `
            -SourceOperationCount 1 -TransformationReason $recordTransformationReason)

        # Copilotを待たず、記録できた事実だけから編集可能な初稿を作る。
        $draft = Get-MbLocalStepDraft -ActionKind ([string]$record.kind) -TargetName $targetName `
            -TargetType $targetType -WindowTitle ([string]$record.windowTitle) `
            -TargetSource $targetSource -TargetConfidence $targetConfidence
        [void](Set-MbStepDraft -Project $Project -StepId $stepId -Title ([string]$draft.title) `
            -Description ([string]$draft.description) -Note ([string]$draft.note))
        $generated++
        if ([bool]$draft.reviewRequired) {
            [void](Set-MbStepReview -Project $Project -StepId $stepId -Action 'review' -Reason ([string]$draft.reviewReason))
            $needsReview++
        } else {
            [void](Set-MbStepReview -Project $Project -StepId $stepId -Action '' -Reason '')
        }
        $added++
    }
    return [pscustomobject]@{
        added = $added; skipped = $skipped; generated = $generated; needsReview = $needsReview
        sourceOperations = $(if ($null -ne $evidenceSession) { [int]$evidenceSession.operationCount } else { $events.Count })
        archivedEvidence = $(if ($null -ne $evidenceSession) { [int]$evidenceSession.operationCount } else { 0 })
        sourceSessionId = $sessionId
    }
}

function Import-MbRecordedEvidenceSession {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath
    )
    if ($null -eq $script:MbRecordingJob) { return $null }
    if ($script:MbRecordingJob.PSObject.Properties.Name -notcontains 'JobId' -or
        $script:MbRecordingJob.PSObject.Properties.Name -notcontains 'LedgerPath' -or
        $script:MbRecordingJob.PSObject.Properties.Name -notcontains 'EvidenceDirectory') {
        throw 'この記録には操作証拠がありません。新しい形式で記録し直してください。'
    }
    $jobId = [string]$script:MbRecordingJob.JobId
    if ($jobId -notmatch '^record-[a-f0-9]{32}$') { throw '証拠セッションIDが正しくありません。' }
    $ledgerSource = [string]$script:MbRecordingJob.LedgerPath
    $evidenceSource = [string]$script:MbRecordingJob.EvidenceDirectory
    if (-not (Test-Path -LiteralPath $ledgerSource -PathType Leaf) -or
        -not (Test-Path -LiteralPath $evidenceSource -PathType Container)) {
        throw '操作証拠の台帳または画像がありません。新しい形式で記録し直してください。'
    }
    $ledgerEntries = New-Object System.Collections.ArrayList
    foreach ($line in @([IO.File]::ReadAllLines($ledgerSource, [Text.Encoding]::UTF8))) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try { [void]$ledgerEntries.Add(($line | ConvertFrom-Json)) }
        catch { throw '操作証拠の台帳が壊れています。記録を取り込まずに停止しました。' }
    }
    $captureStart = @($ledgerEntries | Where-Object { [string]$_.recordType -eq 'capture-start' } | Select-Object -First 1)
    $captureEnd = @($ledgerEntries | Where-Object { [string]$_.recordType -eq 'capture-end' } | Select-Object -Last 1)
    if ($captureStart.Count -ne 1 -or $captureEnd.Count -ne 1 -or
        [int]$captureStart[0].formatVersion -ne 2 -or [int]$captureEnd[0].formatVersion -ne 2) {
        throw 'この記録は現在の証拠形式ではありません。新しい形式で記録し直してください。'
    }
    if ([string]$captureStart[0].sessionId -ne $jobId -or [string]$captureEnd[0].sessionId -ne $jobId -or
        @($ledgerEntries | Where-Object { [string]$_.recordType -eq 'capture-gap' -and [string]$_.sessionId -ne $jobId }).Count -gt 0) {
        throw '操作証拠のセッション情報が一致しません。取り込まずに停止しました。'
    }
    $operations = @($ledgerEntries | Where-Object { [string]$_.recordType -eq 'operation' })
    if ($operations.Count -eq 0) { throw '取り込める操作証拠がありません。' }
    foreach ($operation in $operations) {
        $evidenceId = [string]$operation.id
        $imageName = [string]$operation.image
        if ([string]$operation.sessionId -ne $jobId -or
            $evidenceId -notmatch '^evidence-[a-f0-9]{32}$' -or
            $imageName -ne ($evidenceId + '.jpg') -or
            -not (Test-Path -LiteralPath (Join-Path $evidenceSource $imageName) -PathType Leaf)) {
            throw '操作証拠の画像が不足しています。欠けたまま手順へ変換せずに停止しました。'
        }
    }
    if ($Project.PSObject.Properties.Name -notcontains 'evidenceSessions') {
        $Project | Add-Member -NotePropertyName 'evidenceSessions' -NotePropertyValue @() -Force
    }
    $existing = @($Project.evidenceSessions | Where-Object { [string]$_.id -eq $jobId } | Select-Object -First 1)
    if ($existing.Count -gt 0) { return $existing[0] }

    $projectDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))
    $archiveRoot = Join-Path $projectDirectory 'evidence'
    $sessionRoot = Join-Path $archiveRoot $jobId
    [void](New-Item -ItemType Directory -Path $sessionRoot -Force)
    $ledgerDestination = Join-Path $sessionRoot 'evidence-ledger.jsonl'
    Copy-Item -LiteralPath $ledgerSource -Destination $ledgerDestination -Force
    $imagesDestination = Join-Path $sessionRoot 'images'
    [void](New-Item -ItemType Directory -Path $imagesDestination -Force)
    foreach ($operation in $operations) {
        $imageName = [string]$operation.image
        Copy-Item -LiteralPath (Join-Path $evidenceSource $imageName) `
            -Destination (Join-Path $imagesDestination $imageName) -Force
    }
    $operationCount = 0; $undoneCount = 0
    foreach ($entry in $ledgerEntries) {
        if ([string]$entry.recordType -eq 'operation') { $operationCount++ }
        elseif ([string]$entry.recordType -eq 'decision' -and [string]$entry.action -eq 'undo') { $undoneCount++ }
    }
    $captureGaps = @($ledgerEntries | Where-Object { [string]$_.recordType -eq 'capture-gap' })
    $captureCompleteness = if ($captureGaps.Count -gt 0 -or
        [string]$captureStart[0].completeness -eq 'known-gaps' -or
        [string]$captureEnd[0].completeness -eq 'known-gaps') { 'known-gaps' } else { 'no-known-gaps' }
    $captureWarning = if ($captureCompleteness -eq 'known-gaps') {
        $warning = [string]$captureEnd[0].warning
        if ([string]::IsNullOrWhiteSpace($warning)) { '一部の操作を記録できなかった可能性があります。' } else { $warning }
    } else { '' }
    $session = [pscustomobject]@{
        id = $jobId
        operationCount = $operationCount
        undoneCount = $undoneCount
        ledgerFile = ('evidence/' + $jobId + '/evidence-ledger.jsonl')
        imageDirectory = ('evidence/' + $jobId + '/images')
        decisionsFile = ('evidence/' + $jobId + '/transformations.jsonl')
        formatVersion = 2
        captureCompleteness = $captureCompleteness
        captureWarning = $captureWarning
        retention = 'project-lifetime'
        importedAt = [DateTime]::UtcNow.ToString('o')
    }
    $Project.evidenceSessions = @($Project.evidenceSessions) + @($session)
    return $session
}

function Save-MbRecordedTransformationDecisions {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$SessionId,
        [AllowEmptyCollection()][object[]]$AllProposals = @(),
        [AllowEmptyCollection()][object[]]$AcceptedItems = @(),
        [AllowEmptyCollection()][object[]]$DecisionItems = @()
    )
    if ($SessionId -notmatch '^record-[a-f0-9]{32}$') { return }
    $acceptedIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $acceptedItemsById = @{}
    foreach ($item in @($AcceptedItems)) {
        if ($null -ne $item -and $item.PSObject.Properties.Name -contains 'id') {
            $acceptedId = [string]$item.id
            [void]$acceptedIds.Add($acceptedId)
            $acceptedItemsById[$acceptedId] = $item
        }
    }
    $decisionItemsById = @{}
    foreach ($item in @($DecisionItems)) {
        if ($null -ne $item -and $item.PSObject.Properties.Name -contains 'id') {
            $decisionItemsById[[string]$item.id] = $item
        }
    }
    $projectDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))
    $path = Join-Path (Join-Path (Join-Path $projectDirectory 'evidence') $SessionId) 'transformations.jsonl'
    $revisionId = 'revision-' + [guid]::NewGuid().ToString('N')
    $lines = New-Object System.Collections.ArrayList
    foreach ($proposal in @($AllProposals)) {
        if ($null -eq $proposal -or $proposal.PSObject.Properties.Name -notcontains 'id') { continue }
        $proposalId = [string]$proposal.id
        $evidenceIds = if ($proposal.PSObject.Properties.Name -contains 'evidenceIds') { @($proposal.evidenceIds) } else { @() }
        $sourceOperationCount = if ($proposal.PSObject.Properties.Name -contains 'sourceOperationCount') {
            [int]$proposal.sourceOperationCount
        } else { $evidenceIds.Count }
        $reason = if ($proposal.PSObject.Properties.Name -contains 'transformationReason') {
            [string]$proposal.transformationReason
        } else { '' }
        $acceptedItem = if ($acceptedItemsById.ContainsKey($proposalId)) { $acceptedItemsById[$proposalId] } else { $null }
        $decisionItem = if ($decisionItemsById.ContainsKey($proposalId)) { $decisionItemsById[$proposalId] } else { $acceptedItem }
        $reviewed = $null -ne $decisionItem -and $decisionItem.PSObject.Properties.Name -contains 'reviewed' -and [bool]$decisionItem.reviewed
        $finalTitle = if ($null -ne $decisionItem -and $decisionItem.PSObject.Properties.Name -contains 'title') { [string]$decisionItem.title } else { '' }
        $finalDescription = if ($null -ne $decisionItem -and $decisionItem.PSObject.Properties.Name -contains 'description') { [string]$decisionItem.description } else { '' }
        $record = [ordered]@{
            recordType = 'transformation'
            id = 'transformation-' + [guid]::NewGuid().ToString('N')
            revisionId = $revisionId
            sessionId = $SessionId
            proposalId = $proposalId
            evidenceIds = $evidenceIds
            sourceOperationCount = $sourceOperationCount
            accepted = $acceptedIds.Contains($proposalId)
            reason = $reason
            reviewed = $reviewed
            finalTitle = $finalTitle
            finalDescription = $finalDescription
            decisionSource = $(if ($reviewed) { 'user-review' } else { 'automatic' })
            transformer = 'local-v1'
            decidedAt = [DateTime]::UtcNow.ToString('o')
        }
        [void]$lines.Add(($record | ConvertTo-Json -Depth 8 -Compress))
    }
    if ($lines.Count -gt 0) {
        [IO.File]::AppendAllText($path, ((@($lines) -join [Environment]::NewLine) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
    }
}

function Save-MbRecordedRawTransformationDecisions {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$SessionId,
        [AllowEmptyCollection()][object[]]$Events = @(),
        [Parameter(Mandatory = $true)]$AcceptedIndexes
    )
    if ($SessionId -notmatch '^record-[a-f0-9]{32}$') { return }
    $proposals = New-Object System.Collections.ArrayList
    $accepted = New-Object System.Collections.ArrayList
    foreach ($event in @($Events)) {
        $index = 0
        try { $index = [int]$event.index } catch { $index = 0 }
        if ($index -le 0) { continue }
        $evidenceIds = @()
        if ($event.PSObject.Properties.Name -contains 'evidenceId' -and
            [string]$event.evidenceId -match '^evidence-[a-f0-9]{32}$') {
            $evidenceIds = @([string]$event.evidenceId)
        }
        $proposal = [pscustomobject]@{
            id = ('raw-event-{0:d5}' -f $index)
            evidenceIds = $evidenceIds
            sourceOperationCount = 1
            transformationReason = Get-MbRecorderTransformationReason -Events @($event) `
                -ActionKind ([string]$event.kind) -HasAfterImage ($event.PSObject.Properties.Name -contains 'resultImage' -and -not [string]::IsNullOrWhiteSpace([string]$event.resultImage))
        }
        [void]$proposals.Add($proposal)
        if ($AcceptedIndexes.Contains($index)) { [void]$accepted.Add([pscustomobject]@{ id = [string]$proposal.id }) }
    }
    Save-MbRecordedTransformationDecisions -ProjectPath $ProjectPath -SessionId $SessionId `
        -AllProposals @($proposals) -AcceptedItems @($accepted)
}

# このPCで選んだ原画像と操作群を、編集可能な手順へ変換する。
function Import-MbRecordedLocalSelections {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$SelectionJson
    )
    if ([string]::IsNullOrWhiteSpace($SelectionJson)) { throw '取り込む手順候補がありません。' }
    try { $selection = $SelectionJson | ConvertFrom-Json } catch { throw '手順候補の形式が正しくありません。' }
    $items = @($selection)
    if ($selection.PSObject.Properties.Name -contains 'accept') { $items = @($selection.accept) }
    elseif ($selection.PSObject.Properties.Name -contains 'steps') { $items = @($selection.steps) }
    $decisionItems = if ($selection.PSObject.Properties.Name -contains 'decisions') { @($selection.decisions) } else { @() }
    if ($items.Count -gt 300) { throw '一度に取り込める手順は300件までです。' }

    $evidenceSession = Import-MbRecordedEvidenceSession -Project $Project -ProjectPath $ProjectPath
    $sessionId = if ($null -ne $evidenceSession) { [string]$evidenceSession.id } else { '' }
    $allProposals = @(Get-MbRecordedLocalProposals)
    Save-MbRecordedTransformationDecisions -ProjectPath $ProjectPath -SessionId $sessionId `
        -AllProposals $allProposals -AcceptedItems $items -DecisionItems $decisionItems

    $frameMap = @{}
    foreach ($frame in @(Get-MbRecordedFrames)) { $frameMap[[string]$frame.id] = $frame }
    $eventMap = @{}
    foreach ($event in @(Get-MbRecordedRawEvents)) { $eventMap[[int]$event.index] = $event }
    $added = 0; $skipped = 0; $needsReview = 0
    foreach ($item in $items) {
        if ($null -eq $item) { $skipped++; continue }
        $beforeId = ([string]$item.beforeFrame).Trim().ToUpperInvariant()
        if (-not $frameMap.ContainsKey($beforeId)) { $skipped++; continue }
        $beforeFrame = $frameMap[$beforeId]
        $beforePath = Get-MbRecordedFrameImagePath -FileName ([string]$beforeFrame.image)
        if ([string]::IsNullOrWhiteSpace($beforePath)) { $skipped++; continue }

        $afterId = if ($item.PSObject.Properties.Name -contains 'afterFrame') { ([string]$item.afterFrame).Trim().ToUpperInvariant() } else { '' }
        $afterFrame = $null
        if (-not [string]::IsNullOrWhiteSpace($afterId) -and $afterId -ne $beforeId -and $frameMap.ContainsKey($afterId)) {
            $afterFrame = $frameMap[$afterId]
            $beforeApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$beforeFrame.windowTitle)
            $afterApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$afterFrame.windowTitle)
            if ([string]::IsNullOrWhiteSpace($beforeApp) -or $beforeApp -ne $afterApp -or
                [int]$afterFrame.timeMs -le [int]$beforeFrame.timeMs) { $skipped++; continue }
        } else { $afterId = '' }

        $eventIds = New-Object System.Collections.ArrayList
        if ($item.PSObject.Properties.Name -contains 'eventIds') {
            foreach ($value in @($item.eventIds)) {
                $eventId = 0
                try { $eventId = [int]$value } catch { $eventId = 0 }
                if ($eventMap.ContainsKey($eventId) -and -not $eventIds.Contains($eventId)) { [void]$eventIds.Add($eventId) }
            }
        }
        $requestedTargetEventId = 0
        if ($item.PSObject.Properties.Name -contains 'targetEventId') { try { $requestedTargetEventId = [int]$item.targetEventId } catch { } }
        $validAnchorIds = @($eventIds | Where-Object {
            $eventMap.ContainsKey([int]$_) -and
            (Test-MbRecordedSelectionAnchorContext -Event $eventMap[[int]$_] -BeforeFrame $beforeFrame -AfterFrame $afterFrame)
        })
        $targetEventId = if ($requestedTargetEventId -gt 0 -and -not $eventIds.Contains($requestedTargetEventId)) {
            0
        } elseif ($eventIds.Contains($requestedTargetEventId) -and $requestedTargetEventId -in $validAnchorIds) {
            $requestedTargetEventId
        } elseif ($validAnchorIds.Count -eq 1) { [int]$validAnchorIds[0] } else { 0 }
        $anchor = $(if ($targetEventId -gt 0) { $eventMap[$targetEventId] } else { $null })

        $result = Add-MbImageStep -Project $Project -ProjectPath $ProjectPath -SheetId $SheetId `
            -Bytes ([IO.File]::ReadAllBytes($beforePath)) -Source 'recorder' -AllowDuplicateStep
        if ([string]$result.Status -ne 'added') { $skipped++; continue }
        $stepId = [string]$result.Step.id

        if ($null -ne $afterFrame) {
            $afterPath = Get-MbRecordedFrameImagePath -FileName ([string]$afterFrame.image)
            if (-not [string]::IsNullOrWhiteSpace($afterPath)) {
                try {
                    $asset = Add-MbImageAsset -Project $Project -ProjectPath $ProjectPath -Bytes ([IO.File]::ReadAllBytes($afterPath)) -Source 'recorder'
                    $result.Step.resultImageId = [string]$asset.Image.id
                    $result.Step.imageLayout = 'before'; $result.Step.imageOrder = 'before-after'
                } catch { }
            }
        }

        $annotations = @()
        if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'rect' -and
            $null -ne $anchor.rect -and (Test-MbNormalizedRect -Rect $anchor.rect)) {
            $annotations = @([pscustomobject]@{
                id = New-MbRecorderAnnotationId; type = 'rect'
                x1 = [Math]::Round([double]$anchor.rect.x1, 6); y1 = [Math]::Round([double]$anchor.rect.y1, 6)
                x2 = [Math]::Round([double]$anchor.rect.x2, 6); y2 = [Math]::Round([double]$anchor.rect.y2, 6); label = 0
            })
            [void](Set-MbStepAnnotations -Project $Project -StepId $stepId -AnnotationsJson (ConvertTo-Json -InputObject $annotations -Depth 5))
        }
        $title = ([string]$item.title).Trim()
        $description = ([string]$item.description).Trim()
        if ($title.Length -gt 100) { $title = $title.Substring(0, 100) }
        if ($description.Length -gt 500) { $description = $description.Substring(0, 500) }
        if ([string]::IsNullOrWhiteSpace($title)) { $title = '手順を確認' }
        if ([string]::IsNullOrWhiteSpace($description)) { $description = '画面を確認して操作します。' }
        [void](Set-MbStepDraft -Project $Project -StepId $stepId -Title $title -Description $description -Note '')

        $targetName = $(if ($null -ne $anchor) { [string]$anchor.targetName } else { '' })
        $targetType = $(if ($null -ne $anchor) { [string]$anchor.targetType } else { '' })
        $targetSource = $(if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'targetSource') { [string]$anchor.targetSource } else { '' })
        $targetConfidence = $(if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'confidence') { [string]$anchor.confidence } else { '' })
        $candidateId = $(if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'targetCandidateId') { [string]$anchor.targetCandidateId } else { '' })
        $candidatesJson = $(if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'targetCandidates') { ConvertTo-Json -InputObject @($anchor.targetCandidates) -Depth 8 -Compress } else { '' })
        $itemActionKind = if ($item.PSObject.Properties.Name -contains 'actionKind') { [string]$item.actionKind } else { '' }
        if (Test-MbRecorderUnreliableOcrLabel -Label $targetName -ActionKind $itemActionKind -Source $targetSource) {
            # 不採用にしたOCR入力値を、手順の表示外メタデータへも残さない。
            # 矩形とクリック点は赤枠の確認用に維持する。
            $targetName = ''
            $candidateId = ''
            $candidatesJson = ''
        }
        $clickPointJson = $(if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'clickPoint') {
            ConvertTo-Json -InputObject ([pscustomobject]@{ x = [double]$anchor.clickPoint.x; y = [double]$anchor.clickPoint.y }) -Compress
        } else { '' })
        [object[]]$evidenceIds = @(if ($item.PSObject.Properties.Name -contains 'evidenceIds') { @($item.evidenceIds) } else { @() })
        $evidenceIdsJson = if ($evidenceIds.Count -gt 0) { ConvertTo-Json -InputObject ([object[]]@($evidenceIds)) -Compress } else { '' }
        $sourceOperationCount = if ($item.PSObject.Properties.Name -contains 'sourceOperationCount') { [int]$item.sourceOperationCount } else { $eventIds.Count }
        $transformationReason = if ($item.PSObject.Properties.Name -contains 'transformationReason') { [string]$item.transformationReason } else {
            Get-MbRecorderTransformationReason -Events @($eventIds | ForEach-Object { $eventMap[[int]$_] }) `
                -ActionKind $itemActionKind -HasAfterImage ($null -ne $afterFrame)
        }
        [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind 'recorded-local' -VideoTimeMs ([int]$beforeFrame.timeMs) `
            -ClickLabel $targetName -WindowTitle ([string]$beforeFrame.windowTitle) -TargetType $targetType `
            -TargetSource $targetSource -TargetConfidence $targetConfidence -TargetCandidateId $candidateId `
            -TargetCandidatesJson $candidatesJson -ClickPointJson $clickPointJson -SourceSessionId $sessionId `
            -EvidenceIdsJson $evidenceIdsJson -SourceOperationCount $sourceOperationCount `
            -TransformationReason $transformationReason)
        $confidence = if ($item.PSObject.Properties.Name -contains 'confidence') { [string]$item.confidence } else { 'low' }
        $userReviewed = $item.PSObject.Properties.Name -contains 'reviewed' -and [bool]$item.reviewed
        if (($confidence -ne 'high' -or @($annotations).Count -eq 0) -and -not $userReviewed) {
            $reason = if ($item.PSObject.Properties.Name -contains 'reason') { [string]$item.reason } else { '' }
            if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'このPCで選んだ画像と手順です。文章と赤枠を確認してください。' }
            [void](Set-MbStepReview -Project $Project -StepId $stepId -Action 'review' -Reason $reason)
            $needsReview++
        } else { [void](Set-MbStepReview -Project $Project -StepId $stepId -Action '' -Reason '') }
        $added++
    }
    return [pscustomobject]@{
        added = $added; skipped = $skipped; generated = $added; needsReview = $needsReview
        sourceOperations = $(if ($null -ne $evidenceSession) { [int]$evidenceSession.operationCount } else { $eventMap.Count })
        archivedEvidence = $(if ($null -ne $evidenceSession) { [int]$evidenceSession.operationCount } else { 0 })
        sourceSessionId = $sessionId
    }
}

function Get-MbRecordedFrameImagePath {
    param([Parameter(Mandatory = $true)][string]$FileName)
    if ($null -eq $script:MbRecordingJob) { return '' }
    if ($FileName -notmatch '^frame-\d{5}\.jpg$') { return '' }
    $path = Join-Path ([string]$script:MbRecordingJob.FramesDirectory) $FileName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    return $path
}

function Get-MbRecordingSourceInfo {
    if ($null -eq $script:MbRecordingJob) { return $null }
    return [pscustomobject]@{
        jobId = [string]$script:MbRecordingJob.JobId
        jobDirectory = [string]$script:MbRecordingJob.JobDirectory
        eventsPath = [string]$script:MbRecordingJob.EventsPath
        eventsDirectory = [string]$script:MbRecordingJob.EventsDirectory
        evidenceDirectory = [string]$script:MbRecordingJob.EvidenceDirectory
        ledgerPath = [string]$script:MbRecordingJob.LedgerPath
        framesPath = [string]$script:MbRecordingJob.FramesPath
        framesDirectory = [string]$script:MbRecordingJob.FramesDirectory
    }
}

function Remove-MbRecordingJob {
    $script:MbRecorderOcrSnapshotCache = @{}
    if ($null -eq $script:MbRecordingJob) { return }
    $job = $script:MbRecordingJob
    $directory = [string]$job.JobDirectory
    $processIdentities = New-Object System.Collections.ArrayList
    foreach ($property in @('ProcessIdentity', 'UiaWorkerProcessIdentity', 'ControllerProcessIdentity')) {
        if ($job.PSObject.Properties.Name -contains $property -and $null -ne $job.$property) {
            [void]$processIdentities.Add($job.$property)
        }
    }
    $stopPath = if ($job.PSObject.Properties.Name -contains 'StopPath') { [string]$job.StopPath } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($stopPath)) {
        try { [IO.File]::WriteAllText($stopPath, 'stop', (New-Object Text.UTF8Encoding($false))) } catch { }
    }
    $script:MbRecordingJob = $null
    # まだ記録プロセスが動いていたら止める。放置すると画面を撮り続ける。
    foreach ($identity in @($processIdentities)) { [void](Stop-MbRecorderOwnedProcess -Identity $identity) }
    if ([string]::IsNullOrWhiteSpace($directory)) { return }
    try { Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}

# Recorder モジュールはこのモジュールの内側にしか読み込まれないため、
# 本体からは この関数を通して記録できるかどうかを受け取る。
function Get-MbRecordingCapability {
    $recorder = Get-MbRecorderCapability
    return [pscustomobject]@{
        available = $recorder.available
        reason    = $recorder.reason
    }
}

# Test-MbNormalizedRect はこのモジュールの内部でだけ使う。
# CopilotServer も同名の関数を公開しており、両方を公開すると
# 本体へ読み込んだときに後勝ちで上書きされる。
Export-ModuleMember -Function @(
    'Initialize-MbRecorderServer',
    'Remove-MbOrphanedRecordingJobs',
    'Start-MbRecordingJob',
    'Read-MbRecordingStatus',
    'Stop-MbRecordingJob',
    'Set-MbRecordingPaused',
    'Undo-MbLastRecordingEvent',
    'Get-MbRecordedEvents',
    'Get-MbRecordedRawEvents',
    'Get-MbRecordedFrames',
    'Get-MbRecordedLocalProposals',
    'Merge-MbRecordedEditInteractions',
    'Get-MbRecordedEventImagePath',
    'Get-MbRecordedFrameImagePath',
    'Get-MbRecordingSourceInfo',
    'Import-MbRecordedEvents',
    'Import-MbRecordedLocalSelections',
    'Remove-MbRecordingJob',
    'Get-MbRecordingCapability'
)
