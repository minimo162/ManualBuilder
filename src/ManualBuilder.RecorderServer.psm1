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
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Dictation.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.LocalDraft.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.RecorderCopilot.psm1')

$script:MbRecordingJobsRoot = ''
$script:MbRecordingScriptRoot = ''
$script:MbRecordingJob = $null

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

function Get-MbRecordingIdleStatus {
    return [pscustomobject]@{
        jobId = ''; state = 'idle'; count = 0; message = ''; lastTarget = ''; updatedAt = ''
    }
}

function Read-MbRecordingStatus {
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $statusPath = [string]$script:MbRecordingJob.StatusPath
    if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { return (Get-MbRecordingIdleStatus) }
    $status = $null
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
        # 差し替えと同時になった場合は、次の巡回で読み直す。
        return (Get-MbRecordingIdleStatus)
    }
    if ($null -eq $status) { return (Get-MbRecordingIdleStatus) }

    # 記録プロセスが落ちたまま recording / paused が残らないようにする。
    if ([string]$status.state -in @('recording', 'paused')) {
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
        [string[]]$IgnoreTitlePatterns = @('ManualBuilder'),
        [switch]$WithNarration
    )

    $current = Read-MbRecordingStatus
    if ([string]$current.state -in @('recording', 'paused')) { return $current }

    $capability = Get-MbRecorderCapability
    if (-not $capability.available) { throw ([string]$capability.reason) }

    # 前回の記録が残っていれば片付けてから始める。
    Remove-MbRecordingJob

    $jobId = 'record-' + [guid]::NewGuid().ToString('N')
    $jobDirectory = Join-Path $script:MbRecordingJobsRoot $jobId
    $eventsDirectory = Join-Path $jobDirectory 'events'
    [void](New-Item -ItemType Directory -Path $eventsDirectory -Force)
    $framesDirectory = Join-Path $jobDirectory 'frames'
    [void](New-Item -ItemType Directory -Path $framesDirectory -Force)
    $statusPath = Join-Path $jobDirectory 'status.json'
    $eventsPath = Join-Path $jobDirectory 'events.jsonl'
    $framesPath = Join-Path $jobDirectory 'frames.jsonl'
    $stopPath = Join-Path $jobDirectory 'stop.requested'
    $pausePath = Join-Path $jobDirectory 'pause.requested'
    $undoPath = Join-Path $jobDirectory 'undo.requested'
    $narrationPath = Join-Path $jobDirectory 'narration.jsonl'
    $narrationStatusPath = Join-Path $jobDirectory 'narration-status.json'
    $uiaTargetPath = Join-Path $jobDirectory 'uia-target.json'
    $uiaLogPath = Join-Path $jobDirectory 'uia-monitor.log'

    $queued = [pscustomobject]@{
        jobId = $jobId; state = 'recording'; count = 0
        message = '記録の準備をしています'
        lastTarget = ''; updatedAt = [DateTime]::UtcNow.ToString('o')
    }
    [IO.File]::WriteAllText($statusPath, ($queued | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))

    # 記録プロセスと音声プロセスが同じ時計を使うよう、開始時刻を揃えて渡す。
    $startedAtUtc = [DateTime]::UtcNow
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
        '-FramesDirectory', (& $quote $framesDirectory),
        '-FramesPath', (& $quote $framesPath),
        '-StatusPath', (& $quote $statusPath),
        '-StopPath', (& $quote $stopPath),
        '-PausePath', (& $quote $pausePath),
        '-UndoPath', (& $quote $undoPath),
        '-JobId', (& $quote $jobId),
        '-UiaTargetPath', (& $quote $uiaTargetPath)
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

    # 音声の聞き取りは記録ループと同居できない。並走する別プロセスにする。
    # 起動に失敗しても操作の記録は続けられるので、ここでは止めない。
    $dictationProcessId = 0; $dictationProcessIdentity = $null
    if ($WithNarration) {
        try {
            $dictationWorkerPath = Join-Path $script:MbRecordingScriptRoot 'Invoke-ManualBuilderDictation.ps1'
            $dictationArguments = @(
                '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $dictationWorkerPath),
                '-OutputPath', (& $quote $narrationPath),
                '-StopPath', (& $quote $stopPath),
                '-PausePath', (& $quote $pausePath),
                '-StatusPath', (& $quote $narrationStatusPath),
                '-StartedAtUtcTicks', ([string]$startedAtUtc.Ticks)
            )
            $dictationWorker = Start-Process -FilePath $powerShellPath -ArgumentList $dictationArguments -WindowStyle Hidden -PassThru
            $dictationProcessId = [int]$dictationWorker.Id
            $dictationProcessIdentity = New-MbRecorderProcessIdentity -Process $dictationWorker
            $dictationWorker.Dispose()
        } catch {
            $dictationProcessId = 0; $dictationProcessIdentity = $null
        }
    }

    $script:MbRecordingJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $processId; JobDirectory = $jobDirectory
        ProcessIdentity = $processIdentity
        EventsDirectory = $eventsDirectory; EventsPath = $eventsPath
        FramesDirectory = $framesDirectory; FramesPath = $framesPath
        StatusPath = $statusPath; StopPath = $stopPath; PausePath = $pausePath; UndoPath = $undoPath; StartedAt = Get-Date
        NarrationPath = $narrationPath; NarrationStatusPath = $narrationStatusPath
        DictationProcessId = $dictationProcessId
        DictationProcessIdentity = $dictationProcessIdentity
        UiaTargetPath = $uiaTargetPath; UiaLogPath = $uiaLogPath
        UiaWorkerProcessId = $uiaWorkerProcessId
        UiaWorkerProcessIdentity = $uiaWorkerProcessIdentity
    }
    return (Read-MbRecordingStatus)
}

function Stop-MbRecordingJob {
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $status = Read-MbRecordingStatus
    if ([string]$status.state -in @('recording', 'paused')) {
        [IO.File]::WriteAllText([string]$script:MbRecordingJob.StopPath, 'stop', (New-Object Text.UTF8Encoding($false)))
        # 記録プロセスが停止を見て後始末を終えるまで少しだけ待つ。
        for ($i = 0; $i -lt 40; $i++) {
            Start-Sleep -Milliseconds 100
            $status = Read-MbRecordingStatus
            if ([string]$status.state -notin @('recording', 'paused')) { break }
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
    [IO.File]::WriteAllText([string]$script:MbRecordingJob.UndoPath, 'undo', (New-Object Text.UTF8Encoding($false)))
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
                $isEditClick = [string]$candidate.kind -eq 'click' -and
                    [string]$candidate.targetType -in @('ControlType.Edit', 'ControlType.DataItem')
                $sameField = -not [string]::IsNullOrWhiteSpace([string]$candidate.targetName) -and
                    [string]$candidate.targetName -eq [string]$current.targetName -and
                    [string]$candidate.windowTitle -eq [string]$current.windowTitle -and
                    ((Test-SamePosition -First $candidate -Second $current -Tolerance 0.02) -or
                        (-not (Test-HasPosition -Event $candidate) -and -not (Test-HasPosition -Event $current)))
                $gapMs = [int]$current.timeMs - [int]$candidate.timeMs
                if (-not ($isEditClick -and $sameField -and $gapMs -ge 0 -and $gapMs -le $MaxGapMs)) { break }
                $result.RemoveAt($result.Count - 1)
            }
        }
        [void]$result.Add($current)
    }
    return @($result)
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
    return @($events)
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
        [void]$result.Add([pscustomobject]@{
            id = [string]$candidate.id
            beforeFrame = [string]$candidate.beforeFrame
            afterFrame = [string]$candidate.afterFrame
            eventIds = @($candidate.eventIds)
            targetEventId = [int]$candidate.targetEventId
            title = [string]$draft.title
            description = [string]$draft.description
            confidence = $(if ([bool]$draft.reviewRequired) { 'low' } else { 'medium' })
            reason = $(if ([bool]$draft.reviewRequired) { [string]$draft.reviewReason } else { 'このPCで記録した操作前後から作成しました。' })
            timeMs = [int]$candidate.timeMs
            beforeImage = [string]$candidate.beforeImage
            afterImage = [string]$candidate.afterImage
            source = 'local'
        })
    }
    return @($result)
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

function Get-MbRecordedNarration {
    if ($null -eq $script:MbRecordingJob) { return @() }
    $path = [string]$script:MbRecordingJob.NarrationPath
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { return @() }

    $phrases = New-Object System.Collections.ArrayList
    foreach ($line in [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $record = $null
        try { $record = $line | ConvertFrom-Json } catch { continue }
        if ($null -eq $record) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$record.text)) { continue }
        [void]$phrases.Add($record)
    }
    return @($phrases)
}

# 話した内容を、どの操作の説明かで振り分ける。
#
# 人の喋り方は2通りある。
#   「ここで申請ボタンを押します」→ 操作する    … 発話のあとに操作が来る
#   操作する →「これで一覧に出ました」          … 操作のあとに発話が来る
#
# 直前の操作からすぐ喋り始めた場合はその操作への補足とみなし、
# そうでなければ次に来る操作の説明とみなす。
# 1つの発話は1つの操作にしか付けない。同じ文が複数の手順に出ると読みにくいため。
function Merge-MbNarrationIntoEvents {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Events,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Phrases,
        [int]$TrailMs = 2000,
        [int]$LeadMs = 3000
    )

    $assigned = @{}
    if (@($Events).Count -eq 0 -or @($Phrases).Count -eq 0) { return $assigned }

    $ordered = @($Events | Sort-Object @{ Expression = { [int]$_.timeMs } })
    foreach ($phrase in (@($Phrases) | Sort-Object @{ Expression = { [int]$_.startMs } })) {
        $startMs = [int]$phrase.startMs
        $endMs = [int]$phrase.endMs
        $target = $null

        # 直後の補足。操作してからすぐに喋り始めたもの。
        foreach ($item in $ordered) {
            $gap = $startMs - [int]$item.timeMs
            if ($gap -ge 0 -and $gap -le $TrailMs) { $target = $item }
        }
        # そうでなければ、これから行う操作の説明とみなす。
        if ($null -eq $target) {
            foreach ($item in $ordered) {
                $time = [int]$item.timeMs
                if ($time -ge $startMs -and $time -le ($endMs + $LeadMs)) { $target = $item; break }
            }
        }
        if ($null -eq $target) { continue }

        $index = [int]$target.index
        $text = ([string]$phrase.text).Trim()
        if ($assigned.ContainsKey($index)) {
            $assigned[$index] = [string]$assigned[$index] + ' ' + $text
        } else {
            $assigned[$index] = $text
        }
    }
    return $assigned
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

    # 話した内容を、どの操作の説明かで振り分けておく。
    $narration = Merge-MbNarrationIntoEvents -Events $events -Phrases (@(Get-MbRecordedNarration))

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
                    $result.Step.imageLayout = 'side-by-side'
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
        $spoken = ''
        if ($narration.ContainsKey($index)) { $spoken = [string]$narration[$index] }
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
        [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind $kind `
            -VideoTimeMs ([int]$record.timeMs) -ClickLabel $targetName `
            -WindowTitle ([string]$record.windowTitle) -Narration $spoken -TargetType $targetType `
            -TargetSource $targetSource -TargetConfidence $targetConfidence `
            -TargetCandidateId $candidateId -TargetCandidatesJson $candidatesJson -ClickPointJson $clickPointJson)

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
    return [pscustomobject]@{ added = $added; skipped = $skipped; generated = $generated; needsReview = $needsReview }
}

# Copilotがコンタクトシートから選んだ原画像と操作群を、編集可能な手順へ変換する。
function Import-MbRecordedCopilotSelections {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$SelectionJson
    )
    if ([string]::IsNullOrWhiteSpace($SelectionJson)) { throw '取り込むAI手順候補がありません。' }
    try { $selection = $SelectionJson | ConvertFrom-Json } catch { throw 'AI手順候補の形式が正しくありません。' }
    $items = @($selection)
    if ($selection.PSObject.Properties.Name -contains 'accept') { $items = @($selection.accept) }
    elseif ($selection.PSObject.Properties.Name -contains 'steps') { $items = @($selection.steps) }
    if ($items.Count -gt 300) { throw '一度に取り込める手順は300件までです。' }

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
                    $result.Step.imageLayout = 'side-by-side'; $result.Step.imageOrder = 'before-after'
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
        $clickPointJson = $(if ($null -ne $anchor -and $anchor.PSObject.Properties.Name -contains 'clickPoint') {
            ConvertTo-Json -InputObject ([pscustomobject]@{ x = [double]$anchor.clickPoint.x; y = [double]$anchor.clickPoint.y }) -Compress
        } else { '' })
        [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind 'recorded-ai' -VideoTimeMs ([int]$beforeFrame.timeMs) `
            -ClickLabel $targetName -WindowTitle ([string]$beforeFrame.windowTitle) -TargetType $targetType `
            -TargetSource $targetSource -TargetConfidence $targetConfidence -TargetCandidateId $candidateId `
            -TargetCandidatesJson $candidatesJson -ClickPointJson $clickPointJson)
        $confidence = if ($item.PSObject.Properties.Name -contains 'confidence') { [string]$item.confidence } else { 'low' }
        if ($confidence -ne 'high' -or @($annotations).Count -eq 0) {
            $reason = if ($item.PSObject.Properties.Name -contains 'reason') { [string]$item.reason } else { '' }
            if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'AIが選んだ画像と手順です。文章と赤枠を確認してください。' }
            [void](Set-MbStepReview -Project $Project -StepId $stepId -Action 'review' -Reason $reason)
            $needsReview++
        } else { [void](Set-MbStepReview -Project $Project -StepId $stepId -Action '' -Reason '') }
        $added++
    }
    return [pscustomobject]@{ added = $added; skipped = $skipped; generated = $added; needsReview = $needsReview }
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
        framesPath = [string]$script:MbRecordingJob.FramesPath
        framesDirectory = [string]$script:MbRecordingJob.FramesDirectory
    }
}

function Remove-MbRecordingJob {
    if ($null -eq $script:MbRecordingJob) { return }
    $job = $script:MbRecordingJob
    $directory = [string]$job.JobDirectory
    $processIdentities = New-Object System.Collections.ArrayList
    foreach ($property in @('ProcessIdentity', 'DictationProcessIdentity', 'UiaWorkerProcessIdentity')) {
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
    # 音声は任意。使えなくても操作の記録はできるので、別々に返す。
    $dictation = $null
    try { $dictation = Get-MbDictationCapability } catch {
        $dictation = [pscustomobject]@{ available = $false; reason = 'この環境では音声入力を利用できません。'; language = '' }
    }
    return [pscustomobject]@{
        available = $recorder.available
        reason    = $recorder.reason
        narration = $dictation
    }
}

# Test-MbNormalizedRect はこのモジュールの内部でだけ使う。
# CopilotServer も同名の関数を公開しており、両方を公開すると
# 本体へ読み込んだときに後勝ちで上書きされる。
Export-ModuleMember -Function @(
    'Initialize-MbRecorderServer',
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
    'Import-MbRecordedCopilotSelections',
    'Get-MbRecordedNarration',
    'Merge-MbNarrationIntoEvents',
    'Remove-MbRecordingJob',
    'Get-MbRecordingCapability'
)
