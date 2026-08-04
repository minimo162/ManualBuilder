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

$script:MbRecordingJobsRoot = ''
$script:MbRecordingScriptRoot = ''
$script:MbRecordingJob = $null

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

    # 記録プロセスが落ちたまま recording が残らないようにする。
    if ([string]$status.state -eq 'recording') {
        $alive = $false
        try { $alive = $null -ne (Get-Process -Id ([int]$script:MbRecordingJob.ProcessId) -ErrorAction SilentlyContinue) } catch { $alive = $false }
        if (-not $alive) {
            $status.state = 'failed'
            $status.message = '記録が途中で終わりました。もう一度実行してください。'
        }
        if ($script:MbRecordingJob.PSObject.Properties.Name -contains 'UiaWorkerProcessId') {
            $uiaWorkerAlive = $false
            $uiaWorkerProcessId = [int]$script:MbRecordingJob.UiaWorkerProcessId
            if ($uiaWorkerProcessId -gt 0) {
                try {
                    $uiaWorkerAlive = $null -ne (Get-Process -Id $uiaWorkerProcessId -ErrorAction SilentlyContinue)
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
    if ([string]$current.state -eq 'recording') { return $current }

    $capability = Get-MbRecorderCapability
    if (-not $capability.available) { throw ([string]$capability.reason) }

    # 前回の記録が残っていれば片付けてから始める。
    Remove-MbRecordingJob

    $jobId = 'record-' + [guid]::NewGuid().ToString('N')
    $jobDirectory = Join-Path $script:MbRecordingJobsRoot $jobId
    $eventsDirectory = Join-Path $jobDirectory 'events'
    [void](New-Item -ItemType Directory -Path $eventsDirectory -Force)
    $statusPath = Join-Path $jobDirectory 'status.json'
    $eventsPath = Join-Path $jobDirectory 'events.jsonl'
    $stopPath = Join-Path $jobDirectory 'stop.requested'
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
    $uiaWorkerProcessId = 0
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
        $uiaWorker.Dispose()
    } catch {
        $uiaWorkerProcessId = 0
    }

    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $workerPath),
        '-EventsDirectory', (& $quote $eventsDirectory),
        '-EventsPath', (& $quote $eventsPath),
        '-StatusPath', (& $quote $statusPath),
        '-StopPath', (& $quote $stopPath),
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
        $worker.Dispose()
    } catch {
        if ($uiaWorkerProcessId -gt 0) {
            try { (Get-Process -Id $uiaWorkerProcessId -ErrorAction SilentlyContinue).Kill() } catch { }
        }
        try { Remove-Item -LiteralPath $jobDirectory -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        throw
    }

    # 音声の聞き取りは記録ループと同居できない。並走する別プロセスにする。
    # 起動に失敗しても操作の記録は続けられるので、ここでは止めない。
    $dictationProcessId = 0
    if ($WithNarration) {
        try {
            $dictationWorkerPath = Join-Path $script:MbRecordingScriptRoot 'Invoke-ManualBuilderDictation.ps1'
            $dictationArguments = @(
                '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $dictationWorkerPath),
                '-OutputPath', (& $quote $narrationPath),
                '-StopPath', (& $quote $stopPath),
                '-StatusPath', (& $quote $narrationStatusPath),
                '-StartedAtUtcTicks', ([string]$startedAtUtc.Ticks)
            )
            $dictationWorker = Start-Process -FilePath $powerShellPath -ArgumentList $dictationArguments -WindowStyle Hidden -PassThru
            $dictationProcessId = [int]$dictationWorker.Id
            $dictationWorker.Dispose()
        } catch {
            $dictationProcessId = 0
        }
    }

    $script:MbRecordingJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $processId; JobDirectory = $jobDirectory
        EventsDirectory = $eventsDirectory; EventsPath = $eventsPath
        StatusPath = $statusPath; StopPath = $stopPath; StartedAt = Get-Date
        NarrationPath = $narrationPath; NarrationStatusPath = $narrationStatusPath
        DictationProcessId = $dictationProcessId
        UiaTargetPath = $uiaTargetPath; UiaLogPath = $uiaLogPath
        UiaWorkerProcessId = $uiaWorkerProcessId
    }
    return (Read-MbRecordingStatus)
}

function Stop-MbRecordingJob {
    if ($null -eq $script:MbRecordingJob) { return (Get-MbRecordingIdleStatus) }
    $status = Read-MbRecordingStatus
    if ([string]$status.state -eq 'recording') {
        [IO.File]::WriteAllText([string]$script:MbRecordingJob.StopPath, 'stop', (New-Object Text.UTF8Encoding($false)))
        # 記録プロセスが停止を見て後始末を終えるまで少しだけ待つ。
        for ($i = 0; $i -lt 40; $i++) {
            Start-Sleep -Milliseconds 100
            $status = Read-MbRecordingStatus
            if ([string]$status.state -ne 'recording') { break }
        }
    }
    return $status
}

# 記録した操作を読み出す。取り込む前に一覧を見せて選んでもらうために使う。
function Get-MbRecordedEvents {
    if ($null -eq $script:MbRecordingJob) { return @() }
    $eventsPath = [string]$script:MbRecordingJob.EventsPath
    if (-not (Test-Path -LiteralPath $eventsPath -PathType Leaf)) { return @() }

    $events = New-Object System.Collections.ArrayList
    foreach ($line in [IO.File]::ReadAllLines($eventsPath, [Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $record = $null
        try { $record = $line | ConvertFrom-Json } catch { continue }
        if ($null -eq $record) { continue }
        [void]$events.Add($record)
    }
    return @($events)
}

function Get-MbRecordedEventImagePath {
    param([Parameter(Mandatory = $true)][string]$FileName)
    if ($null -eq $script:MbRecordingJob) { return '' }
    if ($FileName -notmatch '^event-\d{3}\.jpg$') { return '' }
    $path = Join-Path ([string]$script:MbRecordingJob.EventsDirectory) $FileName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    return $path
}

function New-MbRecorderAnnotationId {
    return 'annotation-' + [guid]::NewGuid().ToString('N')
}

# 元画像はウィンドウ全体のまま残し、通常表示だけを操作対象の周辺へ寄せる。
# クリック位置だけのフォールバックは空クリックの可能性があるため、自動切り抜きしない。
function Get-MbRecorderTargetCrop {
    param(
        [AllowNull()]$Rect,
        [string]$TargetType = ''
    )

    if ($TargetType -eq 'ControlType.ClickPoint' -or -not (Test-MbNormalizedRect -Rect $Rect)) { return $null }
    $targetWidth = [double]$Rect.x2 - [double]$Rect.x1
    $targetHeight = [double]$Rect.y2 - [double]$Rect.y1
    if ($targetWidth -le 0 -or $targetHeight -le 0) { return $null }

    # 小さな文字や赤枠を約1.8倍で見せつつ、周辺の文脈も半画面以上残す。
    $width = [Math]::Min(1.0, [Math]::Max(0.55, $targetWidth + 0.24))
    $height = [Math]::Min(1.0, [Math]::Max(0.55, $targetHeight + 0.24))
    if ($width -ge 0.999999 -and $height -ge 0.999999) { return $null }

    $centerX = ([double]$Rect.x1 + [double]$Rect.x2) / 2.0
    $centerY = ([double]$Rect.y1 + [double]$Rect.y2) / 2.0
    $x = [Math]::Max(0.0, [Math]::Min(1.0 - $width, $centerX - ($width / 2.0)))
    $y = [Math]::Max(0.0, [Math]::Min(1.0 - $height, $centerY - ($height / 2.0)))
    return [pscustomobject]@{
        x = [Math]::Round($x, 6); y = [Math]::Round($y, 6)
        width = [Math]::Round($width, 6); height = [Math]::Round($height, 6)
    }
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
    if ($events.Count -eq 0) { return [pscustomobject]@{ added = 0; skipped = 0 } }

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
    foreach ($record in $events) {
        $index = 0
        try { $index = [int]$record.index } catch { $index = 0 }
        if ($null -ne $wanted -and -not $wanted.Contains($index)) { continue }

        $imagePath = Get-MbRecordedEventImagePath -FileName ([string]$record.image)
        if ([string]::IsNullOrWhiteSpace($imagePath)) { $skipped++; continue }

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

        $targetType = if ($record.PSObject.Properties.Name -contains 'targetType') { [string]$record.targetType } else { '' }
        $crop = Get-MbRecorderTargetCrop -Rect $rect -TargetType $targetType
        if ($null -ne $crop) {
            [void](Set-MbStepImageEdits -Project $Project -StepId $stepId `
                -AnnotationsJson (ConvertTo-Json -InputObject $annotation -Depth 5) `
                -CropJson (ConvertTo-Json -InputObject $crop -Compress))
        } elseif (@($annotation).Count -gt 0) {
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
        [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind $kind `
            -VideoTimeMs ([int]$record.timeMs) -ClickLabel $targetName `
            -WindowTitle ([string]$record.windowTitle) -Narration $spoken -TargetType $targetType `
            -TargetSource $targetSource -TargetConfidence $targetConfidence `
            -TargetCandidateId $candidateId -TargetCandidatesJson $candidatesJson)
        $added++
    }
    return [pscustomobject]@{ added = $added; skipped = $skipped }
}

function Remove-MbRecordingJob {
    if ($null -eq $script:MbRecordingJob) { return }
    $job = $script:MbRecordingJob
    $directory = [string]$job.JobDirectory
    $processId = [int]$job.ProcessId
    $dictationProcessId = 0
    if ($job.PSObject.Properties.Name -contains 'DictationProcessId') { $dictationProcessId = [int]$job.DictationProcessId }
    $uiaWorkerProcessId = 0
    if ($job.PSObject.Properties.Name -contains 'UiaWorkerProcessId') { $uiaWorkerProcessId = [int]$job.UiaWorkerProcessId }
    $stopPath = if ($job.PSObject.Properties.Name -contains 'StopPath') { [string]$job.StopPath } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($stopPath)) {
        try { [IO.File]::WriteAllText($stopPath, 'stop', (New-Object Text.UTF8Encoding($false))) } catch { }
    }
    $script:MbRecordingJob = $null
    # まだ記録プロセスが動いていたら止める。放置すると画面を撮り続ける。
    foreach ($id in @($processId, $dictationProcessId, $uiaWorkerProcessId)) {
        if ($id -le 0) { continue }
        try {
            $process = Get-Process -Id $id -ErrorAction SilentlyContinue
            if ($null -ne $process) { $process.Kill() }
        } catch { }
    }
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
    'Get-MbRecordedEvents',
    'Get-MbRecordedEventImagePath',
    'Import-MbRecordedEvents',
    'Get-MbRecorderTargetCrop',
    'Get-MbRecordedNarration',
    'Merge-MbNarrationIntoEvents',
    'Remove-MbRecordingJob',
    'Get-MbRecordingCapability'
)
