# 操作記録のジョブ管理と、記録した操作の取り込み。
#
# 記録そのものは別プロセス（Invoke-ManualBuilderRecorder.ps1）で行う。
# サーバーは1本のスレッドでHTTPを捌いているため、60Hzのポーリングを同居させると
# 画面の操作が止まる。Excel出力などと同じく、状態は status.json、停止は
# stop.requested のファイルで受け渡す。

Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Recorder.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Dictation.psm1') -Force

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
        $raw = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8)
        $status = $raw | ConvertFrom-Json
    } catch {
        # 書き換えの最中に読むと壊れて見えることがある。次の巡回で読み直す。
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
    }
    return $status
}

function Start-MbRecordingJob {
    param([string[]]$IgnoreTitlePatterns = @('ManualBuilder'), [switch]$WithNarration)

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

    $queued = [pscustomobject]@{
        jobId = $jobId; state = 'recording'; count = 0; message = '記録の準備をしています'
        lastTarget = ''; updatedAt = [DateTime]::UtcNow.ToString('o')
    }
    [IO.File]::WriteAllText($statusPath, ($queued | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))

    # 記録プロセスと音声プロセスが同じ時計を使うよう、開始時刻を揃えて渡す。
    $startedAtUtc = [DateTime]::UtcNow
    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { throw 'Windows PowerShell 5.1が見つかりません。' }
    $workerPath = Join-Path $script:MbRecordingScriptRoot 'Invoke-ManualBuilderRecorder.ps1'
    $quote = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $workerPath),
        '-EventsDirectory', (& $quote $eventsDirectory),
        '-EventsPath', (& $quote $eventsPath),
        '-StatusPath', (& $quote $statusPath),
        '-StopPath', (& $quote $stopPath),
        '-JobId', (& $quote $jobId)
    )
    if (@($IgnoreTitlePatterns).Count -gt 0) {
        $arguments += '-IgnoreTitlePatterns'
        $arguments += (@($IgnoreTitlePatterns) | ForEach-Object { & $quote $_ }) -join ','
    }

    $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $processId = [int]$worker.Id
    $worker.Dispose()

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
# 録画からの取り込みと違い、赤枠の位置も操作対象の名前も UI Automation の確定値なので、
# 推定も文字認識も要らない。ここでは受け取った矩形をそのまま注釈にする。
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
        $result = Add-MbImageStep -Project $Project -ProjectPath $ProjectPath -SheetId $SheetId -Bytes $bytes -Source 'recorder'
        if ($result.Status -ne 'added') {
            # まったく同じ画面が続いた場合。手順を二重に作らない。
            $skipped++
            continue
        }
        $stepId = [string]$result.Step.id

        $rect = $null
        if ($record.PSObject.Properties.Name -contains 'rect') { $rect = $record.rect }
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
            [void](Set-MbStepAnnotations -Project $Project -StepId $stepId -AnnotationsJson (ConvertTo-Json -InputObject $annotation -Depth 5))
        }

        $kind = if ([string]$record.kind -eq 'input') { 'recorded-input' } else { 'recorded-click' }
        $targetName = [string]$record.targetName
        # 入力の手順は「どの欄に入れたか」を示す。入力した文字は記録していない。
        if ($kind -eq 'recorded-input' -and -not [string]::IsNullOrWhiteSpace($targetName)) {
            $targetName = $targetName + '（入力）'
        }
        $spoken = ''
        if ($narration.ContainsKey($index)) { $spoken = [string]$narration[$index] }
        [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind $kind `
            -VideoTimeMs ([int]$record.timeMs) -ClickLabel $targetName `
            -WindowTitle ([string]$record.windowTitle) -Narration $spoken)
        $added++
    }
    return [pscustomobject]@{ added = $added; skipped = $skipped }
}

# 正規化された矩形かどうか。CopilotServer と同じ判定を使いたいが、
# モジュールをまたいで公開すると読み込み順に縛られるため、ここにも持つ。
function Test-MbNormalizedRect {
    param([AllowNull()]$Rect)
    if ($null -eq $Rect) { return $false }
    foreach ($name in @('x1', 'y1', 'x2', 'y2')) {
        if ($Rect.PSObject.Properties.Name -notcontains $name) { return $false }
        $value = 0.0
        try { $value = [double]$Rect.$name } catch { return $false }
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { return $false }
        if ($value -lt 0 -or $value -gt 1) { return $false }
    }
    if (([double]$Rect.x2 - [double]$Rect.x1) -lt 0.004) { return $false }
    if (([double]$Rect.y2 - [double]$Rect.y1) -lt 0.004) { return $false }
    return $true
}

function Remove-MbRecordingJob {
    if ($null -eq $script:MbRecordingJob) { return }
    $job = $script:MbRecordingJob
    $directory = [string]$job.JobDirectory
    $processId = [int]$job.ProcessId
    $dictationProcessId = 0
    if ($job.PSObject.Properties.Name -contains 'DictationProcessId') { $dictationProcessId = [int]$job.DictationProcessId }
    $script:MbRecordingJob = $null
    # まだ記録プロセスが動いていたら止める。放置すると画面を撮り続ける。
    foreach ($id in @($processId, $dictationProcessId)) {
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
    'Get-MbRecordedNarration',
    'Merge-MbNarrationIntoEvents',
    'Remove-MbRecordingJob',
    'Get-MbRecordingCapability'
)
