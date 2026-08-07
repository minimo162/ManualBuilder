# ManualBuilder recorder timeline selection worker.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$FramesDirectory,
    [Parameter(Mandatory = $true)][string]$FramesPath,
    [Parameter(Mandatory = $true)][string]$EventsPath,
    [Parameter(Mandatory = $true)][string]$WorkDirectory,
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [Parameter(Mandatory = $true)][string]$CancelPath,
    [Parameter(Mandatory = $true)][string]$JobId,
    [Parameter(Mandatory = $true)][string]$ProfileDirectory,
    [AllowEmptyString()][string]$ConfigPath = '',
    [AllowEmptyString()][string]$LogPath = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.RecorderCopilot.psm1') -Force

$script:Utf8NoBom = New-Object Text.UTF8Encoding($false)
$startedAt = [DateTime]::UtcNow.ToString('o')
$totalPackets = 0

function Write-MbRecorderAiLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    try { [IO.File]::AppendAllText($LogPath, ('[{0}] [{1}] {2}{3}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message, [Environment]::NewLine), $script:Utf8NoBom) } catch { }
}

function Write-MbRecorderAiStatus {
    param([string]$State, [string]$Phase, [string]$Message, [int]$Percent, [int]$CurrentPacket = 0, [int]$ProposalCount = 0, [string]$ErrorCode = '')
    $status = [pscustomobject]@{
        jobId = $JobId; state = $State; phase = $Phase; message = $Message
        percent = [Math]::Max(0, [Math]::Min(100, $Percent))
        currentPacket = $CurrentPacket; totalPackets = $totalPackets; proposalCount = $ProposalCount
        resultPath = $ResultPath; errorCode = $ErrorCode
        startedAt = $startedAt; updatedAt = [DateTime]::UtcNow.ToString('o')
        completedAt = $(if ($State -in @('completed', 'failed', 'cancelled')) { [DateTime]::UtcNow.ToString('o') } else { '' })
    }
    $temporary = $StatusPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $backup = $StatusPath + '.' + [guid]::NewGuid().ToString('N') + '.bak'
    try {
        [IO.File]::WriteAllText($temporary, ($status | ConvertTo-Json -Depth 6), $script:Utf8NoBom)
        foreach ($delay in @(0, 25, 50, 100, 200)) {
            if ($delay -gt 0) { Start-Sleep -Milliseconds $delay }
            try {
                if ([IO.File]::Exists($StatusPath)) { [IO.File]::Replace($temporary, $StatusPath, $backup, $true) }
                else { [IO.File]::Move($temporary, $StatusPath) }
                break
            } catch [IO.IOException] {
                if ($delay -eq 200) { throw }
            } catch [UnauthorizedAccessException] {
                if ($delay -eq 200) { throw }
            }
        }
    } finally {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue
    }
}

function Test-MbRecorderAiCancelled { return (Test-Path -LiteralPath $CancelPath -PathType Leaf) }

try {
    Set-MbCopilotLogger -Logger { param($Message, $Level) Write-MbRecorderAiLog $Message $Level }
    Write-MbRecorderAiStatus -State 'running' -Phase 'preparing' -Message '記録したコマを並べています' -Percent 2
    $allFrames = @(Read-MbRecorderJsonLines -Path $FramesPath)
    $events = @(Read-MbRecorderJsonLines -Path $EventsPath)
    if ($allFrames.Count -lt 2) { throw 'AIが比較できる画面が足りません。もう一度操作を記録してください。' }
    $allFrames = @(Add-MbRecorderFrameVisualMetrics -Frames $allFrames -FramesDirectory $FramesDirectory)
    # 開始準備と記録終了後に前面へ戻った画面は、操作イベントの範囲外なので除く。
    # ManualBuilderやターミナルが一覧へ混ざるとAIが架空手順として扱いやすい。
    $eventWindowFrames = @(Select-MbRecorderEventWindowFrames -Frames $allFrames -Events $events)
    # まずこのPCで操作前／操作後を組み立て、Copilotにはその代表コマだけを渡す。
    # 生の最大30コマを渡すより、入力途中・同一結果・ツールチップの重複が減り、
    # AIは曖昧な候補の取捨選択と文章化へ集中できる。
    $localCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $eventWindowFrames -Events $events -MaximumFrames 30)
    $frames = @(Select-MbRecorderCandidateFrames -Frames $eventWindowFrames -Candidates $localCandidates)
    if ($frames.Count -lt 2) {
        $frames = @(Select-MbRecorderTimelineFrames -Frames $eventWindowFrames -Events $events -Maximum 30)
    }
    $contactDirectory = Join-Path $WorkDirectory 'contact-sheets'
    $sheets = @(New-MbRecorderContactSheets -Frames $frames -FramesDirectory $FramesDirectory -OutputDirectory $contactDirectory)
    if ($sheets.Count -lt 1) { throw 'コンタクトシートを作れませんでした。' }

    $settings = Get-MbCopilotSettings -ConfigPath $ConfigPath
    # Scene selection is a constrained JSON extraction task.  The automatic
    # model is faster and proved more stable than forcing Think Deeper for two
    # large contact sheets.
    $settings.copilot_model = '自動,Automatic,Auto'
    $settings.request_timeout = [Math]::Min(90, [int]$settings.request_timeout)
    # 実機で大きな一覧画像2枚の同時添付がM365汎用エラーになったため1枚ずつ送る。
    $perPacket = 1
    $totalPackets = [int][Math]::Ceiling($sheets.Count / [double]$perPacket)
    $proposals = New-Object System.Collections.ArrayList
    $serviceUnavailable = $false
    $marker = Get-MbCopilotPromptTailAnchor

    for ($packetIndex = 0; $packetIndex -lt $totalPackets; $packetIndex++) {
        if (Test-MbRecorderAiCancelled) {
            Write-MbRecorderAiStatus -State 'cancelled' -Phase 'cancelled' -Message '中止しました' -Percent 100 -ProposalCount $proposals.Count
            exit 0
        }
        $packetNumber = $packetIndex + 1
        $packetSheets = @($sheets | Select-Object -Skip ($packetIndex * $perPacket) -First $perPacket)
        $packetFrames = @($packetSheets | ForEach-Object { @($_.frames) })
        $firstTime = [int]($packetFrames | Measure-Object -Property timeMs -Minimum).Minimum
        $lastTime = [int]($packetFrames | Measure-Object -Property timeMs -Maximum).Maximum
        $packetEvents = @($events | Where-Object { [int]$_.timeMs -ge ($firstTime - 1000) -and [int]$_.timeMs -le ($lastTime + 1000) })
        $prompt = New-MbRecorderCopilotPrompt -Frames $packetFrames -Events $packetEvents -Marker $marker
        $basePercent = 8 + [int](($packetIndex / [double]$totalPackets) * 84)
        Write-MbRecorderAiStatus -State 'running' -Phase 'attaching' `
            -Message ("時系列画像をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets) `
            -Percent $basePercent -CurrentPacket $packetNumber -ProposalCount $proposals.Count
        $phaseCallback = {
            param([string]$Phase)
            $message = switch ($Phase) {
                'attaching' { "時系列画像をCopilotへ渡しています（$packetNumber/$totalPackets）" }
                'sending' { "必要な手順の選択を依頼しています（$packetNumber/$totalPackets）" }
                'waiting' { "Copilotが手順と代表コマを選んでいます（$packetNumber/$totalPackets）" }
                default { "Copilotを準備しています（$packetNumber/$totalPackets）" }
            }
            Write-MbRecorderAiStatus -State 'running' -Phase $Phase -Message $message -Percent $basePercent `
                -CurrentPacket $packetNumber -ProposalCount $proposals.Count
        }
        $maximumAttempts = 2
        for ($attempt = 1; $attempt -le $maximumAttempts; $attempt++) {
            $response = Invoke-MbCopilotRequest -Settings $settings -ProfileDirectory $ProfileDirectory -Prompt $prompt `
                -AttachPaths @($packetSheets | ForEach-Object { [string]$_.path }) -Marker $marker `
                -OnPhase $phaseCallback -ShouldCancel { Test-MbRecorderAiCancelled } -AllowEmptySteps
            if ($response.cancelled) {
                Write-MbRecorderAiStatus -State 'cancelled' -Phase 'cancelled' -Message '中止しました' -Percent 100 -ProposalCount $proposals.Count
                exit 0
            }
            if (-not $response.ok -or $null -eq $response.answer) {
                $tail = if ($response.PSObject.Properties.Name -contains 'tail') { ([string]$response.tail).Trim() } else { '' }
                if ($tail.Length -gt 160) { $tail = $tail.Substring($tail.Length - 160) }
                Write-MbRecorderAiLog ("パケット {0} の回答を読み取れませんでした（{1}/{2}）。{3}" -f `
                    $packetNumber, $attempt, $maximumAttempts, $tail) 'WARN'
                $responseErrorCode = if ($response.PSObject.Properties.Name -contains 'errorCode') {
                    [string]$response.errorCode
                } else { '' }
                if ([string]$response.completedBy -eq 'service-error' -or
                    $responseErrorCode -eq 'COPILOT_SERVICE_UNAVAILABLE' -or
                    (Test-MbCopilotServiceErrorText -Text $tail)) {
                    # The M365 service already returned a terminal error. Sending
                    # the same large image again only doubles the user's wait;
                    # the UI immediately falls back to locally captured events.
                    $serviceUnavailable = $true
                    break
                }
                if ([string]$response.completedBy -eq 'no-json') { break }
                continue
            }
            $converted = @(ConvertFrom-MbRecorderCopilotAnswer -Answer $response.answer -Frames $packetFrames -Events $packetEvents)
            foreach ($proposal in $converted) { [void]$proposals.Add($proposal) }
            if ($converted.Count -gt 0) { break }
            Write-MbRecorderAiLog ("パケット {0} の回答に採用できる手順がありませんでした（{1}/{2}）。" -f `
                $packetNumber, $attempt, $maximumAttempts) 'WARN'
        }
        if ($serviceUnavailable) { break }
    }

    $ordered = @($proposals | Sort-Object { [int]$_.timeMs })
    if ($ordered.Count -lt 1 -and $serviceUnavailable) {
        Write-MbRecorderAiStatus -State 'failed' -Phase 'fallback' `
            -Message 'Copilotは現在応答しないため、AIによる絞り込みだけ省略しました。記録内容は失われていません。' `
            -Percent 100 -ProposalCount 0 -ErrorCode 'COPILOT_SERVICE_UNAVAILABLE'
        exit 1
    }
    if ($ordered.Count -lt 1) { throw 'Copilotが必要な手順を選べませんでした。' }
    [IO.File]::WriteAllText($ResultPath, (([pscustomobject]@{ jobId = $JobId; proposals = $ordered }) | ConvertTo-Json -Depth 10), $script:Utf8NoBom)
    Write-MbRecorderAiStatus -State 'completed' -Phase 'completed' `
        -Message ("Copilotが {0} 件の手順候補を選びました" -f $ordered.Count) -Percent 100 -CurrentPacket $totalPackets -ProposalCount $ordered.Count
    exit 0
} catch {
    Write-MbRecorderAiLog $_.Exception.ToString() 'ERROR'
    try {
        Write-MbRecorderAiStatus -State 'failed' -Phase 'failed' -Message ('AIで手順を整理できませんでした: ' + $_.Exception.Message) `
            -Percent 100 -ProposalCount 0 -ErrorCode 'RECORDER_AI_FAILED'
    } catch { }
    exit 1
}
