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
    $events = @(Repair-MbRecorderExcelInputEventAnchors -Events $events)
    if ($allFrames.Count -lt 2) { throw 'AIが比較できる画面が足りません。もう一度操作を記録してください。' }
    $allFrames = @(Add-MbRecorderFrameVisualMetrics -Frames $allFrames -FramesDirectory $FramesDirectory)
    # Copilotへ渡す原本はローカル候補で先に決めない。画面変化前後を保護した最大
    # 20コマを時系列に残し、2列×10コマの2枚へ収める。3列表示で読めなかった
    # 小さなセル値や電卓表示を960px幅まで拡大し、待ち時間も2回に抑える。
    $frames = @(Select-MbRecorderCopilotSourceFrames -Frames $allFrames -Events $events -Maximum 20)
    # 画像差分で場面を決めるのではなく、実際のクリック／入力を操作境界の
    # アンカーとしてだけ使う。AIは各グループの代表画像と文章を精査する。
    $interactionGroups = @(New-MbRecorderLocalFrameCandidates -Frames $allFrames -Events $events -MaximumFrames 30)
    $eventMap = @{}
    foreach ($event in $events) { $eventMap[[int]$event.index] = $event }
    $contactDirectory = Join-Path $WorkDirectory 'contact-sheets'
    $sheets = @(New-MbRecorderContactSheets -Frames $frames -FramesDirectory $FramesDirectory -OutputDirectory $contactDirectory)
    if ($sheets.Count -lt 1) { throw 'コンタクトシートを作れませんでした。' }

    $settings = Get-MbCopilotSettings -ConfigPath $ConfigPath
    # Scene selection is a constrained JSON extraction task.  The automatic
    # model is faster and proved more stable than forcing Think Deeper for two
    # large contact sheets.
    $settings.copilot_model = '自動,Automatic,Auto'
    # 大きな一覧画像ではM365側の画像理解だけで90秒を超えることがある。
    # ローカル候補は先に確認できるため、背景処理だけ150秒まで待って途中回答を
    # 同じチャットへ重ねて再送しない。
    $settings.request_timeout = 150
    # 実機で大きな一覧画像2枚の同時添付がM365汎用エラーになったため1枚ずつ送る。
    $perPacket = 1
    $totalPackets = [int][Math]::Ceiling($sheets.Count / [double]$perPacket)
    $proposals = New-Object System.Collections.ArrayList
    $incompletePackets = New-Object System.Collections.ArrayList
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
        $lowerBoundary = -1
        if ($packetIndex -gt 0) {
            $previousPacketSheets = @($sheets | Select-Object -Skip (($packetIndex - 1) * $perPacket) -First $perPacket)
            $previousPacketFrames = @($previousPacketSheets | ForEach-Object { @($_.frames) })
            $previousLastTime = [int]($previousPacketFrames | Measure-Object -Property timeMs -Maximum).Maximum
            $lowerBoundary = [int][Math]::Floor(($previousLastTime + $firstTime) / 2.0)
        }
        $upperBoundary = $lastTime + 1500
        if ($packetIndex + 1 -lt $totalPackets) {
            $nextPacketSheets = @($sheets | Select-Object -Skip (($packetIndex + 1) * $perPacket) -First $perPacket)
            $nextPacketFrames = @($nextPacketSheets | ForEach-Object { @($_.frames) })
            $nextFirstTime = [int]($nextPacketFrames | Measure-Object -Property timeMs -Minimum).Minimum
            $upperBoundary = [int][Math]::Floor(($lastTime + $nextFirstTime) / 2.0)
        }
        $packetGroups = New-Object System.Collections.ArrayList
        foreach ($group in $interactionGroups) {
            $groupIds = @($group.eventIds | ForEach-Object { [int]$_ })
            $groupEvents = @($groupIds | Where-Object { $eventMap.ContainsKey($_) } | ForEach-Object { $eventMap[$_] })
            if ($groupEvents.Count -lt 1) { continue }
            $completionTime = [int]($groupEvents | Measure-Object -Property timeMs -Maximum).Maximum
            if ($completionTime -le $lowerBoundary -or $completionTime -gt $upperBoundary) { continue }
            $firstGroupEvent = @($groupEvents | Sort-Object { [int]$_.timeMs } | Select-Object -First 1)[0]
            [void]$packetGroups.Add([pscustomobject]@{
                eventIds = $groupIds
                targetEventId = [int]$group.targetEventId
                actionKind = [string]$group.actionKind
                targetName = [string]$firstGroupEvent.targetName
                beforeFrame = [string]$group.beforeFrame
                afterFrame = [string]$group.afterFrame
            })
        }
        $packetEventIds = @($packetGroups | ForEach-Object { @($_.eventIds) } | Select-Object -Unique)
        $packetEvents = @($packetEventIds | Where-Object { $eventMap.ContainsKey([int]$_) } | ForEach-Object { $eventMap[[int]$_] } |
            Sort-Object { [int]$_.timeMs }, { [int]$_.index })
        $previousFrameId = ''
        if ($packetIndex -gt 0) {
            $previousSheets = @($sheets | Select-Object -Skip (($packetIndex * $perPacket) - 1) -First 1)
            if ($previousSheets.Count -gt 0 -and @($previousSheets[0].frames).Count -gt 0) {
                $previousFrameId = [string]@($previousSheets[0].frames)[-1].id
            }
        }
        $prompt = New-MbRecorderCopilotPrompt -Frames $packetFrames -Events $packetEvents -InteractionGroups @($packetGroups) `
            -PacketNumber $packetNumber -TotalPackets $totalPackets -PreviousFrameId $previousFrameId -Marker $marker
        # 一覧だけでは読めない「確定後に表示が変わる数式」を原寸で1枚だけ追加する。
        # 実機では一覧1＋原本2の3枚を同時に送ると、M365が汎用サービスエラーを
        # 返すことがある。通常の値は2列一覧で十分読めるため1枚送信にし、
        # paste証拠がある数式だけ一覧＋原本の2枚に抑える。
        $detailFrames = New-Object System.Collections.ArrayList
        $lastInputEvidence = @($packetFrames | Where-Object {
            $_.PSObject.Properties.Name -contains 'role' -and [string]$_.role -eq 'input-evidence' -and
            $_.PSObject.Properties.Name -contains 'evidenceKind' -and [string]$_.evidenceKind -eq 'paste'
        } | Sort-Object { [int]$_.timeMs } -Descending | Select-Object -First 1)
        if ($lastInputEvidence.Count -gt 0) {
            # 数式は確定後のセルに結果値しか残らない。入力中の最終証拠を原寸で
            # 渡し、Copilotが数式バーの文字列を読めるようにする。
            [void]$detailFrames.Add($lastInputEvidence[0])
        }
        $attachPaths = New-Object System.Collections.ArrayList
        foreach ($sheet in $packetSheets) { [void]$attachPaths.Add([string]$sheet.path) }
        foreach ($detailFrame in $detailFrames) {
            $detailPath = Join-Path $FramesDirectory ([string]$detailFrame.image)
            if (Test-Path -LiteralPath $detailPath -PathType Leaf) { [void]$attachPaths.Add($detailPath) }
        }
        if ($detailFrames.Count -gt 0) {
            $detailNote = '追加添付の原寸画像はフレーム ' + (@($detailFrames | ForEach-Object { [string]$_.id }) -join '、') +
                ' です。一覧画像と同じIDの高解像度原本として、表示値・選択状態・ボタン名を確認してください。'
            $prompt = $prompt.Replace($marker, ($detailNote + "`r`n" + $marker))
        }
        $packetNeedsSteps = Test-MbRecorderFrameSetHasMeaningfulChange -Frames $packetFrames
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
        $packetAccepted = $false
        for ($attempt = 1; $attempt -le $maximumAttempts; $attempt++) {
            $attemptPrompt = $prompt
            if ($attempt -gt 1) {
                $retryInstruction = @'
前回はこの一覧から有効な手順を確定できませんでした。先頭と末尾だけでなく全コマを見直し、安定した画面変化、同じブラウザー内のページ遷移、最後の操作結果を漏らさずJSONへ含めてください。曖昧な総称ではなく、画面で読める名称と値を使ってください。
'@
                $attemptPrompt = $prompt.Replace($marker, ($retryInstruction.Trim() + "`r`n" + $marker))
            }
            $response = Invoke-MbCopilotRequest -Settings $settings -ProfileDirectory $ProfileDirectory -Prompt $attemptPrompt `
                -AttachPaths @($attachPaths) -Marker $marker `
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
                if ($tail -match '応答を生成しています|お待ちください|generating') {
                    throw 'Copilotがまだ回答を生成しています。送信中の回答へ重ねて再送せず、しばらく待ってから「Copilotを再確認」を押してください。'
                }
                if ([string]$response.completedBy -eq 'no-json') { break }
                continue
            }
            $converted = @(ConvertFrom-MbRecorderCopilotAnswer -Answer $response.answer -Frames $packetFrames -Events $packetEvents)
            if ($packetGroups.Count -gt 0 -and $converted.Count -ne $packetGroups.Count) {
                Write-MbRecorderAiLog ("パケット {0} は操作グループ {1} 件に対して {2} 件でした（{3}/{4}）。" -f `
                    $packetNumber, $packetGroups.Count, $converted.Count, $attempt, $maximumAttempts) 'WARN'
                if ($attempt -lt $maximumAttempts) { continue }
                $converted = @()
            }
            if ($packetGroups.Count -gt 0 -and $converted.Count -eq $packetGroups.Count) {
                # AIは画像から文章と代表コマを選ぶ。赤枠だけは、同数・同順で返った
                # 場合に限り、実際に記録したクリック座標へ確実に結び直す。
                $orderedConverted = @($converted | Sort-Object { [int]$_.timeMs })
                for ($convertedIndex = 0; $convertedIndex -lt $orderedConverted.Count; $convertedIndex++) {
                    $group = $packetGroups[$convertedIndex]
                    $orderedConverted[$convertedIndex].eventIds = @($group.eventIds)
                    $orderedConverted[$convertedIndex].targetEventId = [int]$group.targetEventId
                }
                $converted = $orderedConverted
            }
            foreach ($proposal in $converted) { [void]$proposals.Add($proposal) }
            if ($converted.Count -gt 0) { $packetAccepted = $true; break }
            $rawStepCount = if ($response.answer.PSObject.Properties.Name -contains 'steps') { @($response.answer.steps).Count } else { 0 }
            if ($rawStepCount -eq 0 -and -not $packetNeedsSteps) { $packetAccepted = $true; break }
            Write-MbRecorderAiLog ("パケット {0} の回答に採用できる手順がありませんでした（{1}/{2}）。" -f `
                $packetNumber, $attempt, $maximumAttempts) 'WARN'
        }
        if ($serviceUnavailable) { break }
        if (-not $packetAccepted -and $packetNeedsSteps) { [void]$incompletePackets.Add($packetNumber) }
    }

    if ($serviceUnavailable) {
        Write-MbRecorderAiStatus -State 'failed' -Phase 'fallback' `
            -Message 'Copilotは途中で応答しなくなったため、不完全な結果は採用しませんでした。記録内容は失われていません。' `
            -Percent 100 -ProposalCount 0 -ErrorCode 'COPILOT_SERVICE_UNAVAILABLE'
        exit 1
    }
    if ($incompletePackets.Count -gt 0) {
        throw ("Copilotが時系列画像 {0} 枚目の操作を確定できませんでした。前半だけを採用せず、もう一度やり直してください。" -f `
            (@($incompletePackets) -join '、'))
    }
    $ordered = @(Add-MbRecorderTitleTransitionProposals -Frames $frames -Proposals @($proposals))
    $ordered = @(Merge-MbRecorderDuplicateTransitionProposals -Frames $frames -Proposals $ordered)
    $ordered = @(Repair-MbRecorderTransientAfterFrames -Frames $frames -Events $events -Proposals $ordered)
    $ordered = @(Expand-MbRecorderExcelRangeSelectionProposals -Events $events -Proposals $ordered)
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
