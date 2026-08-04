# ManualBuilder Copilot draft worker.
#
# 手順の下書きをCopilotへ依頼する専用プロセス。Excel出力と同じ形で、
# 進捗は status.json、結果は result.json、中止は cancel.requested で受け渡す。
# 本体のサーバーとは別プロセスにして、Copilot待ちの間も編集を続けられるようにする。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [AllowEmptyString()][string]$SourceProjectPath = '',
    [Parameter(Mandatory = $true)][string]$WorkDirectory,
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [Parameter(Mandatory = $true)][string]$CancelPath,
    [Parameter(Mandatory = $true)][string]$JobId,
    [Parameter(Mandatory = $true)][string]$ProfileDirectory,
    [AllowEmptyString()][string]$ConfigPath = '',
    [AllowEmptyString()][string]$LogPath = '',
    [ValidateSet('draft', 'review', 'operation')][string]$Mode = 'draft',
    [switch]$IncludeWritten
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.CopilotJob.psm1') -Force

$script:Utf8NoBom = New-Object Text.UTF8Encoding($false)

function Write-MbJobLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    try {
        $line = ('[{0}] [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message)
        [IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine, $script:Utf8NoBom)
    } catch { }
}

function Write-MbJobStatus {
    param([Parameter(Mandatory = $true)][hashtable]$Fields)
    $status = [pscustomobject]$Fields
    $directory = Split-Path -Parent $StatusPath
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    # 読み取り中の破損を避けるため、一時ファイルへ書いてから置き換える。
    $temporary = $StatusPath + '.tmp'
    [IO.File]::WriteAllText($temporary, ($status | ConvertTo-Json -Depth 8), $script:Utf8NoBom)
    [IO.File]::Copy($temporary, $StatusPath, $true)
    Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
}

function Test-MbJobCancelled {
    return (Test-Path -LiteralPath $CancelPath -PathType Leaf)
}

$startedAt = [DateTime]::UtcNow.ToString('o')
$totalPackets = 0
$totalSteps = 0

function New-MbStatusFields {
    param(
        [string]$State, [string]$Phase, [string]$Message, [int]$Percent,
        [int]$CurrentPacket = 0, [int]$DraftCount = 0, [string]$ErrorCode = '', [string]$CompletedAt = ''
    )
    return @{
        jobId         = $JobId
        state         = $State
        phase         = $Phase
        message       = $Message
        percent       = [Math]::Max(0, [Math]::Min(100, $Percent))
        currentPacket = $CurrentPacket
        totalPackets  = $totalPackets
        totalSteps    = $totalSteps
        draftCount    = $DraftCount
        resultPath    = $ResultPath
        startedAt     = $startedAt
        updatedAt     = [DateTime]::UtcNow.ToString('o')
        completedAt   = $CompletedAt
        errorCode     = $ErrorCode
    }
}

try {
    Set-MbCopilotLogger -Logger { param($Message, $Level) Write-MbJobLog $Message $Level }
    Write-MbJobLog "Copilot下書きジョブを開始します: $JobId"
    Write-MbJobStatus -Fields (New-MbStatusFields -State 'running' -Phase 'preparing' -Message '手順を整理しています' -Percent 2)

    $settings = Get-MbCopilotSettings -ConfigPath $ConfigPath
    if ([string]::IsNullOrWhiteSpace($SourceProjectPath)) { $SourceProjectPath = $ProjectPath }
    if ($Mode -eq 'operation') {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'running' -Phase 'preflight' -Message '画像添付の事前確認をしています' -Percent 3)
        $preflight = Test-MbCopilotOperationPreflight -WorkDirectory $WorkDirectory
        $preflightMarker = 'MB_PREFLIGHT_END'
        $preflightPrompt = @"
添付された合成テスト画像を読み取れるか確認します。実際の業務画面ではありません。
画像が3枚添付され、青色と緑色の長方形を確認できた場合は {"steps":[{"id":"preflight","ok":true}]}、確認できない場合はokをfalseにしたJSONだけを返してください。
JSONの後、最後の行に $preflightMarker とだけ書いてください。
$(Get-MbCopilotPromptTailAnchor)
"@
        $preflightResponse = Invoke-MbCopilotRequest -Settings $settings -ProfileDirectory $ProfileDirectory `
            -Prompt $preflightPrompt -AttachPaths @($preflight.attachments) -Marker $preflightMarker `
            -ShouldCancel { Test-MbJobCancelled }
        $preflightOk = $preflightResponse.ok -and $null -ne $preflightResponse.answer -and
            $preflightResponse.answer.PSObject.Properties.Name -contains 'steps' -and
            @($preflightResponse.answer.steps).Count -gt 0 -and
            $preflightResponse.answer.steps[0].PSObject.Properties.Name -contains 'ok' -and
            [bool]$preflightResponse.answer.steps[0].ok
        if (-not $preflightOk) { throw 'Copilotの画像添付セルフテストに失敗しました。Copilot画面の変更、サインイン、添付制限を確認してください。' }
    }
    $project = Get-MbProject -Path $ProjectPath
    $allSteps = Get-MbCopilotStepList -Project $project
    $stepsPerPacket = [int]$settings.steps_per_packet
    if ($Mode -eq 'review') { $stepsPerPacket = [int]$settings.review_steps_per_packet }
    if ($Mode -eq 'operation') { $stepsPerPacket = 1 }
    $packets = Get-MbCopilotPackets -Steps $allSteps -StepsPerPacket $stepsPerPacket -IncludeWritten:$IncludeWritten -Mode $Mode
    $totalPackets = @($packets).Count
    $totalSteps = 0
    foreach ($packet in $packets) { $totalSteps += @($packet).Count }

    if ($totalPackets -eq 0) {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'completed' -Phase 'completed' `
            -Message $(if ($Mode -eq 'review') { '整える文章がありませんでした' } elseif ($Mode -eq 'operation') { '解析待ちの操作がありませんでした' } else { '下書きが必要な手順がありませんでした' }) -Percent 100 -CompletedAt ([DateTime]::UtcNow.ToString('o')))
        [IO.File]::WriteAllText($ResultPath, (([pscustomobject]@{ jobId = $JobId; drafts = @() }) | ConvertTo-Json -Depth 8), $script:Utf8NoBom)
        exit 0
    }

    $styleSamples = Get-MbCopilotStyleSamples -Steps $allSteps -Maximum 3
    $marker = [string]$settings.response_end_marker
    $drafts = New-Object System.Collections.ArrayList
    $failures = New-Object System.Collections.ArrayList
    $shouldCancel = { Test-MbJobCancelled }

    for ($packetIndex = 0; $packetIndex -lt $totalPackets; $packetIndex++) {
        if (Test-MbJobCancelled) { break }
        $packet = @($packets[$packetIndex])
        $packetNumber = $packetIndex + 1
        # 準備・送信・待機の3段でだいたい進むので、パケット単位で均等に割り当てる。
        $basePercent = 5 + [int](($packetIndex / [double]$totalPackets) * 90)

        $startMessage = "画面をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets
        if ($Mode -eq 'review') { $startMessage = "文章をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets }
        elseif ($Mode -eq 'operation') { $startMessage = "操作の前後画像をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets }
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'running' -Phase 'attaching' `
            -Message $startMessage -Percent $basePercent -CurrentPacket $packetNumber -DraftCount $drafts.Count)

        # 添付用の画像を作る。赤枠を焼き込んでおくと、Copilotが操作対象を取り違えない。
        # 校正は文章だけを見るので、画像は作らない。
        $packetDirectory = Join-Path $WorkDirectory ('packet-{0:d3}' -f $packetNumber)
        $attachments = New-Object System.Collections.ArrayList
        $attachmentNames = @{}
        $usableSteps = New-Object System.Collections.ArrayList
        $operationEvidenceNames = $null
        if ($Mode -eq 'review') {
            foreach ($step in $packet) { [void]$usableSteps.Add($step) }
        } elseif ($Mode -eq 'operation') {
            foreach ($step in $packet) {
                $sourcePath = Get-MbImageFilePath -Project $project -ProjectPath $SourceProjectPath -ImageId ([string]$step.imageId)
                if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { continue }
                $afterPath = ''
                if (-not [string]::IsNullOrWhiteSpace([string]$step.afterImageId)) {
                    $afterPath = Get-MbImageFilePath -Project $project -ProjectPath $SourceProjectPath -ImageId ([string]$step.afterImageId)
                }
                $evidence = New-MbCopilotOperationEvidence -Step $step -BeforeSourcePath $sourcePath `
                    -AfterSourcePath $afterPath -WorkDirectory $packetDirectory
                foreach ($attachment in @($evidence.attachments)) { [void]$attachments.Add($attachment) }
                $operationEvidenceNames = $evidence.names
                [void]$usableSteps.Add($step)
            }
        } else {
            foreach ($step in $packet) {
                $sourcePath = Get-MbImageFilePath -Project $project -ProjectPath $SourceProjectPath -ImageId ([string]$step.imageId)
                if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                    Write-MbJobLog ("画像が見つからないため手順を飛ばします: " + [string]$step.id) 'WARN'
                    continue
                }
                $fileName = ('step-{0:d3}.jpg' -f [int]$step.order)
                $attachmentPath = New-MbCopilotAttachment -Step $step -SourcePath $sourcePath -WorkDirectory $packetDirectory -FileName $fileName
                [void]$attachments.Add($attachmentPath)
                $attachmentNames[[string]$step.id] = $fileName
                [void]$usableSteps.Add($step)
            }
        }
        if ($usableSteps.Count -eq 0) {
            $emptyReason = "パケット {0} に使える画像がありませんでした。" -f $packetNumber
            if ($Mode -eq 'review') { $emptyReason = "パケット {0} に整える文章がありませんでした。" -f $packetNumber }
            Write-MbJobLog $emptyReason 'WARN'
            continue
        }

        if ($Mode -eq 'review') {
            $prompt = New-MbCopilotReviewPrompt -Project $project -PacketSteps @($usableSteps) `
                -TotalSteps @($allSteps).Count -Marker $marker
        } elseif ($Mode -eq 'operation') {
            $prompt = New-MbCopilotOperationPrompt -Project $project -Step $usableSteps[0] `
                -EvidenceNames $operationEvidenceNames -StyleSamples $styleSamples -Marker $marker
        } else {
            $prompt = New-MbCopilotStepPrompt -Project $project -PacketSteps @($usableSteps) `
                -AttachmentNames $attachmentNames -StyleSamples $styleSamples -TotalSteps @($allSteps).Count -Marker $marker
        }

        $onPhase = {
            param([string]$Phase)
            $message = switch ($Phase) {
                'attaching' { "画面をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets }
                'sending'   { if ($Mode -eq 'review') { "文章の確認を依頼しています（{0}/{1}）" -f $packetNumber, $totalPackets } elseif ($Mode -eq 'operation') { "操作対象と手順の解析を依頼しています（{0}/{1}）" -f $packetNumber, $totalPackets } else { "手順の下書きを依頼しています（{0}/{1}）" -f $packetNumber, $totalPackets } }
                'waiting'   { "Copilotの回答を待っています（{0}/{1}）" -f $packetNumber, $totalPackets }
                default     { "Copilotの画面を準備しています（{0}/{1}）" -f $packetNumber, $totalPackets }
            }
            Write-MbJobStatus -Fields (New-MbStatusFields -State 'running' -Phase $Phase -Message $message `
                -Percent $basePercent -CurrentPacket $packetNumber -DraftCount $drafts.Count)
        }

        $response = $null
        try {
            $response = Invoke-MbCopilotRequest -Settings $settings -ProfileDirectory $ProfileDirectory `
                -Prompt $prompt -AttachPaths @($attachments) -Marker $marker `
                -OnPhase $onPhase -ShouldCancel $shouldCancel
        } catch {
            Write-MbJobLog ("パケット {0} の依頼に失敗しました: {1}" -f $packetNumber, $_.Exception.Message) 'ERROR'
            [void]$failures.Add([pscustomobject]@{ packet = $packetNumber; message = [string]$_.Exception.Message })
            continue
        }

        if ($response.cancelled) { break }
        if (-not $response.ok -or $null -eq $response.answer) {
            $reason = switch ([string]$response.completedBy) {
                'timeout' { 'Copilotの回答が時間内に終わりませんでした。' }
                'no-json' { 'Copilotが手順の形で答えませんでした。' }
                default   { 'Copilotの回答を読み取れませんでした。' }
            }
            Write-MbJobLog ("パケット {0}: {1} tail={2}" -f $packetNumber, $reason, [string]$response.tail) 'WARN'
            [void]$failures.Add([pscustomobject]@{ packet = $packetNumber; message = $reason })
            continue
        }

        $packetDrafts = ConvertFrom-MbCopilotStepAnswer -Answer $response.answer -PacketSteps @($usableSteps) -Mode $Mode
        foreach ($draft in $packetDrafts) { [void]$drafts.Add($draft) }
        # 1操作ごとに結果を保存し、プロセス中断時も完了済みイベントを復元できるようにする。
        $checkpoint = [pscustomobject]@{ jobId = $JobId; mode = $Mode; drafts = @($drafts); failures = @($failures); completedPackets = $packetNumber }
        [IO.File]::WriteAllText($ResultPath, ($checkpoint | ConvertTo-Json -Depth 12), $script:Utf8NoBom)
        Write-MbJobLog ("パケット {0}/{1} 完了 drafts={2}" -f $packetNumber, $totalPackets, @($packetDrafts).Count)
    }

    $cancelled = Test-MbJobCancelled
    $result = [pscustomobject]@{
        jobId    = $JobId
        drafts   = @($drafts)
        failures = @($failures)
        mode     = $Mode
    }
    [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json -Depth 8), $script:Utf8NoBom)

    if ($cancelled) {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'cancelled' -Phase 'cancelled' `
            -Message '中止しました' -Percent 100 -DraftCount $drafts.Count -CompletedAt ([DateTime]::UtcNow.ToString('o')))
        exit 0
    }

    $message = "{0} 件の下書きができました" -f $drafts.Count
    if ($Mode -eq 'review') { $message = "{0} 件の直したい箇所が見つかりました" -f $drafts.Count }
    elseif ($Mode -eq 'operation') { $message = "{0} 件の操作を解析しました" -f $drafts.Count }
    if ($failures.Count -gt 0) {
        $message += "（{0} 件のまとまりは失敗しました）" -f $failures.Count
    }
    if ($drafts.Count -eq 0) {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'failed' -Phase 'failed' `
            -Message $(if ($Mode -eq 'review') { '直すところは見つかりませんでした' } elseif ($Mode -eq 'operation') { 'Copilotから操作の解析結果を受け取れませんでした' } else { 'Copilotから手順の下書きを受け取れませんでした' }) -Percent 100 -ErrorCode 'NO_DRAFT' `
            -CompletedAt ([DateTime]::UtcNow.ToString('o')))
        exit 1
    }
    Write-MbJobStatus -Fields (New-MbStatusFields -State 'completed' -Phase 'completed' `
        -Message $message -Percent 100 -DraftCount $drafts.Count -CompletedAt ([DateTime]::UtcNow.ToString('o')))
    exit 0
} catch {
    Write-MbJobLog ('ジョブが失敗しました: ' + $_.Exception.Message) 'ERROR'
    try {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'failed' -Phase 'failed' `
            -Message ('手順の下書きを作れませんでした: ' + $_.Exception.Message) -Percent 100 `
            -ErrorCode 'WORKER_FAILED' -CompletedAt ([DateTime]::UtcNow.ToString('o')))
    } catch { }
    exit 1
}
