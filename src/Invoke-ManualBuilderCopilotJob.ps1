# ManualBuilder Copilot draft worker.
#
# 手順の下書きをCopilotへ依頼する専用プロセス。Excel出力と同じ形で、
# 進捗は status.json、結果は result.json、中止は cancel.requested で受け渡す。
# 本体のサーバーとは別プロセスにして、Copilot待ちの間も編集を続けられるようにする。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [Parameter(Mandatory = $true)][string]$WorkDirectory,
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [Parameter(Mandatory = $true)][string]$CancelPath,
    [Parameter(Mandatory = $true)][string]$JobId,
    [Parameter(Mandatory = $true)][string]$ProfileDirectory,
    [AllowEmptyString()][string]$ConfigPath = '',
    [AllowEmptyString()][string]$LogPath = '',
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
    $project = Get-MbProject -Path $ProjectPath
    $allSteps = Get-MbCopilotStepList -Project $project
    $packets = Get-MbCopilotPackets -Steps $allSteps -StepsPerPacket ([int]$settings.steps_per_packet) -IncludeWritten:$IncludeWritten
    $totalPackets = @($packets).Count
    $totalSteps = 0
    foreach ($packet in $packets) { $totalSteps += @($packet).Count }

    if ($totalPackets -eq 0) {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'completed' -Phase 'completed' `
            -Message '下書きが必要な手順がありませんでした' -Percent 100 -CompletedAt ([DateTime]::UtcNow.ToString('o')))
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

        Write-MbJobStatus -Fields (New-MbStatusFields -State 'running' -Phase 'attaching' `
            -Message ("画面をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets) `
            -Percent $basePercent -CurrentPacket $packetNumber -DraftCount $drafts.Count)

        # 添付用の画像を作る。赤枠を焼き込んでおくと、Copilotが操作対象を取り違えない。
        $packetDirectory = Join-Path $WorkDirectory ('packet-{0:d3}' -f $packetNumber)
        $attachments = New-Object System.Collections.ArrayList
        $attachmentNames = @{}
        $usableSteps = New-Object System.Collections.ArrayList
        foreach ($step in $packet) {
            $sourcePath = Get-MbImageFilePath -Project $project -ProjectPath $ProjectPath -ImageId ([string]$step.imageId)
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
        if ($usableSteps.Count -eq 0) {
            Write-MbJobLog ("パケット {0} に使える画像がありませんでした。" -f $packetNumber) 'WARN'
            continue
        }

        $prompt = New-MbCopilotStepPrompt -Project $project -PacketSteps @($usableSteps) `
            -AttachmentNames $attachmentNames -StyleSamples $styleSamples -TotalSteps @($allSteps).Count -Marker $marker

        $onPhase = {
            param([string]$Phase)
            $message = switch ($Phase) {
                'attaching' { "画面をCopilotへ渡しています（{0}/{1}）" -f $packetNumber, $totalPackets }
                'sending'   { "手順の下書きを依頼しています（{0}/{1}）" -f $packetNumber, $totalPackets }
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

        $packetDrafts = ConvertFrom-MbCopilotStepAnswer -Answer $response.answer -PacketSteps @($usableSteps)
        foreach ($draft in $packetDrafts) { [void]$drafts.Add($draft) }
        Write-MbJobLog ("パケット {0}/{1} 完了 drafts={2}" -f $packetNumber, $totalPackets, @($packetDrafts).Count)
    }

    $cancelled = Test-MbJobCancelled
    $result = [pscustomobject]@{
        jobId    = $JobId
        drafts   = @($drafts)
        failures = @($failures)
    }
    [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json -Depth 8), $script:Utf8NoBom)

    if ($cancelled) {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'cancelled' -Phase 'cancelled' `
            -Message '中止しました' -Percent 100 -DraftCount $drafts.Count -CompletedAt ([DateTime]::UtcNow.ToString('o')))
        exit 0
    }

    $message = "{0} 件の下書きができました" -f $drafts.Count
    if ($failures.Count -gt 0) {
        $message += "（{0} 件のまとまりは失敗しました）" -f $failures.Count
    }
    if ($drafts.Count -eq 0) {
        Write-MbJobStatus -Fields (New-MbStatusFields -State 'failed' -Phase 'failed' `
            -Message 'Copilotから手順の下書きを受け取れませんでした' -Percent 100 -ErrorCode 'NO_DRAFT' `
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
