# 録画の場面取り込みと、Copilot下書きジョブの管理。
#
# 本体（Start-ManualBuilder.ps1）を大きくしないため、状態と手順をここへ寄せる。
# ジョブの状態はモジュール変数として保持し、リクエストをまたいで残す。

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Ocr.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1')

$script:MbCopilotJobsRoot = ''
$script:MbCopilotScriptRoot = ''
$script:MbCopilotProfileRoot = ''
$script:MbCopilotConfigPath = ''
$script:MbCopilotJob = $null

function Initialize-MbCopilotServer {
    param(
        [Parameter(Mandatory = $true)][string]$JobsRoot,
        [Parameter(Mandatory = $true)][string]$ScriptRoot,
        [Parameter(Mandatory = $true)][string]$ProfileRoot,
        [AllowEmptyString()][string]$ConfigPath = ''
    )
    $script:MbCopilotJobsRoot = $JobsRoot
    $script:MbCopilotScriptRoot = $ScriptRoot
    $script:MbCopilotProfileRoot = $ProfileRoot
    $script:MbCopilotConfigPath = $ConfigPath
}

function Get-MbCopilotServerSettings {
    return (Get-MbCopilotSettings -ConfigPath $script:MbCopilotConfigPath)
}

# ---------------------------------------------------------------------
# 録画の場面を手順として取り込む
# ---------------------------------------------------------------------
function New-MbAnnotationId {
    return 'annotation-' + [guid]::NewGuid().ToString('N')
}

function Get-MbVideoSceneHash {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
}

# 1コマを手順として追加し、操作位置の赤枠と、読み取った文字を書き込む。
function Import-MbVideoScene {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [int]$TimeMs = 0,
        [AllowEmptyString()][string]$RectJson = '',
        [AllowEmptyString()][string]$CandidatesJson = '',
        [switch]$SkipOcr
    )

    # 途中失敗後の再実行で、同じ時刻の同じ場面を二重に追加しない。
    $sceneHash = Get-MbVideoSceneHash -Bytes $Bytes
    $existingImage = @($Project.images | Where-Object { [string]$_.sha256 -eq $sceneHash }) | Select-Object -First 1
    if ($null -ne $existingImage) {
        $targetSheet = @($Project.sheets | Where-Object { [string]$_.id -eq $SheetId }) | Select-Object -First 1
        if ($null -ne $targetSheet) {
            foreach ($existingStep in @($targetSheet.steps | Where-Object { [string]$_.imageId -eq [string]$existingImage.id })) {
                if ($existingStep.PSObject.Properties.Name -contains 'capture' -and $null -ne $existingStep.capture -and
                    [string]$existingStep.capture.kind -eq 'video-scene' -and [int]$existingStep.capture.videoTimeMs -eq $TimeMs) {
                    return [pscustomobject]@{ status = 'duplicate'; stepId = [string]$existingStep.id; clickLabel = [string]$existingStep.capture.clickLabel; ocrAvailable = $false }
                }
            }
        }
    }

    # 同じ画面へ戻る操作も、時刻が異なれば別の手順として残す。画像実体は Add-MbImageAsset の
    # ハッシュ重複排除で共有されるため、ファイルが二重に保存されることはない。
    $added = Add-MbImageStep -Project $Project -ProjectPath $ProjectPath -SheetId $SheetId -Bytes $Bytes `
        -Source 'video' -AllowDuplicateStep
    if ($added.Status -ne 'added') {
        # 同じ画面がすでに取り込まれている。場面分割で拾い切れなかった重複。
        return [pscustomobject]@{ status = $added.Status; stepId = ''; clickLabel = ''; ocrAvailable = $false }
    }
    $step = $added.Step
    $stepId = [string]$step.id

    $rect = $null
    if (-not [string]::IsNullOrWhiteSpace($RectJson)) {
        try { $rect = $RectJson | ConvertFrom-Json } catch { $rect = $null }
        if (-not (Test-MbNormalizedRect -Rect $rect)) { $rect = $null }
    }

    $candidates = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($CandidatesJson)) {
        $parsedCandidates = $null
        try { $parsedCandidates = $CandidatesJson | ConvertFrom-Json } catch { $parsedCandidates = $null }
        foreach ($candidate in @(@($parsedCandidates) | Select-Object -First 4)) {
            if ($null -eq $candidate -or $candidate.PSObject.Properties.Name -notcontains 'rect' -or
                -not (Test-MbNormalizedRect -Rect $candidate.rect)) { continue }
            $candidateId = if ($candidate.PSObject.Properties.Name -contains 'id') { [string]$candidate.id } else { '' }
            if ($candidateId -notmatch '^video-diff-[1-4]$') { $candidateId = 'video-diff-' + ($candidates.Count + 1) }
            [void]$candidates.Add([pscustomobject]@{
                id = $candidateId
                source = 'video-diff'
                confidence = $(if ($candidate.PSObject.Properties.Name -contains 'confidence' -and [string]$candidate.confidence -in @('high', 'medium', 'low')) { [string]$candidate.confidence } else { 'low' })
                label = ''
                targetType = ''
                rect = $candidate.rect
            })
        }
    }
    if ($candidates.Count -eq 0 -and $null -ne $rect) {
        [void]$candidates.Add([pscustomobject]@{
            id = 'video-diff-1'; source = 'video-diff'; confidence = 'low'; label = ''; targetType = ''; rect = $rect
        })
    }

    $clickLabel = ''
    $screenText = ''
    $ocrAvailable = $false
    $imagePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$step.imageId)
    if (-not $SkipOcr -and -not [string]::IsNullOrWhiteSpace($imagePath)) {
        $snapshot = Get-MbOcrSnapshot -Path $imagePath
        $ocrAvailable = [bool]$snapshot.available
        if ($ocrAvailable) { $screenText = [string]$snapshot.text }
        if ($ocrAvailable -and $candidates.Count -gt 0) {
            # 各候補をOCR文字へ寄せる。最大領域だけを確定せず、Copilotが比較できる形で残す。
            foreach ($candidate in @($candidates)) {
                $resolved = Resolve-MbOperationRect -Rect $candidate.rect -Snapshot $snapshot
                $candidate.rect = $resolved.rect
                $candidate.label = [string]$resolved.label
            }
        }
    }

    if ($candidates.Count -gt 0) {
        $rect = $candidates[0].rect
        $clickLabel = [string]$candidates[0].label
    }

    if ($null -ne $rect) {
        $annotation = @([pscustomobject]@{
            id    = New-MbAnnotationId
            type  = 'rect'
            x1    = [Math]::Round([double]$rect.x1, 6)
            y1    = [Math]::Round([double]$rect.y1, 6)
            x2    = [Math]::Round([double]$rect.x2, 6)
            y2    = [Math]::Round([double]$rect.y2, 6)
            label = 0
        })
        [void](Set-MbStepAnnotations -Project $Project -StepId $stepId -AnnotationsJson (ConvertTo-Json -InputObject $annotation -Depth 5))
    }

    # 候補数は確からしさではない。複数あるほど曖昧な場合もあるため、
    # 現在採用している候補自身の評価をそのまま引き継ぐ。
    [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind 'video-scene' -VideoTimeMs $TimeMs `
        -ClickLabel $clickLabel -ScreenText $screenText -TargetSource 'video-diff' `
        -TargetConfidence $(if ($candidates.Count -gt 0) { [string]$candidates[0].confidence } else { '' }) `
        -TargetCandidateId $(if ($candidates.Count -gt 0) { [string]$candidates[0].id } else { '' }) `
        -TargetCandidatesJson $(if ($candidates.Count -gt 0) { ConvertTo-Json -InputObject @($candidates) -Depth 8 -Compress } else { '' }))

    return [pscustomobject]@{
        status       = 'added'
        stepId       = $stepId
        clickLabel   = $clickLabel
        hasRect      = ($null -ne $rect)
        ocrAvailable = $ocrAvailable
    }
}

# ---------------------------------------------------------------------
# Copilot下書きジョブ
# ---------------------------------------------------------------------
function Get-MbCopilotIdleStatus {
    return [pscustomobject]@{
        jobId = ''; state = 'idle'; phase = 'idle'; message = ''; percent = 0
        currentPacket = 0; totalPackets = 0; totalSteps = 0; draftCount = 0
        resultPath = ''; startedAt = ''; updatedAt = ''; completedAt = ''; errorCode = ''
    }
}

function Read-MbCopilotDraftStatus {
    if ($null -eq $script:MbCopilotJob) { return Get-MbCopilotIdleStatus }
    $statusPath = [string]$script:MbCopilotJob.StatusPath
    if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { return Get-MbCopilotIdleStatus }
    $status = $null
    try {
        $raw = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8)
        $status = $raw | ConvertFrom-Json
    } catch {
        # 書き換えの最中に読むと壊れて見えることがある。次の巡回で読み直す。
        return Get-MbCopilotIdleStatus
    }
    if ($null -eq $status) { return Get-MbCopilotIdleStatus }

    # ワーカーが落ちて状態が running のまま残ることがある。プロセスの生死で補正する。
    if ([string]$status.state -in @('queued', 'running')) {
        $alive = $false
        try { $alive = $null -ne (Get-Process -Id ([int]$script:MbCopilotJob.ProcessId) -ErrorAction SilentlyContinue) } catch { $alive = $false }
        if (-not $alive) {
            $status.state = 'failed'
            $status.phase = 'failed'
            $status.message = '下書きの処理が途中で終わりました。もう一度実行してください。'
            $status.errorCode = 'WORKER_LOST'
        }
    }
    return $status
}

# ワーカーはスナップショットの project.json と同じ場所にある images を読む。
# 下書き対象の手順が参照する画像だけをジョブ側へ複製し、実行中の編集や削除から隔離する。
function Copy-MbCopilotDraftSnapshotImages {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$SourceProjectPath,
        [Parameter(Mandatory = $true)][string]$SnapshotProjectPath,
        [switch]$IncludeWritten,
        [ValidateSet('draft', 'review')][string]$Mode = 'draft'
    )

    # 校正は文章だけを渡すため、画像のスナップショットは不要。
    if ($Mode -eq 'review') { return 0 }

    $targetImageIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            $imageId = [string]$step.imageId
            if ([string]::IsNullOrWhiteSpace($imageId)) { continue }
            $needsDraft = [string]::IsNullOrWhiteSpace([string]$step.title) -or
                [string]::IsNullOrWhiteSpace([string]$step.description)
            if ($IncludeWritten -or $needsDraft) { [void]$targetImageIds.Add($imageId) }
        }
    }

    if ($targetImageIds.Count -eq 0) { return 0 }
    $snapshotImageDirectory = Join-Path (Split-Path -Parent ([IO.Path]::GetFullPath($SnapshotProjectPath))) 'images'
    [void](New-Item -ItemType Directory -Path $snapshotImageDirectory -Force)

    $copied = 0
    foreach ($imageId in $targetImageIds) {
        $image = @($Project.images | Where-Object { [string]$_.id -eq $imageId }) | Select-Object -First 1
        if ($null -eq $image) { throw "下書きに使う画像の情報が見つかりません: $imageId" }
        $fileName = [string]$image.fileName
        # Get-MbProject の検証に加え、コピー先でもファイル名だけを受け入れて経路逸脱を防ぐ。
        if ($fileName -notmatch '^image-[a-f0-9]{32}\.(png|jpg|bmp)$' -or
            [IO.Path]::GetFileName($fileName) -ne $fileName) {
            throw "下書きに使う画像ファイル名が不正です: $imageId"
        }
        $sourcePath = Get-MbImageFilePath -Project $Project -ProjectPath $SourceProjectPath -ImageId $imageId
        if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw "下書きに使う画像ファイルが見つかりません: $fileName"
        }
        [IO.File]::Copy($sourcePath, (Join-Path $snapshotImageDirectory $fileName), $true)
        $copied++
    }
    return $copied
}

function Start-MbCopilotDraftJob {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [switch]$IncludeWritten,
        [ValidateSet('draft', 'review')][string]$Mode = 'draft',
        [scriptblock]$WorkerStarter = $null
    )

    $current = Read-MbCopilotDraftStatus
    if ([string]$current.state -in @('queued', 'running')) { return $current }

    $project = Get-MbProject -Path $ProjectPath
    $steps = Get-MbCopilotStepListFromProject -Project $project
    if ($Mode -eq 'review') {
        if ($steps.withText -lt 1) { throw '文章が書かれた手順がありません。先に手順の文章を作ってください。' }
    } else {
        if ($steps.withImage -lt 1) { throw '画像のある手順が1件もありません。録画かスクリーンショットから手順を作ってください。' }
        if (-not $IncludeWritten -and $steps.needsDraft -lt 1) {
            throw 'すべての手順に文章が入っています。書き直したい場合は「すでに書いた手順も対象にする」を選んでください。'
        }
    }
    [void](Save-MbProject -Project $project -Path $ProjectPath)

    $jobId = 'copilot-' + [guid]::NewGuid().ToString('N')
    $jobDirectory = Join-Path $script:MbCopilotJobsRoot $jobId
    [void](New-Item -ItemType Directory -Path $jobDirectory -Force)
    $snapshotPath = Join-Path $jobDirectory 'project.json'
    try {
        [IO.File]::Copy($ProjectPath, $snapshotPath, $true)
        [void](Copy-MbCopilotDraftSnapshotImages -Project $project -SourceProjectPath $ProjectPath `
            -SnapshotProjectPath $snapshotPath -IncludeWritten:$IncludeWritten -Mode $Mode)
    } catch {
        # 不完全なスナップショットを残さず、ワーカーも起動しない。
        Remove-Item -LiteralPath $jobDirectory -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }

    $statusPath = Join-Path $jobDirectory 'status.json'
    $resultPath = Join-Path $jobDirectory 'result.json'
    $cancelPath = Join-Path $jobDirectory 'cancel.requested'
    $logPath = Join-Path $jobDirectory 'copilot.log'

    $queued = [pscustomobject]@{
        jobId = $jobId; state = 'queued'; phase = 'queued'; message = 'Copilotの準備をしています'; percent = 0
        currentPacket = 0; totalPackets = 0; totalSteps = $(if ($Mode -eq 'review') { $steps.withText } else { $steps.needsDraft }); draftCount = 0
        resultPath = $resultPath; startedAt = [DateTime]::UtcNow.ToString('o')
        updatedAt = [DateTime]::UtcNow.ToString('o'); completedAt = ''; errorCode = ''
    }
    [IO.File]::WriteAllText($statusPath, ($queued | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))

    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { throw 'Windows PowerShell 5.1が見つかりません。' }
    $workerPath = Join-Path $script:MbCopilotScriptRoot 'Invoke-ManualBuilderCopilotJob.ps1'
    $quote = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $workerPath),
        '-ProjectPath', (& $quote $snapshotPath),
        '-WorkDirectory', (& $quote $jobDirectory),
        '-StatusPath', (& $quote $statusPath),
        '-ResultPath', (& $quote $resultPath),
        '-CancelPath', (& $quote $cancelPath),
        '-JobId', (& $quote $jobId),
        '-ProfileDirectory', (& $quote $script:MbCopilotProfileRoot),
        '-ConfigPath', (& $quote $script:MbCopilotConfigPath),
        '-LogPath', (& $quote $logPath)
    )
    if ($IncludeWritten) { $arguments += '-IncludeWritten' }
    $arguments += '-Mode'
    $arguments += $Mode

    if ($null -eq $WorkerStarter) {
        $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    } else {
        # プロセス起動を伴わずにジョブ境界を検査できるよう、テスト時だけ起動処理を差し替える。
        $worker = & $WorkerStarter $powerShellPath $arguments
    }
    if ($null -eq $worker -or $worker.PSObject.Properties.Name -notcontains 'Id') {
        throw 'Copilot下書きワーカーを起動できませんでした。'
    }
    $processId = [int]$worker.Id
    if ($worker.PSObject.Methods.Name -contains 'Dispose') { $worker.Dispose() }

    $script:MbCopilotJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $processId; JobDirectory = $jobDirectory
        StatusPath = $statusPath; ResultPath = $resultPath; CancelPath = $cancelPath
        SnapshotPath = $snapshotPath; LogPath = $logPath; StartedAt = Get-Date
    }
    return (Read-MbCopilotDraftStatus)
}

# 画像つきの手順と、下書きが要る手順の数を数える。
function Get-MbCopilotStepListFromProject {
    param([Parameter(Mandatory = $true)][object]$Project)
    $withImage = 0
    $needsDraft = 0
    $withText = 0
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            # 校正は文字だけの手順も対象にするため、画像の有無とは別に数える。
            if ((-not [string]::IsNullOrWhiteSpace([string]$step.title)) -or
                (-not [string]::IsNullOrWhiteSpace([string]$step.description)) -or
                (-not [string]::IsNullOrWhiteSpace([string]$step.note))) {
                $withText++
            }
            if ([string]::IsNullOrWhiteSpace([string]$step.imageId)) { continue }
            $withImage++
            if ([string]::IsNullOrWhiteSpace([string]$step.title) -or [string]::IsNullOrWhiteSpace([string]$step.description)) {
                $needsDraft++
            }
        }
    }
    return [pscustomobject]@{ withImage = $withImage; needsDraft = $needsDraft; withText = $withText }
}

function Request-MbCopilotDraftCancel {
    if ($null -eq $script:MbCopilotJob) { return (Get-MbCopilotIdleStatus) }
    $status = Read-MbCopilotDraftStatus
    if ([string]$status.state -in @('queued', 'running')) {
        [IO.File]::WriteAllText([string]$script:MbCopilotJob.CancelPath, 'cancel', (New-Object Text.UTF8Encoding($false)))
    }
    return (Read-MbCopilotDraftStatus)
}

function Get-MbCopilotDraftResult {
    if ($null -eq $script:MbCopilotJob) { return [pscustomobject]@{ drafts = @(); failures = @() } }
    $resultPath = [string]$script:MbCopilotJob.ResultPath
    if (-not (Test-Path -LiteralPath $resultPath -PathType Leaf)) { return [pscustomobject]@{ drafts = @(); failures = @() } }
    try {
        $raw = [IO.File]::ReadAllText($resultPath, [Text.Encoding]::UTF8)
        $parsed = $raw | ConvertFrom-Json
        if ($null -eq $parsed) { return [pscustomobject]@{ drafts = @(); failures = @() } }
        return $parsed
    } catch {
        return [pscustomobject]@{ drafts = @(); failures = @() }
    }
}

function Remove-MbCopilotDraftJob {
    if ($null -eq $script:MbCopilotJob) { return }
    $directory = [string]$script:MbCopilotJob.JobDirectory
    $script:MbCopilotJob = $null
    if ([string]::IsNullOrWhiteSpace($directory)) { return }
    # 消せなくても次のジョブに支障はない。掃除は次回起動時に任せる。
    try { Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}

function Get-MbCopilotCandidateCrop {
    param([Parameter(Mandatory = $true)]$Rect)
    if (-not (Test-MbNormalizedRect -Rect $Rect)) { return [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 } }
    $targetWidth = [double]$Rect.x2 - [double]$Rect.x1
    $targetHeight = [double]$Rect.y2 - [double]$Rect.y1
    $width = [Math]::Min(1.0, [Math]::Max(0.55, $targetWidth + 0.24))
    $height = [Math]::Min(1.0, [Math]::Max(0.55, $targetHeight + 0.24))
    $centerX = ([double]$Rect.x1 + [double]$Rect.x2) / 2.0
    $centerY = ([double]$Rect.y1 + [double]$Rect.y2) / 2.0
    return [pscustomobject]@{
        x = [Math]::Round([Math]::Max(0.0, [Math]::Min(1.0 - $width, $centerX - ($width / 2.0))), 6)
        y = [Math]::Round([Math]::Max(0.0, [Math]::Min(1.0 - $height, $centerY - ($height / 2.0))), 6)
        width = [Math]::Round($width, 6); height = [Math]::Round($height, 6)
    }
}

function Test-MbCopilotSameRect {
    param([AllowNull()]$First, [AllowNull()]$Second)
    if (-not (Test-MbNormalizedRect -Rect $First) -or -not (Test-MbNormalizedRect -Rect $Second)) { return $false }
    return ([Math]::Abs([double]$First.x1 - [double]$Second.x1) -lt 0.00001 -and
        [Math]::Abs([double]$First.y1 - [double]$Second.y1) -lt 0.00001 -and
        [Math]::Abs([double]$First.x2 - [double]$Second.x2) -lt 0.00001 -and
        [Math]::Abs([double]$First.y2 - [double]$Second.y2) -lt 0.00001)
}

function Set-MbCopilotVisualSelection {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [AllowEmptyString()][string]$TargetCandidateId = '',
        [ValidateSet('keep', 'focus', 'full')][string]$Zoom = 'keep'
    )
    # 候補IDがない拡大指示は、現在の自動赤枠だけを外す危険があるため一切適用しない。
    if ([string]::IsNullOrWhiteSpace($TargetCandidateId)) { return $false }
    $step = Get-MbStepById -Project $Project -StepId $StepId
    if ($null -eq $step -or $null -eq $step.capture) { return $false }
    $candidates = @($step.capture.targetCandidates)
    # 視覚候補がない手順では none や zoom も受け付けない。手動のcropや対象名を守る。
    if ($candidates.Count -eq 0) { return $false }
    $selected = $null
    if ($TargetCandidateId -ne 'none' -and -not [string]::IsNullOrWhiteSpace($TargetCandidateId)) {
        $selected = @($candidates | Where-Object { [string]$_.id -eq $TargetCandidateId }) | Select-Object -First 1
        if ($null -eq $selected) { return $false }
    }

    # 以前の自動候補と一致する矩形だけを外す。利用者が追加した別の赤枠は保持する。
    $oldCandidate = @($candidates | Where-Object { [string]$_.id -eq [string]$step.capture.targetCandidateId }) | Select-Object -First 1
    $annotations = New-Object System.Collections.ArrayList
    $oldRemoved = $false
    foreach ($annotation in @($step.annotations)) {
        if (-not $oldRemoved -and [string]$annotation.type -eq 'rect' -and $null -ne $oldCandidate -and
            (Test-MbCopilotSameRect -First $annotation -Second $oldCandidate.rect)) {
            $oldRemoved = $true
            continue
        }
        [void]$annotations.Add($annotation)
    }
    if ($null -ne $selected) {
        [void]$annotations.Add([pscustomobject]@{
            id = New-MbAnnotationId; type = 'rect'
            x1 = $selected.rect.x1; y1 = $selected.rect.y1; x2 = $selected.rect.x2; y2 = $selected.rect.y2; label = 0
        })
    }

    $crop = $step.crop
    if ($Zoom -eq 'full') { $crop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 } }
    elseif ($Zoom -eq 'focus' -and $null -ne $selected) { $crop = Get-MbCopilotCandidateCrop -Rect $selected.rect }
    [void](Set-MbStepImageEdits -Project $Project -StepId $StepId `
        -AnnotationsJson (ConvertTo-Json -InputObject @($annotations) -Depth 6) `
        -CropJson (ConvertTo-Json -InputObject $crop -Compress))

    if (-not [string]::IsNullOrWhiteSpace($TargetCandidateId)) {
        $step.capture.targetCandidateId = $TargetCandidateId
        if ($null -eq $selected) {
            $step.capture.clickLabel = ''
        } else {
            $step.capture.clickLabel = [string]$selected.label
            $step.capture.targetSource = [string]$selected.source
            $step.capture.targetConfidence = [string]$selected.confidence
            $step.capture.targetType = [string]$selected.targetType
        }
    }
    return $true
}

# 採用された下書きだけをプロジェクトへ書き込む。
function Set-MbCopilotDraftSelection {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$SelectionJson
    )

    if ([string]::IsNullOrWhiteSpace($SelectionJson)) { return 0 }
    $selection = $null
    try { $selection = $SelectionJson | ConvertFrom-Json } catch { throw '採用する下書きの形式が正しくありません。' }
    if ($null -eq $selection) { return 0 }
    $items = @($selection)
    if ($selection.PSObject.Properties.Name -contains 'accept') { $items = @($selection.accept) }
    if ($items.Count -gt 1000) { throw '一度に採用できる手順は1000件までです。' }
    $attentionItems = @()
    if ($selection.PSObject.Properties.Name -contains 'attention') { $attentionItems = @($selection.attention) }
    if ($attentionItems.Count -gt 1000) { throw '一度に確認待ちにできる手順は1000件までです。' }

    $applied = 0
    foreach ($item in $items) {
        if ($null -eq $item) { continue }
        if ($item.PSObject.Properties.Name -notcontains 'id') { continue }
        $title = if ($item.PSObject.Properties.Name -contains 'title') { [string]$item.title } else { '' }
        $description = if ($item.PSObject.Properties.Name -contains 'description') { [string]$item.description } else { '' }
        $note = if ($item.PSObject.Properties.Name -contains 'note') { [string]$item.note } else { '' }
        $targetCandidateId = if ($item.PSObject.Properties.Name -contains 'targetCandidateId') { [string]$item.targetCandidateId } else { '' }
        $zoom = if ($item.PSObject.Properties.Name -contains 'zoom' -and [string]$item.zoom -in @('keep', 'focus', 'full')) { [string]$item.zoom } else { 'keep' }
        $changed = $false
        try {
            $changed = Set-MbStepDraft -Project $Project -StepId ([string]$item.id) -Title $title -Description $description -Note $note
            $visualChanged = Set-MbCopilotVisualSelection -Project $Project -StepId ([string]$item.id) `
                -TargetCandidateId $targetCandidateId -Zoom $zoom
            $changed = $changed -or $visualChanged
            [void](Set-MbStepReview -Project $Project -StepId ([string]$item.id))
        } catch {
            # 採用の途中で手順が消えていた場合。その1件だけ飛ばして続ける。
            continue
        }
        if ($changed) { $applied++ }
    }
    foreach ($item in $attentionItems) {
        if ($null -eq $item -or $item.PSObject.Properties.Name -notcontains 'id') { continue }
        $action = if ($item.PSObject.Properties.Name -contains 'action') { [string]$item.action } else { 'review' }
        if ($action -notin @('review', 'delete')) { continue }
        $reason = if ($item.PSObject.Properties.Name -contains 'reason') { [string]$item.reason } else { '' }
        try { [void](Set-MbStepReview -Project $Project -StepId ([string]$item.id) -Action $action -Reason $reason) } catch { continue }
    }
    return $applied
}

# 画面の文字認識が使えるかを返す。
# Ocrモジュールはこのモジュールの内側にしか読み込まれないため、
# 本体からはこの関数を通して状態を受け取る。
function Get-MbCopilotCapabilities {
    return [pscustomobject]@{ ocr = Get-MbOcrStatus }
}

function Show-MbCopilotSignInWindow {
    $settings = Get-MbCopilotServerSettings
    return (Show-MbCopilotWindow -Settings $settings -ProfileDirectory $script:MbCopilotProfileRoot)
}

Export-ModuleMember -Function @(
    'Initialize-MbCopilotServer',
    'Get-MbCopilotServerSettings',
    'Import-MbVideoScene',
    'Start-MbCopilotDraftJob',
    'Read-MbCopilotDraftStatus',
    'Request-MbCopilotDraftCancel',
    'Get-MbCopilotDraftResult',
    'Remove-MbCopilotDraftJob',
    'Set-MbCopilotDraftSelection',
    'Show-MbCopilotSignInWindow',
    'Get-MbCopilotCapabilities',
    'Get-MbCopilotStepListFromProject'
)
