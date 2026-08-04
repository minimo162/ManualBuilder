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
$script:MbCopilotWarmup = $null
$script:MbCopilotRuntimeStatusPath = ''

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
    if (-not (Test-Path -LiteralPath $JobsRoot)) { [void](New-Item -ItemType Directory -Path $JobsRoot -Force) }
    $script:MbCopilotRuntimeStatusPath = Join-Path $JobsRoot 'runtime-status.json'
}

function Get-MbCopilotServerSettings {
    return (Get-MbCopilotSettings -ConfigPath $script:MbCopilotConfigPath)
}

function Get-MbCopilotRuntimeIdleStatus {
    return [pscustomobject]@{
        state = 'idle'; message = 'Copilot画面をまだ準備していません。'; errorCode = ''
        updatedAt = ''; visible = $false; canOpen = $true
    }
}

function Read-MbCopilotRuntimeStatus {
    $settings = Get-MbCopilotServerSettings
    $ownedProcessIds = @(Get-MbCopilotEdgeProcessIds -Settings $settings -ProfileDirectory $script:MbCopilotProfileRoot)
    $edgeAlive = (Test-MbDevTools -Port ([int]$settings.cdp_port)) -and $ownedProcessIds.Count -gt 0
    $status = $null
    if (-not [string]::IsNullOrWhiteSpace($script:MbCopilotRuntimeStatusPath) -and
        (Test-Path -LiteralPath $script:MbCopilotRuntimeStatusPath -PathType Leaf)) {
        try {
            $raw = [IO.File]::ReadAllText($script:MbCopilotRuntimeStatusPath, [Text.Encoding]::UTF8)
            $status = $raw | ConvertFrom-Json
        } catch { $status = $null }
    }
    if ($null -eq $status) { $status = Get-MbCopilotRuntimeIdleStatus }

    if (-not $edgeAlive) {
        $warmupAlive = $false
        if ($null -ne $script:MbCopilotWarmup) {
            try {
                $warmupAlive = $null -ne (Get-Process -Id ([int]$script:MbCopilotWarmup.ProcessId) -ErrorAction SilentlyContinue)
            } catch { $warmupAlive = $false }
        }
        if (-not $warmupAlive -and [string]$status.state -in @('starting', 'loading', 'ready', 'signin-required')) {
            $status.state = 'closed'
            $status.message = 'Copilot画面が閉じられました。クリックすると再起動します。'
            $status.errorCode = 'COPILOT_EDGE_CLOSED'
        }
    }
    $status | Add-Member -NotePropertyName 'visible' -NotePropertyValue $edgeAlive -Force
    $status | Add-Member -NotePropertyName 'canOpen' -NotePropertyValue $true -Force
    return $status
}

function Start-MbCopilotWarmup {
    param([switch]$Force)

    if ($null -ne $script:MbCopilotWarmup) {
        $alive = $false
        try { $alive = $null -ne (Get-Process -Id ([int]$script:MbCopilotWarmup.ProcessId) -ErrorAction SilentlyContinue) } catch { $alive = $false }
        if ($alive) { return (Read-MbCopilotRuntimeStatus) }
    }

    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { throw 'Windows PowerShell 5.1が見つかりません。' }
    $workerPath = Join-Path $script:MbCopilotScriptRoot 'Invoke-ManualBuilderCopilotWarmup.ps1'
    if (-not (Test-Path -LiteralPath $workerPath -PathType Leaf)) { throw 'Copilot画面の準備プログラムが見つかりません。' }
    $quote = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quote $workerPath),
        '-ProfileDirectory', (& $quote $script:MbCopilotProfileRoot),
        '-ConfigPath', (& $quote $script:MbCopilotConfigPath),
        '-StatusPath', (& $quote $script:MbCopilotRuntimeStatusPath)
    )
    $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $processId = [int]$worker.Id
    $worker.Dispose()
    $script:MbCopilotWarmup = [pscustomobject]@{ ProcessId = $processId; StartedAt = Get-Date }
    return [pscustomobject]@{
        state = 'starting'; message = 'Copilot用Edgeを起動しています。'; errorCode = ''
        updatedAt = [DateTime]::UtcNow.ToString('o'); visible = $false; canOpen = $true
    }
}

function Get-MbCopilotRecorderExclusions {
    $settings = Get-MbCopilotServerSettings
    return [pscustomobject]@{
        processIds = @(Get-MbCopilotEdgeProcessIds -Settings $settings -ProfileDirectory $script:MbCopilotProfileRoot)
        titlePatterns = @('ManualBuilder', 'ManualBuilder Copilot')
    }
}

# ---------------------------------------------------------------------
# 録画の場面を手順として取り込む
# ---------------------------------------------------------------------
function New-MbAnnotationId {
    return 'annotation-' + [guid]::NewGuid().ToString('N')
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
        [switch]$SkipOcr
    )

    $added = Add-MbImageStep -Project $Project -ProjectPath $ProjectPath -SheetId $SheetId -Bytes $Bytes -Source 'video'
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

    $clickLabel = ''
    $screenText = ''
    $ocrAvailable = $false
    $imagePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$step.imageId)
    if (-not $SkipOcr -and -not [string]::IsNullOrWhiteSpace($imagePath)) {
        $snapshot = Get-MbOcrSnapshot -Path $imagePath
        $ocrAvailable = [bool]$snapshot.available
        if ($ocrAvailable) { $screenText = [string]$snapshot.text }
        if ($null -ne $rect -and $ocrAvailable) {
            # 変化領域を、そこにある文字の矩形と突き合わせて締める。
            $resolved = Resolve-MbOperationRect -Rect $rect -Snapshot $snapshot
            $rect = $resolved.rect
            $clickLabel = [string]$resolved.label
        }
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

    [void](Set-MbStepCapture -Project $Project -StepId $stepId -Kind 'video-scene' -VideoTimeMs $TimeMs `
        -ClickLabel $clickLabel -ScreenText $screenText)

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
        jobId = ''; mode = ''; state = 'idle'; phase = 'idle'; message = ''; percent = 0
        currentPacket = 0; totalPackets = 0; totalSteps = 0; draftCount = 0
        resultPath = ''; startedAt = ''; updatedAt = ''; completedAt = ''; errorCode = ''
    }
}

function Restore-MbCopilotDraftJob {
    param([Parameter(Mandatory = $true)][string]$ProjectPath)

    if ($null -ne $script:MbCopilotJob -or -not (Test-Path -LiteralPath $script:MbCopilotJobsRoot -PathType Container)) { return }
    $expectedProjectPath = [IO.Path]::GetFullPath($ProjectPath)
    foreach ($directory in @(Get-ChildItem -LiteralPath $script:MbCopilotJobsRoot -Directory -Filter 'copilot-*' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending)) {
        $metadataPath = Join-Path $directory.FullName 'job.json'
        $statusPath = Join-Path $directory.FullName 'status.json'
        if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf) -or
            -not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { continue }
        try {
            $metadata = [IO.File]::ReadAllText($metadataPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
            if ([IO.Path]::GetFullPath([string]$metadata.sourceProjectPath) -ne $expectedProjectPath) { continue }
            $script:MbCopilotJob = [pscustomobject]@{
                JobId = [string]$metadata.jobId; ProcessId = [int]$metadata.processId
                Mode = [string]$metadata.mode
                JobDirectory = [string]$directory.FullName; StatusPath = $statusPath
                ResultPath = (Join-Path $directory.FullName 'result.json')
                CancelPath = (Join-Path $directory.FullName 'cancel.requested')
                SnapshotPath = (Join-Path $directory.FullName 'project.json')
                LogPath = (Join-Path $directory.FullName 'copilot.log'); StartedAt = $directory.CreationTime
            }
            return
        } catch { $script:MbCopilotJob = $null }
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
    $jobMode = ''
    if ($script:MbCopilotJob.PSObject.Properties.Name -contains 'Mode') { $jobMode = [string]$script:MbCopilotJob.Mode }
    $status | Add-Member -NotePropertyName 'mode' -NotePropertyValue $jobMode -Force

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

function Start-MbCopilotDraftJob {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [switch]$IncludeWritten,
        [ValidateSet('draft', 'review', 'operation')][string]$Mode = 'draft'
    )

    Restore-MbCopilotDraftJob -ProjectPath $ProjectPath
    $current = Read-MbCopilotDraftStatus
    if ([string]$current.state -in @('queued', 'running', 'completed')) {
        if (-not [string]::IsNullOrWhiteSpace([string]$current.mode) -and [string]$current.mode -ne $Mode) {
            throw '別のCopilot処理または確認待ちの結果があります。先にその画面で完了または破棄してください。'
        }
        return $current
    }

    $project = Get-MbProject -Path $ProjectPath
    $steps = Get-MbCopilotStepListFromProject -Project $project
    if ($Mode -eq 'review') {
        if ($steps.withText -lt 1) { throw '文章が書かれた手順がありません。先に手順の文章を作ってください。' }
    } elseif ($Mode -eq 'operation') {
        if ($steps.pendingOperations -lt 1) { throw 'Copilotの解析待ちになっている記録操作がありません。' }
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
    [IO.File]::Copy($ProjectPath, $snapshotPath, $true)

    # 画像は元のプロジェクト側を読む。Copilotへ渡すのは焼き込んだ複製なので、
    # 出力ジョブのように画像一式を退避する必要はない。
    $statusPath = Join-Path $jobDirectory 'status.json'
    $resultPath = Join-Path $jobDirectory 'result.json'
    $cancelPath = Join-Path $jobDirectory 'cancel.requested'
    $logPath = Join-Path $jobDirectory 'copilot.log'

    $queued = [pscustomobject]@{
        jobId = $jobId; mode = $Mode; state = 'queued'; phase = 'queued'; message = 'Copilotの準備をしています'; percent = 0
        currentPacket = 0; totalPackets = 0; totalSteps = $(if ($Mode -eq 'review') { $steps.withText } elseif ($Mode -eq 'operation') { $steps.pendingOperations } else { $steps.needsDraft }); draftCount = 0
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
        '-SourceProjectPath', (& $quote $ProjectPath),
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

    $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $processId = [int]$worker.Id
    $worker.Dispose()

    $script:MbCopilotJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $processId; Mode = $Mode; JobDirectory = $jobDirectory
        StatusPath = $statusPath; ResultPath = $resultPath; CancelPath = $cancelPath
        SnapshotPath = $snapshotPath; LogPath = $logPath; StartedAt = Get-Date
    }
    $metadata = [pscustomobject]@{
        jobId = $jobId; processId = $processId; mode = $Mode
        sourceProjectPath = [IO.Path]::GetFullPath($ProjectPath)
        createdAt = [DateTime]::UtcNow.ToString('o')
    }
    [IO.File]::WriteAllText((Join-Path $jobDirectory 'job.json'), ($metadata | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($false)))
    return (Read-MbCopilotDraftStatus)
}

# 画像つきの手順と、下書きが要る手順の数を数える。
function Get-MbCopilotStepListFromProject {
    param([Parameter(Mandatory = $true)][object]$Project)
    $withImage = 0
    $needsDraft = 0
    $withText = 0
    $pendingOperations = 0
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
            if ($step.capture -and $step.capture.PSObject.Properties.Name -contains 'analysisState' -and
                [string]$step.capture.analysisState -in @('pending', 'needs-review')) { $pendingOperations++ }
            if ([string]::IsNullOrWhiteSpace([string]$step.title) -or [string]::IsNullOrWhiteSpace([string]$step.description)) {
                $needsDraft++
            }
        }
    }
    return [pscustomobject]@{ withImage = $withImage; needsDraft = $needsDraft; withText = $withText; pendingOperations = $pendingOperations }
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

    # 操作解析の矩形・対象名はクライアントから受け取らず、ジョブの検証済み結果を参照する。
    $draftById = @{}
    $jobResult = Get-MbCopilotDraftResult
    foreach ($draft in @($jobResult.drafts)) { $draftById[[string]$draft.id] = $draft }
    $operationMode = $jobResult.PSObject.Properties.Name -contains 'mode' -and [string]$jobResult.mode -eq 'operation'

    $applied = 0
    foreach ($item in $items) {
        if ($null -eq $item) { continue }
        if ($item.PSObject.Properties.Name -notcontains 'id') { continue }
        $title = if ($item.PSObject.Properties.Name -contains 'title') { [string]$item.title } else { '' }
        $description = if ($item.PSObject.Properties.Name -contains 'description') { [string]$item.description } else { '' }
        $note = if ($item.PSObject.Properties.Name -contains 'note') { [string]$item.note } else { '' }
        $changed = $false
        try {
            $changed = Set-MbStepDraft -Project $Project -StepId ([string]$item.id) -Title $title -Description $description -Note $note
        } catch {
            # 採用の途中で手順が消えていた場合。その1件だけ飛ばして続ける。
            continue
        }
        if ($operationMode -and $draftById.ContainsKey([string]$item.id)) {
            $draft = $draftById[[string]$item.id]
            $step = Get-MbStepById -Project $Project -StepId ([string]$item.id)
            if ($null -ne $step) {
                $rect = if ($draft.PSObject.Properties.Name -contains 'targetRect') { $draft.targetRect } else { $null }
                if ($null -ne $rect -and (Test-MbNormalizedRect -Rect $rect)) {
                    $annotations = @($step.annotations | Where-Object { [string]$_.type -eq 'blackout' })
                    $annotations += [pscustomobject]@{
                        id = New-MbAnnotationId; type = 'rect'; label = 0
                        x1 = [Math]::Round([double]$rect.x1, 6); y1 = [Math]::Round([double]$rect.y1, 6)
                        x2 = [Math]::Round([double]$rect.x2, 6); y2 = [Math]::Round([double]$rect.y2, 6)
                    }
                    [void](Set-MbStepAnnotations -Project $Project -StepId ([string]$item.id) `
                        -AnnotationsJson (ConvertTo-Json -InputObject $annotations -Depth 6))
                }
                $label = if ($draft.PSObject.Properties.Name -contains 'targetLabel') { [string]$draft.targetLabel } else { [string]$step.capture.clickLabel }
                $targetType = if ($draft.PSObject.Properties.Name -contains 'targetType') { [string]$draft.targetType } else { '' }
                [void](Set-MbStepCapture -Project $Project -StepId ([string]$item.id) `
                    -Kind ([string]$step.capture.kind) -VideoTimeMs ([int]$step.capture.videoTimeMs) `
                    -ClickLabel $label -WindowTitle ([string]$step.capture.windowTitle) `
                    -ScreenText ([string]$step.capture.screenText) -Narration ([string]$step.capture.narration) `
                    -TargetType $targetType -ClickX ([double]$step.capture.clickX) -ClickY ([double]$step.capture.clickY) `
                    -AfterImageId ([string]$step.capture.afterImageId) -TargetCandidates @($step.capture.targetCandidates) `
                    -AnalysisState 'confirmed' -AnalysisReason ([string]$draft.reason) -NeedsReview:$false)
                $changed = $true
            }
        }
        if ($changed) { $applied++ }
    }
    if ($operationMode -and $selection.PSObject.Properties.Name -contains 'reject') {
        $rejectIds = New-Object System.Collections.ArrayList
        foreach ($rejectIdValue in @($selection.reject)) {
            $rejectId = [string]$rejectIdValue
            if (-not $draftById.ContainsKey($rejectId)) { continue }
            $rejectStep = Get-MbStepById -Project $Project -StepId $rejectId
            if ($null -eq $rejectStep -or $null -eq $rejectStep.capture -or
                [string]$rejectStep.capture.analysisState -notin @('pending', 'needs-review')) { continue }
            [void]$rejectIds.Add($rejectId)
        }
        if ($rejectIds.Count -gt 0) {
            Remove-MbSteps -Project $Project -StepIds @($rejectIds)
            $applied += $rejectIds.Count
        }
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
    'Start-MbCopilotWarmup',
    'Read-MbCopilotRuntimeStatus',
    'Get-MbCopilotRecorderExclusions',
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
