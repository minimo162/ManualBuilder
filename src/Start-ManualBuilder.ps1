# ManualBuilder Phase 1 foundation server.

[CmdletBinding()]
param(
    [ValidateRange(1024, 65500)][int]$Port = 8765,
    [string]$DataRoot,
    [string]$ProjectPath,
    [string]$LegacyAppRoot,
    # ブラウザーは非表示タブのタイマーを1分に1回まで間引く（Chrome/Edgeの集中スロットリング）。
    # 撮影中はManualBuilderのタブが必ず裏へ回るため、30秒では正常なタブでも失効する。
    [ValidateRange(10, 600)][int]$HeartbeatTimeoutSec = 90,
    # 失効後も、この秒数までは保存先の新着を保留して、タブが戻った時点で取り込む。
    [ValidateRange(0, 3600)][int]$CaptureStandbySec = 900,
    [ValidateRange(60, 1800)][int]$ExcelExportTimeoutSec = 300,
    [ValidateRange(60, 1800)][int]$WordExportTimeoutSec = 300,
    [switch]$DisableScreenshotWatcher,
    [switch]$NoBrowser,
    # 自動テスト専用。NoBrowserはManualBuilderのタブだけを開かない指定で、
    # 通常のローカル起動でもCopilot用Edgeの事前準備は止めない。
    [switch]$SkipCopilotWarmup,
    # 自動E2Eテスト専用。明示したProjectPathとPortごとに別のmutexを使い、
    # 利用中の通常インスタンスを停止せず隔離プロジェクトを検証する。
    [switch]$AllowParallelTestInstance
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$appRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$webRoot = Join-Path $appRoot 'web'
$usesExplicitProjectPath = -not [string]::IsNullOrWhiteSpace($ProjectPath)

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Storage.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Workspace.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Web.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Excel.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.CopilotServer.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.RecorderServer.psm1') -Force

$storageLayout = Get-MbStorageLayout -AppRoot $appRoot -DataRoot $DataRoot -ProjectPath $ProjectPath -LegacyAppRoot $LegacyAppRoot
$DataRoot = [string]$storageLayout.DataRoot
$ProjectPath = [string]$storageLayout.ProjectPath
$runtimePath = [string]$storageLayout.RuntimePath

$script:CaptureVersion = 0
$script:CaptureOwnerTab = $null
$script:CaptureOwnerLastHeartbeat = $null
$script:CaptureOwnerSheetId = $null
$script:Watcher = $null
$script:WatcherState = 'disabled'
$script:WatchDirectory = if ($DisableScreenshotWatcher) { $null } else { Resolve-MbScreenshotDirectory }
$script:ImportWatermark = Get-Date
$script:PendingImages = New-Object System.Collections.ArrayList
# 保留が積み上がり続けないよう上限を設ける（超えたぶんは古い順に捨てる）。
$script:PendingImageLimit = 200
$script:WatcherEventIds = @('ManualBuilder.Capture.Created.' + $PID, 'ManualBuilder.Capture.Renamed.' + $PID)
$script:ExcelExportJob = $null
$script:ExcelExportCancelRequestedAt = $null
$script:ExcelExportCancelReason = ''
$script:ExcelExportJobsRoot = [string]$storageLayout.ExportJobsRoot
$script:ExcelExportWorkerPath = Join-Path $PSScriptRoot 'Export-ManualBuilderExcel.ps1'
$script:WordExportJob = $null
$script:WordExportCancelRequestedAt = $null
$script:WordExportCancelReason = ''
$script:WordExportJobsRoot = [string]$storageLayout.ExportJobsRoot
$script:WordExportWorkerPath = Join-Path $PSScriptRoot 'Export-ManualBuilderWord.ps1'
# Copilot連携。Copilot操作用のプロファイルと設定はユーザーごとのローカル領域に置き、
# 共有フォルダーへアプリを置いても利用者どうしで混ざらないようにする。
$script:CopilotJobsRoot = Join-Path $DataRoot 'copilot-jobs'
$script:CopilotProfileRoot = Join-Path $DataRoot 'copilot-edge-profile'
$script:CopilotConfigPath = Join-Path $DataRoot 'copilot.json'
Initialize-MbCopilotServer -JobsRoot $script:CopilotJobsRoot -ScriptRoot $PSScriptRoot `
    -ProfileRoot $script:CopilotProfileRoot -ConfigPath $script:CopilotConfigPath
# 操作記録。記録した画面はジョブ配下に置き、取り込んだ時点でプロジェクトへ移る。
$script:RecordingJobsRoot = Join-Path $DataRoot 'recording-jobs'
Initialize-MbRecorderServer -JobsRoot $script:RecordingJobsRoot -ScriptRoot $PSScriptRoot
$script:ImageReplacementHistory = @{}
$script:DeletionUndo = $null
$script:ProjectHomeVisible = -not $usesExplicitProjectPath
$script:ActiveProjectKey = if ($usesExplicitProjectPath) { '' } else { 'default' }

function Write-MbLog {
    param([string]$Message, [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $color = switch ($Level) {
        'OK' { 'Green' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Gray' }
    }
    Write-Host ((Get-Date).ToString('HH:mm:ss') + " [$Level] " + $Message) -ForegroundColor $color
}

function Get-MbMutex {
    param([AllowEmptyString()][string]$Scope = '')
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $name = "Local\ManualBuilder-$sid"
    if (-not [string]::IsNullOrWhiteSpace($Scope)) {
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            $scopeBytes = [Text.Encoding]::UTF8.GetBytes($Scope)
            $scopeHash = ([BitConverter]::ToString($sha256.ComputeHash($scopeBytes))).Replace('-', '').Substring(0, 16)
            $name += "-Test-$scopeHash"
        } finally {
            $sha256.Dispose()
        }
    }
    $created = $false
    $abandoned = $false
    $mutex = New-Object System.Threading.Mutex($true, $name, [ref]$created)
    if (-not $created) {
        try {
            if ($mutex.WaitOne(0)) { $created = $true }
        } catch [System.Threading.AbandonedMutexException] {
            $created = $true
            $abandoned = $true
        }
    }
    return [pscustomobject]@{ Mutex = $mutex; Acquired = $created; Abandoned = $abandoned }
}

function Show-MbExistingInstanceNotice {
    if (-not (Test-Path -LiteralPath $runtimePath)) { return $false }
    try {
        $runtime = [IO.File]::ReadAllText($runtimePath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $existingUrl = [string]$runtime.url
        if (($existingUrl -match '^http://localhost:\d+/$') -and
            (Get-Process -Id ([int]$runtime.pid) -ErrorAction SilentlyContinue)) {
            Start-Process $existingUrl
            return $true
        }
    } catch {
        return $false
    }
    return $false
}

function Get-MbCaptureRole {
    param([AllowEmptyString()][string]$TabId)
    if (-not $script:CaptureOwnerTab) { return 'available' }
    if ($TabId -and $TabId -eq $script:CaptureOwnerTab) { return 'owner' }
    return 'viewer'
}

function Update-MbCaptureHeartbeatState {
    if (-not $script:Watcher) {
        $script:WatcherState = 'disabled'
        return
    }
    if (-not $script:CaptureOwnerTab -or -not $script:CaptureOwnerLastHeartbeat) {
        $script:WatcherState = 'suspended'
        return
    }

    $silenceSec = ((Get-Date) - $script:CaptureOwnerLastHeartbeat).TotalSeconds
    if ($silenceSec -le $HeartbeatTimeoutSec) {
        if ($script:WatcherState -eq 'standby') {
            # 保留中の新着は破棄せず、そのまま取り込みへ戻す。取り込み基準時刻も動かさない。
            $script:WatcherState = 'active'
            $heldCount = $script:PendingImages.Count
            if ($heldCount -gt 0) {
                Write-MbLog "ブラウザーが戻ったため、保留していたスクリーンショットを取り込みます: ${heldCount}件" 'OK'
            } else {
                Write-MbLog 'ブラウザーが戻ったためスクリーンショット監視を再開しました。' 'OK'
            }
        } elseif ($script:WatcherState -ne 'active') {
            $script:WatcherState = 'active'
            $script:ImportWatermark = Get-Date
            Write-MbLog "スクリーンショット監視を開始しました: $script:WatchDirectory" 'OK'
        }
        return
    }

    if ($silenceSec -le ($HeartbeatTimeoutSec + $CaptureStandbySec)) {
        if ($script:WatcherState -eq 'active') {
            $script:WatcherState = 'standby'
            Write-MbLog 'ブラウザーのハートビートが届きません。新しいスクリーンショットは取り込まず保留します。' 'WARN'
        }
        return
    }

    if ($script:WatcherState -ne 'suspended') {
        $script:WatcherState = 'suspended'
        $script:PendingImages.Clear()
        $script:ImportWatermark = Get-Date
        Write-MbLog 'ブラウザーのハートビートが途絶えたため監視を一時停止しました。' 'WARN'
    }
}

function Set-MbCaptureHeartbeat {
    param([Parameter(Mandatory = $true)][string]$TabId, [AllowEmptyString()][string]$SheetId = '')

    $ownerExpired = -not $script:CaptureOwnerLastHeartbeat -or
        ((Get-Date) - $script:CaptureOwnerLastHeartbeat).TotalSeconds -gt $HeartbeatTimeoutSec
    if (-not $script:CaptureOwnerTab -or $ownerExpired) {
        if ($script:CaptureOwnerTab -and $script:CaptureOwnerTab -ne $TabId) {
            # 別タブへ撮影対象が移ったときだけ、前のタブ向けの保留を捨てて基準時刻を引き直す。
            $script:PendingImages.Clear()
            $script:ImportWatermark = Get-Date
        }
        $script:CaptureOwnerTab = $TabId
        $script:CaptureOwnerLastHeartbeat = Get-Date
        if ($SheetId) { $script:CaptureOwnerSheetId = $SheetId }
        Write-MbLog "撮影対象タブを設定しました: $TabId" 'OK'
    } elseif ($script:CaptureOwnerTab -eq $TabId) {
        $script:CaptureOwnerLastHeartbeat = Get-Date
        if ($SheetId) { $script:CaptureOwnerSheetId = $SheetId }
    }
    Update-MbCaptureHeartbeatState
    return Get-MbCaptureRole -TabId $TabId
}

function Initialize-MbScreenshotWatcher {
    if (-not $script:WatchDirectory) {
        $script:WatcherState = 'disabled'
        Write-MbLog 'スクリーンショット保存先を検出できません。貼り付け・ドロップ・画像選択は利用できます。' 'WARN'
        return
    }
    try {
        $script:Watcher = New-Object IO.FileSystemWatcher
        $script:Watcher.Path = $script:WatchDirectory
        $script:Watcher.Filter = '*.*'
        $script:Watcher.IncludeSubdirectories = $false
        Register-ObjectEvent -InputObject $script:Watcher -EventName Created -SourceIdentifier $script:WatcherEventIds[0] | Out-Null
        Register-ObjectEvent -InputObject $script:Watcher -EventName Renamed -SourceIdentifier $script:WatcherEventIds[1] | Out-Null
        $script:Watcher.EnableRaisingEvents = $true
        $script:WatcherState = 'suspended'
        Write-MbLog "スクリーンショット監視の準備ができました: $script:WatchDirectory" 'OK'
    } catch {
        $script:WatcherState = 'disabled'
        if ($script:Watcher) { try { $script:Watcher.Dispose() } catch { }; $script:Watcher = $null }
        Write-MbLog "スクリーンショット監視を開始できません: $($_.Exception.Message)" 'WARN'
    }
}

function Test-MbWatchCandidate {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([IO.Path]::GetFileName($Path) -like '~$*') { return $false }
    return [IO.Path]::GetExtension($Path).ToLowerInvariant() -in @('.png', '.jpg', '.jpeg', '.bmp')
}

function Invoke-MbWatcherFlush {
    if (-not $script:Watcher) { return }
    foreach ($sourceId in $script:WatcherEventIds) {
        foreach ($eventItem in @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue)) {
            $path = $null
            try { $path = [string]$eventItem.SourceEventArgs.FullPath } catch { }
            Remove-Event -EventIdentifier $eventItem.EventIdentifier -ErrorAction SilentlyContinue
            if (-not $path -or -not (Test-MbWatchCandidate -Path $path)) { continue }
            # standbyでも受け付ける。裏へ回ったタブのハートビートが遅れただけの可能性があるため、
            # ここで捨てると撮影したスクリーンショットが二度と取り込めなくなる。
            if ($script:WatcherState -notin @('active', 'standby')) { continue }
            if (@($script:PendingImages | Where-Object { $_.Path -eq $path }).Count -eq 0) {
                [void]$script:PendingImages.Add([pscustomobject]@{ Path = $path; Size = [long]-1; Tries = 0 })
            }
        }
    }
    while ($script:PendingImages.Count -gt $script:PendingImageLimit) {
        $script:PendingImages.RemoveAt(0)
    }

    # 保留中は取り込まない。再試行回数も進めず、タブが戻るまでそのまま待つ。
    if ($script:WatcherState -ne 'active') { return }
    if ($script:PendingImages.Count -eq 0) { return }
    $completed = New-Object System.Collections.ArrayList
    foreach ($pending in @($script:PendingImages)) {
        $pending.Tries = [int]$pending.Tries + 1
        if (-not (Test-Path -LiteralPath $pending.Path -PathType Leaf)) {
            if ($pending.Tries -gt 60) { [void]$completed.Add($pending) }
            continue
        }
        $file = $null
        try { $file = Get-Item -LiteralPath $pending.Path -ErrorAction Stop } catch { }
        if (-not $file) { continue }
        if ("$($file.Attributes)" -match 'Offline|RecallOnDataAccess') {
            if ($pending.Tries -gt 100) { [void]$completed.Add($pending) }
            continue
        }
        if ($file.CreationTime -lt $script:ImportWatermark) {
            [void]$completed.Add($pending)
            continue
        }
        if ($file.Length -lt 1 -or $file.Length -ne $pending.Size) {
            $pending.Size = [long]$file.Length
            if ($pending.Tries -gt 60) { [void]$completed.Add($pending) }
            continue
        }
        if ($file.Length -gt (20 * 1024 * 1024)) {
            Write-MbLog "20MBを超えるため取り込みません: $($file.Name)" 'WARN'
            [void]$completed.Add($pending)
            continue
        }
        try {
            $stream = [IO.File]::Open($pending.Path, 'Open', 'Read', 'None')
            try {
                $bytes = New-Object byte[] ([int]$stream.Length)
                $offset = 0
                while ($offset -lt $bytes.Length) {
                    $read = $stream.Read($bytes, $offset, $bytes.Length - $offset)
                    if ($read -le 0) { break }
                    $offset += $read
                }
                if ($offset -ne $bytes.Length) { throw '画像ファイルを最後まで読み込めません。' }
            } finally {
                $stream.Dispose()
            }
            $project = Get-MbProject -Path $ProjectPath
            $targetSheetId = [string]$script:CaptureOwnerSheetId
            if (-not $targetSheetId -or @($project.sheets | Where-Object { $_.id -eq $targetSheetId }).Count -eq 0) {
                $targetSheetId = [string]$project.selectedSheetId
            }
            $result = Add-MbImageStep -Project $project -ProjectPath $ProjectPath -SheetId $targetSheetId -Bytes $bytes -Source watcher
            if ($result.Status -eq 'added') {
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                $script:CaptureVersion++
                Write-MbLog "スクリーンショットを追加しました: $($file.Name)" 'OK'
            }
        } catch {
            Write-MbLog "画像を取り込めません: $($file.Name) / $($_.Exception.Message)" 'WARN'
        }
        [void]$completed.Add($pending)
    }
    foreach ($item in $completed) { [void]$script:PendingImages.Remove($item) }
}

function Stop-MbScreenshotWatcher {
    foreach ($sourceId in $script:WatcherEventIds) {
        try { Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue } catch { }
        foreach ($eventItem in @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue)) {
            Remove-Event -EventIdentifier $eventItem.EventIdentifier -ErrorAction SilentlyContinue
        }
    }
    if ($script:Watcher) {
        try { $script:Watcher.EnableRaisingEvents = $false } catch { }
        try { $script:Watcher.Dispose() } catch { }
        $script:Watcher = $null
    }
}

function Get-MbExcelOutputDirectory {
    $documents = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($documents)) { $documents = Join-Path $appRoot 'exports' }
    return [IO.Path]::GetFullPath((Join-Path $documents 'ManualBuilder'))
}

function Write-MbExportServerStatus {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][object]$Status)
    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    $temporaryPath = Join-Path $directory ('.server-status-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = Join-Path $directory ('.server-status-backup-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($temporaryPath, ($Status | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Path) {
            [IO.File]::Replace($temporaryPath, $Path, $backupPath, $true)
        } else {
            [IO.File]::Move($temporaryPath, $Path)
        }
    } finally {
        foreach ($temporaryFile in @($temporaryPath, $backupPath)) {
            if (Test-Path -LiteralPath $temporaryFile) { Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Stop-MbExcelExportWorkerSafely {
    param([Parameter(Mandatory = $true)][object]$Status)
    if (-not $script:ExcelExportJob) { return }

    if ([bool]$Status.ownershipProven -and [int]$Status.ownedExcelPid -gt 0 -and $Status.ownedExcelStartTimeUtc) {
        $excelProcess = Get-Process -Id ([int]$Status.ownedExcelPid) -ErrorAction SilentlyContinue
        if ($excelProcess -and $excelProcess.ProcessName -eq 'EXCEL') {
            $expectedStart = [DateTime]::Parse([string]$Status.ownedExcelStartTimeUtc).ToUniversalTime()
            $actualStart = $excelProcess.StartTime.ToUniversalTime()
            if ([Math]::Abs(($actualStart - $expectedStart).TotalSeconds) -lt 2) {
                Stop-Process -Id $excelProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
    Start-Sleep -Milliseconds 500
    $worker = Get-Process -Id ([int]$script:ExcelExportJob.ProcessId) -ErrorAction SilentlyContinue
    if ($worker -and $worker.ProcessName -match '^powershell$') {
        Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue
    }
}

function Remove-MbExcelExportSnapshot {
    if (-not $script:ExcelExportJob -or $script:ExcelExportJob.CleanupDone) { return }
    $root = [IO.Path]::GetFullPath($script:ExcelExportJobsRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $jobDirectory = [IO.Path]::GetFullPath([string]$script:ExcelExportJob.JobDirectory)
    if (-not $jobDirectory.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { return }
    foreach ($target in @('project.json', 'project.json.bak', 'images', 'rendered-images')) {
        $path = Join-Path $jobDirectory $target
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $script:ExcelExportJob.CleanupDone = $true
}

function Read-MbExcelExportStatus {
    if (-not $script:ExcelExportJob) {
        return [pscustomobject]@{
            jobId = ''; state = 'idle'; phase = 'idle'; message = 'Excel出力を開始できます'; percent = 0
            currentStep = 0; totalSteps = 0; outputPath = ''; outputName = ''
            outputFolder = ''; outputFolderName = ''; videoCount = 0
            outputDirectory = (Get-MbExcelOutputDirectory); sheetNameMappings = @()
            startedAt = ''; updatedAt = ''; completedAt = ''; errorCode = ''
        }
    }

    $status = $null
    for ($attempt = 0; $attempt -lt 3 -and -not $status; $attempt++) {
        try {
            if (Test-Path -LiteralPath $script:ExcelExportJob.StatusPath -PathType Leaf) {
                $status = [IO.File]::ReadAllText($script:ExcelExportJob.StatusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
            }
        } catch {
            if ($attempt -lt 2) { Start-Sleep -Milliseconds 30 }
        }
    }
    if (-not $status) {
        $status = [pscustomobject]@{
            jobId = [string]$script:ExcelExportJob.JobId; state = 'queued'; phase = 'queued'
            message = 'Excel出力を開始しています'; percent = 0; currentStep = 0
            totalSteps = [int]$script:ExcelExportJob.TotalSteps; outputPath = ''; outputName = ''
            outputFolder = ''; outputFolderName = ''; videoCount = 0
            outputDirectory = [string]$script:ExcelExportJob.OutputDirectory
            sheetNameMappings = @()
            ownedExcelPid = 0; ownedExcelStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false
            startedAt = [string]$script:ExcelExportJob.StartedAt; updatedAt = ''; completedAt = ''; errorCode = ''
        }
    }

    $process = Get-Process -Id ([int]$script:ExcelExportJob.ProcessId) -ErrorAction SilentlyContinue
    $active = [string]$status.state -in @('queued', 'running', 'finalizing')
    if ($active -and $process -and -not $script:ExcelExportCancelRequestedAt) {
        $elapsed = (Get-Date) - [DateTime]$script:ExcelExportJob.StartedAt
        if ($elapsed.TotalSeconds -gt $ExcelExportTimeoutSec) {
            [IO.File]::WriteAllText($script:ExcelExportJob.CancelPath, 'timeout', (New-Object Text.UTF8Encoding($false)))
            $script:ExcelExportCancelRequestedAt = Get-Date
            $script:ExcelExportCancelReason = 'timeout'
            $status.message = '規定時間を超えたため、安全に中止しています'
        }
    }
    if ($active -and $process -and $script:ExcelExportCancelRequestedAt -and
        ((Get-Date) - $script:ExcelExportCancelRequestedAt).TotalSeconds -gt 30) {
        Stop-MbExcelExportWorkerSafely -Status $status
        $status.state = if ($script:ExcelExportCancelReason -eq 'timeout') { 'failed' } else { 'cancelled' }
        $status.phase = [string]$status.state
        $status.message = if ($script:ExcelExportCancelReason -eq 'timeout') {
            'Excelが規定時間内に応答しなかったため、出力処理を停止しました'
        } else { 'Excel作成を中止しました' }
        $status.errorCode = if ($script:ExcelExportCancelReason -eq 'timeout') { 'EXPORT_TIMEOUT' } else { 'CANCELLED' }
        $status.completedAt = [DateTime]::UtcNow.ToString('o')
        try { Write-MbExportServerStatus -Path $script:ExcelExportJob.StatusPath -Status $status } catch { }
        $process = $null
        $active = $false
    }
    if ($active -and -not $process) {
        $status.state = 'failed'
        $status.phase = 'failed'
        $status.message = 'Excel出力プロセスが完了結果を返さず終了しました'
        $status.errorCode = 'WORKER_EXITED_WITHOUT_RESULT'
        $status.completedAt = [DateTime]::UtcNow.ToString('o')
        try { Write-MbExportServerStatus -Path $script:ExcelExportJob.StatusPath -Status $status } catch { }
    }
    if ([string]$status.state -in @('completed', 'cancelled', 'failed')) {
        Remove-MbExcelExportSnapshot
    }
    return $status
}

function Start-MbExcelExportJob {
    $current = Read-MbExcelExportStatus
    if ([string]$current.state -in @('queued', 'running', 'finalizing')) { return $current }
    if ($script:WordExportJob) {
        $wordStatus = Read-MbWordExportStatus
        if ([string]$wordStatus.state -in @('queued', 'running', 'finalizing')) {
            throw 'Word作成中です。完了または中止してからExcelを作成してください。'
        }
    }

    $project = Get-MbProject -Path $ProjectPath
    $totalSteps = 0
    foreach ($sheet in @($project.sheets)) { $totalSteps += @($sheet.steps).Count }
    if ($totalSteps -lt 1) { throw 'Excelへ出力する手順を1件以上追加してください。' }
    [void](Save-MbProject -Project $project -Path $ProjectPath)

    $jobId = 'export-' + [guid]::NewGuid().ToString('N')
    $jobDirectory = Join-Path $script:ExcelExportJobsRoot $jobId
    $snapshotPath = Join-Path $jobDirectory 'project.json'
    $snapshotImageDirectory = Join-Path $jobDirectory 'images'
    $statusPath = Join-Path $jobDirectory 'status.json'
    $cancelPath = Join-Path $jobDirectory 'cancel.requested'
    $outputDirectory = Get-MbExcelOutputDirectory
    [void](New-Item -ItemType Directory -Path $snapshotImageDirectory -Force)
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    [IO.File]::Copy($ProjectPath, $snapshotPath, $true)

    $sourceImageDirectory = Join-Path (Split-Path -Parent $ProjectPath) 'images'
    foreach ($image in @($project.images)) {
        $fileName = [string]$image.fileName
        if ($fileName -notmatch '^image-[a-f0-9]{32}\.(png|jpg|bmp)$') { throw "画像ファイル名が不正です: $fileName" }
        $sourcePath = Join-Path $sourceImageDirectory $fileName
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "画像ファイルが見つかりません: $fileName" }
        [IO.File]::Copy($sourcePath, (Join-Path $snapshotImageDirectory $fileName), $true)
    }

    # 動画つきの手順があるとExcelはフォルダー出力になり、別プロセスが動画を読む。
    # 出力中に元が差し替わっても影響しないよう、画像と同じくスナップショットへ複製する。
    $referencedVideoIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($sheet in @($project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            if ($step.PSObject.Properties.Name -contains 'videoId' -and -not [string]::IsNullOrWhiteSpace([string]$step.videoId)) {
                [void]$referencedVideoIds.Add([string]$step.videoId)
            }
        }
    }
    if ($referencedVideoIds.Count -gt 0) {
        $sourceVideoDirectory = Join-Path (Split-Path -Parent $ProjectPath) 'videos'
        $snapshotVideoDirectory = Join-Path $jobDirectory 'videos'
        [void](New-Item -ItemType Directory -Path $snapshotVideoDirectory -Force)
        foreach ($video in @($project.videos)) {
            if (-not $referencedVideoIds.Contains([string]$video.id)) { continue }
            $fileName = [string]$video.fileName
            if ($fileName -notmatch '^video-[a-f0-9]{32}\.(mp4|webm)$') { throw "動画ファイル名が不正です: $fileName" }
            $sourcePath = Join-Path $sourceVideoDirectory $fileName
            if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "動画ファイルが見つかりません: $fileName" }
            [IO.File]::Copy($sourcePath, (Join-Path $snapshotVideoDirectory $fileName), $true)
        }
    }

    $queuedStatus = [pscustomobject]@{
        jobId = $jobId; state = 'queued'; phase = 'queued'; message = 'Excel出力を開始しています'; percent = 0
        currentStep = 0; totalSteps = $totalSteps; outputPath = ''; outputName = ''; outputDirectory = $outputDirectory
        outputFolder = ''; outputFolderName = ''; videoCount = 0
        sheetNameMappings = @()
        ownedExcelPid = 0; ownedExcelStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false
        startedAt = [DateTime]::UtcNow.ToString('o'); updatedAt = [DateTime]::UtcNow.ToString('o'); completedAt = ''; errorCode = ''
    }
    Write-MbExportServerStatus -Path $statusPath -Status $queuedStatus

    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { throw 'Windows PowerShell 5.1が見つかりません。' }
    $quoted = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quoted $script:ExcelExportWorkerPath),
        '-ProjectPath', (& $quoted $snapshotPath), '-OutputDirectory', (& $quoted $outputDirectory),
        '-StatusPath', (& $quoted $statusPath), '-CancelPath', (& $quoted $cancelPath), '-JobId', (& $quoted $jobId)
    )
    $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $workerId = [int]$worker.Id
    $worker.Dispose()
    $script:ExcelExportJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $workerId; StatusPath = $statusPath; CancelPath = $cancelPath
        JobDirectory = $jobDirectory; OutputDirectory = $outputDirectory; StartedAt = Get-Date; TotalSteps = $totalSteps
        CleanupDone = $false
    }
    $script:ExcelExportCancelRequestedAt = $null
    $script:ExcelExportCancelReason = ''
    Write-MbLog "Excel出力を開始しました: $jobId" 'OK'
    return Read-MbExcelExportStatus
}

function Request-MbExcelExportCancel {
    $status = Read-MbExcelExportStatus
    if ([string]$status.state -in @('queued', 'running')) {
        [IO.File]::WriteAllText($script:ExcelExportJob.CancelPath, 'cancel', (New-Object Text.UTF8Encoding($false)))
        $script:ExcelExportCancelRequestedAt = Get-Date
        $script:ExcelExportCancelReason = 'user'
        $status.message = 'Excel作成を安全に中止しています'
    }
    return $status
}

function Open-MbExcelExportResult {
    param([ValidateSet('file', 'folder')][string]$Mode)
    $status = Read-MbExcelExportStatus
    if ([string]$status.state -ne 'completed') { throw '完成したExcelファイルがありません。' }
    $root = [IO.Path]::GetFullPath([string]$status.outputDirectory).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $path = [IO.Path]::GetFullPath([string]$status.outputPath)
    if (-not $path.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw '出力ファイルの場所を確認できません。'
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw '出力したExcelファイルが見つかりません。' }
    if ($Mode -eq 'file') {
        Start-Process -FilePath $path
        return $status
    }
    # 動画つきはフォルダー出力になる。その場合は保存先の親ではなく、そのフォルダー自体を開く。
    $folderPath = $root
    $outputFolder = if ($status.PSObject.Properties.Name -contains 'outputFolder') { [string]$status.outputFolder } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($outputFolder)) {
        $candidate = [IO.Path]::GetFullPath($outputFolder)
        if ($candidate.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
            (Test-Path -LiteralPath $candidate -PathType Container)) {
            $folderPath = $candidate
        }
    }
    Start-Process -FilePath $folderPath
    return $status
}

function Stop-MbWordExportWorkerSafely {
    param([Parameter(Mandatory = $true)][object]$Status)
    if (-not $script:WordExportJob) { return }
    if ([bool]$Status.ownershipProven -and [int]$Status.ownedWordPid -gt 0 -and $Status.ownedWordStartTimeUtc) {
        $wordProcess = Get-Process -Id ([int]$Status.ownedWordPid) -ErrorAction SilentlyContinue
        if ($wordProcess -and $wordProcess.ProcessName -eq 'WINWORD') {
            $expectedStart = [DateTime]::Parse([string]$Status.ownedWordStartTimeUtc).ToUniversalTime()
            $actualStart = $wordProcess.StartTime.ToUniversalTime()
            if ([Math]::Abs(($actualStart - $expectedStart).TotalSeconds) -lt 2 -and [string]$Status.ownershipMode -eq 'Hwnd') {
                Stop-Process -Id $wordProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
    Start-Sleep -Milliseconds 500
    $worker = Get-Process -Id ([int]$script:WordExportJob.ProcessId) -ErrorAction SilentlyContinue
    if ($worker -and $worker.ProcessName -match '^powershell$') { Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue }
}

function Remove-MbWordExportSnapshot {
    if (-not $script:WordExportJob -or $script:WordExportJob.CleanupDone) { return }
    $root = [IO.Path]::GetFullPath($script:WordExportJobsRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $jobDirectory = [IO.Path]::GetFullPath([string]$script:WordExportJob.JobDirectory)
    if (-not $jobDirectory.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { return }
    foreach ($target in @('project.json', 'project.json.bak', 'images', 'rendered-word-images')) {
        $path = Join-Path $jobDirectory $target
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $script:WordExportJob.CleanupDone = $true
}

function Read-MbWordExportStatus {
    if (-not $script:WordExportJob) {
        return [pscustomobject]@{
            jobId = ''; state = 'idle'; phase = 'idle'; message = 'Word出力を開始できます'; percent = 0
            currentStep = 0; totalSteps = 0; outputPath = ''; outputName = ''; outputDirectory = (Get-MbExcelOutputDirectory)
            ownedWordPid = 0; ownedWordStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false; pageCount = 0
            startedAt = ''; updatedAt = ''; completedAt = ''; errorCode = ''
        }
    }
    $status = $null
    for ($attempt = 0; $attempt -lt 3 -and -not $status; $attempt++) {
        try {
            if (Test-Path -LiteralPath $script:WordExportJob.StatusPath -PathType Leaf) {
                $status = [IO.File]::ReadAllText($script:WordExportJob.StatusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
            }
        } catch { if ($attempt -lt 2) { Start-Sleep -Milliseconds 30 } }
    }
    if (-not $status) {
        $status = [pscustomobject]@{
            jobId = [string]$script:WordExportJob.JobId; state = 'queued'; phase = 'queued'; message = 'Word出力を開始しています'
            percent = 0; currentStep = 0; totalSteps = [int]$script:WordExportJob.TotalSteps; outputPath = ''; outputName = ''
            outputDirectory = [string]$script:WordExportJob.OutputDirectory
            ownedWordPid = 0; ownedWordStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false; pageCount = 0
            startedAt = [string]$script:WordExportJob.StartedAt; updatedAt = ''; completedAt = ''; errorCode = ''
        }
    }
    $process = Get-Process -Id ([int]$script:WordExportJob.ProcessId) -ErrorAction SilentlyContinue
    $active = [string]$status.state -in @('queued', 'running', 'finalizing')
    if ($active -and $process -and -not $script:WordExportCancelRequestedAt) {
        if (((Get-Date) - [DateTime]$script:WordExportJob.StartedAt).TotalSeconds -gt $WordExportTimeoutSec) {
            [IO.File]::WriteAllText($script:WordExportJob.CancelPath, 'timeout', (New-Object Text.UTF8Encoding($false)))
            $script:WordExportCancelRequestedAt = Get-Date; $script:WordExportCancelReason = 'timeout'
            $status.message = '規定時間を超えたため、安全に中止しています'
        }
    }
    if ($active -and $process -and $script:WordExportCancelRequestedAt -and ((Get-Date) - $script:WordExportCancelRequestedAt).TotalSeconds -gt 30) {
        Stop-MbWordExportWorkerSafely -Status $status
        $status.state = if ($script:WordExportCancelReason -eq 'timeout') { 'failed' } else { 'cancelled' }
        $status.phase = [string]$status.state
        $status.message = if ($script:WordExportCancelReason -eq 'timeout') { 'Wordが規定時間内に応答しなかったため、出力処理を停止しました' } else { 'Word作成を中止しました' }
        $status.errorCode = if ($script:WordExportCancelReason -eq 'timeout') { 'EXPORT_TIMEOUT' } else { 'CANCELLED' }
        $status.completedAt = [DateTime]::UtcNow.ToString('o')
        try { Write-MbExportServerStatus -Path $script:WordExportJob.StatusPath -Status $status } catch { }
        $process = $null; $active = $false
    }
    if ($active -and -not $process) {
        $status.state = 'failed'; $status.phase = 'failed'; $status.message = 'Word出力プロセスが完了結果を返さず終了しました'
        $status.errorCode = 'WORKER_EXITED_WITHOUT_RESULT'; $status.completedAt = [DateTime]::UtcNow.ToString('o')
        try { Write-MbExportServerStatus -Path $script:WordExportJob.StatusPath -Status $status } catch { }
    }
    if ([string]$status.state -in @('completed', 'cancelled', 'failed')) { Remove-MbWordExportSnapshot }
    return $status
}

function Start-MbWordExportJob {
    $current = Read-MbWordExportStatus
    if ([string]$current.state -in @('queued', 'running', 'finalizing')) { return $current }
    $excelStatus = Read-MbExcelExportStatus
    if ([string]$excelStatus.state -in @('queued', 'running', 'finalizing')) { throw 'Excel作成中です。完了または中止してからWordを作成してください。' }
    $project = Get-MbProject -Path $ProjectPath
    $totalSteps = 0
    foreach ($sheet in @($project.sheets)) { $totalSteps += @($sheet.steps).Count }
    if ($totalSteps -lt 1) { throw 'Wordへ出力する手順を1件以上追加してください。' }
    [void](Save-MbProject -Project $project -Path $ProjectPath)
    $jobId = 'word-export-' + [guid]::NewGuid().ToString('N')
    $jobDirectory = Join-Path $script:WordExportJobsRoot $jobId
    $snapshotPath = Join-Path $jobDirectory 'project.json'
    $snapshotImageDirectory = Join-Path $jobDirectory 'images'
    $statusPath = Join-Path $jobDirectory 'status.json'
    $cancelPath = Join-Path $jobDirectory 'cancel.requested'
    $outputDirectory = Get-MbExcelOutputDirectory
    [void](New-Item -ItemType Directory -Path $snapshotImageDirectory -Force)
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    [IO.File]::Copy($ProjectPath, $snapshotPath, $true)
    $sourceImageDirectory = Join-Path (Split-Path -Parent $ProjectPath) 'images'
    foreach ($image in @($project.images)) {
        $fileName = [string]$image.fileName
        if ($fileName -notmatch '^image-[a-f0-9]{32}\.(png|jpg|bmp)$') { throw "画像ファイル名が不正です: $fileName" }
        $sourcePath = Join-Path $sourceImageDirectory $fileName
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "画像ファイルが見つかりません: $fileName" }
        [IO.File]::Copy($sourcePath, (Join-Path $snapshotImageDirectory $fileName), $true)
    }
    $queuedStatus = [pscustomobject]@{
        jobId = $jobId; state = 'queued'; phase = 'queued'; message = 'Word出力を開始しています'; percent = 0
        currentStep = 0; totalSteps = $totalSteps; outputPath = ''; outputName = ''; outputDirectory = $outputDirectory
        ownedWordPid = 0; ownedWordStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false; pageCount = 0
        startedAt = [DateTime]::UtcNow.ToString('o'); updatedAt = [DateTime]::UtcNow.ToString('o'); completedAt = ''; errorCode = ''
    }
    Write-MbExportServerStatus -Path $statusPath -Status $queuedStatus
    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { throw 'Windows PowerShell 5.1が見つかりません。' }
    $quoted = { param([string]$Value) '"' + $Value.Replace('"', '\"') + '"' }
    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (& $quoted $script:WordExportWorkerPath),
        '-ProjectPath', (& $quoted $snapshotPath), '-OutputDirectory', (& $quoted $outputDirectory),
        '-StatusPath', (& $quoted $statusPath), '-CancelPath', (& $quoted $cancelPath), '-JobId', (& $quoted $jobId)
    )
    $worker = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $workerId = [int]$worker.Id; $worker.Dispose()
    $script:WordExportJob = [pscustomobject]@{
        JobId = $jobId; ProcessId = $workerId; StatusPath = $statusPath; CancelPath = $cancelPath
        JobDirectory = $jobDirectory; OutputDirectory = $outputDirectory; StartedAt = Get-Date; TotalSteps = $totalSteps; CleanupDone = $false
    }
    $script:WordExportCancelRequestedAt = $null; $script:WordExportCancelReason = ''
    Write-MbLog "Word出力を開始しました: $jobId" 'OK'
    return Read-MbWordExportStatus
}

function Request-MbWordExportCancel {
    $status = Read-MbWordExportStatus
    if ([string]$status.state -in @('queued', 'running')) {
        [IO.File]::WriteAllText($script:WordExportJob.CancelPath, 'cancel', (New-Object Text.UTF8Encoding($false)))
        $script:WordExportCancelRequestedAt = Get-Date; $script:WordExportCancelReason = 'user'
        $status.message = 'Word作成を安全に中止しています'
    }
    return $status
}

function Open-MbWordExportResult {
    param([ValidateSet('file', 'folder')][string]$Mode)
    $status = Read-MbWordExportStatus
    if ([string]$status.state -ne 'completed') { throw '完成したWordファイルがありません。' }
    $root = [IO.Path]::GetFullPath([string]$status.outputDirectory).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $path = [IO.Path]::GetFullPath([string]$status.outputPath)
    if (-not $path.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw '出力ファイルの場所を確認できません。' }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw '出力したWordファイルが見つかりません。' }
    if ($Mode -eq 'file') { Start-Process -FilePath $path } else { Start-Process -FilePath $root }
    return $status
}

function Write-MbResponse {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerContext]$Context,
        [AllowEmptyString()][string]$Body,
        [int]$StatusCode = 200,
        [string]$ContentType = 'text/html; charset=utf-8'
    )

    $bytes = [Text.Encoding]::UTF8.GetBytes($Body)
    $response = $Context.Response
    $response.StatusCode = $StatusCode
    $response.ContentType = $ContentType
    $response.Headers.Add('Cache-Control', 'no-store')
    $response.Headers.Add('X-Content-Type-Options', 'nosniff')
    $response.Headers.Add('Referrer-Policy', 'no-referrer')
    $response.Headers.Add('X-Frame-Options', 'DENY')
    # media-srcを省くとdefault-srcへフォールバックし、動画ダイアログのblob:再生がCSPで止まる。
    # 動画はブラウザー内で作ったblob:のみを再生し、外部URLは読み込まない。
    $response.Headers.Add('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data: blob:; media-src 'self' blob:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")
    $response.ContentLength64 = $bytes.Length
    if ($bytes.Length -gt 0) {
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
    }
}

function Write-MbFile {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ContentType
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-MbResponse -Context $Context -Body 'not found' -StatusCode 404 -ContentType 'text/plain; charset=utf-8'
        return
    }
    $bytes = [IO.File]::ReadAllBytes($Path)
    $response = $Context.Response
    $response.StatusCode = 200
    $response.ContentType = $ContentType
    $response.Headers.Add('Cache-Control', 'no-cache')
    $response.Headers.Add('X-Content-Type-Options', 'nosniff')
    $response.Headers.Add('Referrer-Policy', 'no-referrer')
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Write-MbDownload {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$FileName
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-MbResponse -Context $Context -Body '書き出しファイルが見つかりません。' -StatusCode 404 -ContentType 'text/plain; charset=utf-8'
        return
    }
    $file = Get-Item -LiteralPath $Path
    $response = $Context.Response
    $response.StatusCode = 200
    $response.ContentType = 'application/zip'
    $response.Headers.Add('Cache-Control', 'no-store')
    $response.Headers.Add('X-Content-Type-Options', 'nosniff')
    $response.Headers.Add('Referrer-Policy', 'no-referrer')
    $response.Headers.Add('Content-Disposition', 'attachment; filename="ManualBuilder-project.zip"')
    $response.Headers.Add('X-Mb-Download-Name', [Uri]::EscapeDataString($FileName))
    $response.ContentLength64 = $file.Length
    $stream = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $stream.CopyTo($response.OutputStream)
    } finally {
        $stream.Dispose()
    }
}

function Read-MbForm {
    param([Parameter(Mandatory = $true)][System.Net.HttpListenerRequest]$Request)

    if ($Request.ContentLength64 -lt 0 -or $Request.ContentLength64 -gt (1024 * 1024)) {
        throw 'リクエストサイズが不正です。'
    }
    $contentType = [string]$Request.ContentType
    if (-not $contentType.StartsWith('application/x-www-form-urlencoded', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'フォーム形式のリクエストだけを受け付けます。'
    }

    $reader = New-Object IO.StreamReader($Request.InputStream, [Text.Encoding]::UTF8, $true, 1024, $true)
    try { $body = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $form = @{}
    if ($body.Length -eq 0) { return $form }
    foreach ($pair in $body.Split('&')) {
        $parts = $pair.Split('=', 2)
        $key = [Uri]::UnescapeDataString($parts[0].Replace('+', ' '))
        $value = if ($parts.Count -gt 1) { [Uri]::UnescapeDataString($parts[1].Replace('+', ' ')) } else { '' }
        $form[$key] = $value
    }
    return $form
}

function Get-MbFormValue {
    param([hashtable]$Form, [string]$Name)
    if ($Form.ContainsKey($Name)) { return [string]$Form[$Name] }
    return ''
}

function Test-MbRequestSecurity {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerRequest]$Request,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$BoundPort,
        [Parameter(Mandatory = $true)][string]$Token
    )

    $expectedHosts = @("localhost:$BoundPort", "127.0.0.1:$BoundPort")
    $hostHeader = [string]$Request.Headers['Host']
    if ($hostHeader -and ($expectedHosts -notcontains $hostHeader)) {
        throw [System.UnauthorizedAccessException]::new('Hostヘッダーが不正です。')
    }

    $isImage = $Path.StartsWith('/images/', [StringComparison]::OrdinalIgnoreCase)
    $isProtected = $Path.StartsWith('/api/', [StringComparison]::OrdinalIgnoreCase) -or
        $Path.StartsWith('/ui/', [StringComparison]::OrdinalIgnoreCase) -or $isImage
    if ($isProtected -and $Path -ne '/api/health') {
        $providedToken = [string]$Request.Headers['X-Manual-Token']
        if ($isImage -and -not $providedToken) { $providedToken = [string]$Request.QueryString['token'] }
        if ($providedToken -ne $Token) {
            throw [System.UnauthorizedAccessException]::new('セッショントークンが一致しません。')
        }
    }

    if ($Request.HttpMethod -ne 'GET' -and $Request.HttpMethod -ne 'HEAD') {
        $origin = [string]$Request.Headers['Origin']
        if ($origin -and $origin -notin @("http://localhost:$BoundPort", "http://127.0.0.1:$BoundPort")) {
            throw [System.UnauthorizedAccessException]::new('Originが不正です。')
        }
    }
}

function ConvertTo-MbCurrentWorkspaceHtml {
    param([Parameter(Mandatory = $true)][object]$Project, [AllowEmptyString()][string]$TabId)
    return ConvertTo-MbWorkspaceHtml -Project $Project -Token $script:Token -CaptureState $script:WatcherState `
        -CaptureRole (Get-MbCaptureRole -TabId $TabId) -CaptureDirectory ([string]$script:WatchDirectory) `
        -CaptureVersion $script:CaptureVersion -UndoImageStepIds @($script:ImageReplacementHistory.Keys)
}

function ConvertTo-MbCurrentProjectLibraryHtml {
    $settings = Get-MbWorkspaceSettings -DataRoot $DataRoot
    return ConvertTo-MbProjectLibraryHtml -Projects @(Get-MbProjectCatalog -DataRoot $DataRoot) `
        -LastOpenedProjectKey ([string]$settings.lastOpenedProjectKey)
}

function Test-MbOfficeExportActive {
    $excel = Read-MbExcelExportStatus
    $word = Read-MbWordExportStatus
    return ([string]$excel.state -in @('queued', 'running', 'finalizing')) -or
        ([string]$word.state -in @('queued', 'running', 'finalizing'))
}

function Copy-MbDetachedValue {
    param([Parameter(Mandatory = $true)][object]$Value)
    return ($Value | ConvertTo-Json -Depth 100 | ConvertFrom-Json)
}

function Remove-MbDeletionUndoFiles {
    param([AllowNull()][object]$Entry)
    if (-not $Entry) { return }
    foreach ($path in @($Entry.removedPaths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)) {
        if (Test-Path -LiteralPath ([string]$path) -PathType Leaf) {
            Remove-Item -LiteralPath ([string]$path) -Force -ErrorAction SilentlyContinue
        }
    }
}

function Clear-MbDeletionUndo {
    param([switch]$DeleteFiles)
    $previous = $script:DeletionUndo
    $script:DeletionUndo = $null
    if ($DeleteFiles) { Remove-MbDeletionUndoFiles -Entry $previous }
}

function Set-MbDeletionUndo {
    param([Parameter(Mandatory = $true)][object]$Entry)
    $previous = $script:DeletionUndo
    $script:DeletionUndo = $Entry
    Remove-MbDeletionUndoFiles -Entry $previous
}

function New-MbDeletionUndoEntry {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][ValidateSet('step', 'steps', 'sheet')][string]$Kind,
        [Parameter(Mandatory = $true)][string]$Label,
        [string[]]$StepIds = @(),
        [AllowEmptyString()][string]$SheetId = ''
    )

    $items = New-Object System.Collections.ArrayList
    $targetSteps = New-Object System.Collections.ArrayList
    $sheetSnapshot = $null
    $sheetIndex = -1
    if ($Kind -eq 'sheet') {
        for ($index = 0; $index -lt @($Project.sheets).Count; $index++) {
            if ([string]$Project.sheets[$index].id -ne $SheetId) { continue }
            $sheetIndex = $index
            $sheetSnapshot = Copy-MbDetachedValue -Value $Project.sheets[$index]
            foreach ($step in @($Project.sheets[$index].steps)) { [void]$targetSteps.Add($step) }
            break
        }
        if (-not $sheetSnapshot) { throw '対象シートが見つかりません。' }
    } else {
        foreach ($stepId in @($StepIds)) {
            $found = $false
            foreach ($sheet in @($Project.sheets)) {
                for ($index = 0; $index -lt @($sheet.steps).Count; $index++) {
                    if ([string]$sheet.steps[$index].id -ne [string]$stepId) { continue }
                    [void]$targetSteps.Add($sheet.steps[$index])
                    [void]$items.Add([pscustomobject]@{
                        sheetId = [string]$sheet.id
                        index   = $index
                        step    = Copy-MbDetachedValue -Value $sheet.steps[$index]
                    })
                    $found = $true
                    break
                }
                if ($found) { break }
            }
            if (-not $found) { throw '対象手順が見つかりません。' }
        }
    }

    $imageIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $videoIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $history = New-Object System.Collections.ArrayList
    foreach ($step in @($targetSteps)) {
        $stepId = [string]$step.id
        $resultImageId = if ($step.PSObject.Properties.Name -contains 'resultImageId') { [string]$step.resultImageId } else { '' }
        foreach ($imageId in @([string]$step.imageId, $resultImageId)) {
            if (-not [string]::IsNullOrWhiteSpace($imageId)) { [void]$imageIds.Add($imageId) }
        }
        $videoId = [string]$step.videoId
        if (-not [string]::IsNullOrWhiteSpace($videoId)) { [void]$videoIds.Add($videoId) }
        if ($script:ImageReplacementHistory.ContainsKey($stepId)) {
            $historyValue = Copy-MbDetachedValue -Value $script:ImageReplacementHistory[$stepId]
            [void]$history.Add([pscustomobject]@{ stepId = $stepId; value = $historyValue })
            $historyImageId = [string]$historyValue.imageId
            if (-not [string]::IsNullOrWhiteSpace($historyImageId)) { [void]$imageIds.Add($historyImageId) }
        }
    }

    $images = @($Project.images | Where-Object { $imageIds.Contains([string]$_.id) } | ForEach-Object { Copy-MbDetachedValue -Value $_ })
    $videos = @($Project.videos | Where-Object { $videoIds.Contains([string]$_.id) } | ForEach-Object { Copy-MbDetachedValue -Value $_ })
    return [pscustomobject]@{
        projectPath    = [string]$ProjectPath
        kind           = $Kind
        label          = $Label
        selectedSheetId = [string]$Project.selectedSheetId
        items          = @($items)
        sheet          = $sheetSnapshot
        sheetIndex     = $sheetIndex
        images         = $images
        videos         = $videos
        history        = @($history)
        removedPaths   = @()
    }
}

function Remove-MbDeletionAssets {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][object]$Entry
    )

    foreach ($item in @($Entry.history)) { [void]$script:ImageReplacementHistory.Remove([string]$item.stepId) }
    $paths = New-Object System.Collections.ArrayList
    foreach ($image in @($Entry.images)) {
        $path = Remove-MbUnusedImage -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$image.id)
        if ($path) { [void]$paths.Add([string]$path) }
    }
    foreach ($video in @($Entry.videos)) {
        $path = Remove-MbUnusedVideo -Project $Project -ProjectPath $ProjectPath -VideoId ([string]$video.id)
        if ($path) { [void]$paths.Add([string]$path) }
    }
    $Entry.removedPaths = @($paths | Select-Object -Unique)
}

function Restore-MbDeletionHistory {
    param([Parameter(Mandatory = $true)][object]$Entry)
    foreach ($item in @($Entry.history)) {
        $script:ImageReplacementHistory[[string]$item.stepId] = $item.value
    }
}

function Invoke-MbDeletionUndo {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [AllowEmptyString()][string]$TabId
    )

    $entry = $script:DeletionUndo
    if (-not $entry -or [string]$entry.projectPath -ne [string]$ProjectPath) { throw '元に戻せる削除はありません。' }
    foreach ($path in @($entry.removedPaths)) {
        if (-not (Test-Path -LiteralPath ([string]$path) -PathType Leaf)) {
            throw '削除した画像または動画が見つからないため、元に戻せません。'
        }
    }
    if ([string]$entry.kind -eq 'sheet') {
        [void](Restore-MbSheet -Project $Project -Sheet $entry.sheet -Index ([int]$entry.sheetIndex))
    } else {
        [void](Restore-MbSteps -Project $Project -Items @($entry.items))
        if (@($Project.sheets | Where-Object { [string]$_.id -eq [string]$entry.selectedSheetId }).Count -gt 0) {
            $Project.selectedSheetId = [string]$entry.selectedSheetId
        }
    }
    foreach ($image in @($entry.images)) {
        if (@($Project.images | Where-Object { [string]$_.id -eq [string]$image.id }).Count -eq 0) {
            $Project.images = @($Project.images) + @($image)
        }
    }
    foreach ($video in @($entry.videos)) {
        if (@($Project.videos | Where-Object { [string]$_.id -eq [string]$video.id }).Count -eq 0) {
            $Project.videos = @($Project.videos) + @($video)
        }
    }
    $saved = Save-MbProject -Project $Project -Path $ProjectPath
    Restore-MbDeletionHistory -Entry $entry
    $html = ConvertTo-MbCurrentWorkspaceHtml -Project $saved -TabId $TabId
    Clear-MbDeletionUndo
    return $html
}

function Reset-MbActiveProjectSession {
    Clear-MbDeletionUndo -DeleteFiles
    $script:CaptureOwnerTab = $null
    $script:CaptureOwnerLastHeartbeat = $null
    $script:CaptureOwnerSheetId = $null
    $script:WatcherState = if ($script:Watcher) { 'suspended' } else { 'disabled' }
    $script:ImportWatermark = Get-Date
    $script:PendingImages.Clear()
    $script:ImageReplacementHistory.Clear()
    $script:ExcelExportJob = $null
    $script:ExcelExportCancelRequestedAt = $null
    $script:ExcelExportCancelReason = ''
    $script:WordExportJob = $null
    $script:WordExportCancelRequestedAt = $null
    $script:WordExportCancelReason = ''
    $script:CaptureVersion++
}

function Open-MbCatalogProject {
    param([Parameter(Mandatory = $true)][string]$ProjectKey)
    if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中はマニュアルを切り替えられません。' }
    if (Test-MbOfficeExportActive) { throw 'Office出力中はマニュアルを切り替えられません。' }
    $path = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'マニュアルが見つかりません。' }
    [void](Get-MbProject -Path $path)
    $script:ProjectPath = $path
    $script:ActiveProjectKey = $ProjectKey
    $script:ProjectHomeVisible = $false
    Set-MbLastOpenedProject -DataRoot $DataRoot -ProjectKey $ProjectKey
    Reset-MbActiveProjectSession
    return Get-MbProject -Path $script:ProjectPath
}

function Save-MbAndRenderWorkspace {
    param([object]$Project, [AllowEmptyString()][string]$TabId)
    $saved = Save-MbProject -Project $Project -Path $ProjectPath
    return ConvertTo-MbCurrentWorkspaceHtml -Project $saved -TabId $TabId
}

function Invoke-MbRoute {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory = $true)][int]$BoundPort,
        [Parameter(Mandatory = $true)][string]$Token
    )

    $request = $Context.Request
    $path = [Uri]::UnescapeDataString($request.Url.LocalPath)
    $tabId = [string]$request.Headers['X-Tab-Id']
    Test-MbRequestSecurity -Request $request -Path $path -BoundPort $BoundPort -Token $Token

    if ($request.HttpMethod -eq 'GET') {
        if ($path -match '^/images/(?<id>image-[a-f0-9]{32})$') {
            $project = Get-MbProject -Path $ProjectPath
            $imageId = [string]$Matches['id']
            $image = @($project.images | Where-Object { $_.id -eq $imageId }) | Select-Object -First 1
            $imagePath = Get-MbImageFilePath -Project $project -ProjectPath $ProjectPath -ImageId $imageId
            if (-not $image -or -not $imagePath -or -not (Test-Path -LiteralPath $imagePath -PathType Leaf)) {
                Write-MbResponse $Context 'not found' 404 'text/plain; charset=utf-8'
                return
            }
            Write-MbFile -Context $Context -Path $imagePath -ContentType ([string]$image.mimeType)
            return
        }
        # 記録した操作の確認用サムネイル。imgタグはヘッダーを送れないため、
        # クエリのトークンが使える /images/ 配下に置く。
        if ($path -match '^/images/recording/(?<name>event-\d{3}(?:-result)?\.jpg)$') {
            $recordedPath = Get-MbRecordedEventImagePath -FileName ([string]$Matches['name'])
            if ([string]::IsNullOrWhiteSpace($recordedPath)) {
                Write-MbResponse $Context 'not found' 404 'text/plain; charset=utf-8'
                return
            }
            Write-MbFile -Context $Context -Path $recordedPath -ContentType 'image/jpeg'
            return
        }
        if ($path -match '^/images/recording/(?<name>frame-\d{5}\.jpg)$') {
            $recordedPath = Get-MbRecordedFrameImagePath -FileName ([string]$Matches['name'])
            if ([string]::IsNullOrWhiteSpace($recordedPath)) {
                Write-MbResponse $Context 'not found' 404 'text/plain; charset=utf-8'
                return
            }
            Write-MbFile -Context $Context -Path $recordedPath -ContentType 'image/jpeg'
            return
        }
        switch ($path) {
            '/' {
                $templatePath = Join-Path $webRoot 'index.html'
                $html = [IO.File]::ReadAllText($templatePath, [Text.Encoding]::UTF8).Replace('__TOKEN__', $Token)
                Write-MbResponse -Context $Context -Body $html
                return
            }
            '/assets/css/app.css' { Write-MbFile $Context (Join-Path $webRoot 'assets\css\app.css') 'text/css; charset=utf-8'; return }
            '/assets/js/app.js' { Write-MbFile $Context (Join-Path $webRoot 'assets\js\app.js') 'application/javascript; charset=utf-8'; return }
            '/assets/js/video-scenes.js' { Write-MbFile $Context (Join-Path $webRoot 'assets\js\video-scenes.js') 'application/javascript; charset=utf-8'; return }
            '/assets/js/heartbeat-worker.js' { Write-MbFile $Context (Join-Path $webRoot 'assets\js\heartbeat-worker.js') 'application/javascript; charset=utf-8'; return }
            '/vendor/htmx-2.0.10.min.js' { Write-MbFile $Context (Join-Path $webRoot 'vendor\htmx-2.0.10.min.js') 'application/javascript; charset=utf-8'; return }
            '/api/health' { Write-MbResponse $Context '{"status":"ok"}' 200 'application/json; charset=utf-8'; return }
            '/api/deletions/status' {
                $available = $null -ne $script:DeletionUndo -and [string]$script:DeletionUndo.projectPath -eq [string]$ProjectPath
                $status = [pscustomobject]@{
                    available = $available
                    label     = if ($available) { [string]$script:DeletionUndo.label } else { '' }
                }
                Write-MbResponse $Context ($status | ConvertTo-Json -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/ui/workspace' {
                if ($script:ProjectHomeVisible) {
                    Write-MbResponse -Context $Context -Body (ConvertTo-MbCurrentProjectLibraryHtml)
                } else {
                    $project = Get-MbProject -Path $ProjectPath
                    Write-MbResponse -Context $Context -Body (ConvertTo-MbCurrentWorkspaceHtml -Project $project -TabId $tabId)
                }
                return
            }
            '/api/capture/poll' {
                $clientVersion = -1
                if ($request.QueryString['version']) {
                    [void][int]::TryParse([string]$request.QueryString['version'], [ref]$clientVersion)
                }
                if ($clientVersion -eq $script:CaptureVersion) {
                    $Context.Response.StatusCode = 204
                    return
                }
                $project = Get-MbProject -Path $ProjectPath
                $sheetId = [string]$request.QueryString['sheetId']
                Write-MbResponse -Context $Context -Body (ConvertTo-MbCaptureSnapshotHtml -Project $project -SheetId $sheetId -Token $Token -Version $script:CaptureVersion -UndoImageStepIds @($script:ImageReplacementHistory.Keys))
                return
            }
            '/api/export/excel/status' {
                $status = Read-MbExcelExportStatus
                Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/export/word/status' {
                $status = Read-MbWordExportStatus
                Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/copilot/draft/status' {
                $status = Read-MbCopilotDraftStatus
                Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/copilot/draft/result' {
                $result = Get-MbCopilotDraftResult
                Write-MbResponse $Context ($result | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/recorder/status' {
                $status = Read-MbRecordingStatus
                Write-MbResponse $Context ($status | ConvertTo-Json -Depth 6 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/recorder/events' {
                # 画像そのものは別の口から出す。ここでは一覧だけを返す。
                $events = @(Get-MbRecordedEvents)
                $localProposals = @(Get-MbRecordedLocalProposals)
                Write-MbResponse $Context (([pscustomobject]@{ events = $events; localProposals = $localProposals } | ConvertTo-Json -Depth 10 -Compress)) 200 'application/json; charset=utf-8'
                return
            }
            '/api/recorder/analyze/status' {
                Write-MbResponse $Context ((Read-MbRecorderCopilotStatus) | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/recorder/analyze/result' {
                Write-MbResponse $Context ((Get-MbRecorderCopilotResult) | ConvertTo-Json -Depth 10 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/recorder/capabilities' {
                $capability = Get-MbRecordingCapability
                Write-MbResponse $Context ($capability | ConvertTo-Json -Depth 4 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            '/api/copilot/capabilities' {
                # 画面の文字認識と音声の文字起こしが使えるかを先に伝え、
                # 使えない機能のボタンを押させないようにする。
                $capabilities = Get-MbCopilotCapabilities
                Write-MbResponse $Context ($capabilities | ConvertTo-Json -Depth 6 -Compress) 200 'application/json; charset=utf-8'
                return
            }
            default { Write-MbResponse $Context 'not found' 404 'text/plain; charset=utf-8'; return }
        }
    }

    if ($request.HttpMethod -ne 'POST') {
        Write-MbResponse $Context 'method not allowed' 405 'text/plain; charset=utf-8'
        return
    }

    if ($path -eq '/api/shutdown') {
        Write-MbResponse $Context 'ManualBuilderを終了しました。' 200 'text/plain; charset=utf-8'
        $script:Running = $false
        return
    }

    if ($path -eq '/api/capture/heartbeat') {
        if ($script:ProjectHomeVisible) {
            Write-MbResponse $Context (ConvertTo-MbWatchStatusHtml -State $script:WatcherState -Role 'available' -Directory ([string]$script:WatchDirectory))
            return
        }
        if (-not $tabId) {
            Write-MbResponse $Context 'tab id required' 400 'text/plain; charset=utf-8'
            return
        }
        $role = Set-MbCaptureHeartbeat -TabId $tabId -SheetId ([string]$request.Headers['X-Sheet-Id'])
        Write-MbResponse $Context (ConvertTo-MbWatchStatusHtml -State $script:WatcherState -Role $role -Directory ([string]$script:WatchDirectory))
        return
    }

    if ($path -eq '/api/images/import') {
        if (-not $tabId) {
            Write-MbResponse $Context 'tab id required' 400 'text/plain; charset=utf-8'
            return
        }
        $role = Set-MbCaptureHeartbeat -TabId $tabId -SheetId ([string]$request.Headers['X-Sheet-Id'])
        if ($role -ne 'owner') {
            Write-MbResponse $Context 'このタブは閲覧専用です。撮影対象のタブで追加してください。' 409 'text/plain; charset=utf-8'
            return
        }
        $length = [long]$request.ContentLength64
        if ($length -lt 1) { Write-MbResponse $Context '画像データが空です。' 400 'text/plain; charset=utf-8'; return }
        if ($length -gt (20 * 1024 * 1024)) { Write-MbResponse $Context '画像は20MB以下にしてください。' 400 'text/plain; charset=utf-8'; return }
        $memory = New-Object IO.MemoryStream
        try {
            $request.InputStream.CopyTo($memory)
            $bytes = $memory.ToArray()
        } finally {
            $memory.Dispose()
        }
        $source = [string]$request.Headers['X-Image-Source']
        if ($source -notin @('paste', 'drop', 'file', 'video')) { $source = 'file' }
        $sheetId = [string]$request.Headers['X-Sheet-Id']
        $project = Get-MbProject -Path $ProjectPath
        $result = Add-MbImageStep -Project $project -ProjectPath $ProjectPath -SheetId $sheetId -Bytes $bytes -Source $source
        if ($result.Status -eq 'added') {
            $project = Save-MbProject -Project $project -Path $ProjectPath
            $script:CaptureVersion++
            Write-MbLog "画像を追加しました: $source / $($result.Image.width)x$($result.Image.height)px" 'OK'
        }
        Write-MbResponse $Context (ConvertTo-MbCaptureSnapshotHtml -Project $project -SheetId $sheetId -Token $Token -Version $script:CaptureVersion -Status $result.Status -UndoImageStepIds @($script:ImageReplacementHistory.Keys))
        return
    }

    # 録画から切り出した1コマを、操作位置の赤枠つきで手順にする。
    if ($path -eq '/api/videos/scenes/import') {
        if (-not $tabId) {
            Write-MbResponse $Context 'tab id required' 400 'text/plain; charset=utf-8'
            return
        }
        $role = Set-MbCaptureHeartbeat -TabId $tabId -SheetId ([string]$request.Headers['X-Sheet-Id'])
        if ($role -ne 'owner') {
            Write-MbResponse $Context 'このタブは閲覧専用です。撮影対象のタブで取り込んでください。' 409 'text/plain; charset=utf-8'
            return
        }
        $length = [long]$request.ContentLength64
        if ($length -lt 1) { Write-MbResponse $Context '画像データが空です。' 400 'text/plain; charset=utf-8'; return }
        if ($length -gt (20 * 1024 * 1024)) { Write-MbResponse $Context '画像は20MB以下にしてください。' 400 'text/plain; charset=utf-8'; return }
        $memory = New-Object IO.MemoryStream
        try {
            $request.InputStream.CopyTo($memory)
            $bytes = $memory.ToArray()
        } finally {
            $memory.Dispose()
        }
        $sheetId = [string]$request.Headers['X-Sheet-Id']
        $timeMs = 0
        try { $timeMs = [int][string]$request.Headers['X-Scene-Time-Ms'] } catch { $timeMs = 0 }
        $rectJson = [string]$request.Headers['X-Scene-Rect']
        $candidatesJson = [string]$request.Headers['X-Scene-Candidates']
        try {
            $project = Get-MbProject -Path $ProjectPath
            $imported = Import-MbVideoScene -Project $project -ProjectPath $ProjectPath -SheetId $sheetId `
                -Bytes $bytes -TimeMs $timeMs -RectJson $rectJson -CandidatesJson $candidatesJson
            if ($imported.status -eq 'added') {
                $project = Save-MbProject -Project $project -Path $ProjectPath
                $script:CaptureVersion++
                Write-MbLog "録画の場面を手順にしました: $([string]$imported.clickLabel)" 'OK'
            }
        } catch {
            Write-MbResponse $Context ('場面を取り込めませんでした: ' + $_.Exception.Message) 400 'text/plain; charset=utf-8'
            return
        }
        Write-MbResponse $Context (ConvertTo-MbCaptureSnapshotHtml -Project $project -SheetId $sheetId -Token $Token -Version $script:CaptureVersion -Status ([string]$imported.status) -UndoImageStepIds @($script:ImageReplacementHistory.Keys))
        return
    }

    if ($path -eq '/api/recorder/start') {
        try {
            $form = Read-MbForm -Request $request
            # 音声はマイクを入れ、Microsoftのオンライン音声認識へ送る。既定では行わない。
            $withNarration = ([string](Get-MbFormValue -Form $form -Name 'withNarration')) -match '^(?i:true|1|on|yes)$'
            # 普段使っているEdgeを含む、現在のデスクトップだけを記録する。
            # クライアントから旧modeが送られても専用プロファイルは起動しない。
            $status = Start-MbRecordingJob -WithNarration:$withNarration
            Write-MbLog '操作の記録を開始しました。' 'OK'
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 6 -Compress) 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context (([pscustomobject]@{ message = [string]$_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/recorder/stop') {
        $status = Stop-MbRecordingJob
        Write-MbResponse $Context ($status | ConvertTo-Json -Depth 6 -Compress) 200 'application/json; charset=utf-8'
        return
    }

    if ($path -eq '/api/recorder/pause') {
        $form = Read-MbForm -Request $request
        $paused = ([string](Get-MbFormValue -Form $form -Name 'paused')) -match '^(?i:true|1|on|yes)$'
        $status = Set-MbRecordingPaused -Paused:$paused
        Write-MbResponse $Context ($status | ConvertTo-Json -Depth 6 -Compress) 200 'application/json; charset=utf-8'
        return
    }

    if ($path -eq '/api/recorder/undo') {
        $status = Undo-MbLastRecordingEvent
        Write-MbResponse $Context ($status | ConvertTo-Json -Depth 6 -Compress) 200 'application/json; charset=utf-8'
        return
    }

    if ($path -eq '/api/recorder/analyze/start') {
        try {
            $source = Get-MbRecordingSourceInfo
            if ($null -eq $source) { throw '記録した画面が見つかりません。' }
            $status = Start-MbRecorderCopilotJob -SourceInfo $source
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context (([pscustomobject]@{ message = [string]$_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/recorder/analyze/cancel') {
        Write-MbResponse $Context ((Request-MbRecorderCopilotCancel) | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
        return
    }

    # 記録した操作のうち、選ばれたものだけを手順にする。
    if ($path -eq '/api/recorder/import') {
        if (-not $tabId) {
            Write-MbResponse $Context 'tab id required' 400 'text/plain; charset=utf-8'
            return
        }
        $length = [long]$request.ContentLength64
        if ($length -gt (1024 * 1024)) { Write-MbResponse $Context '取り込む操作が多すぎます。' 400 'text/plain; charset=utf-8'; return }
        $selectionJson = ''
        if ($length -gt 0) {
            $reader = New-Object IO.StreamReader($request.InputStream, [Text.Encoding]::UTF8, $true, 4096, $true)
            try { $selectionJson = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        try {
            $project = Get-MbProject -Path $ProjectPath
            $sheetId = [string]$request.Headers['X-Sheet-Id']
            if ([string]::IsNullOrWhiteSpace($sheetId)) { $sheetId = [string]$project.selectedSheetId }
            $imported = if ($selectionJson -match '"beforeFrame"') {
                Import-MbRecordedCopilotSelections -Project $project -ProjectPath $ProjectPath -SheetId $sheetId -SelectionJson $selectionJson
            } else {
                Import-MbRecordedEvents -Project $project -ProjectPath $ProjectPath -SheetId $sheetId -SelectionJson $selectionJson
            }
            if ([int]$imported.added -gt 0) {
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                $script:CaptureVersion++
            }
            Remove-MbRecorderCopilotJob
            Remove-MbRecordingJob
            Write-MbLog "記録した操作を $([int]$imported.added) 件の手順にしました。" 'OK'
            Write-MbResponse $Context ($imported | ConvertTo-Json -Depth 4 -Compress) 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context ('記録した操作を取り込めませんでした: ' + $_.Exception.Message) 400 'text/plain; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/recorder/discard') {
        Remove-MbRecorderCopilotJob
        Remove-MbRecordingJob
        Write-MbResponse $Context '{"status":"ok"}' 200 'application/json; charset=utf-8'
        return
    }

    if ($path -eq '/api/copilot/draft/start') {
        try {
            $form = Read-MbForm -Request $request
            $includeWritten = ([string](Get-MbFormValue -Form $form -Name 'includeWritten')) -match '^(?i:true|1|on|yes)$'
            $mode = [string](Get-MbFormValue -Form $form -Name 'mode')
            if ($mode -notin @('draft', 'review')) { $mode = 'draft' }
            $status = Start-MbCopilotDraftJob -ProjectPath $ProjectPath -IncludeWritten:$includeWritten -Mode $mode
            Write-MbLog 'Copilotへ手順の下書きを依頼しました。' 'OK'
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context (([pscustomobject]@{ message = [string]$_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/copilot/draft/cancel') {
        $status = Request-MbCopilotDraftCancel
        Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
        return
    }

    # 確認画面で採用された下書きだけを書き込む。却下したものは残さない。
    # 日本語をURLエンコードすると1文字9バイトになり、フォームの上限（1MiB）に届きうる。
    # そのためここだけはJSONの本文をそのまま受け取る。
    if ($path -eq '/api/copilot/draft/apply') {
        $length = [long]$request.ContentLength64
        if ($length -lt 1) { Write-MbResponse $Context '採用する手順がありません。' 400 'text/plain; charset=utf-8'; return }
        if ($length -gt (8 * 1024 * 1024)) { Write-MbResponse $Context '採用する手順が多すぎます。' 400 'text/plain; charset=utf-8'; return }
        $selectionJson = ''
        $reader = New-Object IO.StreamReader($request.InputStream, [Text.Encoding]::UTF8, $true, 4096, $true)
        try { $selectionJson = $reader.ReadToEnd() } finally { $reader.Dispose() }
        try {
            $project = Get-MbProject -Path $ProjectPath
            $applied = Set-MbCopilotDraftSelection -Project $project -SelectionJson $selectionJson
            # 採用0件でも、不要候補・自信なし候補の「要確認」をproject.jsonへ残す。
            [void](Save-MbProject -Project $project -Path $ProjectPath)
            Remove-MbCopilotDraftJob
            Write-MbLog "Copilotの下書きを $applied 件採用しました。" 'OK'
            # 画面の作り直しは /ui/workspace に任せる。ここでHTMLを返すと、
            # 呼び出し側が編集画面の初期化を通らず、操作が効かなくなる。
            Write-MbResponse $Context (([pscustomobject]@{ applied = $applied } | ConvertTo-Json -Compress)) 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context ('下書きを反映できませんでした: ' + $_.Exception.Message) 400 'text/plain; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/copilot/draft/discard') {
        Remove-MbCopilotDraftJob
        Write-MbResponse $Context '{"status":"ok"}' 200 'application/json; charset=utf-8'
        return
    }

    # サインインや様子の確認のためにCopilotの画面を前面に出す。
    if ($path -eq '/api/copilot/window') {
        try {
            [void](Show-MbCopilotSignInWindow)
            Write-MbResponse $Context '{"status":"ok"}' 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context (([pscustomobject]@{ message = [string]$_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/images/replace') {
        if (-not $tabId) {
            Write-MbResponse $Context 'tab id required' 400 'text/plain; charset=utf-8'
            return
        }
        $role = Set-MbCaptureHeartbeat -TabId $tabId -SheetId ([string]$request.Headers['X-Sheet-Id'])
        if ($role -ne 'owner') {
            Write-MbResponse $Context 'このタブは閲覧専用です。撮影対象のタブで差し替えてください。' 409 'text/plain; charset=utf-8'
            return
        }
        $length = [long]$request.ContentLength64
        if ($length -lt 1) { Write-MbResponse $Context '画像データが空です。' 400 'text/plain; charset=utf-8'; return }
        if ($length -gt (20 * 1024 * 1024)) { Write-MbResponse $Context '画像は20MB以下にしてください。' 400 'text/plain; charset=utf-8'; return }
        $memory = New-Object IO.MemoryStream
        try {
            $request.InputStream.CopyTo($memory)
            $bytes = $memory.ToArray()
        } finally {
            $memory.Dispose()
        }
        try {
            $source = [string]$request.Headers['X-Image-Source']
            if ($source -notin @('paste', 'drop', 'file', 'video')) { $source = 'file' }
            $stepId = [string]$request.Headers['X-Step-Id']
            if ([string]::IsNullOrWhiteSpace($stepId)) { throw '対象手順が指定されていません。' }
            $project = Get-MbProject -Path $ProjectPath
            $result = Set-MbStepImage -Project $project -ProjectPath $ProjectPath -StepId $stepId -Bytes $bytes -Source $source
            if ($result.Status -eq 'duplicate') {
                $body = [pscustomobject]@{ state = 'duplicate'; message = '同じ画像が設定されています。'; stepId = $stepId } | ConvertTo-Json -Compress
                Write-MbResponse $Context $body 200 'application/json; charset=utf-8'
                return
            }

            $discardedImagePath = $null
            if ($script:ImageReplacementHistory.ContainsKey($stepId)) {
                $discarded = $script:ImageReplacementHistory[$stepId]
                [void]$script:ImageReplacementHistory.Remove($stepId)
                $discardedImagePath = Remove-MbUnusedImage -Project $project -ProjectPath $ProjectPath -ImageId ([string]$discarded.imageId)
            }
            $canUndoReplacement = -not [string]::IsNullOrWhiteSpace([string]$result.Previous.imageId)
            if ($canUndoReplacement) {
                $script:ImageReplacementHistory[$stepId] = $result.Previous
            }
            $project = Save-MbProject -Project $project -Path $ProjectPath
            if ($discardedImagePath -and (Test-Path -LiteralPath $discardedImagePath -PathType Leaf)) {
                Remove-Item -LiteralPath $discardedImagePath -Force -ErrorAction SilentlyContinue
            }
            $script:CaptureVersion++
            $body = [pscustomobject]@{
                state       = 'replaced'
                message     = if ($canUndoReplacement) { '画像を差し替えました。元の画像へ戻すこともできます。' } else { '画像を追加しました。' }
                stepId      = $stepId
                imageId     = [string]$result.Image.id
                imageUrl    = '/images/' + [string]$result.Image.id + '?token=' + $Token
                width       = [int]$result.Image.width
                height      = [int]$result.Image.height
                annotations = @()
                crop        = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
                canUndo     = $canUndoReplacement
            } | ConvertTo-Json -Depth 6 -Compress
            Write-MbLog "画像を差し替えました: $stepId / $($result.Image.width)x$($result.Image.height)px" 'OK'
            Write-MbResponse $Context $body 200 'application/json; charset=utf-8'
        } catch {
            $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/images/result') {
        if (-not $tabId) {
            Write-MbResponse $Context 'tab id required' 400 'text/plain; charset=utf-8'
            return
        }
        $role = Set-MbCaptureHeartbeat -TabId $tabId -SheetId ([string]$request.Headers['X-Sheet-Id'])
        if ($role -ne 'owner') {
            Write-MbResponse $Context 'このタブは閲覧専用です。撮影対象のタブで画像を追加してください。' 409 'text/plain; charset=utf-8'
            return
        }
        $length = [long]$request.ContentLength64
        if ($length -lt 1) { Write-MbResponse $Context '画像データが空です。' 400 'text/plain; charset=utf-8'; return }
        if ($length -gt (20 * 1024 * 1024)) { Write-MbResponse $Context '画像は20MB以下にしてください。' 400 'text/plain; charset=utf-8'; return }
        $memory = New-Object IO.MemoryStream
        try {
            $request.InputStream.CopyTo($memory)
            $bytes = $memory.ToArray()
        } finally { $memory.Dispose() }
        try {
            $source = [string]$request.Headers['X-Image-Source']
            if ($source -notin @('paste', 'drop', 'file', 'recorder')) { $source = 'file' }
            $stepId = [string]$request.Headers['X-Step-Id']
            if ([string]::IsNullOrWhiteSpace($stepId)) { throw '対象手順が指定されていません。' }
            $project = Get-MbProject -Path $ProjectPath
            $result = Set-MbStepResultImage -Project $project -ProjectPath $ProjectPath -StepId $stepId -Bytes $bytes -Source $source
            if ($result.Status -eq 'duplicate') {
                Write-MbResponse $Context (([pscustomobject]@{ state = 'duplicate'; message = '同じ操作後画像が設定されています。'; stepId = $stepId } | ConvertTo-Json -Compress)) 200 'application/json; charset=utf-8'
                return
            }
            $project = Save-MbProject -Project $project -Path $ProjectPath
            if ($result.RemovedPath -and (Test-Path -LiteralPath $result.RemovedPath -PathType Leaf)) {
                Remove-Item -LiteralPath $result.RemovedPath -Force -ErrorAction SilentlyContinue
            }
            $script:CaptureVersion++
            $body = [pscustomobject]@{
                state         = 'set'
                message       = '操作後画像を追加しました。見せ方を選べます。'
                stepId        = $stepId
                resultImageId = [string]$result.Image.id
                resultImageUrl = '/images/' + [string]$result.Image.id + '?token=' + $Token
                imageLayout   = [string]$result.Step.imageLayout
                imageOrder    = [string]$result.Step.imageOrder
            } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 200 'application/json; charset=utf-8'
        } catch {
            Write-MbResponse $Context (([pscustomobject]@{ state = 'failed'; message = $_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/videos/attach') {
        # 動画は手順へ添付するだけで、ブラウザーへは返さない。Excel出力のときだけ読む。
        $bytes = $null
        try {
            $stepId = [string]$request.Headers['X-Step-Id']
            if ([string]::IsNullOrWhiteSpace($stepId)) { throw '対象手順が指定されていません。' }
            $length = [long]$request.ContentLength64
            if ($length -lt 1) { throw '動画データが空です。' }
            if ($length -gt (30 * 1024 * 1024)) { throw '動画は30MB以下にしてください。短く撮り直すか、解像度を下げてください。' }
            $memory = New-Object IO.MemoryStream
            try {
                $request.InputStream.CopyTo($memory)
                $bytes = $memory.ToArray()
            } finally {
                $memory.Dispose()
            }
            $duration = [double]0
            [void][double]::TryParse([string]$request.Headers['X-Video-Duration'],
                [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$duration)

            $project = Get-MbProject -Path $ProjectPath
            $result = Set-MbStepVideo -Project $project -ProjectPath $ProjectPath -StepId $stepId -Bytes $bytes -DurationSec $duration
            $project = Save-MbProject -Project $project -Path $ProjectPath
            if ($result.RemovedPath -and (Test-Path -LiteralPath $result.RemovedPath -PathType Leaf)) {
                Remove-Item -LiteralPath $result.RemovedPath -Force -ErrorAction SilentlyContinue
            }
            $script:CaptureVersion++
            $totalBytes = Get-MbVideoTotalBytes -Project $project
            $body = [pscustomobject]@{
                state       = 'attached'
                message     = '動画を添付しました。Excelで作成すると再生できます。'
                stepId      = $stepId
                videoId     = [string]$result.Video.id
                byteLength  = [long]$result.Video.byteLength
                durationSec = [double]$result.Video.durationSec
                totalBytes  = [long]$totalBytes
            } | ConvertTo-Json -Compress
            Write-MbLog "動画を添付しました: $stepId / $([Math]::Round([long]$result.Video.byteLength / 1MB, 1))MB" 'OK'
            Write-MbResponse $Context $body 200 'application/json; charset=utf-8'
        } catch {
            $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 400 'application/json; charset=utf-8'
        }
        return
    }

    if ($path -eq '/api/projects/import') {
        $packagePath = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-import-' + [guid]::NewGuid().ToString('N') + '.zip')
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は取り込みできません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は取り込みできません。' }
            $contentType = [string]$request.ContentType
            if (-not ($contentType.StartsWith('application/zip', [StringComparison]::OrdinalIgnoreCase) -or
                $contentType.StartsWith('application/octet-stream', [StringComparison]::OrdinalIgnoreCase))) {
                throw 'ManualBuilderから書き出したZIPを選択してください。'
            }
            $declaredLength = [long]$request.ContentLength64
            if ($declaredLength -eq 0 -or $declaredLength -gt (250 * 1024 * 1024)) { throw '取り込めるZIPは250MBまでです。' }
            Write-MbLog "マニュアルZIPを受信しています: $declaredLength bytes" 'INFO'
            $destination = [IO.File]::Open($packagePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try {
                $buffer = New-Object byte[] 81920
                $totalLength = [long]0
                while ($totalLength -lt $declaredLength) {
                    $remainingLength = $declaredLength - $totalLength
                    $requestedLength = [int][Math]::Min([long]$buffer.Length, $remainingLength)
                    $readLength = $request.InputStream.Read($buffer, 0, $requestedLength)
                    if ($readLength -le 0) { throw 'ZIPデータを最後まで受信できませんでした。' }
                    $totalLength += [long]$readLength
                    if ($totalLength -gt (250 * 1024 * 1024)) { throw '取り込めるZIPは250MBまでです。' }
                    $destination.Write($buffer, 0, $readLength)
                }
                if ($totalLength -lt 1) { throw 'ZIPデータが空です。' }
            } finally {
                $destination.Dispose()
            }
            Write-MbLog 'マニュアルZIPを検証しています。' 'INFO'
            [void](Import-MbCatalogProjectPackage -DataRoot $DataRoot -PackagePath $packagePath)
            Write-MbLog 'マニュアルZIPを取り込みました。' 'OK'
            Write-MbResponse $Context (ConvertTo-MbCurrentProjectLibraryHtml)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
        } finally {
            if (Test-Path -LiteralPath $packagePath -PathType Leaf) {
                Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue
            }
        }
        return
    }

    $form = Read-MbForm -Request $request

    if ($path -eq '/api/projects/export') {
        $packagePath = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-export-' + [guid]::NewGuid().ToString('N') + '.zip')
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は書き出しできません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は書き出しできません。' }
            $package = Export-MbCatalogProjectPackage -DataRoot $DataRoot -ProjectKey (Get-MbFormValue $form 'projectKey') -OutputPath $packagePath
            Write-MbDownload -Context $Context -Path ([string]$package.Path) -FileName ([string]$package.FileName)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
        } finally {
            if (Test-Path -LiteralPath $packagePath -PathType Leaf) {
                Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue
            }
        }
        return
    }

    if ($path -eq '/api/projects/home') {
        if ($usesExplicitProjectPath) {
            Write-MbResponse $Context '明示プロジェクト指定中は一覧を開けません。' 409 'text/plain; charset=utf-8'
            return
        }
        if (Test-MbOfficeExportActive) {
            Write-MbResponse $Context 'Office出力中は一覧へ戻れません。' 409 'text/plain; charset=utf-8'
            return
        }
        $script:ProjectHomeVisible = $true
        Reset-MbActiveProjectSession
        Write-MbResponse $Context (ConvertTo-MbCurrentProjectLibraryHtml)
        return
    }
    if ($path -eq '/api/projects/open') {
        try {
            $project = Open-MbCatalogProject -ProjectKey (Get-MbFormValue $form 'projectKey')
            Write-MbResponse $Context (ConvertTo-MbCurrentWorkspaceHtml -Project $project -TabId $tabId)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 409 'text/plain; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/projects/create') {
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は新規作成できません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は新規作成できません。' }
            $created = New-MbCatalogProject -DataRoot $DataRoot -Title (Get-MbFormValue $form 'title')
            $project = Open-MbCatalogProject -ProjectKey ([string]$created.Key)
            Write-MbResponse $Context (ConvertTo-MbCurrentWorkspaceHtml -Project $project -TabId $tabId)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/projects/duplicate') {
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は複製できません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は複製できません。' }
            [void](Copy-MbCatalogProject -DataRoot $DataRoot -ProjectKey (Get-MbFormValue $form 'projectKey'))
            Write-MbResponse $Context (ConvertTo-MbCurrentProjectLibraryHtml)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 409 'text/plain; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/projects/archive') {
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は削除できません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は削除できません。' }
            $key = Get-MbFormValue $form 'projectKey'
            [void](Move-MbCatalogProjectToArchive -DataRoot $DataRoot -ProjectKey $key)
            if ($script:ActiveProjectKey -eq $key) {
                $script:ActiveProjectKey = ''
                $script:ProjectHomeVisible = $true
                Reset-MbActiveProjectSession
            }
            Write-MbResponse $Context (ConvertTo-MbCurrentProjectLibraryHtml)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 409 'text/plain; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/projects/delete') {
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は削除できません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は削除できません。' }
            $key = Get-MbFormValue $form 'projectKey'
            [void](Remove-MbCatalogProject -DataRoot $DataRoot -ProjectKey $key)
            if ($script:ActiveProjectKey -eq $key) {
                $script:ActiveProjectKey = ''
                $script:ProjectHomeVisible = $true
                Reset-MbActiveProjectSession
            }
            Write-MbResponse $Context (ConvertTo-MbCurrentProjectLibraryHtml)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 409 'text/plain; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/projects/restore') {
        try {
            if ($usesExplicitProjectPath) { throw '明示プロジェクト指定中は復元できません。' }
            if (Test-MbOfficeExportActive) { throw 'Office出力中は復元できません。' }
            [void](Restore-MbCatalogProject -DataRoot $DataRoot -ProjectKey (Get-MbFormValue $form 'projectKey'))
            Write-MbResponse $Context (ConvertTo-MbCurrentProjectLibraryHtml)
        } catch {
            Write-MbResponse $Context $_.Exception.Message 409 'text/plain; charset=utf-8'
        }
        return
    }

    if ($script:ProjectHomeVisible) {
        Write-MbResponse $Context '編集するマニュアルを開いてください。' 409 'text/plain; charset=utf-8'
        return
    }

    if ($path -eq '/api/export/excel/start') {
        try {
            $status = Start-MbExcelExportJob
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 202 'application/json; charset=utf-8'
        } catch {
            $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message; errorCode = 'START_REJECTED' } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 400 'application/json; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/export/excel/cancel') {
        $status = Request-MbExcelExportCancel
        Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 202 'application/json; charset=utf-8'
        return
    }
    if ($path -eq '/api/export/excel/open') {
        try {
            $mode = Get-MbFormValue $form 'mode'
            if ($mode -notin @('file', 'folder')) { throw '開く対象が不正です。' }
            $status = Open-MbExcelExportResult -Mode $mode
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
        } catch {
            $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message; errorCode = 'OPEN_FAILED' } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 400 'application/json; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/export/word/start') {
        try {
            $status = Start-MbWordExportJob
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 202 'application/json; charset=utf-8'
        } catch {
            $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message; errorCode = 'START_REJECTED' } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 400 'application/json; charset=utf-8'
        }
        return
    }
    if ($path -eq '/api/export/word/cancel') {
        $status = Request-MbWordExportCancel
        Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 202 'application/json; charset=utf-8'
        return
    }
    if ($path -eq '/api/export/word/open') {
        try {
            $mode = Get-MbFormValue $form 'mode'
            if ($mode -notin @('file', 'folder')) { throw '開く対象が不正です。' }
            $status = Open-MbWordExportResult -Mode $mode
            Write-MbResponse $Context ($status | ConvertTo-Json -Depth 8 -Compress) 200 'application/json; charset=utf-8'
        } catch {
            $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message; errorCode = 'OPEN_FAILED' } | ConvertTo-Json -Compress
            Write-MbResponse $Context $body 400 'application/json; charset=utf-8'
        }
        return
    }
    $project = Get-MbProject -Path $ProjectPath

    switch ($path) {
        '/api/deletions/undo' {
            try {
                Write-MbResponse $Context (Invoke-MbDeletionUndo -Project $project -TabId $tabId)
            } catch {
                Write-MbResponse $Context $_.Exception.Message 409 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/project/title' {
            try {
                Set-MbProjectTitle -Project $project -Title (Get-MbFormValue $form 'title')
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error)
            }
            return
        }
        '/api/sheets/add' {
            [void](Add-MbSheet -Project $project)
            Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            return
        }
        '/api/sheets/duplicate' {
            try {
                [void](Copy-MbSheet -Project $project -SheetId (Get-MbFormValue $form 'sheetId'))
                Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            } catch {
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/sheets/select' {
            Select-MbSheet -Project $project -SheetId (Get-MbFormValue $form 'sheetId')
            Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            return
        }
        '/api/sheets/reorder' {
            try {
                $orderedIds = @((Get-MbFormValue $form 'orderedIds').Split(',') | Where-Object { $_ })
                Set-MbSheetOrder -Project $project -SheetIds $orderedIds
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error) 400
            }
            return
        }
        '/api/sheets/rename' {
            try {
                Rename-MbSheet -Project $project -SheetId (Get-MbFormValue $form 'sheetId') -Name (Get-MbFormValue $form 'name')
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error)
            }
            return
        }
        '/api/sheets/delete' {
            $entry = $null
            try {
                $sheetId = Get-MbFormValue $form 'sheetId'
                $sheet = @($project.sheets | Where-Object { [string]$_.id -eq $sheetId }) | Select-Object -First 1
                $sheetName = if ($sheet) { [string]$sheet.name } else { 'シート' }
                $entry = New-MbDeletionUndoEntry -Project $project -Kind sheet -Label "「$sheetName」を削除しました" -SheetId $sheetId
                Remove-MbSheet -Project $project -SheetId $sheetId
                Remove-MbDeletionAssets -Project $project -Entry $entry
                $html = Save-MbAndRenderWorkspace -Project $project -TabId $tabId
                Set-MbDeletionUndo -Entry $entry
                Write-MbResponse $Context $html
            } catch {
                if ($entry) { Restore-MbDeletionHistory -Entry $entry }
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/steps/add' {
            try {
                [void](Add-MbStep -Project $project -SheetId (Get-MbFormValue $form 'sheetId') -AfterStepId (Get-MbFormValue $form 'afterStepId'))
                Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            } catch {
                Write-MbResponse $Context $_.Exception.Message 400
            }
            return
        }
        '/api/steps/update' {
            try {
                Update-MbStep -Project $project -StepId (Get-MbFormValue $form 'stepId') -Title (Get-MbFormValue $form 'title') -Description (Get-MbFormValue $form 'description') -Note (Get-MbFormValue $form 'note')
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error)
            }
            return
        }
        '/api/steps/review/resolve' {
            try {
                [void](Set-MbStepReview -Project $project -StepId (Get-MbFormValue $form 'stepId'))
                Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            } catch {
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/steps/reorder' {
            try {
                $orderedIds = @((Get-MbFormValue $form 'orderedIds').Split(',') | Where-Object { $_ })
                Set-MbStepOrder -Project $project -SheetId (Get-MbFormValue $form 'sheetId') -StepIds $orderedIds
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error) 400
            }
            return
        }
        '/api/steps/move' {
            try {
                $stepId = Get-MbFormValue $form 'stepId'
                [void](Move-MbStepToSheet -Project $project -StepId $stepId -TargetSheetId (Get-MbFormValue $form 'targetSheetId'))
                Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            } catch {
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/steps/move-many' {
            try {
                $stepIds = @((Get-MbFormValue $form 'stepIds').Split(',') | Where-Object { $_ })
                [void](Move-MbStepsToSheet -Project $project -StepIds $stepIds -TargetSheetId (Get-MbFormValue $form 'targetSheetId'))
                Write-MbResponse $Context (Save-MbAndRenderWorkspace -Project $project -TabId $tabId)
            } catch {
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/steps/annotations' {
            try {
                Set-MbStepImageEdits -Project $project -StepId (Get-MbFormValue $form 'stepId') `
                    -AnnotationsJson (Get-MbFormValue $form 'annotations') -CropJson (Get-MbFormValue $form 'crop') `
                    -Target $(if ((Get-MbFormValue $form 'target') -eq 'result') { 'result' } else { 'before' })
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error) 400
            }
            return
        }
        '/api/steps/image-layout' {
            try {
                [void](Set-MbStepImageLayout -Project $project -StepId (Get-MbFormValue $form 'stepId') `
                    -Layout (Get-MbFormValue $form 'layout') -Order (Get-MbFormValue $form 'order'))
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                Write-MbResponse $Context '{"state":"saved"}' 200 'application/json; charset=utf-8'
            } catch {
                Write-MbResponse $Context (([pscustomobject]@{ state = 'failed'; message = $_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
            }
            return
        }
        '/api/images/result/remove' {
            try {
                $result = Remove-MbStepResultImage -Project $project -ProjectPath $ProjectPath -StepId (Get-MbFormValue $form 'stepId')
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                if ($result.RemovedPath -and (Test-Path -LiteralPath $result.RemovedPath -PathType Leaf)) {
                    Remove-Item -LiteralPath $result.RemovedPath -Force -ErrorAction SilentlyContinue
                }
                $script:CaptureVersion++
                Write-MbResponse $Context '{"state":"removed"}' 200 'application/json; charset=utf-8'
            } catch {
                Write-MbResponse $Context (([pscustomobject]@{ state = 'failed'; message = $_.Exception.Message } | ConvertTo-Json -Compress)) 400 'application/json; charset=utf-8'
            }
            return
        }
        '/api/images/replace/undo' {
            try {
                $stepId = Get-MbFormValue $form 'stepId'
                if (-not $script:ImageReplacementHistory.ContainsKey($stepId)) {
                    throw '元に戻せる画像がありません。'
                }
                $previous = $script:ImageReplacementHistory[$stepId]
                $result = Restore-MbStepImage -Project $project -StepId $stepId -Previous $previous
                [void]$script:ImageReplacementHistory.Remove($stepId)
                $unusedImagePath = Remove-MbUnusedImage -Project $project -ProjectPath $ProjectPath -ImageId ([string]$result.ReplacedImageId)
                $project = Save-MbProject -Project $project -Path $ProjectPath
                if ($unusedImagePath -and (Test-Path -LiteralPath $unusedImagePath -PathType Leaf)) {
                    Remove-Item -LiteralPath $unusedImagePath -Force -ErrorAction SilentlyContinue
                }
                $restoredImage = @($project.images | Where-Object { $_.id -eq [string]$result.Step.imageId }) | Select-Object -First 1
                $script:CaptureVersion++
                $body = [pscustomobject]@{
                    state       = 'restored'
                    message     = '元の画像と編集内容へ戻しました。'
                    stepId      = $stepId
                    imageId     = [string]$restoredImage.id
                    imageUrl    = '/images/' + [string]$restoredImage.id + '?token=' + $Token
                    width       = [int]$restoredImage.width
                    height      = [int]$restoredImage.height
                    annotations = @($result.Step.annotations)
                    crop        = $result.Step.crop
                    canUndo     = $false
                } | ConvertTo-Json -Depth 8 -Compress
                Write-MbResponse $Context $body 200 'application/json; charset=utf-8'
            } catch {
                $body = [pscustomobject]@{ state = 'failed'; message = $_.Exception.Message } | ConvertTo-Json -Compress
                Write-MbResponse $Context $body 409 'application/json; charset=utf-8'
            }
            return
        }
        '/api/videos/detach' {
            try {
                $result = Remove-MbStepVideo -Project $project -ProjectPath $ProjectPath -StepId (Get-MbFormValue $form 'stepId')
                [void](Save-MbProject -Project $project -Path $ProjectPath)
                if ($result.RemovedPath -and (Test-Path -LiteralPath $result.RemovedPath -PathType Leaf)) {
                    Remove-Item -LiteralPath $result.RemovedPath -Force -ErrorAction SilentlyContinue
                }
                $script:CaptureVersion++
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml)
            } catch {
                Write-MbResponse $Context (ConvertTo-MbSaveStatusHtml -Message $_.Exception.Message -State error) 400
            }
            return
        }
        '/api/steps/delete' {
            $entry = $null
            try {
                $stepId = Get-MbFormValue $form 'stepId'
                $entry = New-MbDeletionUndoEntry -Project $project -Kind step -Label '手順を削除しました' -StepIds @($stepId)
                Remove-MbStep -Project $project -StepId $stepId
                Remove-MbDeletionAssets -Project $project -Entry $entry
                $html = Save-MbAndRenderWorkspace -Project $project -TabId $tabId
                Set-MbDeletionUndo -Entry $entry
                Write-MbResponse $Context $html
            } catch {
                if ($entry) { Restore-MbDeletionHistory -Entry $entry }
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        '/api/steps/delete-many' {
            $entry = $null
            try {
                $stepIds = @((Get-MbFormValue $form 'stepIds').Split(',') | Where-Object { $_ })
                $entry = New-MbDeletionUndoEntry -Project $project -Kind steps -Label "$($stepIds.Count)件の手順を削除しました" -StepIds $stepIds
                Remove-MbSteps -Project $project -StepIds $stepIds
                Remove-MbDeletionAssets -Project $project -Entry $entry
                $html = Save-MbAndRenderWorkspace -Project $project -TabId $tabId
                Set-MbDeletionUndo -Entry $entry
                Write-MbResponse $Context $html
            } catch {
                if ($entry) { Restore-MbDeletionHistory -Entry $entry }
                Write-MbResponse $Context $_.Exception.Message 400 'text/plain; charset=utf-8'
            }
            return
        }
        default { Write-MbResponse $Context 'not found' 404 'text/plain; charset=utf-8'; return }
    }
}

$mutexScope = if ($AllowParallelTestInstance) { "$ProjectPath|$Port" } else { '' }
$mutexInfo = Get-MbMutex -Scope $mutexScope
if (-not $mutexInfo.Acquired) {
    if (-not (Show-MbExistingInstanceNotice)) {
        Write-Host 'ManualBuilderはすでに起動しています。ブラウザーから既存画面を開いてください。' -ForegroundColor Yellow
    }
    $mutexInfo.Mutex.Dispose()
    exit 0
}

$listener = $null
$projectReady = $false
$boundPort = $Port
$token = [guid]::NewGuid().ToString('N')
$script:Token = $token
$script:Running = $true

try {
    $storageState = Initialize-MbUserStorage -Layout $storageLayout -MigrateLegacy:$storageLayout.UsesDefaultUserData
    if ($storageState.Migrated) {
        Write-MbLog "既存プロジェクトをユーザーのローカル保存先へコピーしました: $ProjectPath" 'OK'
        Write-MbLog "移行元は削除していません: $($storageState.LegacyProjectPath)" 'INFO'
    }
    $initialProject = Get-MbProject -Path $ProjectPath
    $projectReady = $true
    $orphanedImagePaths = @(
        Remove-MbUnreferencedImages -Project $initialProject -ProjectPath $ProjectPath
        Remove-MbUnreferencedVideos -Project $initialProject -ProjectPath $ProjectPath
    ) | Where-Object { $_ }
    $orphanedImagePaths = @($orphanedImagePaths)
    if ($orphanedImagePaths.Count -gt 0) {
        [void](Save-MbProject -Project $initialProject -Path $ProjectPath)
        foreach ($orphanedImagePath in $orphanedImagePaths) {
            if (Test-Path -LiteralPath $orphanedImagePath -PathType Leaf) {
                Remove-Item -LiteralPath $orphanedImagePath -Force -ErrorAction SilentlyContinue
            }
        }
        Write-MbLog "未参照の画像を整理しました: $($orphanedImagePaths.Count)件" 'INFO'
    }
    Initialize-MbScreenshotWatcher

    $started = $false
    for ($offset = 0; $offset -lt 10; $offset++) {
        $candidate = $Port + $offset
        $candidateListener = New-Object System.Net.HttpListener
        $candidateListener.Prefixes.Add("http://localhost:$candidate/")
        try {
            $candidateListener.Start()
            $listener = $candidateListener
            $boundPort = $candidate
            $started = $true
            break
        } catch {
            $candidateListener.Close()
            Write-MbLog "ポート $candidate は使用できません。" 'WARN'
        }
    }
    if (-not $started) { throw '利用可能なlocalhostポートが見つかりません。' }

    $url = "http://localhost:$boundPort/"
    $runtimeDirectory = Split-Path -Parent $runtimePath
    if (-not (Test-Path -LiteralPath $runtimeDirectory)) { [void](New-Item -ItemType Directory -Path $runtimeDirectory -Force) }
    $runtime = [pscustomobject]@{ pid = $PID; url = $url; startedAt = [DateTime]::UtcNow.ToString('o') } | ConvertTo-Json
    [IO.File]::WriteAllText($runtimePath, $runtime, (New-Object Text.UTF8Encoding($false)))

    Write-Host ''
    Write-Host '======================================================================' -ForegroundColor Cyan
    Write-Host '  ManualBuilder Phase 1 foundation' -ForegroundColor Cyan
    Write-Host '======================================================================' -ForegroundColor Cyan
    Write-MbLog "起動しました: $url" 'OK'
    Write-MbLog "ユーザーデータ: $DataRoot" 'INFO'
    Write-MbLog "プロジェクト: $ProjectPath" 'INFO'
    Write-Host '  停止するには画面の「終了」または Ctrl+C を使用してください。' -ForegroundColor Yellow
    Write-Host ''

    # Copilotは記録後の場面選択で使うため、利用者が記録を終えてからEdgeの起動を
    # 待たなくてよいよう、サーバーの待受開始後に非同期で準備する。初期化worker側で
    # 既存の普段使いEdgeを再利用し、サインイン確認が必要な場合も画面上へ表示する。
    if (-not $SkipCopilotWarmup) {
        try {
            if (Start-MbCopilotWarmup) {
                Write-MbLog 'Microsoft 365 Copilotの画面を準備しています。' 'INFO'
            }
        } catch {
            # Copilotが使えなくても、ローカル初稿と編集・Office出力は続行できる。
            Write-MbLog ('Copilotの事前準備を開始できませんでした: ' + $_.Exception.Message) 'WARN'
        }
    }

    if (-not $NoBrowser) {
        Start-Process $url
    }

    while ($script:Running -and $listener.IsListening) {
        $context = $null
        $contextTask = $listener.GetContextAsync()
        while (-not $contextTask.AsyncWaitHandle.WaitOne(200)) {
            try { Update-MbCaptureHeartbeatState } catch { Write-MbLog $_.Exception.Message 'WARN' }
            try { Invoke-MbWatcherFlush } catch { Write-MbLog $_.Exception.Message 'WARN' }
            if (-not $script:Running) { break }
        }
        if (-not $script:Running) { break }
        if (-not $contextTask.IsCompleted) { continue }
        try {
            $context = $contextTask.GetAwaiter().GetResult()
            Invoke-MbRoute -Context $context -BoundPort $boundPort -Token $token
        } catch [System.UnauthorizedAccessException] {
            # 応答の送信途中で失敗した場合、再送信も失敗する。ここで握り潰さないとサーバー全体が停止する。
            if ($context) {
                try { Write-MbResponse $context $_.Exception.Message 403 'text/plain; charset=utf-8' } catch { }
            }
            Write-MbLog $_.Exception.Message 'WARN'
        } catch {
            if ($context) {
                try { Write-MbResponse $context '処理中にエラーが発生しました。入力内容はプロジェクトファイルを確認してください。' 500 'text/plain; charset=utf-8' } catch { }
            }
            Write-MbLog $_.Exception.Message 'ERROR'
        } finally {
            if ($context) {
                try { $context.Response.OutputStream.Close() } catch { }
                try { $context.Response.Close() } catch { }
            }
        }
    }
} catch {
    Write-MbLog $_.Exception.Message 'ERROR'
    exit 1
} finally {
    if ($script:ExcelExportJob) {
        try { [void](Request-MbExcelExportCancel) } catch { }
    }
    if ($script:WordExportJob) {
        try { [void](Request-MbWordExportCancel) } catch { }
    }
    Stop-MbScreenshotWatcher
    try {
        $script:ImageReplacementHistory.Clear()
        if ($projectReady -and (Test-Path -LiteralPath $ProjectPath -PathType Leaf)) {
            $finalProject = Get-MbProject -Path $ProjectPath
            $unusedImagePaths = @(
                Remove-MbUnreferencedImages -Project $finalProject -ProjectPath $ProjectPath
                Remove-MbUnreferencedVideos -Project $finalProject -ProjectPath $ProjectPath
            ) | Where-Object { $_ }
            $unusedImagePaths = @($unusedImagePaths)
            if ($unusedImagePaths.Count -gt 0) {
                [void](Save-MbProject -Project $finalProject -Path $ProjectPath)
                foreach ($unusedImagePath in $unusedImagePaths) {
                    if (Test-Path -LiteralPath $unusedImagePath -PathType Leaf) {
                        Remove-Item -LiteralPath $unusedImagePath -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    } catch {
        Write-MbLog "未参照画像の終了時整理に失敗しました: $($_.Exception.Message)" 'WARN'
    }
    if ($listener) {
        try { $listener.Stop() } catch { }
        try { $listener.Close() } catch { }
    }
    if (Test-Path -LiteralPath $runtimePath) {
        Remove-Item -LiteralPath $runtimePath -Force -ErrorAction SilentlyContinue
    }
    if ($mutexInfo.Acquired) {
        try { $mutexInfo.Mutex.ReleaseMutex() } catch { }
    }
    try { $mutexInfo.Mutex.Dispose() } catch { }
    Write-MbLog '停止しました。' 'INFO'
}
