# ManualBuilder shared-folder launcher and local-cache updater.

[CmdletBinding()]
param(
    # HTMLマニュアルに同梱した元データ（_source フォルダー）。「編集する.cmd」から渡される。
    [string]$ImportFrom,
    # そのHTMLマニュアルが置かれているフォルダー。次回の「共有フォルダーへ反映」先として覚える。
    [string]$PublishTo
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Launcher.psm1') -Force
$cacheRoot = Get-MbDefaultAppCacheRoot
$cachedStartScript = Join-Path $cacheRoot 'src\Start-ManualBuilder.ps1'
$localDataRoot = Join-Path (Split-Path -Parent $cacheRoot) 'data'
$runtimePath = Join-Path $localDataRoot 'runtime.json'

function Write-MbLauncherLog {
    param([string]$Message, [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $color = switch ($Level) {
        'OK' { 'Green' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Gray' }
    }
    Write-Host ((Get-Date).ToString('HH:mm:ss') + " [$Level] " + $Message) -ForegroundColor $color
}

function Test-MbApplicationIsRunning {
    if (-not (Test-Path -LiteralPath $runtimePath -PathType Leaf)) { return $false }
    try {
        $runtime = [IO.File]::ReadAllText($runtimePath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $url = [string]$runtime.url
        if ($url -notmatch '^http://localhost:\d+/$') { return $false }
        $process = Get-Process -Id ([int]$runtime.pid) -ErrorAction SilentlyContinue
        if (-not $process) { return $false }
        try {
            $health = Invoke-WebRequest -UseBasicParsing -Uri ($url + 'api/health') -TimeoutSec 1
            return $health.StatusCode -eq 200
        } catch {
            return $true
        }
    } catch {
        return $false
    }
}

function Show-MbAlreadyRunningNotice {
    $message = if ($ImportFrom) {
        "ManualBuilderはすでに起動しています。`r`n既存のブラウザータブでManualBuilderを終了してから、もう一度「編集する」を実行してください。"
    } else {
        "ManualBuilderはすでに起動しています。`r`n既存のブラウザータブへ戻ってください。"
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [void][System.Windows.Forms.MessageBox]::Show(
            $message,
            'ManualBuilder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Write-MbLauncherLog $message 'WARN'
    }
}

$startArguments = @{}
if (-not [string]::IsNullOrWhiteSpace($ImportFrom)) { $startArguments['ImportFrom'] = $ImportFrom }
if (-not [string]::IsNullOrWhiteSpace($PublishTo)) { $startArguments['PublishTo'] = $PublishTo }

if ($sourceRoot.Equals($cacheRoot, [StringComparison]::OrdinalIgnoreCase)) {
    & (Join-Path $sourceRoot 'src\Start-ManualBuilder.ps1') @startArguments
    return
}

$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$updateMutex = New-Object System.Threading.Mutex($false, ("Local\ManualBuilder-Update-$sid"))
$updateMutexAcquired = $false
$startScript = $null
try {
    try {
        $updateMutexAcquired = $updateMutex.WaitOne(30000)
    } catch [System.Threading.AbandonedMutexException] {
        $updateMutexAcquired = $true
    }
    if (-not $updateMutexAcquired) {
        throw '別の更新処理が完了するまで30秒以内に待機できませんでした。'
    }

    if (Test-MbApplicationIsRunning) {
        Show-MbAlreadyRunningNotice
        return
    }

    try {
        $install = Install-MbLocalApplication -SourceRoot $sourceRoot -CacheRoot $cacheRoot
        if ($install.Updated) {
            Write-MbLauncherLog "ローカル実行版を更新しました: v$($install.Version)" 'OK'
        } else {
            Write-MbLauncherLog "ローカル実行版は最新です: v$($install.Version)" 'INFO'
        }
        $startScript = Join-Path ([string]$install.CacheRoot) 'src\Start-ManualBuilder.ps1'
    } catch {
        Write-MbLauncherLog "配布元から更新できませんでした: $($_.Exception.Message)" 'WARN'
        if (Test-MbCachedApplication -CacheRoot $cacheRoot) {
            Write-MbLauncherLog '検証済みの既存ローカル版で起動します。' 'INFO'
            $startScript = $cachedStartScript
        } elseif (Test-MbCachedApplication -CacheRoot ($cacheRoot + '.previous')) {
            Write-MbLauncherLog '直前の検証済みローカル版で起動します。' 'INFO'
            $startScript = Join-Path ($cacheRoot + '.previous') 'src\Start-ManualBuilder.ps1'
        } else {
            throw '起動できる検証済みローカル版がありません。配布元への接続を確認してください。'
        }
    }
} catch {
    Write-MbLauncherLog $_.Exception.Message 'ERROR'
    exit 1
} finally {
    if ($updateMutexAcquired) {
        try { $updateMutex.ReleaseMutex() } catch { }
    }
    try { $updateMutex.Dispose() } catch { }
}

if (-not $startScript -or -not (Test-Path -LiteralPath $startScript -PathType Leaf)) {
    Write-MbLauncherLog 'ローカル実行版の起動ファイルが見つかりません。' 'ERROR'
    exit 1
}
Write-MbLauncherLog "ローカル版から起動します: $startScript" 'INFO'
& $startScript -LegacyAppRoot $sourceRoot @startArguments
