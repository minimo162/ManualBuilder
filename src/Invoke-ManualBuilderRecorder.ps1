# ManualBuilder operation recorder worker.
#
# 操作を記録する専用プロセス。60Hzでマウスとキーの状態を見続けるため、
# HTTPを捌く本体とは必ず分ける。停止は stop.requested、進捗は status.json で受け渡す。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$EventsDirectory,
    [Parameter(Mandatory = $true)][string]$EventsPath,
    [Parameter(Mandatory = $true)][string]$FramesDirectory,
    [Parameter(Mandatory = $true)][string]$FramesPath,
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][string]$StopPath,
    [Parameter(Mandatory = $true)][string]$PausePath,
    [Parameter(Mandatory = $true)][string]$UndoPath,
    [Parameter(Mandatory = $true)][string]$JobId,
    [AllowEmptyString()][string]$IgnoreTitlePatterns = '',
    [AllowEmptyString()][string]$DomTargetPath = '',
    [AllowEmptyString()][string]$UiaTargetPath = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Recorder.psm1') -Force

try {
    $patterns = @()
    if (-not [string]::IsNullOrWhiteSpace($IgnoreTitlePatterns)) {
        $patterns = @($IgnoreTitlePatterns -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if (-not (Test-Path -LiteralPath $EventsDirectory)) {
        [void](New-Item -ItemType Directory -Path $EventsDirectory -Force)
    }
    if (-not (Test-Path -LiteralPath $FramesDirectory)) {
        [void](New-Item -ItemType Directory -Path $FramesDirectory -Force)
    }

    [void](Invoke-MbRecordingLoop -EventsDirectory $EventsDirectory -EventsPath $EventsPath `
        -FramesDirectory $FramesDirectory -FramesPath $FramesPath `
        -StatusPath $StatusPath -StopPath $StopPath -PausePath $PausePath -UndoPath $UndoPath `
        -JobId $JobId -IgnoreTitlePatterns $patterns `
        -DomTargetPath $DomTargetPath -UiaTargetPath $UiaTargetPath)
    exit 0
} catch {
    try {
        Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'failed' -Count 0 `
            -Message ('操作を記録できませんでした: ' + $_.Exception.Message)
    } catch { }
    exit 1
}
