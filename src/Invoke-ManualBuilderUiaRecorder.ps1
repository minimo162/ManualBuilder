# クリック前のWindows UI Automation対象を保持するワーカー。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CachePath,
    [Parameter(Mandatory = $true)][string]$StopPath,
    [AllowEmptyString()][string]$LogPath = '',
    [AllowEmptyString()][string]$IgnoreTitlePatterns = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Recorder.psm1') -Force

try {
    $patterns = @()
    if (-not [string]::IsNullOrWhiteSpace($IgnoreTitlePatterns)) {
        $patterns = @($IgnoreTitlePatterns -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    Invoke-MbUiaTargetCacheLoop -CachePath $CachePath -StopPath $StopPath -LogPath $LogPath `
        -IgnoreTitlePatterns $patterns
    exit 0
} catch {
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        try {
            $line = [DateTime]::UtcNow.ToString('o') + "`tFATAL`t" + ($_ | Out-String) + [Environment]::NewLine
            [IO.File]::AppendAllText($LogPath, $line, (New-Object Text.UTF8Encoding($false)))
        } catch { }
    }
    exit 1
}
