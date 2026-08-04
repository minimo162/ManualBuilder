# 記録用EdgeのDOM監視ワーカー。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CachePath,
    [Parameter(Mandatory = $true)][string]$StopPath,
    [AllowEmptyString()][string]$LogPath = '',
    [ValidateRange(1024, 65500)][int]$Port = 9465
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.EdgeRecorder.psm1') -Force

try {
    Invoke-MbEdgeRecorderCacheLoop -CachePath $CachePath -StopPath $StopPath -LogPath $LogPath -Port $Port
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
