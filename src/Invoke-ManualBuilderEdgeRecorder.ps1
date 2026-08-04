# 記録用EdgeのDOM監視ワーカー。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CachePath,
    [Parameter(Mandatory = $true)][string]$StopPath,
    [ValidateRange(1024, 65500)][int]$Port = 9465
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.EdgeRecorder.psm1') -Force

try {
    Invoke-MbEdgeRecorderCacheLoop -CachePath $CachePath -StopPath $StopPath -Port $Port
    exit 0
} catch {
    exit 1
}
