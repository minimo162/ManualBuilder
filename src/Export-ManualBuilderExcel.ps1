# ManualBuilder Excel export worker.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][string]$CancelPath,
    [Parameter(Mandatory = $true)][string]$JobId
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Excel.psm1') -Force

try {
    $project = Get-MbProject -Path $ProjectPath
    $result = Invoke-MbExcelExport -Project $project -ProjectPath $ProjectPath -OutputDirectory $OutputDirectory `
        -StatusPath $StatusPath -CancelPath $CancelPath -JobId $JobId
    if ($result.state -eq 'completed' -or $result.state -eq 'cancelled') { exit 0 }
    exit 1
} catch {
    $status = [pscustomobject]@{
        jobId = $JobId
        state = 'failed'
        phase = 'failed'
        message = 'Excel出力処理を開始できませんでした: ' + $_.Exception.Message
        percent = 0
        currentStep = 0
        totalSteps = 0
        outputPath = ''
        outputName = ''
        outputFolder = ''
        outputFolderName = ''
        videoCount = 0
        outputDirectory = $OutputDirectory
        sheetNameMappings = @()
        ownedExcelPid = 0
        ownedExcelStartTimeUtc = ''
        ownershipMode = ''
        ownershipProven = $false
        startedAt = [DateTime]::UtcNow.ToString('o')
        updatedAt = [DateTime]::UtcNow.ToString('o')
        completedAt = [DateTime]::UtcNow.ToString('o')
        errorCode = 'WORKER_START_FAILED'
    }
    $directory = Split-Path -Parent $StatusPath
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    [IO.File]::WriteAllText($StatusPath, ($status | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
    exit 1
}
