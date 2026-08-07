[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class MbOfficeVisualNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

$workbookFullPath = [IO.Path]::GetFullPath($WorkbookPath)
$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $workbookFullPath -PathType Leaf)) {
    throw "Workbook not found: $workbookFullPath"
}
$outputDirectory = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
}

$existingPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = $null
$workbook = $null
$ownedPid = 0
$ownershipProven = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $pidValue = [uint32]0
    [void][MbOfficeVisualNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $ownedPid = [int]$pidValue
    $ownershipProven = $ownedPid -gt 0 -and $ownedPid -notin $existingPids
    if (-not $ownershipProven) {
        throw 'MB_CONNECTED_TO_EXISTING_EXCEL'
    }

    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $true)
    $workbook.ExportAsFixedFormat(0, $outputFullPath)
    if (-not (Test-Path -LiteralPath $outputFullPath -PathType Leaf)) {
        throw 'Excel did not create the PDF.'
    }
    [pscustomobject]@{
        state = 'completed'
        outputPath = $outputFullPath
        ownedExcelPid = $ownedPid
        preexistingExcelPids = $existingPids
        ownershipProven = $ownershipProven
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
    }
    if ($null -ne $excel -and $ownershipProven) {
        try { $excel.Quit() } catch { }
    }
    foreach ($comObject in @($workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch { }
        }
    }
}
