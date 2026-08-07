[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$DocumentPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class MbWordVisualNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

$documentFullPath = [IO.Path]::GetFullPath($DocumentPath)
$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $documentFullPath -PathType Leaf)) {
    throw "Document not found: $documentFullPath"
}
$outputDirectory = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
}

$existingPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
$document = $null
$ownedPid = 0
$ownershipProven = $false
try {
    if ($existingPids.Count -gt 0) {
        throw 'MB_WORD_RUNNING'
    }
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($documentFullPath, $false, $true)
    $pidValue = [uint32]0
    [void][MbWordVisualNative]::GetWindowThreadProcessId([IntPtr]$word.ActiveWindow.Hwnd, [ref]$pidValue)
    $ownedPid = [int]$pidValue
    $ownershipProven = $ownedPid -gt 0 -and $ownedPid -notin $existingPids
    if (-not $ownershipProven) {
        throw 'MB_CONNECTED_TO_EXISTING_WORD'
    }

    $document.ExportAsFixedFormat($outputFullPath, 17)
    if (-not (Test-Path -LiteralPath $outputFullPath -PathType Leaf)) {
        throw 'Word did not create the PDF.'
    }
    [pscustomobject]@{
        state = 'completed'
        outputPath = $outputFullPath
        ownedWordPid = $ownedPid
        preexistingWordPids = $existingPids
        ownershipProven = $ownershipProven
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $document) {
        try { $document.Close($false) } catch { }
    }
    if ($null -ne $word -and $ownershipProven) {
        try { $word.Quit() } catch { }
    }
    foreach ($comObject in @($document, $word)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch { }
        }
    }
}
