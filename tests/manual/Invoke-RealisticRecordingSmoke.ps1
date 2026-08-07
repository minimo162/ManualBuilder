[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$FixturePath,
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [string]$ReadyPath = '',
    [string]$StartPath = '',
    [string]$DonePath = '',
    [string]$ReleasePath = '',
    [switch]$AllowExistingExcel
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if (-not $AllowExistingExcel -and (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
    throw 'Existing Excel process detected; stopping the smoke test.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class MbRecordingSmokeNative {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int command);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint attach, uint attachTo, bool value);
    [DllImport("user32.dll")] public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
}
'@

function Set-MbSmokeForeground {
    param([Parameter(Mandatory = $true)][IntPtr]$Handle)
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        $foregroundHandle = [MbRecordingSmokeNative]::GetForegroundWindow()
        $foregroundPid = [uint32]0
        $foregroundThread = [MbRecordingSmokeNative]::GetWindowThreadProcessId($foregroundHandle, [ref]$foregroundPid)
        $currentThread = [MbRecordingSmokeNative]::GetCurrentThreadId()
        $attached = $false
        try {
            if ($foregroundThread -ne 0 -and $foregroundThread -ne $currentThread) {
                $attached = [MbRecordingSmokeNative]::AttachThreadInput($currentThread, $foregroundThread, $true)
            }
            [void][MbRecordingSmokeNative]::ShowWindow($Handle, 3)
            [void][MbRecordingSmokeNative]::BringWindowToTop($Handle)
            [void][MbRecordingSmokeNative]::SetForegroundWindow($Handle)
        } finally {
            if ($attached) { [void][MbRecordingSmokeNative]::AttachThreadInput($currentThread, $foregroundThread, $false) }
        }
        Start-Sleep -Milliseconds 350
        if ([MbRecordingSmokeNative]::GetForegroundWindow() -eq $Handle) { return }

        # A short Alt press gives this foreground process permission to switch windows.
        [MbRecordingSmokeNative]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
        [MbRecordingSmokeNative]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
        [void][MbRecordingSmokeNative]::SetForegroundWindow($Handle)
        Start-Sleep -Milliseconds 350
        if ([MbRecordingSmokeNative]::GetForegroundWindow() -eq $Handle) { return }
    }
    throw 'Could not bring the test window to the foreground.'
}

function Invoke-MbSmokeClick {
    param([Parameter(Mandatory = $true)][IntPtr]$Handle, [int]$RelativeX, [int]$RelativeY)
    $rect = New-Object MbRecordingSmokeNative+RECT
    if (-not [MbRecordingSmokeNative]::GetWindowRect($Handle, [ref]$rect)) { throw 'Could not read the test window bounds.' }
    [void][MbRecordingSmokeNative]::SetCursorPos($rect.Left + $RelativeX, $rect.Top + $RelativeY)
    Start-Sleep -Milliseconds 160
    [MbRecordingSmokeNative]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 80
    [MbRecordingSmokeNative]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 650
}

function Send-MbSmokeKeys {
    param([Parameter(Mandatory = $true)][string]$Keys)
    [Windows.Forms.SendKeys]::SendWait($Keys)
    Start-Sleep -Milliseconds 700
}

function Send-MbSmokeText {
    param([Parameter(Mandatory = $true)][string]$Value, [switch]$Enter)
    foreach ($character in $Value.ToCharArray()) {
        [Windows.Forms.SendKeys]::SendWait([string]$character)
        Start-Sleep -Milliseconds 90
    }
    if ($Enter) { [Windows.Forms.SendKeys]::SendWait('{ENTER}') }
    Start-Sleep -Milliseconds 700
}

$fixtureFullPath = [IO.Path]::GetFullPath($FixturePath)
$workbookFullPath = [IO.Path]::GetFullPath($WorkbookPath)
$workbookDirectory = Split-Path -Parent $workbookFullPath
[void](New-Item -ItemType Directory -Path $workbookDirectory -Force)
$fixtureUri = ([Uri]$fixtureFullPath).AbsoluteUri

$edgeHandle = [IntPtr]::Zero
$excel = $null
$workbook = $null
$sheet = $null
try {
$existingEdgeHandles = @(Get-Process -Name msedge -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    ForEach-Object { [long]$_.MainWindowHandle })
Start-Process -FilePath 'msedge.exe' -ArgumentList '--new-window', $fixtureUri | Out-Null
$edgeWindow = $null
for ($attempt = 0; $attempt -lt 40; $attempt++) {
    Start-Sleep -Milliseconds 250
    $edgeWindow = Get-Process -Name msedge -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 -and
            $existingEdgeHandles -notcontains [long]$_.MainWindowHandle -and
            $_.MainWindowTitle -like '*ManualBuilder Recorder Smoke*' } |
        Select-Object -First 1
    if ($edgeWindow) { break }
}
if (-not $edgeWindow) { throw 'Could not find the Edge smoke-test window.' }

$edgeHandle = [IntPtr]$edgeWindow.MainWindowHandle
Set-MbSmokeForeground -Handle $edgeHandle
Start-Sleep -Milliseconds 2000

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = 'Order Entry'
$sheet.Range('A1').Value2 = 'Item'
$sheet.Range('B1').Value2 = 'Value'
$sheet.Range('A2').Value2 = 'Product A'
$sheet.Range('A3').Value2 = 'Product B'
$sheet.Range('A4').Value2 = 'Status'
$sheet.Columns.Item('A:B').ColumnWidth = 20
$workbook.SaveAs($workbookFullPath, 51)
$excel.WindowState = -4137
$excelHandle = [IntPtr]$excel.Hwnd
$excelPidValue = [uint32]0
[void][MbRecordingSmokeNative]::GetWindowThreadProcessId($excelHandle, [ref]$excelPidValue)

$session = [pscustomobject]@{
    EdgeWindowHandle = [long]$edgeHandle
    EdgeProcessId = [int]$edgeWindow.Id
    ExcelWindowHandle = [long]$excelHandle
    ExcelProcessId = [int]$excelPidValue
    WorkbookPath = $workbookFullPath
}

if ($ReadyPath) {
    $readyFullPath = [IO.Path]::GetFullPath($ReadyPath)
    $readyDirectory = Split-Path -Parent $readyFullPath
    if ($readyDirectory) { [void](New-Item -ItemType Directory -Path $readyDirectory -Force) }
    [IO.File]::WriteAllText($readyFullPath, ($session | ConvertTo-Json -Compress))
}

if ($StartPath) {
    $startFullPath = [IO.Path]::GetFullPath($StartPath)
    $started = $false
    for ($attempt = 0; $attempt -lt 360; $attempt++) {
        if (Test-Path -LiteralPath $startFullPath) {
            $started = $true
            break
        }
        Start-Sleep -Milliseconds 250
    }
    if (-not $started) { throw 'Timed out waiting for the recording start signal.' }
}

# Edge: enter filters, search, and open the resulting details.
Set-MbSmokeForeground -Handle $edgeHandle
Start-Sleep -Milliseconds 700
# GetWindowRect is in physical pixels while Edge content on this 125%-scaled test PC
# is rendered at 1600x960. These points target the visible control centres after scaling.
Invoke-MbSmokeClick -Handle $edgeHandle -RelativeX 185 -RelativeY 384
Send-MbSmokeText -Value 'C-1042'
Invoke-MbSmokeClick -Handle $edgeHandle -RelativeX 435 -RelativeY 384
Send-MbSmokeKeys -Keys '{HOME}{DOWN}{ENTER}'
Invoke-MbSmokeClick -Handle $edgeHandle -RelativeX 620 -RelativeY 384
Invoke-MbSmokeClick -Handle $edgeHandle -RelativeX 1070 -RelativeY 555
Start-Sleep -Milliseconds 2200

Set-MbSmokeForeground -Handle $excelHandle
Start-Sleep -Milliseconds 700

# Excel: record three clicks and three different results.
Invoke-MbSmokeClick -Handle $excelHandle -RelativeX 280 -RelativeY 285
Send-MbSmokeText -Value '1200' -Enter
Invoke-MbSmokeClick -Handle $excelHandle -RelativeX 280 -RelativeY 311
Send-MbSmokeText -Value '350' -Enter
Invoke-MbSmokeClick -Handle $excelHandle -RelativeX 280 -RelativeY 336
Send-MbSmokeText -Value 'ready' -Enter
Send-MbSmokeKeys -Keys '^s'

if (-not [string]::IsNullOrWhiteSpace($DonePath)) {
    $doneFullPath = [IO.Path]::GetFullPath($DonePath)
    $doneDirectory = Split-Path -Parent $doneFullPath
    if ($doneDirectory) { [void](New-Item -ItemType Directory -Path $doneDirectory -Force) }
    [IO.File]::WriteAllText($doneFullPath, 'done', [Text.UTF8Encoding]::new($false))
    if (-not [string]::IsNullOrWhiteSpace($ReleasePath)) {
        $releaseFullPath = [IO.Path]::GetFullPath($ReleasePath)
        $released = $false
        for ($attempt = 0; $attempt -lt 360; $attempt++) {
            if (Test-Path -LiteralPath $releaseFullPath -PathType Leaf) {
                $released = $true
                break
            }
            Start-Sleep -Milliseconds 250
        }
        if (-not $released) { throw 'Timed out waiting for the recorder to stop.' }
    }
}

$session | ConvertTo-Json -Compress
} finally {
    # The smoke test owns this COM instance and Edge window. Always release both so
    # failed test runs do not leave hidden Excel processes locking temporary books.
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
    }
    foreach ($comObject in @($sheet, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch { }
        }
    }
    if ($edgeHandle -ne [IntPtr]::Zero) {
        try { [void][MbRecordingSmokeNative]::PostMessage($edgeHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) } catch { }
    }
}
