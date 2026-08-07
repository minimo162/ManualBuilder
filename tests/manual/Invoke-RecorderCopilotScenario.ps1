[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('edge-excel-order-transfer', 'excel-multi-operation', 'edge-delayed-transition')]
    [string]$Scenario,
    [string]$WorkbookPath = '',
    [string]$ReadyPath = '',
    [string]$StartPath = '',
    [string]$DonePath = '',
    [string]$ReleasePath = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$existingHarness = Join-Path $PSScriptRoot 'Invoke-RealisticRecordingSmoke.ps1'

if ($Scenario -eq 'edge-excel-order-transfer') {
    if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
        throw '-WorkbookPath is required for edge-excel-order-transfer.'
    }
    & $existingHarness -FixturePath (Join-Path $repoRoot 'tests\fixtures\recorder-realistic.html') `
        -WorkbookPath $WorkbookPath -ReadyPath $ReadyPath -StartPath $StartPath `
        -DonePath $DonePath -ReleasePath $ReleasePath -AllowExistingExcel
    [pscustomobject]@{ scenario = $Scenario; completed = $true } | ConvertTo-Json -Compress
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class MbRecorderScenarioNative {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int command);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint attach, uint attachTo, bool value);
    [DllImport("user32.dll")] public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, UIntPtr extraInfo);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern short VkKeyScan(char character);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
}
'@

function Set-MbScenarioForeground {
    param([Parameter(Mandatory = $true)][IntPtr]$Handle)
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        $foreground = [MbRecorderScenarioNative]::GetForegroundWindow()
        $foregroundPid = [uint32]0
        $foregroundThread = [MbRecorderScenarioNative]::GetWindowThreadProcessId($foreground, [ref]$foregroundPid)
        $currentThread = [MbRecorderScenarioNative]::GetCurrentThreadId()
        $attached = $false
        try {
            if ($foregroundThread -ne 0 -and $foregroundThread -ne $currentThread) {
                $attached = [MbRecorderScenarioNative]::AttachThreadInput($currentThread, $foregroundThread, $true)
            }
            [void][MbRecorderScenarioNative]::ShowWindow($Handle, 3)
            [void][MbRecorderScenarioNative]::BringWindowToTop($Handle)
            [void][MbRecorderScenarioNative]::SetForegroundWindow($Handle)
        } finally {
            if ($attached) { [void][MbRecorderScenarioNative]::AttachThreadInput($currentThread, $foregroundThread, $false) }
        }
        Start-Sleep -Milliseconds 300
        if ([MbRecorderScenarioNative]::GetForegroundWindow() -eq $Handle) { return }
        [MbRecorderScenarioNative]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
        [MbRecorderScenarioNative]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
    }
    throw 'Could not bring the owned scenario window to the foreground.'
}

function Get-MbScenarioControl {
    param(
        [Parameter(Mandatory = $true)][IntPtr]$Handle,
        [Parameter(Mandatory = $true)][string[]]$NamePatterns,
        [int]$TimeoutSeconds = 10
    )
    $root = [Windows.Automation.AutomationElement]::FromHandle($Handle)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $items = $root.FindAll([Windows.Automation.TreeScope]::Descendants,
            [Windows.Automation.Condition]::TrueCondition)
        foreach ($item in $items) {
            $name = [string]$item.Current.Name
            if ([string]::IsNullOrWhiteSpace($name) -or $item.Current.IsOffscreen) { continue }
            foreach ($pattern in $NamePatterns) {
                if ($name -match $pattern) { return $item }
            }
        }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    throw ('Could not find a visible control matching: ' + ($NamePatterns -join ', '))
}

function Invoke-MbScenarioControlClick {
    param([Parameter(Mandatory = $true)]$Control)
    $bounds = $Control.Current.BoundingRectangle
    if ($bounds.IsEmpty -or $bounds.Width -le 1 -or $bounds.Height -le 1) { throw 'The target control has no clickable bounds.' }
    [void][MbRecorderScenarioNative]::SetCursorPos(
        [int][Math]::Round($bounds.Left + ($bounds.Width / 2)),
        [int][Math]::Round($bounds.Top + ($bounds.Height / 2)))
    Start-Sleep -Milliseconds 150
    [MbRecorderScenarioNative]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 80
    [MbRecorderScenarioNative]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 650
}

function Invoke-MbScenarioDrag {
    param([Parameter(Mandatory = $true)]$From, [Parameter(Mandatory = $true)]$To)
    $first = $From.Current.BoundingRectangle
    $last = $To.Current.BoundingRectangle
    $startX = [int][Math]::Round($first.Left + ($first.Width / 2))
    $startY = [int][Math]::Round($first.Top + ($first.Height / 2))
    $endX = [int][Math]::Round($last.Left + ($last.Width / 2))
    $endY = [int][Math]::Round($last.Top + ($last.Height / 2))
    [void][MbRecorderScenarioNative]::SetCursorPos($startX, $startY)
    Start-Sleep -Milliseconds 150
    [MbRecorderScenarioNative]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    for ($step = 1; $step -le 8; $step++) {
        [void][MbRecorderScenarioNative]::SetCursorPos(
            [int][Math]::Round($startX + (($endX - $startX) * $step / 8.0)),
            [int][Math]::Round($startY + (($endY - $startY) * $step / 8.0)))
        Start-Sleep -Milliseconds 70
    }
    [MbRecorderScenarioNative]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 800
}

function Send-MbScenarioText {
    param([Parameter(Mandatory = $true)][string]$Value, [switch]$Enter)
    if ($Value.StartsWith('=')) {
        # Excelの数式記号はキーボード配列によって異なるため、テスト環境の配列差で
        # 欠けた文字を録画品質の不具合と誤判定しないよう、数式だけは貼り付ける。
        # Ctrl+Vはレコーダー側で入力操作として検出され、完成した数式を証拠画像に残せる。
        $clipboardBackup = $null
        try {
            $clipboardBackup = [Windows.Forms.Clipboard]::GetDataObject()
            [Windows.Forms.Clipboard]::SetText($Value)
            [MbRecorderScenarioNative]::keybd_event(0x11, 0, 0, [UIntPtr]::Zero)
            [MbRecorderScenarioNative]::keybd_event(0x56, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 350
            [MbRecorderScenarioNative]::keybd_event(0x56, 0, 2, [UIntPtr]::Zero)
            [MbRecorderScenarioNative]::keybd_event(0x11, 0, 2, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 500
        }
        finally {
            if ($null -ne $clipboardBackup) {
                [Windows.Forms.Clipboard]::SetDataObject($clipboardBackup, $true)
            }
            else {
                [Windows.Forms.Clipboard]::Clear()
            }
        }
    }
    else {
        foreach ($character in $Value.ToCharArray()) {
            $text = [string]$character
            # SendKeysは押下時間が短く、16ms巡回の実機録画では文字キーを見失う。
            # 現在の日本語/英語キーボード配列から実キーを求め、人間のキー押下に近い
            # 90msのkeydownを送る。これでテスト側だけが不自然に速い状態を避ける。
            $mapping = [int][MbRecorderScenarioNative]::VkKeyScan([char]$text)
            if ($mapping -eq -1) { throw "Could not map character to the active keyboard layout: $text" }
            $virtualKey = [byte]($mapping -band 0xFF)
            $modifiers = ($mapping -shr 8) -band 0xFF
            if (($modifiers -band 2) -ne 0) { [MbRecorderScenarioNative]::keybd_event(0x11, 0, 0, [UIntPtr]::Zero) }
            if (($modifiers -band 4) -ne 0) { [MbRecorderScenarioNative]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero) }
            if (($modifiers -band 1) -ne 0) { [MbRecorderScenarioNative]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero) }
            [MbRecorderScenarioNative]::keybd_event($virtualKey, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 90
            [MbRecorderScenarioNative]::keybd_event($virtualKey, 0, 2, [UIntPtr]::Zero)
            if (($modifiers -band 1) -ne 0) { [MbRecorderScenarioNative]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero) }
            if (($modifiers -band 4) -ne 0) { [MbRecorderScenarioNative]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero) }
            if (($modifiers -band 2) -ne 0) { [MbRecorderScenarioNative]::keybd_event(0x11, 0, 2, [UIntPtr]::Zero) }
            Start-Sleep -Milliseconds 80
        }
    }
    # SendWait直後にEnterを重ねると、人間には不可能な速さで最終文字の描画前に
    # 確定される。実際のタイピングに近い短い間を置き、完成状態も評価対象にする。
    if ($Enter) {
        Start-Sleep -Milliseconds 180
        [MbRecorderScenarioNative]::keybd_event(0x0D, 0, 0, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 60
        [MbRecorderScenarioNative]::keybd_event(0x0D, 0, 2, [UIntPtr]::Zero)
    }
    Start-Sleep -Milliseconds 700
}

function Write-MbScenarioReady {
    param([hashtable]$Details)
    if ([string]::IsNullOrWhiteSpace($ReadyPath)) { return }
    $fullPath = [IO.Path]::GetFullPath($ReadyPath)
    $directory = Split-Path -Parent $fullPath
    if ($directory) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    [IO.File]::WriteAllText($fullPath, (($Details | ConvertTo-Json -Compress)), [Text.UTF8Encoding]::new($false))
}

function Wait-MbScenarioStart {
    if ([string]::IsNullOrWhiteSpace($StartPath)) { return }
    $fullPath = [IO.Path]::GetFullPath($StartPath)
    for ($attempt = 0; $attempt -lt 360; $attempt++) {
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) { return }
        Start-Sleep -Milliseconds 250
    }
    throw 'Timed out waiting for the recording start signal.'
}

function Complete-MbScenarioActions {
    if ([string]::IsNullOrWhiteSpace($DonePath)) { return }
    $fullDonePath = [IO.Path]::GetFullPath($DonePath)
    $directory = Split-Path -Parent $fullDonePath
    if ($directory) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    [IO.File]::WriteAllText($fullDonePath, 'done', [Text.UTF8Encoding]::new($false))
    if ([string]::IsNullOrWhiteSpace($ReleasePath)) { return }
    $fullReleasePath = [IO.Path]::GetFullPath($ReleasePath)
    for ($attempt = 0; $attempt -lt 360; $attempt++) {
        if (Test-Path -LiteralPath $fullReleasePath -PathType Leaf) { return }
        Start-Sleep -Milliseconds 250
    }
    throw 'Timed out waiting for the recorder to stop.'
}

$ownedEdgeHandle = [IntPtr]::Zero
$excel = $null
$workbook = $null
$sheet = $null
try {
    if ($Scenario -eq 'excel-multi-operation') {
        if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { throw '-WorkbookPath is required for excel-multi-operation.' }
        $fullWorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
        $workbookDirectory = Split-Path -Parent $fullWorkbookPath
        if ($workbookDirectory) { [void](New-Item -ItemType Directory -Path $workbookDirectory -Force) }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = 'Sales'
        $sheet.Range('A1').Value2 = 'Product'
        $sheet.Range('B1').Value2 = 'Sales'
        $sheet.Range('A2').Value2 = 'Product A'
        $sheet.Range('A3').Value2 = 'Product B'
        $sheet.Range('A4').Value2 = 'Total'
        $sheet.Columns.Item('A:B').ColumnWidth = 20
        $workbook.SaveAs($fullWorkbookPath, 51)
        $excel.WindowState = -4137
        $handle = [IntPtr]$excel.Hwnd
        Set-MbScenarioForeground -Handle $handle
        Write-MbScenarioReady -Details @{ scenario = $Scenario; windowHandle = [long]$handle; workbook = $fullWorkbookPath }
        Wait-MbScenarioStart
        # 録画開始の操作でManualBuilder/Codexへ前面が戻る。開始合図を受けた後に
        # 操作対象をもう一度前面へ出し、検証環境側のクリックを録画しない。
        Set-MbScenarioForeground -Handle $handle

        foreach ($entry in @(@('B2', '1200'), @('B3', '350'), @('B4', '=SUM(B2:B3)'))) {
            Set-MbScenarioForeground -Handle $handle
            $cell = Get-MbScenarioControl -Handle $handle -NamePatterns @(('^' + [regex]::Escape($entry[0]) + '$'))
            Invoke-MbScenarioControlClick -Control $cell
            Send-MbScenarioText -Value $entry[1] -Enter
        }
        Set-MbScenarioForeground -Handle $handle
        $from = Get-MbScenarioControl -Handle $handle -NamePatterns @('^B2$')
        $to = Get-MbScenarioControl -Handle $handle -NamePatterns @('^B4$')
        Invoke-MbScenarioDrag -From $from -To $to
        Set-MbScenarioForeground -Handle $handle
        $currency = Get-MbScenarioControl -Handle $handle -NamePatterns @(
            '通貨.*表示形式', '会計.*表示形式', 'Currency.*Format', 'Accounting.*Format')
        Invoke-MbScenarioControlClick -Control $currency
        [Windows.Forms.SendKeys]::SendWait('^s')
        Start-Sleep -Milliseconds 700
    } else {
        $fixture = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tests\fixtures\recorder-delayed.html'))
        $fixtureUri = ([Uri]$fixture).AbsoluteUri
        $existingHandles = @(Get-Process -Name msedge -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object { [long]$_.MainWindowHandle })
        Start-Process -FilePath 'msedge.exe' -ArgumentList '--new-window', $fixtureUri | Out-Null
        $edgeWindow = $null
        for ($attempt = 0; $attempt -lt 40; $attempt++) {
            Start-Sleep -Milliseconds 250
            $edgeWindow = Get-Process -Name msedge -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne 0 -and
                    $existingHandles -notcontains [long]$_.MainWindowHandle -and
                    $_.MainWindowTitle -like '*ManualBuilder Recorder Smoke*' } |
                Select-Object -First 1
            if ($edgeWindow) { break }
        }
        if (-not $edgeWindow) { throw 'Could not find the owned Edge scenario window.' }
        $ownedEdgeHandle = [IntPtr]$edgeWindow.MainWindowHandle
        Set-MbScenarioForeground -Handle $ownedEdgeHandle
        Write-MbScenarioReady -Details @{ scenario = $Scenario; windowHandle = [long]$ownedEdgeHandle; processId = [int]$edgeWindow.Id }
        Wait-MbScenarioStart
        Set-MbScenarioForeground -Handle $ownedEdgeHandle

        $orderId = Get-MbScenarioControl -Handle $ownedEdgeHandle -NamePatterns @('注文番号')
        Invoke-MbScenarioControlClick -Control $orderId
        Send-MbScenarioText -Value 'ORD-2026-0817'
        Invoke-MbScenarioControlClick -Control (Get-MbScenarioControl -Handle $ownedEdgeHandle -NamePatterns @('^検索$'))
        Start-Sleep -Milliseconds 1600
        Invoke-MbScenarioControlClick -Control (Get-MbScenarioControl -Handle $ownedEdgeHandle -NamePatterns @('詳細を表示'))
        Start-Sleep -Milliseconds 1400
        Invoke-MbScenarioControlClick -Control (Get-MbScenarioControl -Handle $ownedEdgeHandle -NamePatterns @('一覧へ戻る'))
    }
    Complete-MbScenarioActions
    [pscustomobject]@{ scenario = $Scenario; completed = $true } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($comObject in @($sheet, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch { }
        }
    }
    if ($ownedEdgeHandle -ne [IntPtr]::Zero) {
        try { [void][MbRecorderScenarioNative]::PostMessage($ownedEdgeHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) } catch { }
    }
}
