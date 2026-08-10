# WebView2記録モニターを実際に起動し、Windows UI Automationで主要操作を確認する。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$companionRoot = Join-Path $repoRoot 'src\RecorderCompanion'
$controllerPath = Join-Path $companionRoot 'Invoke-RecorderCompanion.ps1'
$controllerSourcePath = Join-Path $companionRoot 'ManualBuilder.RecorderCompanion.cs'
$vendorRoot = Join-Path $companionRoot 'vendor\WebView2'
$webRoot = Join-Path $companionRoot 'web'
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class MbControllerTestNative {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int index);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint x, uint y, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
'@

function Find-UiaElement {
    param($Root, [string]$Name)
    $condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $Name)
    return $Root.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
}

function Invoke-UiaElement {
    param($Element, [IntPtr]$WindowHandle)
    if ($null -eq $Element) { return $false }
    $bounds = $Element.Current.BoundingRectangle
    if ($bounds.Width -le 0 -or $bounds.Height -le 0) { return $false }
    [void][MbControllerTestNative]::SetForegroundWindow($WindowHandle)
    [void][MbControllerTestNative]::SetCursorPos([int]($bounds.Left + $bounds.Width / 2), [int]($bounds.Top + $bounds.Height / 2))
    [MbControllerTestNative]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    [MbControllerTestNative]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
    return $true
}

$signedDependencies = @('Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.Wpf.dll', 'WebView2Loader.dll') | ForEach-Object {
    Get-AuthenticodeSignature -LiteralPath (Join-Path $vendorRoot $_)
}
Add-Result ((Test-Path -LiteralPath $controllerPath -PathType Leaf) -and
    (Test-Path -LiteralPath $controllerSourcePath -PathType Leaf) -and
    (Test-Path -LiteralPath $webRoot -PathType Container) -and
    @($signedDependencies | Where-Object { $_.Status -ne 'Valid' }).Count -eq 0) `
    '未署名EXEを使わずMicrosoft署名済みWebView2で記録モニターを実行できる'

$testRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderCompanion-' + [guid]::NewGuid().ToString('N'))
$eventsDirectory = Join-Path $testRoot 'events'
$statusPath = Join-Path $testRoot 'status.json'
$pausePath = Join-Path $testRoot 'pause.requested'
$undoPath = Join-Path $testRoot 'undo.requested'
$resultPath = Join-Path $testRoot 'result.requested'
$stopPath = Join-Path $testRoot 'stop.requested'
$jobId = 'controller-test'
$process = $null
$bitmap = $null
$compactPanelHeight = 0
$previewPath = Join-Path $repoRoot '.tmp\recorder-controller-preview.png'

try {
    [void](New-Item -ItemType Directory -Path $eventsDirectory -Force)
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $previewPath) -Force)
    foreach ($suffix in @('', '-result')) {
        $bitmap = New-Object Drawing.Bitmap 1280, 720
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear($(if ($suffix) { [Drawing.Color]::FromArgb(237, 247, 239) } else { [Drawing.Color]::FromArgb(235, 241, 252) }))
            $font = New-Object Drawing.Font 'BIZ UDPゴシック', 30
            try { $graphics.DrawString($(if ($suffix) { '操作後の結果' } else { '操作した場所' }), $font, [Drawing.Brushes]::Navy, 60, 52) } finally { $font.Dispose() }
            $bitmap.Save((Join-Path $eventsDirectory ('event-001{0}.jpg' -f $suffix)), [Drawing.Imaging.ImageFormat]::Jpeg)
        } finally { $graphics.Dispose(); $bitmap.Dispose(); $bitmap = $null }
    }
    $status = [pscustomobject]@{
        jobId = $jobId; state = 'recording'; count = 1; message = 'recording'
        lastTarget = 'セル D5 に入力'; updatedAt = [DateTime]::UtcNow.ToString('o')
        resultRequestId = ''; undoRequestId = ''
    }
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $eventRecord = [pscustomobject]@{ index = 1; kind = 'input'; targetName = 'セル D5 に入力'; image = 'event-001.jpg' }
    [IO.File]::WriteAllText((Join-Path $testRoot 'events.jsonl'), (($eventRecord | ConvertTo-Json -Compress) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))

    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA',
        '-File', ('"' + $controllerPath + '"'),
        '-StatusPath', ('"' + $statusPath + '"'), '-EventsDirectory', ('"' + $eventsDirectory + '"'),
        '-PausePath', ('"' + $pausePath + '"'), '-UndoPath', ('"' + $undoPath + '"'),
        '-ResultPath', ('"' + $resultPath + '"'), '-StopPath', ('"' + $stopPath + '"'),
        '-JobId', $jobId, '-WebRoot', ('"' + $webRoot + '"'), '-TestMode'
    )
    $process = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -WorkingDirectory $vendorRoot -WindowStyle Hidden -PassThru

    $handle = [IntPtr]::Zero
    for ($attempt = 0; $attempt -lt 120 -and $handle -eq [IntPtr]::Zero; $attempt++) {
        Start-Sleep -Milliseconds 100
        $process.Refresh()
        $foundHandle = [MbControllerTestNative]::FindWindow($null, 'ManualBuilder Recorder')
        if ($null -eq $foundHandle -or $foundHandle -eq [IntPtr]::Zero) { $foundHandle = $process.MainWindowHandle }
        if ($null -ne $foundHandle -and $foundHandle -ne [IntPtr]::Zero) { $handle = [IntPtr]$foundHandle }
        if ($process.HasExited) { break }
    }
    Add-Result ($handle -ne [IntPtr]::Zero) 'WebView2記録モニターが画面上へ開く'
    if ($handle -eq [IntPtr]::Zero) {
        if (Test-Path -LiteralPath ($statusPath + '.companion.log')) { Write-Host ([IO.File]::ReadAllText(($statusPath + '.companion.log'))) -ForegroundColor Yellow }
        throw 'WebView2記録モニターを開けませんでした。'
    }

    $readyPath = $statusPath + '.companion.ready'
    for ($attempt = 0; $attempt -lt 120 -and -not (Test-Path -LiteralPath $readyPath); $attempt++) {
        Start-Sleep -Milliseconds 100
    }
    Add-Result (Test-Path -LiteralPath $readyPath) '記録モニターの画面準備完了を待てる'
    Start-Sleep -Milliseconds 400
    $extendedStyle = [MbControllerTestNative]::GetWindowLong($handle, -20)
    Add-Result (($extendedStyle -band 0x00000008) -ne 0) '記録モニターが常に手前に表示される'
    $startupLogPath = $statusPath + '.companion.start.log'
    $startupLog = if (Test-Path -LiteralPath $startupLogPath) { [IO.File]::ReadAllText($startupLogPath) } else { '' }
    Add-Result ($startupLog -match 'dpiAwareness=2') '記録モニターがモニターごとのDPI倍率を認識する'

    $root = [Windows.Automation.AutomationElement]::FromHandle($handle)
    $panelBounds = $root.Current.BoundingRectangle

    $status.state = 'ready'
    [IO.File]::WriteAllText($pausePath, 'ready', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $startElement = $null
    for ($attempt = 0; $attempt -lt 20 -and $null -eq $startElement; $attempt++) {
        $startElement = Find-UiaElement -Root $root -Name '記録を開始'
        if ($null -eq $startElement) { Start-Sleep -Milliseconds 100 }
    }
    Add-Result ($null -ne $startElement -and $startElement.Current.IsEnabled) '待機中は「記録を開始」と案内する'
    Add-Result ($null -ne (Find-UiaElement -Root $root -Name '開始待ち')) '待機状態を短く明確に表示する'
    $shortcutGuide = (Find-UiaElement -Root $root -Name 'Ctrl+Alt+Space：開始・一時停止・再開 Ctrl+Alt+Enter：終了確認')
    $shortcutUnavailable = (Find-UiaElement -Root $root -Name 'ショートカットは現在利用できません。画面のボタンをお使いください。')
    Add-Result ($null -ne $shortcutGuide -or $null -ne $shortcutUnavailable) 'ショートカットまたは画面ボタンへの代替案内を表示する'
    [void][MbControllerTestNative]::SendMessage($handle, 0x0312, [IntPtr]0x4D01, [IntPtr]::Zero)
    [void][MbControllerTestNative]::SendMessage($handle, 0x0312, [IntPtr]0x4D01, [IntPtr]::Zero)
    for ($attempt = 0; $attempt -lt 20 -and (Test-Path -LiteralPath $pausePath); $attempt++) { Start-Sleep -Milliseconds 100 }
    Add-Result (-not (Test-Path -LiteralPath $pausePath)) 'Ctrl+Alt+Space相当で待機から記録を開始できる'
    Add-Result (-not (Test-Path -LiteralPath $pausePath)) '開始操作を連打しても待機へ戻らない'

    $status.state = 'recording'
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    Start-Sleep -Milliseconds 500

    $status.state = 'paused'
    [IO.File]::WriteAllText($pausePath, 'pause', [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $resumeElement = $null
    for ($attempt = 0; $attempt -lt 20 -and $null -eq $resumeElement; $attempt++) {
        $resumeElement = Find-UiaElement -Root $root -Name '記録を再開'
        if ($null -eq $resumeElement) { Start-Sleep -Milliseconds 100 }
    }
    Add-Result ($null -ne $resumeElement -and $resumeElement.Current.IsEnabled) '一時停止中は「記録を再開」と案内する'
    [void][MbControllerTestNative]::SendMessage($handle, 0x0312, [IntPtr]0x4D01, [IntPtr]::Zero)
    [void][MbControllerTestNative]::SendMessage($handle, 0x0312, [IntPtr]0x4D01, [IntPtr]::Zero)
    for ($attempt = 0; $attempt -lt 20 -and (Test-Path -LiteralPath $pausePath); $attempt++) { Start-Sleep -Milliseconds 100 }
    Add-Result (-not (Test-Path -LiteralPath $pausePath)) 'Ctrl+Alt+Space相当で一時停止から再開できる'
    Add-Result (-not (Test-Path -LiteralPath $pausePath)) '再開操作を連打しても一時停止へ戻らない'

    $status.state = 'recording'
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    for ($attempt = 0; $attempt -lt 20 -and $null -eq (Find-UiaElement -Root $root -Name '一時停止'); $attempt++) { Start-Sleep -Milliseconds 100 }

    [void][MbControllerTestNative]::SendMessage($handle, 0x0312, [IntPtr]0x4D02, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
    Add-Result ($null -ne (Find-UiaElement -Root $root -Name '記録を終了しますか？')) 'Ctrl+Alt+Enter相当でも終了確認を省略しない'
    [void](Invoke-UiaElement -Element (Find-UiaElement -Root $root -Name '記録を続ける') -WindowHandle $handle)
    Start-Sleep -Milliseconds 200

    foreach ($expected in @('一時停止', '直前を取り消す', '終了して確認', '直前画像を確認')) {
        $element = $null
        for ($attempt = 0; $attempt -lt 20 -and $null -eq $element; $attempt++) {
            $element = Find-UiaElement -Root $root -Name $expected
            if ($null -eq $element) { Start-Sleep -Milliseconds 100 }
        }
        Add-Result ($null -ne $element) ("記録モニターに「$expected」がある")
        if ($null -ne $element) {
            $buttonBounds = $element.Current.BoundingRectangle
            $inside = $buttonBounds.Width -gt 0 -and $buttonBounds.Height -gt 0 -and
                $buttonBounds.Left -ge $panelBounds.Left -and $buttonBounds.Top -ge $panelBounds.Top -and
                $buttonBounds.Right -le $panelBounds.Right -and $buttonBounds.Bottom -le $panelBounds.Bottom
            Add-Result $inside ("記録モニター内に「$expected」が収まる")
        }
    }
    $targetElement = $null
    for ($attempt = 0; $attempt -lt 20 -and $null -eq $targetElement; $attempt++) {
        $targetElement = Find-UiaElement -Root $root -Name 'セル D5 に入力'
        if ($null -eq $targetElement) { Start-Sleep -Milliseconds 100 }
    }
    Add-Result ($null -ne $targetElement) '直前の操作対象をその場で確認できる'
    Add-Result ($panelBounds.Width -ge 440 -and $panelBounds.Width -le 500 -and $panelBounds.Height -ge 200 -and $panelBounds.Height -le 240) `
        '作業中は約480×220の小型記録レシートで開く'
    Add-Result ($null -ne (Find-UiaElement -Root $root -Name '記録レシート')) '記録レシートの見出しを表示する'
    $expandElement = Find-UiaElement -Root $root -Name '直前画像を確認'
    Add-Result ($null -ne $expandElement) '記録レシートから直前画像を開ける'
    [void](Invoke-UiaElement -Element $expandElement -WindowHandle $handle)
    Start-Sleep -Milliseconds 500
    $root = [Windows.Automation.AutomationElement]::FromHandle($handle)
    $panelBounds = $root.Current.BoundingRectangle
    Add-Result ($panelBounds.Width -ge 940 -and $panelBounds.Height -ge 720) '必要なときは画像を細部まで確認できる大きさへ広がる'
    $compareElement = Find-UiaElement -Root $root -Name '結果画像も見る'
    Add-Result ($null -ne $compareElement) '2枚目は必要なときだけ開ける'
    Add-Result ($null -eq (Find-UiaElement -Root $root -Name '操作前')) '操作前・操作後の3択を最初から見せない'
    if (Invoke-UiaElement -Element $compareElement -WindowHandle $handle) {
        Start-Sleep -Milliseconds 400
        $returnToSingle = Find-UiaElement -Root $root -Name '1枚の表示に戻す'
        Add-Result ($null -ne $returnToSingle -and
            $null -ne (Find-UiaElement -Root $root -Name '操作した場所') -and
            $null -ne (Find-UiaElement -Root $root -Name '操作後の結果')) '必要なときだけ画像の役割を付けて2枚を表示する'
        [void](Invoke-UiaElement -Element $returnToSingle -WindowHandle $handle)
        Start-Sleep -Milliseconds 300
    }

    $undoElement = Find-UiaElement -Root $root -Name '直前の操作を取り消す'
    Add-Result ($null -ne $undoElement -and $undoElement.Current.IsEnabled) '記録があれば直前の操作を取り消せる'
    if (Invoke-UiaElement -Element $undoElement -WindowHandle $handle) {
        for ($attempt = 0; $attempt -lt 30 -and -not (Test-Path -LiteralPath $undoPath); $attempt++) { Start-Sleep -Milliseconds 100 }
        $undoRequest = if (Test-Path -LiteralPath $undoPath) { [IO.File]::ReadAllText($undoPath).Trim() } else { '' }
        Add-Result (-not [string]::IsNullOrWhiteSpace($undoRequest)) '直前操作の取り消しに識別子付きの要求を送る'
        $pendingUndo = Find-UiaElement -Root $root -Name '取り消しています…'
        Add-Result ($null -ne $pendingUndo -and -not $pendingUndo.Current.IsEnabled) '取り消しの応答待ち中は二重要求しない'
        $status.undoRequestId = $undoRequest
        [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        Start-Sleep -Milliseconds 500
        $undoElement = Find-UiaElement -Root $root -Name '直前の操作を取り消す'
        Add-Result ($null -ne $undoElement -and $undoElement.Current.IsEnabled) '取り消しの応答後に再操作できる'
    }

    $root = [Windows.Automation.AutomationElement]::FromHandle($handle)
    $compactElement = Find-UiaElement -Root $root -Name '画像確認を閉じる'
    Add-Result ($null -ne $compactElement) '画像確認表示から記録レシートへ戻れる'
    $fullRect = New-Object MbControllerTestNative+RECT
    [void][MbControllerTestNative]::GetWindowRect($handle, [ref]$fullRect)
    if (Invoke-UiaElement -Element $compactElement -WindowHandle $handle) {
        Start-Sleep -Milliseconds 500
        $smallRect = New-Object MbControllerTestNative+RECT
        [void][MbControllerTestNative]::GetWindowRect($handle, [ref]$smallRect)
        $compactPanelHeight = $smallRect.Bottom - $smallRect.Top
        Add-Result ($compactPanelHeight -lt ($fullRect.Bottom - $fullRect.Top)) '画像確認を閉じると記録レシートへ戻る'
        $expandElement = Find-UiaElement -Root $root -Name '直前画像を確認'
        Add-Result ($null -ne $expandElement) '記録レシートから再び画像確認へ戻れる'
        [void](Invoke-UiaElement -Element $expandElement -WindowHandle $handle)
        Start-Sleep -Milliseconds 500
    }

    $rect = New-Object MbControllerTestNative+RECT
    if ([MbControllerTestNative]::GetWindowRect($handle, [ref]$rect)) {
        [void][MbControllerTestNative]::SetForegroundWindow($handle)
        Start-Sleep -Milliseconds 250
        $width = $rect.Right - $rect.Left; $height = $rect.Bottom - $rect.Top
        $screen = New-Object Drawing.Bitmap $width, $height
        $graphics = [Drawing.Graphics]::FromImage($screen)
        try {
            $hdc = $graphics.GetHdc()
            try { $printed = [MbControllerTestNative]::PrintWindow($handle, $hdc, 2) } finally { $graphics.ReleaseHdc($hdc) }
            if ($printed -and $width -gt 10 -and $height -gt 10) {
                $sample = $screen.GetPixel([int]($width / 2), [int]($height / 2))
                if ($sample.R -lt 8 -and $sample.G -lt 8 -and $sample.B -lt 8) { $printed = $false }
            }
            if (-not $printed) { $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $screen.Size) }
            $screen.Save($previewPath, [Drawing.Imaging.ImageFormat]::Png)
        } finally { $graphics.Dispose(); $screen.Dispose() }
        Add-Result (Test-Path -LiteralPath $previewPath -PathType Leaf) 'WebView2記録モニターの見た目を画像で確認できる'
    }

    [void][MbControllerTestNative]::SendMessage($handle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 400
    $closeDialog = Find-UiaElement -Root $root -Name '記録を終了しますか？'
    Add-Result ($null -ne $closeDialog) '閉じる操作をモダンな確認画面で安全に止める'
    $cancelElement = Find-UiaElement -Root $root -Name '記録を続ける'
    [void](Invoke-UiaElement -Element $cancelElement -WindowHandle $handle)
    Start-Sleep -Milliseconds 250

    $finishElement = Find-UiaElement -Root $root -Name '終了して確認'
    [void](Invoke-UiaElement -Element $finishElement -WindowHandle $handle)
    Start-Sleep -Milliseconds 300
    Add-Result ($null -ne (Find-UiaElement -Root $root -Name '記録を終了しますか？')) '終了ボタンも確認画面を経由する'
    [void](Invoke-UiaElement -Element (Find-UiaElement -Root $root -Name '記録を終了する') -WindowHandle $handle)
    for ($attempt = 0; $attempt -lt 30 -and -not (Test-Path -LiteralPath $stopPath); $attempt++) { Start-Sleep -Milliseconds 100 }
    Add-Result (Test-Path -LiteralPath $stopPath) '終了操作で手順候補の作成を開始する'
    [void][MbControllerTestNative]::SendMessage($handle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
    $process.Refresh()
    Add-Result (-not $process.HasExited) '終了処理中に閉じてもモニターだけを終了しない'

    $status.state = 'failed'
    $status.message = 'テスト用の失敗理由です。'
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    Start-Sleep -Milliseconds 700
    $failedLabels = @($root.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition) | ForEach-Object { [string]$_.Current.Name })
    Add-Result (@($failedLabels | Where-Object { $_ -match 'テスト用の失敗理由' }).Count -gt 0) '失敗理由を記録モニター内で確認できる'
    if (@($failedLabels | Where-Object { $_ -match 'テスト用の失敗理由' }).Count -eq 0) {
        Write-Host ('UIA names after failure: ' + ($failedLabels -join ' | ')) -ForegroundColor Yellow
    }
    if ($compactPanelHeight -gt 0) {
        $failedRect = New-Object MbControllerTestNative+RECT
        [void][MbControllerTestNative]::GetWindowRect($handle, [ref]$failedRect)
        Add-Result (($failedRect.Bottom - $failedRect.Top) -gt $compactPanelHeight) '失敗時は画像確認表示へ戻して理由を読める'
    }

    $status.state = 'completed'
    $status.message = 'done'
    [IO.File]::WriteAllText($statusPath, ($status | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    [void]$process.WaitForExit(6000)
    Add-Result ($process.HasExited) '記録終了後にモニターが自動で閉じる'
} finally {
    if ($null -ne $bitmap) { $bitmap.Dispose() }
    if ($null -ne $process) {
        if (-not $process.HasExited) { try { $process.Kill() } catch { } }
        $process.Dispose()
    }
    if ($errors.Count -gt 0 -and (Test-Path -LiteralPath ($statusPath + '.companion.test.log'))) {
        Write-Host ([IO.File]::ReadAllText(($statusPath + '.companion.test.log'))) -ForegroundColor Yellow
    }
    if ($errors.Count -gt 0 -and (Test-Path -LiteralPath ($statusPath + '.companion.log'))) {
        Write-Host ([IO.File]::ReadAllText(($statusPath + '.companion.log'))) -ForegroundColor Yellow
    }
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if ($errors.Count -gt 0) {
    Write-Host "`nRecorder companion tests failed: $($errors.Count)" -ForegroundColor Red
    exit 1
}
Write-Host "`nRecorder companion tests passed." -ForegroundColor Green
