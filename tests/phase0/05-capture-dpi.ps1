# =====================================================================
# 03-capture-dpi.ps1  —  高DPI / マルチモニタでのキャプチャ検証
#
# 検証項目:
#   V-6  DPI awareness 設定後に原寸ピクセルでキャプチャできるか
#   要検証#3  powershell.exe の既定 DPI awareness
#   要検証#4  混在DPIマルチモニタでの座標
#
# 重要: このスクリプトは 2回に分けて動作を比較します。
#   (a) DPI未設定のまま  → 別プロセスを起動して取得
#   (b) DPI設定あり      → このプロセスで取得
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File .\03-capture-dpi.ps1
#   powershell -ExecutionPolicy Bypass -File .\03-capture-dpi.ps1 -Delay 5
# =====================================================================
[CmdletBinding()]
param(
    [int]$Delay = 0,
    [switch]$NoDpiAware   # 内部用: DPI設定をスキップ（比較のため子プロセスで使う）
)

$ErrorActionPreference = 'Stop'
$outDir = Join-Path $PSScriptRoot 'out\capture'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# =====================================================================
# 1. DPI awareness の設定（System.Windows.Forms を読む「前」に行う）
# =====================================================================
$dpiResult = 'スキップ'
if (-not $NoDpiAware) {
    try {
        Add-Type -Namespace MB -Name Dpi -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
'@ -ErrorAction Stop
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        $ok = [MB.Dpi]::SetProcessDpiAwarenessContext([IntPtr](-4))
        if ($ok) {
            $dpiResult = '成功 (PER_MONITOR_AWARE_V2)'
        } else {
            $err = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            # 5 = ERROR_ACCESS_DENIED（既にマニフェスト等で設定済み）
            $dpiResult = "失敗 (Win32Error=$err)"
            if ($err -eq 5) { $dpiResult += ' — 既にDPI対応として設定済みの可能性' }
        }
    } catch {
        $dpiResult = '失敗: ' + $_.Exception.Message
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host ''
Write-Host '=== V-6: キャプチャ / DPI 検証 ===' -ForegroundColor Cyan
Write-Host ''
Write-Host "DPI awareness 設定: $dpiResult" -ForegroundColor $(if ($dpiResult -like '成功*') { 'Green' } elseif ($NoDpiAware) { 'Gray' } else { 'Yellow' })

# =====================================================================
# 2. 画面情報
# =====================================================================
$screens = [System.Windows.Forms.Screen]::AllScreens
$vs = [System.Windows.Forms.SystemInformation]::VirtualScreen

Write-Host ''
Write-Host "ディスプレイ数: $($screens.Count)"
for ($i = 0; $i -lt $screens.Count; $i++) {
    $b = $screens[$i].Bounds
    $tag = if ($screens[$i].Primary) { ' [プライマリ]' } else { '' }
    Write-Host ("  [{0}] {1}x{2} 位置({3},{4}){5}" -f $i, $b.Width, $b.Height, $b.X, $b.Y, $tag)
}
Write-Host ("仮想画面全体: {0}x{1} 位置({2},{3})" -f $vs.Width, $vs.Height, $vs.X, $vs.Y)

$bmp1 = New-Object System.Drawing.Bitmap 1, 1
$g1 = [System.Drawing.Graphics]::FromImage($bmp1)
$reportedDpi = [int]$g1.DpiX
$g1.Dispose(); $bmp1.Dispose()
Write-Host "報告DPI: $reportedDpi (96 = 拡大なし / 120 = 125% / 144 = 150%)"

# 実際の物理解像度をレジストリ相当から取得（プライマリのみ）
try {
    Add-Type -Namespace MB -Name Dev -MemberDefinition @'
[DllImport("user32.dll")]
public static extern int GetSystemMetrics(int nIndex);
'@ -ErrorAction Stop
    $smW = [MB.Dev]::GetSystemMetrics(0)   # SM_CXSCREEN
    $smH = [MB.Dev]::GetSystemMetrics(1)   # SM_CYSCREEN
    Write-Host "GetSystemMetrics プライマリ: ${smW}x${smH}"
} catch { }

# =====================================================================
# 3. キャプチャ
# =====================================================================
if ($Delay -gt 0) {
    Write-Host ''
    for ($s = $Delay; $s -ge 1; $s--) {
        Write-Host "  $s 秒後にキャプチャします..." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
}

function Save-Capture {
    param([int]$X, [int]$Y, [int]$W, [int]$H, [string]$Path)
    $bmp = New-Object System.Drawing.Bitmap $W, $H
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.CopyFromScreen($X, $Y, 0, 0, (New-Object System.Drawing.Size $W, $H))
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        return @{ W = $bmp.Width; H = $bmp.Height; Bytes = (Get-Item -LiteralPath $Path).Length }
    } finally {
        # Dispose を忘れるとPNGがロックされたまま残る
        $g.Dispose(); $bmp.Dispose()
    }
}

$suffix = if ($NoDpiAware) { 'nodpi' } else { 'dpi' }
$results = New-Object System.Collections.ArrayList

# (1) 仮想画面全体を一括
$p = Join-Path $outDir "virtual-$suffix.png"
$r = Save-Capture -X $vs.X -Y $vs.Y -W $vs.Width -H $vs.Height -Path $p
[void]$results.Add([pscustomobject]@{
    対象 = '仮想画面 一括'; 指定 = "$($vs.Width)x$($vs.Height)"
    実画像 = "$($r.W)x$($r.H)"; サイズKB = [int]($r.Bytes / 1KB); ファイル = (Split-Path $p -Leaf)
})

# (2) モニタごと（推奨方式）
for ($i = 0; $i -lt $screens.Count; $i++) {
    $b = $screens[$i].Bounds
    $p = Join-Path $outDir "monitor$i-$suffix.png"
    $r = Save-Capture -X $b.X -Y $b.Y -W $b.Width -H $b.Height -Path $p
    [void]$results.Add([pscustomobject]@{
        対象 = "モニタ[$i] 個別"; 指定 = "$($b.Width)x$($b.Height)"
        実画像 = "$($r.W)x$($r.H)"; サイズKB = [int]($r.Bytes / 1KB); ファイル = (Split-Path $p -Leaf)
    })
}

# (3) アクティブウィンドウ（DwmGetWindowAttribute で正しい矩形を取る）
try {
    Add-Type -Namespace MB -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
[DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr hwnd, int attr, out RECT r, int size);
public struct RECT { public int Left, Top, Right, Bottom; }
'@ -ErrorAction Stop

    $h = [MB.Win]::GetForegroundWindow()
    $rcWin = New-Object 'MB.Win+RECT'
    $rcDwm = New-Object 'MB.Win+RECT'
    $okWin = [MB.Win]::GetWindowRect($h, [ref]$rcWin)
    $hr = [MB.Win]::DwmGetWindowAttribute($h, 9, [ref]$rcDwm, 16)   # DWMWA_EXTENDED_FRAME_BOUNDS = 9

    Write-Host ''
    Write-Host 'アクティブウィンドウの矩形比較:' -ForegroundColor Cyan
    if ($okWin) {
        Write-Host ("  GetWindowRect            : {0},{1} - {2},{3}  ({4}x{5})" -f `
            $rcWin.Left, $rcWin.Top, $rcWin.Right, $rcWin.Bottom, ($rcWin.Right - $rcWin.Left), ($rcWin.Bottom - $rcWin.Top))
    }
    if ($hr -eq 0) {
        Write-Host ("  DwmGetWindowAttribute(9) : {0},{1} - {2},{3}  ({4}x{5})" -f `
            $rcDwm.Left, $rcDwm.Top, $rcDwm.Right, $rcDwm.Bottom, ($rcDwm.Right - $rcDwm.Left), ($rcDwm.Bottom - $rcDwm.Top))
        $diff = ($rcWin.Right - $rcWin.Left) - ($rcDwm.Right - $rcDwm.Left)
        Write-Host ("  → 幅の差: {0}px（GetWindowRect は不可視の余白を含む）" -f $diff) -ForegroundColor Yellow

        $w = $rcDwm.Right - $rcDwm.Left
        $hh = $rcDwm.Bottom - $rcDwm.Top
        if ($w -gt 0 -and $hh -gt 0) {
            $p = Join-Path $outDir "activewindow-$suffix.png"
            $r = Save-Capture -X $rcDwm.Left -Y $rcDwm.Top -W $w -H $hh -Path $p
            [void]$results.Add([pscustomobject]@{
                対象 = 'アクティブウィンドウ'; 指定 = "${w}x${hh}"
                実画像 = "$($r.W)x$($r.H)"; サイズKB = [int]($r.Bytes / 1KB); ファイル = (Split-Path $p -Leaf)
            })
        }
    } else {
        Write-Host "  DwmGetWindowAttribute 失敗 (HRESULT=$hr)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "アクティブウィンドウ取得でエラー: $($_.Exception.Message)" -ForegroundColor Yellow
}

# =====================================================================
# 4. 結果
# =====================================================================
Write-Host ''
Write-Host '=== キャプチャ結果 ===' -ForegroundColor Cyan
$results | Format-Table -AutoSize

# =====================================================================
# 5. DPI未設定版との比較（親プロセスのみ実行）
# =====================================================================
if (-not $NoDpiAware) {
    Write-Host ''
    Write-Host '[比較] DPI未設定の子プロセスで同じ処理を実行します...' -ForegroundColor Cyan
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-NoDpiAware')
    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Wait -PassThru -WindowStyle Hidden
    Write-Host "  子プロセス終了コード: $($proc.ExitCode)"

    Write-Host ''
    Write-Host '=== DPI設定あり／なし の比較 ===' -ForegroundColor Cyan
    $cmp = New-Object System.Collections.ArrayList
    foreach ($f in @(Get-ChildItem -LiteralPath $outDir -Filter '*-dpi.png')) {
        $base = $f.Name -replace '-dpi\.png$', ''
        $noF = Join-Path $outDir "$base-nodpi.png"
        $i1 = [System.Drawing.Image]::FromFile($f.FullName)
        $s1 = "$($i1.Width)x$($i1.Height)"; $i1.Dispose()
        $s2 = '(なし)'
        if (Test-Path -LiteralPath $noF) {
            $i2 = [System.Drawing.Image]::FromFile($noF)
            $s2 = "$($i2.Width)x$($i2.Height)"; $i2.Dispose()
        }
        [void]$cmp.Add([pscustomobject]@{
            対象 = $base; DPI設定あり = $s1; DPI設定なし = $s2
            判定 = if ($s1 -eq $s2) { '同一' } else { '差あり → DPI設定が効いている' }
        })
    }
    $cmp | Format-Table -AutoSize

    $diffFound = @($cmp | Where-Object { $_.判定 -ne '同一' }).Count
    Write-Host ''
    if ($reportedDpi -eq 96) {
        Write-Host '拡大表示なし(100%)の環境なので、差が出ないのが正常です。' -ForegroundColor Green
        Write-Host '拡大表示のある環境（125%など）でも一度実行してください。' -ForegroundColor Yellow
    } elseif ($diffFound -gt 0) {
        Write-Host 'DPI設定によって取得サイズが変わりました。設定は必須です（要件定義書の通り）。' -ForegroundColor Green
    } else {
        Write-Host '拡大表示があるのに差が出ませんでした。powershell.exe が既にDPI対応の可能性があります。' -ForegroundColor Yellow
    }

    $report = Join-Path $outDir 'result-03-capture.txt'
    ($results | Format-Table -AutoSize | Out-String -Width 200) +
    "`r`nDPI設定: $dpiResult`r`n報告DPI: $reportedDpi`r`n`r`n" +
    ($cmp | Format-Table -AutoSize | Out-String -Width 200) |
        Out-File -FilePath $report -Encoding UTF8
    Write-Host ''
    Write-Host "レポート: $report" -ForegroundColor Cyan
    Write-Host "画像    : $outDir" -ForegroundColor Cyan
    Write-Host ''
    Write-Host '画像を開いて、文字がボケていないか（原寸で取れているか）を確認してください。' -ForegroundColor Yellow
    Write-Host 'フォルダを開きますか？ [Y/n] ' -NoNewline -ForegroundColor Yellow
    if ((Read-Host) -ne 'n') { Start-Process explorer.exe $outDir }
}
