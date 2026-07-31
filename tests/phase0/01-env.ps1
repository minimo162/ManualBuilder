# =====================================================================
# 00-check-env.ps1  —  実行環境の確認
# 使い方: powershell -ExecutionPolicy Bypass -File .\00-check-env.ps1
# =====================================================================
$ErrorActionPreference = 'Continue'

$results = New-Object System.Collections.ArrayList
function Add-Result {
    param([string]$Item, [string]$Value, [string]$Judge, [string]$Note = '')
    [void]$results.Add([pscustomobject]@{
        項目 = $Item; 値 = $Value; 判定 = $Judge; 備考 = $Note
    })
}

Write-Host ''
Write-Host '=== ManualBuilder 環境チェック ===' -ForegroundColor Cyan
Write-Host ''

# --- PowerShell -------------------------------------------------------
$psv = $PSVersionTable.PSVersion.ToString()
Add-Result 'PowerShell バージョン' $psv $(if ($PSVersionTable.PSVersion.Major -ge 5) { 'OK' } else { 'NG' }) 'Windows PowerShell 5.1 を想定'
Add-Result 'エディション' "$($PSVersionTable.PSEdition)" 'INFO' 'Desktop=5.1 / Core=7'
Add-Result 'アパートメント状態' "$([System.Threading.Thread]::CurrentThread.GetApartmentState())" 'INFO' 'COM を扱うなら STA が望ましい'

# --- OS ---------------------------------------------------------------
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $isWin11 = [int]$os.BuildNumber -ge 22000
    $judge = if ($isWin11) { 'OK' } else { 'WARN' }
    $note  = if ($isWin11) { 'Win+Shift+S の自動保存が使える' } else { 'Windows 10 は監視経路が使えない可能性' }
    Add-Result 'OS' "$($os.Caption) (Build $($os.BuildNumber))" $judge $note
} catch {
    Add-Result 'OS' '取得失敗' 'WARN' $_.Exception.Message
}

# --- Word / Excel -----------------------------------------------------
foreach ($app in @('Word', 'Excel')) {
    $found = $false
    $ver = ''
    try {
        if ([Type]::GetTypeFromProgID("$app.Application")) { $found = $true }
    } catch { }
    if ($found) {
        try {
            $com = New-Object -ComObject "$app.Application"
            $ver = "$($com.Version)"
            $com.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($com)
            Remove-Variable com -ErrorAction SilentlyContinue
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        } catch {
            $ver = '起動失敗: ' + $_.Exception.Message
        }
    }
    $val   = if ($found) { "利用可能 (Ver $ver)" } else { '見つかりません' }
    $judge = if ($found) { 'OK' } elseif ($app -eq 'Excel') { 'NG' } else { 'WARN' }
    $note  = if ($app -eq 'Excel') { 'Excel主出力に必須' } else { 'Word副出力に必要' }
    Add-Result "$app COM" $val $judge $note
}

# --- 起動中の Office プロセス -----------------------------------------
foreach ($p in @('WINWORD', 'EXCEL')) {
    $procs = @(Get-Process -Name $p -ErrorAction SilentlyContinue)
    $judge = if ($procs.Count -eq 0) { 'OK' } else { 'WARN' }
    $note  = if ($procs.Count -gt 0) { '出力テストの前に閉じることを推奨' } else { '' }
    Add-Result "$p プロセス" "$($procs.Count) 個" $judge $note
}

# --- スクリーンショット自動保存フォルダ -------------------------------
$shotDir = Join-Path $env:USERPROFILE 'Pictures\Screenshots'
$exists = Test-Path -LiteralPath $shotDir
$count = 0
if ($exists) {
    $count = @(Get-ChildItem -LiteralPath $shotDir -Filter *.png -ErrorAction SilentlyContinue).Count
}
$val   = if ($exists) { "$shotDir (PNG $count 件)" } else { "$shotDir が存在しません" }
$judge = if ($exists) { 'OK' } else { 'WARN' }
$note  = if ($exists) { '監視経路が使える見込み' } else { 'Snipping Tool の自動保存を有効にしてください' }
Add-Result '自動保存フォルダ' $val $judge $note

# --- Snipping Tool ----------------------------------------------------
try {
    $pkg = Get-AppxPackage -Name 'Microsoft.ScreenSketch' -ErrorAction Stop
    Add-Result 'Snipping Tool' "$($pkg.Version)" 'OK' ''
} catch {
    Add-Result 'Snipping Tool' '情報取得不可' 'INFO' 'Get-AppxPackage が使えない環境'
}

# --- ディスプレイ / DPI ------------------------------------------------
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop

    $screens = [System.Windows.Forms.Screen]::AllScreens
    Add-Result 'ディスプレイ数' "$($screens.Count) 台" 'INFO' ''
    for ($i = 0; $i -lt $screens.Count; $i++) {
        $b = $screens[$i].Bounds
        $note = if ($screens[$i].Primary) { 'プライマリ' } else { '' }
        Add-Result "  ディスプレイ[$i]" "$($b.Width)x$($b.Height) @($($b.X),$($b.Y))" 'INFO' $note
    }
    $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
    Add-Result '仮想画面全体' "$($vs.Width)x$($vs.Height)" 'INFO' 'DPI未設定状態の報告値'

    $bmp = New-Object System.Drawing.Bitmap 1, 1
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $dpi = [int]$g.DpiX
    $judge = if ($dpi -eq 96) { 'OK' } else { 'WARN' }
    $note  = if ($dpi -eq 96) { '拡大なし(100%)' } else { '拡大表示あり。03-capture-dpi.ps1 で要確認' }
    Add-Result '報告 DPI' "$dpi x $([int]$g.DpiY)" $judge $note
    $g.Dispose(); $bmp.Dispose()
} catch {
    Add-Result 'ディスプレイ情報' '取得失敗' 'NG' $_.Exception.Message
}

# --- ポートの空き -----------------------------------------------------
$port = 8765
$busy = $false
try {
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $port)
    $l.Start(); $l.Stop()
} catch { $busy = $true }
$judge = if ($busy) { 'WARN' } else { 'OK' }
$note  = if ($busy) { '02-server-rawpost.ps1 は -Port で変更してください' } else { '' }
Add-Result "ポート $port" $(if ($busy) { '使用中' } else { '空き' }) $judge $note

# --- HttpListener が非管理者で使えるか ---------------------------------
try {
    $hl = New-Object System.Net.HttpListener
    $hl.Prefixes.Add("http://localhost:$($port + 1)/")
    $hl.Start(); $hl.Stop(); $hl.Close()
    Add-Result 'HttpListener (localhost)' '起動できました' 'OK' '管理者権限は不要'
} catch {
    Add-Result 'HttpListener (localhost)' '起動失敗' 'NG' $_.Exception.Message
}

# --- 管理者権限 -------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Add-Result '管理者権限' $(if ($isAdmin) { 'あり' } else { 'なし' }) 'INFO' 'なしで動くのが要件'

# --- ディスク空き -----------------------------------------------------
try {
    $drive = (Get-Item -LiteralPath $PSScriptRoot).PSDrive.Name
    $free = (Get-PSDrive $drive).Free / 1GB
    Add-Result "空き容量 ($drive`:)" ('{0:N1} GB' -f $free) $(if ($free -gt 1) { 'OK' } else { 'WARN' }) ''
} catch { }

Add-Result '実行ポリシー' "$(Get-ExecutionPolicy)" 'INFO' '-ExecutionPolicy Bypass で回避可'

# --- 出力 -------------------------------------------------------------
Write-Host ''
$results | Format-Table -AutoSize -Wrap
Write-Host ''

$ng   = @($results | Where-Object { $_.判定 -eq 'NG' })
$warn = @($results | Where-Object { $_.判定 -eq 'WARN' })
if ($ng.Count -gt 0) {
    Write-Host "NG が $($ng.Count) 件あります。先に解決してください:" -ForegroundColor Red
    $ng | ForEach-Object { Write-Host "  - $($_.項目): $($_.値) / $($_.備考)" -ForegroundColor Red }
} else {
    Write-Host 'NG はありません。' -ForegroundColor Green
}
if ($warn.Count -gt 0) {
    Write-Host "WARN が $($warn.Count) 件:" -ForegroundColor Yellow
    $warn | ForEach-Object { Write-Host "  - $($_.項目): $($_.値) / $($_.備考)" -ForegroundColor Yellow }
}

$out = Join-Path $PSScriptRoot 'result-00-env.txt'
$results | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Out-File -FilePath $out -Encoding UTF8
Write-Host ''
Write-Host "結果を保存しました: $out" -ForegroundColor Cyan
Write-Host ''
