# =====================================================================
# 00-gate-check.ps1  —  着手可否ゲート（最初にこれだけ実行してください）
#
# 検証項目:
#   V-0   LanguageMode が FullLanguage か  ← これが違えば設計そのものが成立しない
#   V-14  Windows PowerShell 5.1 のパーサで .ps1 が解析できるか
#   V-16  AppLocker / WDAC / 実行ポリシーの影響
#   V-17  ZIP由来の Mark of the Web が付いていないか
#
# 使い方（PowerShell を普通に開いて）:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\00-gate-check.ps1
#
# 所要: 30秒。Word も起動せず、システム設定も変更しません。
# 結果レポートだけを同じフォルダに保存します。
# =====================================================================
$ErrorActionPreference = 'Continue'

$fatal = @()
$warn  = @()

function Show {
    param([string]$Label, [string]$Value, [string]$Judge = 'INFO', [string]$Note = '')
    $c = switch ($Judge) { 'OK' { 'Green' } 'FATAL' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host ('  {0,-34} {1}' -f ($Label + ':'), $Value) -ForegroundColor $c
    if ($Note) { Write-Host ('  {0,-34} → {1}' -f '', $Note) -ForegroundColor DarkGray }
    if ($Judge -eq 'FATAL') { $script:fatal += "$($Label.Trim()) = $Value" }
    if ($Judge -eq 'WARN')  { $script:warn  += "$($Label.Trim()) = $Value" }
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  ManualBuilder 着手可否ゲート (V-0 / V-14 / V-16 / V-17)' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''

# =====================================================================
# V-0  LanguageMode  ← 最重要
# =====================================================================
Write-Host '[V-0] 言語モード（これが FullLanguage でなければ設計が成立しません）' -ForegroundColor Cyan

$lm = 'unknown'
try { $lm = [string]$ExecutionContext.SessionState.LanguageMode } catch { }
Show 'LanguageMode' $lm $(if ($lm -eq 'FullLanguage') { 'OK' } else { 'FATAL' }) `
    $(if ($lm -ne 'FullLanguage') { 'AppLocker / WDAC が有効です。PowerShell から .NET / COM が使えません' } else { '' })

# 実際に必要な機能を1つずつ試す（モード名だけでは判断しきれないため）
Write-Host ''
Write-Host '  実際に必要な機能を試します:' -ForegroundColor Gray

# (a) 任意の .NET 型の New-Object
$okNewObject = $false
try {
    $null = New-Object System.Collections.ArrayList
    $okNewObject = $true
} catch { }
Show '  New-Object（任意の.NET型）' $(if ($okNewObject) { '使える' } else { '使えない' }) `
    $(if ($okNewObject) { 'OK' } else { 'FATAL' })

# (b) HttpListener
$okListener = $false
try {
    $l = New-Object System.Net.HttpListener
    $l.Close()
    $okListener = $true
} catch { }
Show '  HttpListener の生成' $(if ($okListener) { '使える' } else { '使えない' }) `
    $(if ($okListener) { 'OK' } else { 'FATAL' }) `
    $(if (-not $okListener) { 'サーバを起動できません' } else { '' })

# (c) System.Drawing（スクショ・サムネイル）
$okDrawing = $false
try {
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $b = New-Object System.Drawing.Bitmap 2, 2
    $b.Dispose()
    $okDrawing = $true
} catch { }
Show '  System.Drawing.Bitmap' $(if ($okDrawing) { '使える' } else { '使えない' }) `
    $(if ($okDrawing) { 'OK' } else { 'FATAL' }) `
    $(if (-not $okDrawing) { 'スクショの取り込み・サムネイル生成ができません' } else { '' })

# (d) Add-Type によるC#コンパイル（DPI設定・Win32 API）
$okAddType = $false
try {
    Add-Type -Namespace MBGate -Name T -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern uint GetCurrentProcessId();
'@ -ErrorAction Stop
    $null = [MBGate.T]::GetCurrentProcessId()
    $okAddType = $true
} catch { }
Show '  Add-Type（C#／Win32 API）' $(if ($okAddType) { '使える' } else { '使えない' }) `
    $(if ($okAddType) { 'OK' } else { 'FATAL' }) `
    $(if (-not $okAddType) { 'DPI設定とWordプロセスの特定ができません' } else { '' })

# (e) COM（Excel / Word）— 実際には起動せず、型解決だけ試す
# Excelは主出力なので必須。Wordは副出力なので利用できなくても着手を妨げない。
$okExcelComType = $false
try {
    if ([Type]::GetTypeFromProgID('Excel.Application')) { $okExcelComType = $true }
} catch { }
Show '  Excel.Application の型解決' $(if ($okExcelComType) { '解決できる' } else { '解決できない' }) `
    $(if ($okExcelComType) { 'OK' } else { 'FATAL' }) `
    $(if (-not $okExcelComType) { '主出力のExcelを作成できません' } else { '' })

$okWordComType = $false
try {
    if ([Type]::GetTypeFromProgID('Word.Application')) { $okWordComType = $true }
} catch { }
Show '  Word.Application の型解決' $(if ($okWordComType) { '解決できる' } else { '解決できない' }) `
    $(if ($okWordComType) { 'OK' } else { 'WARN' }) `
    $(if (-not $okWordComType) { 'Word副出力は無効になります。Excel主出力は利用できます' } else { '' })

# (g) FileSystemWatcher（監視）
$okWatcher = $false
try {
    $w = New-Object System.IO.FileSystemWatcher
    $w.Dispose()
    $okWatcher = $true
} catch { }
Show '  FileSystemWatcher' $(if ($okWatcher) { '使える' } else { '使えない' }) `
    $(if ($okWatcher) { 'OK' } else { 'WARN' }) `
    $(if (-not $okWatcher) { '撮影の自動取り込みができません（貼り付けのみ）' } else { '' })

# (h) Mutex（二重起動防止）
$okMutex = $false
try {
    $created = $false
    $m = New-Object System.Threading.Mutex($true, 'Local\ManualBuilder-GateTest', [ref]$created)
    if ($created) { $m.ReleaseMutex() }
    $m.Dispose()
    $okMutex = $true
} catch { }
Show '  Mutex' $(if ($okMutex) { '使える' } else { '使えない' }) `
    $(if ($okMutex) { 'OK' } else { 'WARN' })

# =====================================================================
# V-16  アプリケーション制御ポリシー
# =====================================================================
Write-Host ''
Write-Host '[V-16] アプリケーション制御ポリシー' -ForegroundColor Cyan

Show 'PowerShell バージョン' "$($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" `
    $(if ($PSVersionTable.PSVersion.Major -ge 5) { 'OK' } else { 'FATAL' })

Show '実行ポリシー（現行スコープ）' "$(Get-ExecutionPolicy)" 'INFO' '-ExecutionPolicy Bypass で回避可'
try {
    $list = Get-ExecutionPolicy -List | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }
    Write-Host ('  {0,-34} {1}' -f '実行ポリシー（全スコープ）:', ($list -join ' / ')) -ForegroundColor Gray
} catch { }

# AppLocker の有効性
$appLocker = '検出なし'
try {
    $svc = Get-Service -Name AppIDSvc -ErrorAction Stop
    $appLocker = "AppIDSvc = $($svc.Status) / 起動種別 $($svc.StartType)"
} catch { $appLocker = 'AppIDSvc が見つかりません（＝AppLocker未使用の可能性）' }
Show 'AppLocker サービス' $appLocker 'INFO'

try {
    $pol = Get-ChildItem 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2' -ErrorAction Stop
    Show 'AppLocker ポリシー' "存在します（$($pol.Count) 個のコレクション）" 'WARN' `
        'AppLocker のルールが構成されています。V-0 の結果と併せて判断してください'
} catch {
    Show 'AppLocker ポリシー' '未構成' 'OK'
}

# WDAC / Device Guard
try {
    $dg = Get-CimInstance -ClassName Win32_DeviceGuard `
            -Namespace 'root\Microsoft\Windows\DeviceGuard' -ErrorAction Stop
    $running = @($dg.SecurityServicesRunning)
    Show 'Device Guard 稼働中サービス' $(if ($running.Count) { ($running -join ',') } else { 'なし' }) 'INFO'
    # 注: この値が 2（強制）でも、それがカーネルモード（ドライバのブロックリスト等）
    #     だけを対象にしたポリシーなら、PowerShell は FullLanguage のままになります。
    #     ConstrainedLanguage になるのは、ユーザーモードのスクリプト実行（UMCI）を
    #     対象にしたポリシーが適用されている場合です。
    #     したがって最終判断は V-0 の実機テスト（上の8項目）で行います。
    $ciNote = '0=なし 1=監査 2=強制。ただしカーネルモードのみ対象なら PowerShell には影響しません'
    if ([int]$dg.CodeIntegrityPolicyEnforcementStatus -ge 2 -and $lm -eq 'FullLanguage') {
        $ciNote = '強制されていますが、上の8項目を個別に実測できているため各判定結果を優先します'
    }
    Show 'コード整合性ポリシー' "$($dg.CodeIntegrityPolicyEnforcementStatus)" `
        $(if ([int]$dg.CodeIntegrityPolicyEnforcementStatus -ge 2 -and $lm -ne 'FullLanguage') { 'WARN' } else { 'INFO' }) `
        $ciNote
} catch {
    Show 'Device Guard / WDAC' '情報取得不可' 'INFO' 'WMIクラスが無い環境（＝未構成の可能性）'
}

# =====================================================================
# V-17  Mark of the Web
# =====================================================================
Write-Host ''
Write-Host '[V-17] Mark of the Web（ZIP展開由来のブロック）' -ForegroundColor Cyan

$blocked = @()
foreach ($f in @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.ps1 -ErrorAction SilentlyContinue)) {
    try {
        $z = Get-Content -LiteralPath $f.FullName -Stream Zone.Identifier -ErrorAction Stop
        if ($z) { $blocked += $f.Name }
    } catch { }
}
if ($blocked.Count -gt 0) {
    Show 'Zone.Identifier が付いたファイル' "$($blocked.Count) 個" 'WARN' `
        '配布元を確認できた場合だけ、最後の確認で明示的に解除できます'
} else {
    Show 'Zone.Identifier' '付いていません' 'OK'
}

# =====================================================================
# V-14  Windows PowerShell 5.1 のパーサで構文解析
# =====================================================================
Write-Host ''
Write-Host '[V-14] このPowerShellのパーサで .ps1 を構文解析' -ForegroundColor Cyan

$files = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.ps1 -ErrorAction SilentlyContinue)
if ($files.Count -eq 0) {
    Write-Host '  同じフォルダに .ps1 がありません（このファイルのみ）' -ForegroundColor Gray
} else {
    $bad = 0
    foreach ($f in $files) {
        $errs = $null; $toks = $null
        try {
            [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$toks, [ref]$errs)
        } catch {
            Write-Host "  解析できません $($f.Name): $($_.Exception.Message)" -ForegroundColor Red
            $bad++
            continue
        }
        if ($errs -and $errs.Count -gt 0) {
            $bad++
            Write-Host "  NG $($f.Name)  エラー $($errs.Count) 件" -ForegroundColor Red
            $errs | Select-Object -First 5 | ForEach-Object {
                Write-Host ('     L{0}:{1} {2}' -f $_.Extent.StartLineNumber, $_.Extent.StartColumnNumber, $_.Message) -ForegroundColor Red
            }
        } else {
            Write-Host "  OK $($f.Name)" -ForegroundColor Green
        }
    }
    Show '構文エラーのあるファイル' "$bad 個" $(if ($bad -eq 0) { 'OK' } else { 'FATAL' })
}

# BOM の確認
$noBom = @()
foreach ($f in $files) {
    try {
        $fs = [IO.File]::OpenRead($f.FullName)
        $b = New-Object byte[] 3
        [void]$fs.Read($b, 0, 3)
        $fs.Close()
        if (-not ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)) { $noBom += $f.Name }
    } catch { }
}
if ($files.Count -gt 0) {
    Show 'BOMなしの .ps1' $(if ($noBom.Count) { "$($noBom.Count) 個: $($noBom -join ', ')" } else { 'なし' }) `
        $(if ($noBom.Count) { 'WARN' } else { 'OK' }) `
        $(if ($noBom.Count) { 'PowerShell 5.1 は BOM なし UTF-8 の日本語を化けさせます' } else { '' })
}

# 日本語表示の確認
Write-Host ''
Write-Host '  文字化けの確認 → 次の行が正しく読めますか？' -ForegroundColor Cyan
Write-Host '  「経費精算システム 操作マニュアル ①②③ ⚠」' -ForegroundColor White

# =====================================================================
# 判定
# =====================================================================
Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
if ($fatal.Count -eq 0) {
    Write-Host '  判定: 着手できます' -ForegroundColor Green
    Write-Host '======================================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  次の手順:' -ForegroundColor Yellow
    Write-Host '   1. 01-env.ps1でExcel COMの起動と終了を確認'
    Write-Host '   2. 改訂版の検証キットでV-4（Word副出力の安全性）を確認'
    Write-Host '   3. Excelレイアウト検証を追加してから製品実装へ進む'
    if (-not $okWordComType) {
        Write-Host ''
        Write-Host '  Word副出力は利用できません。Excel主出力の構成で進めます。' -ForegroundColor Yellow
    }
} else {
    Write-Host '  判定: この設計は成立しません' -ForegroundColor Red
    Write-Host '======================================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  致命的な項目:' -ForegroundColor Red
    $fatal | ForEach-Object { Write-Host "   - $_" -ForegroundColor Red }
    Write-Host ''
    if ($lm -ne 'FullLanguage') {
        Write-Host '  原因は言語モードです。AppLocker または WDAC により PowerShell が' -ForegroundColor Yellow
        Write-Host '  ConstrainedLanguage で動いており、.NET と COM を使えません。' -ForegroundColor Yellow
        Write-Host '  これは組織のポリシーなので、利用者側では変更できません。' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  代替案:' -ForegroundColor Yellow
        Write-Host '   A. ブラウザだけで完結する構成にする' -ForegroundColor Yellow
        Write-Host '      （単一HTML＋貼り付け取り込み＋HTML/PDF出力。監視とWord出力は不可）'
        Write-Host '   B. 既製品を導入する（Folge は $89 買い切りで Word 出力と連番バッジに対応）' -ForegroundColor Yellow
    } else {
        Write-Host '  言語モードは FullLanguage なので、原因はポリシーではありません。' -ForegroundColor Yellow
        Write-Host '  次を確認してください:' -ForegroundColor Yellow
        Write-Host '   ・Windows 上の Windows PowerShell 5.1 で実行しているか'
        Write-Host '     （PowerShell 7 / Linux / macOS では System.Drawing や Win32 API が使えません）'
        Write-Host '   ・.NET Framework 4.7.2 以降が入っているか'
        Write-Host '   ・セキュリティ製品が Add-Type のコンパイルを妨げていないか'
        Write-Host '     （%TEMP% への書き込みが必要です）'
    }
}

if ($warn.Count -gt 0) {
    Write-Host ''
    Write-Host '  警告:' -ForegroundColor Yellow
    $warn | ForEach-Object { Write-Host "   - $_" -ForegroundColor Yellow }
}

# レポート保存
$report = Join-Path $PSScriptRoot 'result-00-gate.txt'
$lines = @(
    "ManualBuilder 着手可否ゲート  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "マシン: $env:COMPUTERNAME / ユーザー: $env:USERNAME"
    "PowerShell: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    "LanguageMode: $lm"
    "New-Object: $okNewObject / HttpListener: $okListener / Drawing: $okDrawing"
    "Add-Type: $okAddType / ExcelCOM型: $okExcelComType / WordCOM型: $okWordComType"
    "Watcher: $okWatcher / Mutex: $okMutex"
    "実行ポリシー: $(Get-ExecutionPolicy)"
    "致命的: $(if ($fatal.Count) { $fatal -join ' | ' } else { 'なし' })"
    "警告: $(if ($warn.Count) { $warn -join ' | ' } else { 'なし' })"
)
$lines | Out-File -FilePath $report -Encoding UTF8
Write-Host ''
Write-Host "  レポート: $report" -ForegroundColor Cyan

# MOTW の解除を提案（安全側の既定値は「解除しない」）
if ($blocked.Count -gt 0) {
    Write-Host ''
    Write-Host '  配布元と内容を確認できた場合に限り、ブロックを解除してください。' -ForegroundColor Yellow
    Write-Host "  Zone.Identifier が付いたファイル $($blocked.Count) 個を解除しますか？ [y/N] " -NoNewline -ForegroundColor Yellow
    if ((Read-Host) -eq 'y') {
        Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.ps1 | Unblock-File -ErrorAction SilentlyContinue
        Write-Host '  解除しました。' -ForegroundColor Green
    } else {
        Write-Host '  解除しませんでした。' -ForegroundColor Gray
    }
}
Write-Host ''

if ($fatal.Count -gt 0) {
    exit 1
}
exit 0
