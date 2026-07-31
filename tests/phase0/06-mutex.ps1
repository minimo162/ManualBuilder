# =====================================================================
# 06-mutex.ps1  —  二重起動防止の検証（V-12）
#
# ・Local\ Mutex で同一ログオンセッション内の二重起動を防げるか
# ・前のプロセスが強制終了された場合（放棄されたMutex）に次が起動できるか
#
# 使い方:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\06-mutex.ps1
#
# 内部で子プロセスを起動して挙動を確認します。手動操作は不要です。
# =====================================================================
[CmdletBinding()]
param(
    [switch]$ChildHold,      # 内部用: Mutex を取って一定時間保持する
    [switch]$ChildTry,       # 内部用: Mutex を取ろうとして結果を返す
    [int]$HoldSeconds = 8
)

$ErrorActionPreference = 'Continue'

$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$mutexName = "Local\ManualBuilder-Test-$sid"

function Get-MBMutex {
    # 戻り値: @{ Mutex = <obj>; Acquired = $bool; Abandoned = $bool }
    $created = $false
    $abandoned = $false
    $m = New-Object System.Threading.Mutex($true, $mutexName, [ref]$created)
    if (-not $created) {
        try {
            if ($m.WaitOne(0)) { $created = $true }
        } catch [System.Threading.AbandonedMutexException] {
            # 前のプロセスが ReleaseMutex せずに終了した = 強制終了された
            $created = $true
            $abandoned = $true
        }
    }
    return @{ Mutex = $m; Acquired = $created; Abandoned = $abandoned }
}

# ---------------------------------------------------------------------
# 子プロセスモード
# ---------------------------------------------------------------------
if ($ChildHold) {
    $r = Get-MBMutex
    Write-Output "CHILD_HOLD acquired=$($r.Acquired) abandoned=$($r.Abandoned) pid=$PID"
    if ($r.Acquired) {
        Start-Sleep -Seconds $HoldSeconds
        try { $r.Mutex.ReleaseMutex() } catch { }
    }
    exit 0
}
if ($ChildTry) {
    $r = Get-MBMutex
    Write-Output "CHILD_TRY acquired=$($r.Acquired) abandoned=$($r.Abandoned)"
    if ($r.Acquired) { try { $r.Mutex.ReleaseMutex() } catch { } }
    exit 0
}

# ---------------------------------------------------------------------
# 親（本体）
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  V-12: 二重起動防止（Local\ Mutex）' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Mutex 名: $mutexName" -ForegroundColor Gray
Write-Host ''

$results = New-Object System.Collections.ArrayList
function Add-R { param([string]$Id, [string]$Name, [string]$Judge, [string]$Detail = '')
    [void]$results.Add([pscustomobject]@{ 項目 = $Id; 内容 = $Name; 判定 = $Judge; 詳細 = $Detail })
    $c = switch ($Judge) { 'OK' { 'Green' } 'NG' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host ("  [{0,-4}] {1} {2} {3}" -f $Judge, $Id, $Name, $(if ($Detail) { "— $Detail" } else { '' })) -ForegroundColor $c
}

$self = $PSCommandPath
$psExe = (Get-Process -Id $PID).Path
if (-not $psExe) { $psExe = 'powershell.exe' }

# ---- 12-a: 1つ目が取得できる ----
$first = Get-MBMutex
Add-R '12-a' '1つ目が Mutex を取得' `
    $(if ($first.Acquired) { 'OK' } else { 'NG' }) `
    $(if ($first.Abandoned) { '（放棄されたMutexを引き継ぎ）' } else { '' })

# ---- 12-b: 2つ目は取得できない ----
Write-Host '  子プロセスで2つ目の取得を試します...' -ForegroundColor Gray
$out = & $psExe -NoProfile -ExecutionPolicy Bypass -File $self -ChildTry 2>&1 | Out-String
$acq2 = ($out -match 'acquired=True')
Add-R '12-b' '2つ目は取得できない' `
    $(if (-not $acq2) { 'OK' } else { 'NG' }) `
    $out.Trim()

# ---- 12-c: 解放すると取得できる ----
try { $first.Mutex.ReleaseMutex() } catch { }
$out = & $psExe -NoProfile -ExecutionPolicy Bypass -File $self -ChildTry 2>&1 | Out-String
$acq3 = ($out -match 'acquired=True')
Add-R '12-c' '解放後は取得できる' `
    $(if ($acq3) { 'OK' } else { 'NG' }) `
    $out.Trim()
try { $first.Mutex.Dispose() } catch { }

# ---- 12-d: 強制終了された Mutex を引き継げる（放棄されたMutex） ----
Write-Host ''
Write-Host '  子プロセスに Mutex を取らせ、強制終了してから引き継ぎを試します...' -ForegroundColor Gray
$proc = Start-Process -FilePath $psExe `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$self`"",
                        '-ChildHold', '-HoldSeconds', '30') `
        -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 3

# 子が取れていることを確認（3つ目は取れないはず）
$out = & $psExe -NoProfile -ExecutionPolicy Bypass -File $self -ChildTry 2>&1 | Out-String
$blocked = ($out -notmatch 'acquired=True')
Add-R '12-d1' '保持中は他が取得できない' `
    $(if ($blocked) { 'OK' } else { 'NG' }) $out.Trim()

# 強制終了（ReleaseMutex を呼ばせない）
Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$out = & $psExe -NoProfile -ExecutionPolicy Bypass -File $self -ChildTry 2>&1 | Out-String
$acq4 = ($out -match 'acquired=True')
$aband = ($out -match 'abandoned=True')
Add-R '12-d2' '強制終了後に次が起動できる' `
    $(if ($acq4) { 'OK' } else { 'NG' }) `
    $(if ($aband) { 'AbandonedMutexException を捕捉して引き継ぎました' } else { $out.Trim() })

# ---- 12-e: instance.json の鮮度判定 ----
Write-Host ''
$instDir = Join-Path $env:LOCALAPPDATA 'ManualBuilder'
New-Item -ItemType Directory -Force -Path $instDir | Out-Null
$instPath = Join-Path $instDir 'instance-test.json'

# 死んでいるPIDを書く（自分のPIDより大きい、存在しない番号を探す）
$deadPid = 999999
while (Get-Process -Id $deadPid -ErrorAction SilentlyContinue) { $deadPid-- }
@{ url = 'http://localhost:8765/'; pid = $deadPid; startedAt = (Get-Date).ToString('o') } |
    ConvertTo-Json | Out-File -LiteralPath $instPath -Encoding UTF8

$stale = $false
try {
    $inst = Get-Content -LiteralPath $instPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $alive = Get-Process -Id $inst.pid -ErrorAction SilentlyContinue
    $stale = (-not $alive)
} catch { }
Add-R '12-e' '古い instance.json を検出できる' `
    $(if ($stale) { 'OK' } else { 'NG' }) `
    "PID $deadPid は存在しない → 古い情報と判定"
Remove-Item -LiteralPath $instPath -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------
Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
$results | Format-Table -AutoSize -Wrap
$ng = @($results | Where-Object { $_.判定 -eq 'NG' }).Count
if ($ng -eq 0) {
    Write-Host '  V-12 は通りました。' -ForegroundColor Green
} else {
    Write-Host "  NG が $ng 件あります。ポート使用中の検知で代替する方針に切り替えます。" -ForegroundColor Red
}

$report = Join-Path $PSScriptRoot 'result-06-mutex.txt'
$results | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Out-File -FilePath $report -Encoding UTF8
Write-Host ''
Write-Host "  レポート: $report" -ForegroundColor Cyan
Write-Host ''
