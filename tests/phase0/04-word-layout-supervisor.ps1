# =====================================================================
# 04-word-layout-supervisor.ps1 — Word検証の監督・タイムアウト
#
# Word COMがSaveAs2等で戻らなくなっても、メニュー全体を固めないため、
# 04-word-layout.ps1を別PowerShellプロセスで1回ずつ実行する。
# タイムアウト時は子PowerShellだけを終了し、WINWORDは自動終了しない。
# =====================================================================
[CmdletBinding()]
param(
    [ValidateRange(1, 10)][int]$Runs = 1,
    [ValidateRange(20, 600)][int]$TimeoutSeconds = 90,
    [int]$CancelAtStep = 0,
    [switch]$SimulateExisting
)

$ErrorActionPreference = 'Stop'
$worker = Join-Path $PSScriptRoot '04-word-layout.ps1'
$outDir = Join-Path $PSScriptRoot 'out'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$tmpDir = Join-Path $outDir '.tmp'
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
$reportPath = Join-Path $outDir 'result-04-supervisor.txt'
$report = New-Object System.Collections.ArrayList

function Get-WordPids {
    return @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty Id)
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  Word COM 監督実行' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host "  1回あたりのタイムアウト: $TimeoutSeconds 秒" -ForegroundColor Gray

$initialPids = @(Get-WordPids)
if (-not $SimulateExisting -and $Runs -gt 1 -and $initialPids.Count -gt 0) {
    Write-Host ''
    Write-Host "  Word が起動中です（PID: $($initialPids -join ',')）。" -ForegroundColor Yellow
    Write-Host '  10回試験は開始しません。Wordを閉じてから再実行してください。' -ForegroundColor Yellow
    exit 0
}

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

for ($run = 1; $run -le $Runs; $run++) {
    Write-Host ''
    Write-Host "  === 監督ラウンド $run / $Runs ===" -ForegroundColor Cyan
    $pidsBefore = @(Get-WordPids)
    $statusPath = Join-Path $tmpDir ("supervisor-status-{0}-{1}.txt" -f $PID, $run)
    if (Test-Path -LiteralPath $statusPath) {
        Remove-Item -LiteralPath $statusPath -Force
    }

    $quotedWorker = '"' + $worker + '"'
    $quotedStatus = '"' + $statusPath + '"'
    $arguments = @(
        '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass',
        '-File', $quotedWorker, '-Repeat', '1', '-NoOpenPrompt',
        '-StatusFile', $quotedStatus
    )
    if ($CancelAtStep -gt 0) { $arguments += @('-CancelAtStep', "$CancelAtStep") }
    if ($SimulateExisting) { $arguments += '-SimulateExisting' }

    $child = Start-Process -FilePath $psExe -ArgumentList $arguments `
                           -NoNewWindow -PassThru
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while (-not $child.HasExited -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $child.Refresh()
    }

    if (-not $child.HasExited) {
        Write-Host ''
        Write-Host "  タイムアウト: $TimeoutSeconds 秒を超えました。" -ForegroundColor Red
        Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
        try { [void]$child.WaitForExit(5000) } catch { }
        Start-Sleep -Seconds 2
        $pidsAfter = @(Get-WordPids)
        $newWordPids = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
        $detail = if ($newWordPids.Count) {
            "子PowerShellを停止。残ったWINWORD PID: $($newWordPids -join ',')"
        } else {
            '子PowerShellを停止。新しいWINWORDの残留は検出されませんでした'
        }
        [void]$report.Add([pscustomobject]@{
            ラウンド = $run; 判定 = 'TIMEOUT'; 結果コード = '-'; 詳細 = $detail
        })
        if (Test-Path -LiteralPath $statusPath) {
            Remove-Item -LiteralPath $statusPath -Force -ErrorAction SilentlyContinue
        }
        Write-Host "  $detail" -ForegroundColor Yellow
        Write-Host '  自動でWINWORDは終了しません。文書を保存してPCを再起動してください。' -ForegroundColor Yellow
        break
    }

    # 一部環境ではWaitForExit()後もStart-ProcessのExitCodeを取得できないため、
    # workerが明示的に書いた結果ファイルを正とする。
    $child.WaitForExit()
    $child.Refresh()
    $resultCode = $null
    if (Test-Path -LiteralPath $statusPath) {
        $parsedCode = 0
        $statusText = (Get-Content -LiteralPath $statusPath -Raw).Trim()
        if ([int]::TryParse($statusText, [ref]$parsedCode)) {
            $resultCode = $parsedCode
        }
        Remove-Item -LiteralPath $statusPath -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
    $pidsAfter = @(Get-WordPids)
    $newWordPids = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
    $judge = if ($null -eq $resultCode) {
        'ERROR'
    } elseif ($resultCode -ne 0) {
        'ERROR'
    } elseif ($newWordPids.Count) {
        'LEAK'
    } else {
        'OK'
    }
    $detail = if ($null -eq $resultCode) {
        '子処理の結果ファイルがないか、内容が不正です'
    } elseif ($resultCode -ne 0) {
        "子処理がNGを報告（結果コード $resultCode）"
    } elseif ($newWordPids.Count) {
        "新しいWINWORD PIDが残留: $($newWordPids -join ',')"
    } else {
        '新しいWINWORDの残留なし'
    }
    [void]$report.Add([pscustomobject]@{
        ラウンド = $run
        判定 = $judge
        結果コード = $(if ($null -eq $resultCode) { '-' } else { $resultCode })
        詳細 = $detail
    })

    if ($judge -ne 'OK') {
        Write-Host "  監督判定: $judge / $detail" -ForegroundColor Red
        Write-Host '  残りのラウンドは実行しません。' -ForegroundColor Yellow
        break
    }
    Write-Host "  監督判定: OK / $detail" -ForegroundColor Green
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  監督結果' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
$report | Format-Table -AutoSize -Wrap
$report | Format-Table -AutoSize -Wrap | Out-String -Width 200 |
    Out-File -LiteralPath $reportPath -Encoding UTF8
Write-Host "  レポート: $reportPath" -ForegroundColor Cyan
Write-Host ''

if (@($report | Where-Object { $_.判定 -in @('TIMEOUT', 'ERROR', 'LEAK') }).Count -gt 0) {
    exit 1
}

