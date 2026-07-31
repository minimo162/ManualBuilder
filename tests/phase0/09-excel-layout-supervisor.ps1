# =====================================================================
# 09-excel-layout-supervisor.ps1 — Excel検証の監督・タイムアウト
#
# Excel COMがSaveAs等で戻らなくなってもメニュー全体を固めないため、
# 09-excel-layout.ps1を別PowerShellプロセスで1回ずつ実行する。
# タイムアウト時は子PowerShellだけを終了し、EXCELは自動終了しない。
# =====================================================================
[CmdletBinding()]
param(
    [ValidateRange(1, 10)][int]$Runs = 1,
    [ValidateRange(20, 600)][int]$TimeoutSeconds = 120,
    [int]$CancelAtStep = 0,
    [switch]$SimulateExisting
)

$ErrorActionPreference = 'Stop'
$worker = Join-Path $PSScriptRoot '09-excel-layout.ps1'
$outDir = Join-Path $PSScriptRoot 'out\excel'
$tmpDir = Join-Path $outDir '.tmp'
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
$reportPath = Join-Path $outDir 'result-09-supervisor.txt'
$report = New-Object System.Collections.ArrayList

function Get-ExcelPids {
    return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty Id)
}

function Get-CompletedFiles {
    if (-not (Test-Path -LiteralPath $outDir)) { return @() }
    return @(Get-ChildItem -LiteralPath $outDir -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty FullName)
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  Excel COM 監督実行' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host "  1回あたりのタイムアウト: $TimeoutSeconds 秒" -ForegroundColor Gray

$initialPids = @(Get-ExcelPids)
if (-not $SimulateExisting -and $Runs -gt 1 -and $initialPids.Count -gt 0) {
    Write-Host ''
    Write-Host "  Excel が起動中です（PID: $($initialPids -join ',')）。" -ForegroundColor Yellow
    Write-Host '  10回試験は開始しません。Excelを閉じてから再実行してください。' -ForegroundColor Yellow
    exit 0
}

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

for ($run = 1; $run -le $Runs; $run++) {
    Write-Host ''
    Write-Host "  === 監督ラウンド $run / $Runs ===" -ForegroundColor Cyan
    $pidsBefore = @(Get-ExcelPids)
    $filesBefore = @(Get-CompletedFiles)
    $statusPath = Join-Path $tmpDir ("supervisor-status-{0}-{1}.txt" -f $PID, $run)
    if (Test-Path -LiteralPath $statusPath) {
        Remove-Item -LiteralPath $statusPath -Force
    }

    $quotedWorker = '"' + $worker + '"'
    $quotedStatus = '"' + $statusPath + '"'
    $arguments = @(
        '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass',
        '-File', $quotedWorker, '-Round', "$run", '-NoOpenPrompt',
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
        $pidsAfter = @(Get-ExcelPids)
        $newExcelPids = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
        $detail = if ($newExcelPids.Count) {
            "子PowerShellを停止。残ったEXCEL PID: $($newExcelPids -join ',')"
        } else {
            '子PowerShellを停止。新しいEXCELの残留は検出されませんでした'
        }
        [void]$report.Add([pscustomobject]@{
            ラウンド = $run; 判定 = 'TIMEOUT'; 結果コード = '-'; 生成物 = 0; 詳細 = $detail
        })
        if (Test-Path -LiteralPath $statusPath) {
            Remove-Item -LiteralPath $statusPath -Force -ErrorAction SilentlyContinue
        }
        Write-Host "  $detail" -ForegroundColor Yellow
        Write-Host '  自動でEXCELは終了しません。開いているブックを保存してPCを再起動してください。' -ForegroundColor Yellow
        break
    }

    # 一部環境ではStart-ProcessのExitCodeが安定しないため、
    # workerが明示的に書いたASCII結果ファイルを正とする。
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

    # COM参照解放後の遅延終了を最大10秒待つ。
    $exitDeadline = (Get-Date).AddSeconds(10)
    do {
        $pidsAfter = @(Get-ExcelPids)
        $newExcelPids = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
        if ($newExcelPids.Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $exitDeadline)

    $filesAfter = @(Get-CompletedFiles)
    $newFiles = @($filesAfter | Where-Object { $filesBefore -notcontains $_ })
    $expectsFile = (-not $SimulateExisting -and $CancelAtStep -le 0 -and $pidsBefore.Count -eq 0)
    $expectsNoFile = ($SimulateExisting -or $CancelAtStep -gt 0)
    $existingExcelCase = (-not $SimulateExisting -and $CancelAtStep -le 0 -and $pidsBefore.Count -gt 0)

    $judge = if ($null -eq $resultCode) {
        'ERROR'
    } elseif ($resultCode -ne 0) {
        'ERROR'
    } elseif ($newExcelPids.Count) {
        'LEAK'
    } elseif ($expectsFile -and $newFiles.Count -ne 1) {
        'ERROR'
    } elseif ($expectsNoFile -and $newFiles.Count -ne 0) {
        'ERROR'
    } elseif ($existingExcelCase -and $newFiles.Count -gt 1) {
        'ERROR'
    } else {
        'OK'
    }

    $detail = if ($null -eq $resultCode) {
        '子処理の結果ファイルがないか、内容が不正です'
    } elseif ($resultCode -ne 0) {
        "子処理がNGを報告（結果コード $resultCode）"
    } elseif ($newExcelPids.Count) {
        "新しいEXCEL PIDが残留: $($newExcelPids -join ',')"
    } elseif ($expectsFile -and $newFiles.Count -ne 1) {
        "完成xlsxは1件の想定ですが $($newFiles.Count) 件です"
    } elseif ($expectsNoFile -and $newFiles.Count -ne 0) {
        "中止分岐ですが完成xlsxが $($newFiles.Count) 件増えました"
    } elseif ($existingExcelCase -and $newFiles.Count -gt 1) {
        "既存Excel試験で完成xlsxが $($newFiles.Count) 件増えました"
    } elseif ($newFiles.Count -eq 1) {
        "新しいEXCELの残留なし / 生成: $(Split-Path -Leaf $newFiles[0])"
    } else {
        '新しいEXCELの残留なし / 完成xlsxなし（想定どおり）'
    }

    [void]$report.Add([pscustomobject]@{
        ラウンド = $run
        判定 = $judge
        結果コード = $(if ($null -eq $resultCode) { '-' } else { $resultCode })
        生成物 = $newFiles.Count
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
Write-Host '  Excel監督結果' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
$report | Format-Table -AutoSize -Wrap
$report | Format-Table -AutoSize -Wrap | Out-String -Width 220 |
    Out-File -LiteralPath $reportPath -Encoding UTF8
Write-Host "  レポート: $reportPath" -ForegroundColor Cyan
Write-Host ''

if (@($report | Where-Object { $_.判定 -in @('TIMEOUT', 'ERROR', 'LEAK') }).Count -gt 0) {
    exit 1
}
