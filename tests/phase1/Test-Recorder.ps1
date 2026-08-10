# 操作記録の計算部分を検査する。
# 実際にクリックを記録することはしない。座標の扱いと、記録から除く判断だけを見る。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$srcRoot = Join-Path $repoRoot 'src'
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

Import-Module (Join-Path $srcRoot 'ManualBuilder.Recorder.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Project.psm1') -Force
$recorderSourceText = [IO.File]::ReadAllText((Join-Path $srcRoot 'ManualBuilder.Recorder.psm1'), [Text.Encoding]::UTF8)
$recorderServerSourceText = [IO.File]::ReadAllText((Join-Path $srcRoot 'ManualBuilder.RecorderServer.psm1'), [Text.Encoding]::UTF8)
$recorderCompanionRoot = Join-Path $srcRoot 'RecorderCompanion'
$recorderControllerSourceText = [IO.File]::ReadAllText((Join-Path $recorderCompanionRoot 'ManualBuilder.RecorderCompanion.cs'), [Text.Encoding]::UTF8)
$recorderControllerHtmlText = [IO.File]::ReadAllText((Join-Path $recorderCompanionRoot 'web\index.html'), [Text.Encoding]::UTF8)
$recorderControllerCssText = [IO.File]::ReadAllText((Join-Path $recorderCompanionRoot 'web\styles.css'), [Text.Encoding]::UTF8)
$recorderControllerJsText = [IO.File]::ReadAllText((Join-Path $recorderCompanionRoot 'web\app.js'), [Text.Encoding]::UTF8)
Add-Result (($recorderSourceText -match '\$pendingResultWindowHandle') -and
    ($recorderSourceText -match '別アプリへ移った画面を、直前操作の結果として結び付けない')) `
    '操作後画像を別アプリの画面へすり替えない'
Add-Result (($recorderServerSourceText -match 'Invoke-RecorderCompanion\.ps1') -and
    ($recorderServerSourceText -notmatch 'ManualBuilder\.RecorderCompanion\.exe') -and
    ($recorderServerSourceText -match 'ControllerProcessIdentity') -and
    ($recorderServerSourceText -match '\$controllerReadyTimeoutMs\s*=\s*15000') -and
    ($recorderControllerSourceText -match 'Topmost = true') -and
    ($recorderServerSourceText -match '記録レシートを開けませんでした')) `
    '対象アプリ上へWebView2記録モニターだけを起動し、初回準備を十分待って失敗時は記録を開始しない'
# 呼称は本体画面と揃える。このアプリの「削除」は元に戻せる操作なので、戻せない取り消しには使わない。
Add-Result (($recorderControllerHtmlText -match '直前の操作を取り消す') -and
    ($recorderControllerHtmlText -match '結果画像を追加') -and
    ($recorderControllerHtmlText -notmatch '直前の記録を削除') -and
    ($recorderControllerHtmlText -match '終了して確認')) `
    'WebView2記録レシートから取り消し・結果画像追加・終了を操作できる'
Add-Result (($recorderSourceText -match '\$ManualResultPath') -and
    ($recorderSourceText -match 'ResultRequestId') -and
    ($recorderControllerSourceText -match 'SetForegroundWindow') -and
    ($recorderControllerSourceText -match 'SetWindowDisplayAffinity')) `
    '記録モニターを画像へ混ぜず直前手順へ結果画面を追加する'
Add-Result (($recorderControllerHtmlText -match '記録レシート') -and
    ($recorderControllerHtmlText -match '最近記録した操作') -and
    ($recorderControllerHtmlText -match '直前画像を確認') -and
    ($recorderControllerHtmlText -match '画像確認を閉じる') -and
    ($recorderControllerHtmlText -notmatch 'data-view="before"') -and
    ($recorderControllerCssText -match 'grid-template-columns: repeat\(2') -and
    ($recorderControllerSourceText -match 'Width = 480') -and
    ($recorderControllerSourceText -match 'Height = 220') -and
    ($recorderControllerSourceText -match 'Width = 980') -and
    ($recorderControllerSourceText -match 'Height = 760')) `
    '記録モニターが記録レシートと大きな画像確認を切り替えられる'
Add-Result (($recorderControllerCssText -match '"BIZ UDPゴシック"') -and
    ($recorderControllerCssText -match 'border-radius: 10px') -and
    ($recorderControllerCssText -match 'display: grid') -and
    ($recorderControllerSourceText -notmatch 'System\.Windows\.Forms\.Button')) `
    'WebView2のCSSでフォント・角丸・配置を一元管理する'
Add-Result (($recorderControllerCssText -match '\.preview-stage:not\(\.compare\) \.preview-card figcaption') -and
    ($recorderControllerCssText -match '(?s)\.preview-card img.+padding:\s*6px.+object-fit:\s*contain')) `
    '単独表示では画像を外枠いっぱいに広げ、比較時だけ前後ラベルを重ねる'
Add-Result (($recorderControllerSourceText -match 'pendingResultAtUtc\.AddSeconds\(6\)') -and
    ($recorderControllerSourceText -match 'pendingUndoAtUtc\.AddSeconds\(6\)') -and
    ($recorderControllerSourceText -match 'closingRequested') -and
    ($recorderControllerSourceText -match 'pendingUndoId')) `
    '結果追加の応答欠落と終了・取消の競合から回復できる'
Add-Result (($recorderSourceText -match 'processedUndoRequests\.ContainsKey') -and
    ($recorderSourceText -match 'ProcessedRequests\.ContainsKey') -and
    ($recorderSourceText -match 'Invoke-MbManualResultRequest') -and
    ($recorderControllerSourceText -match 'pendingResultPayload')) `
    '応答遅延時に同じ削除・結果画面追加を二重実行しない'
Add-Result (($recorderSourceText -match 'DroppedMouseClicks') -and
    ($recorderSourceText -match 'DroppedKeyboardActivities') -and
    ($recorderSourceText -match "recordType = 'capture-gap'") -and
    ($recorderSourceText -match "recordType = 'capture-start'") -and
    ($recorderSourceText -match "recordType = 'capture-end'")) `
    '記録機能の停止・キューあふれを黙って欠落させず証拠台帳へ残す'
Add-Result (($recorderControllerSourceText -match '(?s)RequestUndo\(\).+pendingResultId') -and
    ($recorderControllerSourceText -match '(?s)RequestResult\(\).+pendingUndoId')) `
    '直前記録の削除と結果画面追加を同時に要求しない'
Add-Result (($recorderSourceText -match '(?s)ManualResultPath.+常駐パネルの未処理要求を先に取り込み') -and
    ($recorderSourceText -match '(?s)UndoPath.+終了要求と取消が重なった場合')) `
    '終了と同時の削除・結果画面追加を先に反映してから停止する'
Add-Result (($recorderControllerSourceText -match 'GetAwarenessFromDpiAwarenessContext') -and
    ($recorderControllerSourceText -match 'dpiAwareness=')) `
    '記録モニターのDPI設定を実状態で確認できる'
Add-Result (($recorderControllerSourceText -match 'LastWriteTimeUtc\.Ticks') -and
    ($recorderControllerJsText -match 'if \(image\.src !== url\)') -and
    ($recorderControllerSourceText -match 'SetVirtualHostNameToFolderMapping')) `
    '記録モニターが変更のない画像を短周期で読み直さない'

# ---------------------------------------------------------------------
# status.json の読み書き競合
# ---------------------------------------------------------------------
# 進捗を読む側が一瞬ファイルを開いていても、記録ワーカーは短時間待って
# 完成済みJSONへ差し替えられることを確認する。旧実装の File.Copy はここで失敗した。
$statusTestRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderStatus-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $statusTestRoot -Force)
$statusTestPath = Join-Path $statusTestRoot 'status.json'
$readyPath = Join-Path $statusTestRoot 'writer.ready'
$goPath = Join-Path $statusTestRoot 'writer.go'
$statusWriterProcess = $null
$statusLock = $null
try {
    Write-MbRecordingStatus -StatusPath $statusTestPath -JobId 'test' -State 'recording' -Count 1 -Message 'initial'

    $recorderModulePath = (Join-Path $srcRoot 'ManualBuilder.Recorder.psm1').Replace("'", "''")
    $escapedStatusPath = $statusTestPath.Replace("'", "''")
    $escapedReadyPath = $readyPath.Replace("'", "''")
    $escapedGoPath = $goPath.Replace("'", "''")
    $childCommand = "Import-Module '$recorderModulePath' -Force; [IO.File]::WriteAllText('$escapedReadyPath', 'ready'); while (-not (Test-Path -LiteralPath '$escapedGoPath')) { Start-Sleep -Milliseconds 10 }; Write-MbRecordingStatus -StatusPath '$escapedStatusPath' -JobId 'test' -State 'completed' -Count 2 -Message 'done'"
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childCommand))
    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $startParameters = @{
        FilePath = $powerShellPath
        ArgumentList = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encodedCommand)
        WindowStyle = 'Hidden'
        PassThru = $true
    }
    $statusWriterProcess = Start-Process @startParameters

    for ($i = 0; $i -lt 100 -and -not (Test-Path -LiteralPath $readyPath); $i++) {
        Start-Sleep -Milliseconds 50
    }
    if (-not (Test-Path -LiteralPath $readyPath)) { throw '進捗更新の競合試験を開始できませんでした。' }

    # FileShare.Readは旧File.Copyと原子的な置換の両方を一時的に拒否する。
    # 新実装は再試行するので、ロックを離したあとに成功する。
    $statusLock = [IO.File]::Open($statusTestPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    [IO.File]::WriteAllText($goPath, 'go')
    Start-Sleep -Milliseconds 250
    $statusLock.Dispose()
    $statusLock = $null

    [void]$statusWriterProcess.WaitForExit(5000)
    Add-Result ($statusWriterProcess.HasExited -and $statusWriterProcess.ExitCode -eq 0) '進捗ファイルが一時的に使用中でも更新を再試行する'
    $writtenStatus = [IO.File]::ReadAllText($statusTestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Add-Result ([string]$writtenStatus.state -eq 'completed' -and [int]$writtenStatus.count -eq 2) '競合後も完全なstatus.jsonへ差し替える'

} finally {
    if ($null -ne $statusLock) { $statusLock.Dispose() }
    if ($null -ne $statusWriterProcess) {
        if (-not $statusWriterProcess.HasExited) { try { $statusWriterProcess.Kill() } catch { } }
        $statusWriterProcess.Dispose()
    }
    Remove-Item -LiteralPath $statusTestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 記録中の直前取消
# ---------------------------------------------------------------------
$undoRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderUndo-' + [guid]::NewGuid().ToString('N'))
try {
    $undoEventsDirectory = Join-Path $undoRoot 'events'
    $undoEvidenceDirectory = Join-Path $undoRoot 'evidence'
    [void](New-Item -ItemType Directory -Path $undoEventsDirectory -Force)
    [void](New-Item -ItemType Directory -Path $undoEvidenceDirectory -Force)
    $undoEventsPath = Join-Path $undoRoot 'events.jsonl'
    $undoLedgerPath = Join-Path $undoRoot 'evidence-ledger.jsonl'
    $firstEvidenceId = 'evidence-' + [guid]::NewGuid().ToString('N')
    $removedEvidenceId = 'evidence-' + [guid]::NewGuid().ToString('N')
    [IO.File]::WriteAllLines($undoEventsPath, @(
        (@{ index=1; targetName='最初'; evidenceId=$firstEvidenceId } | ConvertTo-Json -Compress),
        (@{ index=2; targetName='直前'; evidenceId=$removedEvidenceId } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllLines($undoLedgerPath, @(
        (@{ recordType='operation'; id=$firstEvidenceId; sessionId='record-test' } | ConvertTo-Json -Compress),
        (@{ recordType='operation'; id=$removedEvidenceId; sessionId='record-test' } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $undoEventsDirectory 'event-002.jpg'), 'before')
    [IO.File]::WriteAllText((Join-Path $undoEventsDirectory 'event-002-result.jpg'), 'after')
    [IO.File]::WriteAllText((Join-Path $undoEvidenceDirectory ($removedEvidenceId + '.jpg')), 'original evidence')
    $undoResult = Remove-MbLastRecordingEvent -EventsPath $undoEventsPath -EventsDirectory $undoEventsDirectory `
        -LedgerPath $undoLedgerPath -JobId 'record-test'
    Add-Result ($undoResult.removed -and [int]$undoResult.count -eq 1) '記録中に直前の1操作を取り消せる'
    Add-Result ([string]$undoResult.lastTarget -eq '最初') '取消後の直前対象を戻す'
    Add-Result ([IO.File]::ReadAllLines($undoEventsPath).Count -eq 1) '取消後もそれ以前の操作ログを保つ'
    Add-Result (-not (Test-Path -LiteralPath (Join-Path $undoEventsDirectory 'event-002.jpg')) -and
        -not (Test-Path -LiteralPath (Join-Path $undoEventsDirectory 'event-002-result.jpg'))) '取消した操作前後の画像だけを除く'
    Add-Result (Test-Path -LiteralPath (Join-Path $undoEvidenceDirectory ($removedEvidenceId + '.jpg'))) `
        '直前取消後も元の操作証拠画像を残す'
    $undoLedger = @([IO.File]::ReadAllLines($undoLedgerPath) | ForEach-Object { $_ | ConvertFrom-Json })
    $undoDecision = @($undoLedger | Where-Object { $_.recordType -eq 'decision' -and $_.action -eq 'undo' })
    Add-Result ($undoDecision.Count -eq 1 -and @($undoDecision[0].evidenceIds) -contains $removedEvidenceId) `
        '取消を元操作の削除ではなく追記判断として記録する'
} finally {
    Remove-Item -LiteralPath $undoRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 取り込む範囲の決め方
# ---------------------------------------------------------------------
$window = [pscustomobject]@{ title = '経費申請'; class = 'Window'; left = 100; top = 100; width = 800; height = 600 }

$region = Get-MbCaptureRegion -Window $window -Target $null
Add-Result ([int]$region.left -eq 100 -and [int]$region.width -eq 800) '操作対象が無ければウィンドウの範囲を使う'

# ドロップダウンやメニューはウィンドウの外へ出る。切れないように範囲を広げる。
# 画面より外へは広げないので、期待値は実行環境の画面の大きさで頭打ちにする。
$screen = Get-MbVirtualScreenBounds
$outside = [pscustomobject]@{ left = 850.0; top = 620.0; width = 120.0; height = 200.0 }
$widened = Get-MbCaptureRegion -Window $window -Target $outside
$expectedRight = [Math]::Min(970, [int]$screen.left + [int]$screen.width)
$expectedBottom = [Math]::Min(820, [int]$screen.top + [int]$screen.height)
Add-Result (([int]$widened.left + [int]$widened.width) -ge $expectedRight) 'ウィンドウの外へ出た操作対象まで範囲を広げる'
Add-Result (([int]$widened.top + [int]$widened.height) -ge $expectedBottom) '縦にはみ出した分も範囲へ含める'
Add-Result (([int]$widened.left + [int]$widened.width) -le ([int]$screen.left + [int]$screen.width)) '画面の外までは広げない'

$inside = [pscustomobject]@{ left = 200.0; top = 200.0; width = 80.0; height = 30.0 }
$unchanged = Get-MbCaptureRegion -Window $window -Target $inside
Add-Result ([int]$unchanged.width -eq 800 -and [int]$unchanged.height -eq 600) 'ウィンドウの中の操作対象では範囲を広げない'

# ウィンドウが取れないときは画面全体へ退避する。
$fallback = Get-MbCaptureRegion -Window $null -Target $null
Add-Result ([int]$fallback.width -gt 0 -and [int]$fallback.height -gt 0) 'ウィンドウが無ければ画面全体を使う'

$tiny = [pscustomobject]@{ title = ''; class = 'Window'; left = 0; top = 0; width = 4; height = 4 }
$tinyRegion = Get-MbCaptureRegion -Window $tiny -Target $null
Add-Result ([int]$tinyRegion.width -gt 16) '極端に小さいウィンドウでは画面全体へ退避する'

# クリック後画像は同じイベント番号へ -result を付け、操作前画像を上書きしない。
$resultRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderResult-' + [guid]::NewGuid().ToString('N'))
$resultBitmap = $null
try {
    [void](New-Item -ItemType Directory -Path $resultRoot -Force)
    $resultBitmap = New-Object Drawing.Bitmap -ArgumentList @(80, 60)
    $resultCapture = [pscustomobject]@{ bitmap = $resultBitmap; origin = [pscustomobject]@{ left = 0; top = 0; width = 80; height = 60 } }
    $resultWindow = [pscustomobject]@{ title = '詳細'; left = 0; top = 0; width = 80; height = 60 }
    $resultName = Save-MbRecordingResultImage -Capture $resultCapture -Index 7 -EventsDirectory $resultRoot -Window $resultWindow
    Add-Result ($resultName -eq 'event-007-result.jpg') 'クリック後画像を同じ操作へ関連付ける名前で保存する'
    Add-Result (Test-Path -LiteralPath (Join-Path $resultRoot $resultName) -PathType Leaf) 'クリック後画像の実体を保存する'
} finally {
    if ($resultBitmap) { $resultBitmap.Dispose() }
    Remove-Item -LiteralPath $resultRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 操作対象の矩形を画像の中の位置へ直す
# ---------------------------------------------------------------------
$capturedRegion = [pscustomobject]@{ left = 100; top = 100; width = 800; height = 600 }
$button = [pscustomobject]@{ left = 500.0; top = 400.0; width = 80.0; height = 30.0 }
$rect = ConvertTo-MbRegionRect -Region $capturedRegion -Target $button
Add-Result ($null -ne $rect) '範囲の中の操作対象は矩形になる'
Add-Result ([Math]::Abs([double]$rect.x1 - 0.5) -lt 0.0001) '左端の割合が正しい'
Add-Result ([Math]::Abs([double]$rect.y1 - 0.5) -lt 0.0001) '上端の割合が正しい'
Add-Result ([Math]::Abs([double]$rect.x2 - 0.6) -lt 0.0001) '右端の割合が正しい'
Add-Result ([Math]::Abs([double]$rect.y2 - 0.55) -lt 0.0001) '下端の割合が正しい'

$outsideTarget = [pscustomobject]@{ left = -500.0; top = -500.0; width = 80.0; height = 30.0 }
$clamped = ConvertTo-MbRegionRect -Region $capturedRegion -Target $outsideTarget
Add-Result ($null -eq $clamped) '範囲の外に出た操作対象は矩形にしない'

# ウィンドウ全体を掴んだ場合、赤枠にしても意味がない。
$whole = [pscustomobject]@{ left = 100.0; top = 100.0; width = 800.0; height = 600.0 }
Add-Result ($null -eq (ConvertTo-MbRegionRect -Region $capturedRegion -Target $whole)) '範囲いっぱいの矩形は赤枠にしない'

# ブラウザーのページ領域は、タブやサイドバーを除くため画像の96%未満でも十分大きい。
$largePage = [pscustomobject]@{ left = 120.0; top = 120.0; width = 740.0; height = 520.0 }
Add-Result ($null -eq (ConvertTo-MbRegionRect -Region $capturedRegion -Target $largePage)) 'ページ全体に近い矩形は赤枠にしない'

$sliver = [pscustomobject]@{ left = 500.0; top = 400.0; width = 0.5; height = 30.0 }
Add-Result ($null -eq (ConvertTo-MbRegionRect -Region $capturedRegion -Target $sliver)) '潰れた矩形は赤枠にしない'

$fourKRegion = [pscustomobject]@{ left = 0; top = 0; width = 3840; height = 2160 }
$fourKThinButton = [pscustomobject]@{ left = 1800.0; top = 900.0; width = 6.0; height = 32.0 }
Add-Result ($null -ne (ConvertTo-MbRegionRect -Region $fourKRegion -Target $fourKThinButton)) '4K画面でも物理幅のある細い操作対象を残す'

Add-Result ($null -eq (ConvertTo-MbRegionRect -Region $capturedRegion -Target $null)) '操作対象が無ければ矩形は作らない'

# DOMを暫定採用しても、UIAとクリック位置を別候補として残す。
$domCandidate = [pscustomobject]@{
    name = '送信'; controlType = 'ControlType.Button'; provider = 'DOM'; confidence = 'medium'
    left = 500.0; top = 400.0; width = 80.0; height = 30.0
}
$uiaCandidate = [pscustomobject]@{
    name = '申請を送信'; controlType = 'ControlType.Button'; provider = 'UIA-CACHE'
    left = 490.0; top = 392.0; width = 105.0; height = 45.0
}
$inferredCandidate = [pscustomobject]@{
    name = '送信の文字'; controlType = 'ControlType.Text'; isInferred = $true
    left = 510.0; top = 405.0; width = 55.0; height = 18.0
}
$clickCandidate = New-MbClickPointTargetInfo -X 540 -Y 415 -Window $window
$alternatives = @(ConvertTo-MbRecordingTargetCandidates -Region $capturedRegion -SelectedTarget $domCandidate `
    -Targets @($uiaCandidate, $inferredCandidate, $clickCandidate) -Maximum 4)
Add-Result ($alternatives.Count -eq 4) 'DOM・UIA・推定UIA・クリック位置を最大4候補として残す'
Add-Result ([string]$alternatives[0].source -eq 'DOM' -and [string]$alternatives[0].confidence -eq 'medium') 'DOMを確定扱いせず暫定候補にする'
Add-Result (@($alternatives | Where-Object { [string]$_.source -eq 'UIA-CACHE' }).Count -eq 1) '誤DOMを補えるクリック前UIA候補を残す'
Add-Result (@($alternatives | Where-Object { [string]$_.source -eq 'UIA' -and [string]$_.confidence -eq 'low' }).Count -eq 1) '推定UIA候補を低信頼として残す'
Add-Result (@($alternatives | Where-Object { [string]$_.source -eq 'click-point' -and [string]$_.confidence -eq 'low' }).Count -eq 1) '対象不明時のクリック位置を低信頼候補として残す'
Add-Result (@($alternatives | Select-Object -ExpandProperty id -Unique).Count -eq 4) '候補へ重複しないIDを付ける'

$limitedAlternatives = @(ConvertTo-MbRecordingTargetCandidates -Region $capturedRegion -SelectedTarget $domCandidate `
    -Targets @($uiaCandidate, $inferredCandidate, $clickCandidate, $button) -Maximum 3)
Add-Result ($limitedAlternatives.Count -eq 3) '候補数の上限を守る'

$longTargetName = ('操作対象' * 60)
$limitedTargetName = ConvertTo-MbRecorderTargetName -Value $longTargetName
Add-Result ($limitedTargetName.Length -eq 200 -and $limitedTargetName.EndsWith('…')) '長い操作対象名を200文字へ省略する'
$limitedInputName = ConvertTo-MbRecorderTargetName -Value $longTargetName -Suffix '（入力）'
Add-Result ($limitedInputName.Length -eq 200 -and $limitedInputName.EndsWith('…（入力）')) '入力の補足を含めて操作対象を200文字へ収める'
$normalizedTargetName = ConvertTo-MbRecorderTargetName -Value "  申請`r`n   ボタン  "
Add-Result ([string]$normalizedTargetName -eq '申請 ボタン') '操作対象名の改行と連続空白を整える'

# 入力欄は自動で黒塗りしない。必要な箇所は取り込み後の画像編集で手動マスクする。
$inputImagePath = Join-Path $env:TEMP ('ManualBuilder-RecorderInput-' + [guid]::NewGuid().ToString('N') + '.jpg')
$inputBitmap = New-Object Drawing.Bitmap 100, 80
$inputGraphics = [Drawing.Graphics]::FromImage($inputBitmap)
try {
    $inputGraphics.Clear([Drawing.Color]::White)
    $inputGraphics.FillRectangle([Drawing.Brushes]::Blue, 20, 25, 40, 20)
    $capture = [pscustomobject]@{
        bitmap = $inputBitmap
        origin = [pscustomobject]@{ left = 0; top = 0; width = 100; height = 80 }
    }
    $imageRegion = [pscustomobject]@{ left = 0; top = 0; width = 100; height = 80 }
    [void](Save-MbBitmapRegion -Capture $capture -Region $imageRegion -Path $inputImagePath -MaxEdge 0 -Quality 100)
    $savedBitmap = New-Object Drawing.Bitmap $inputImagePath
    try {
        $inputPixel = $savedBitmap.GetPixel(30, 30)
        $visible = $savedBitmap.GetPixel(5, 5)
        Add-Result ($inputPixel.B -gt 180 -and $inputPixel.R -lt 80) '入力欄の表示を自動で黒塗りしない'
        Add-Result ($visible.R -gt 235 -and $visible.G -gt 235 -and $visible.B -gt 235) '入力欄以外の画面も維持する'
    } finally {
        $savedBitmap.Dispose()
    }
} finally {
    $inputGraphics.Dispose()
    $inputBitmap.Dispose()
    Remove-Item -LiteralPath $inputImagePath -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 記録から除くウィンドウ
# ---------------------------------------------------------------------
$taskbar = [pscustomobject]@{ title = ''; class = 'Shell_TrayWnd'; left = 0; top = 0; width = 1920; height = 40 }
Add-Result ((Test-MbIgnoredWindow -Window $taskbar) -eq $true) 'タスクバーへのクリックは記録しない'

$desktop = [pscustomobject]@{ title = ''; class = 'Progman'; left = 0; top = 0; width = 1920; height = 1080 }
Add-Result ((Test-MbIgnoredWindow -Window $desktop) -eq $true) 'デスクトップへのクリックは記録しない'

$app = [pscustomobject]@{ title = '経費申請システム'; class = 'Window'; left = 0; top = 0; width = 800; height = 600 }
Add-Result ((Test-MbIgnoredWindow -Window $app) -eq $false) '業務アプリへのクリックは記録する'

# 記録を止めるためにManualBuilderへ戻る操作が手順に混ざらないようにする。
$self = [pscustomobject]@{ title = 'ManualBuilder — Edge'; class = 'Chrome_WidgetWin_1'; left = 0; top = 0; width = 800; height = 600 }
Add-Result ((Test-MbIgnoredWindow -Window $self -IgnoreTitlePatterns @('ManualBuilder')) -eq $true) 'ManualBuilder自身への操作は記録しない'
$edgePage = [pscustomobject]@{ title = 'ManualBuilderの使い方 - Microsoft 365 Copilot - Microsoft Edge'; class = 'Chrome_WidgetWin_1'; left = 0; top = 0; width = 800; height = 600 }
Add-Result ((Test-MbIgnoredWindow -Window $edgePage -IgnoreTitlePatterns @('ManualBuilder')) -eq $false) `
    'ページ題名にManualBuilderを含むEdge操作を自己画面と誤認しない'
Add-Result ((Test-MbIgnoredWindow -Window $self) -eq $false) '除外指定が無ければ普通に記録する'
Add-Result ((Test-MbIgnoredWindow -Window $null) -eq $true) 'ウィンドウが取れない場合は記録しない'

# ---------------------------------------------------------------------
# UI Automation から受け取った情報の検査
# ---------------------------------------------------------------------
$good = [pscustomobject]@{ name = '申請'; controlType = 'ControlType.Button'; left = 10.0; top = 10.0; width = 80.0; height = 24.0 }
Add-Result ((Test-MbUsableElementInfo -Info $good) -eq $true) '大きさのある要素は使える'

$zero = [pscustomobject]@{ name = ''; controlType = 'ControlType.Text'; left = 10.0; top = 10.0; width = 0.0; height = 24.0 }
Add-Result ((Test-MbUsableElementInfo -Info $zero) -eq $false) '大きさの無い要素は使わない'

# 画面に出ていない要素の矩形は無限大になる。これを赤枠にすると座標が壊れる。
$infinite = [pscustomobject]@{ name = 'x'; controlType = 'ControlType.Button'; left = [double]::PositiveInfinity; top = 0.0; width = 10.0; height = 10.0 }
Add-Result ((Test-MbUsableElementInfo -Info $infinite) -eq $false) '画面に出ていない要素は使わない'

Add-Result ((Test-MbUsableElementInfo -Info $null) -eq $false) 'nullは使わない'

$nearEdge = [pscustomobject]@{ name = '境界のボタン'; controlType = 'ControlType.Button'; isActionable = $true; left = 107.0; top = 100.0; width = 80.0; height = 30.0 }
$selected = Select-MbUiaTargetInfo -Candidates @($nearEdge) -X 101 -Y 115
Add-Result ($null -ne $selected) 'UIA座標が数ピクセルずれても近傍の操作対象を選ぶ'

# ブラウザーではFromPointがページ全体のDocumentを返すことがある。
# 同じ点を含む候補から、実際に押せるリンク・ボタンの最小矩形を選ぶ。
$page = [pscustomobject]@{
    name = '申請画面'; controlType = 'ControlType.Document'; isActionable = $false
    left = 0.0; top = 0.0; width = 1200.0; height = 800.0
}
$link = [pscustomobject]@{
    name = '添付ファイルをダウンロード'; controlType = 'ControlType.Hyperlink'; isActionable = $true
    left = 140.0; top = 420.0; width = 340.0; height = 28.0
}
$linkText = [pscustomobject]@{
    name = '添付ファイルをダウンロード'; controlType = 'ControlType.Text'; isActionable = $false
    left = 146.0; top = 423.0; width = 300.0; height = 20.0
}
$selected = Select-MbUiaTargetInfo -Candidates @($page, $link, $linkText) -X 200 -Y 430
Add-Result ($null -ne $selected -and [string]$selected.controlType -eq 'ControlType.Hyperlink') 'ページ全体ではなくクリックしたリンクを選ぶ'

$unnamedButton = [pscustomobject]@{
    name = ''; controlType = 'ControlType.Button'; isActionable = $true
    left = 500.0; top = 300.0; width = 90.0; height = 26.0
}
$selected = Select-MbUiaTargetInfo -Candidates @($page, $unnamedButton) -X 520 -Y 312
Add-Result ($null -ne $selected -and [string]$selected.controlType -eq 'ControlType.Button') '名前が空でもクリックしたボタンの矩形を選ぶ'

$invokeText = [pscustomobject]@{
    name = '次へ'; controlType = 'ControlType.Text'; isActionable = $true
    left = 620.0; top = 500.0; width = 60.0; height = 22.0
}
$selected = Select-MbUiaTargetInfo -Candidates @($page, $invokeText) -X 640 -Y 510
Add-Result ($null -ne $selected -and [string]$selected.name -eq '次へ') 'Textとして公開された要素も操作パターンがあれば選ぶ'

$namedText = [pscustomobject]@{
    name = '承認依頼を送信'; controlType = 'ControlType.Text'; isActionable = $false
    left = 700.0; top = 520.0; width = 140.0; height = 26.0
}
$inferred = Select-MbUiaNamedTargetInfo -Candidates @($page, $namedText) -X 740 -Y 530 `
    -Window ([pscustomobject]@{ left = 0; top = 0; width = 1200; height = 800 })
Add-Result ($null -ne $inferred -and [string]$inferred.name -eq '承認依頼を送信') '操作パターンが無い名前付き要素もクリック位置から補う'
Add-Result ([bool]$inferred.isInferred) '推定で補った操作対象を識別できる'

# エクスプローラーのフォルダー移動などでクリック後に元要素が消えても、
# 直前の同じウィンドウ・同じカーソル位置のキャッシュを利用できる。
$uiaCachePath = Join-Path $env:TEMP ('ManualBuilder-UiaTarget-' + [guid]::NewGuid().ToString('N') + '.json')
try {
    $cachedListItem = [pscustomobject]@{
        name = '請求書'; controlType = 'ControlType.ListItem'; isActionable = $true
        left = 300.0; top = 220.0; width = 120.0; height = 28.0
    }
    $cacheValue = [pscustomobject]@{
        updatedAtUtc = [DateTime]::UtcNow.ToString('o'); cursorX = 340; cursorY = 234
        windowHandle = 12345
        window = [pscustomobject]@{ handle = 12345; title = '請求書'; class = 'CabinetWClass'; left = 0; top = 0; width = 1000; height = 700 }
        target = $cachedListItem
    }
    [IO.File]::WriteAllText($uiaCachePath, ($cacheValue | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
    $cacheWindow = [pscustomobject]@{ handle = 12345; left = 0; top = 0; width = 1000; height = 700 }
    $cachedTarget = Get-MbUiaTargetFromCache -Path $uiaCachePath -X 342 -Y 235 -Window $cacheWindow
    Add-Result ($null -ne $cachedTarget -and [string]$cachedTarget.name -eq '請求書' -and
        [long]$cachedTarget.captureWindow.handle -eq 12345) 'クリック前に保持したWindows操作対象とウィンドウを利用する'
    Add-Result ($null -eq (Get-MbUiaTargetFromCache -Path $uiaCachePath -X 340 -Y 211 -Window $cacheWindow)) `
        'クリック座標から外れた隣接UIAキャッシュを赤枠候補にしない'
    $changedPageWindow = [pscustomobject]@{ handle = 12345; title = '支払完了'; left = 0; top = 0; width = 1000; height = 700 }
    Add-Result ($null -eq (Get-MbUiaTargetFromCache -Path $uiaCachePath -X 342 -Y 235 -Window $changedPageWindow)) `
        '同じEdgeウィンドウでも画面遷移後は古いUIAキャッシュを使わない'
    $otherWindow = [pscustomobject]@{ handle = 54321; left = 0; top = 0; width = 1000; height = 700 }
    Add-Result ($null -eq (Get-MbUiaTargetFromCache -Path $uiaCachePath -X 342 -Y 235 -Window $otherWindow)) `
        '十分新しいキャッシュでも別HWNDなら誤適用しない'
    $cacheValue.updatedAtUtc = [DateTime]::UtcNow.AddMilliseconds(-800).ToString('o')
    [IO.File]::WriteAllText($uiaCachePath, ($cacheValue | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
    Add-Result ($null -eq (Get-MbUiaTargetFromCache -Path $uiaCachePath -X 342 -Y 235 -Window $otherWindow)) '別ウィンドウのUIAキャッシュを誤適用しない'
} finally {
    Remove-Item -LiteralPath $uiaCachePath -Force -ErrorAction SilentlyContinue
}

$msaaButton = New-MbMsaaElementInfo -Name '保存' -Role 43 -DefaultAction '押す' `
    -Left 820 -Top 540 -Width 96 -Height 32
Add-Result ([string]$msaaButton.controlType -eq 'ControlType.Button' -and [bool]$msaaButton.isActionable) 'MSAAのボタンを操作対象へ変換する'
$msaaLink = New-MbMsaaElementInfo -Name '詳細を見る' -Role 30 -DefaultAction '' `
    -Left 620 -Top 440 -Width 120 -Height 22
Add-Result ([string]$msaaLink.controlType -eq 'ControlType.Hyperlink' -and [bool]$msaaLink.isActionable) 'MSAAのリンクを操作対象へ変換する'
$msaaText = New-MbMsaaElementInfo -Name '説明だけの文字' -Role 41 -DefaultAction '' `
    -Left 620 -Top 470 -Width 120 -Height 22
Add-Result (-not [bool]$msaaText.isActionable) 'MSAAの静的な文字を操作可能と誤判定しない'

Add-Result ($null -eq (Select-MbUiaTargetInfo -Candidates @($page, $linkText) -X 200 -Y 430)) '操作できる候補が無ければページ全体の赤枠を付けない'

$window = [pscustomobject]@{ left = 100.0; top = 100.0; width = 800.0; height = 600.0 }
$pointFallback = New-MbClickPointTargetInfo -X 420 -Y 360 -Window $window
Add-Result ([string]$pointFallback.controlType -eq 'ControlType.ClickPoint') 'UIAで特定できない操作はクリック位置へフォールバックする'
Add-Result ([Math]::Abs([double]$pointFallback.left - 392.0) -lt 0.001 -and
    [Math]::Abs([double]$pointFallback.top - 342.0) -lt 0.001) 'クリック位置の小さな枠を中央へ置く'
$edgeFallback = New-MbClickPointTargetInfo -X 101 -Y 101 -Window $window
Add-Result ([double]$edgeFallback.left -ge 100.0 -and [double]$edgeFallback.top -ge 100.0) 'クリック位置の枠をウィンドウ外へ出さない'
Add-Result ([double]$edgeFallback.anchorX -eq 101.0 -and [double]$edgeFallback.anchorY -eq 101.0) `
    '画面端で小枠をクランプしても実クリック座標を保持する'

$wideSplitButton = [pscustomobject]@{
    name = '名前'; controlType = 'ControlType.SplitButton'; provider = 'UIA-CACHE'
    left = 340.0; top = 140.0; width = 360.0; height = 28.0
}
$widePoint = New-MbClickPointTargetInfo -X 470 -Y 155 -Window $window
$widePrimary = Select-MbRecordingPrimaryTarget -Target $wideSplitButton -PointTarget $widePoint -Window $window
Add-Result ([bool]$widePrimary.isFallback -and [Math]::Abs([double]$widePrimary.width - 56.0) -lt 0.001) '横長のUIA親要素は既定赤枠をクリック点へ寄せる'
Add-Result ([string]$widePrimary.name -eq '名前') 'クリック点へ寄せても取得した対象名を文章化用に残す'
Add-Result ([string](Get-MbRecorderTargetEvidence -Target $widePrimary).confidence -eq 'medium' -and
    [string](Get-MbRecorderTargetEvidence -Target $wideSplitButton).confidence -eq 'low') `
    'クリック座標を中信頼、横長UIAを低信頼としてCopilotへ渡す'

$observedExplorerSplitButton = [pscustomobject]@{
    name = '名前'; controlType = 'ControlType.SplitButton'; provider = 'UIA'
    left = 340.0; top = 140.0; width = 194.0; height = 28.0
}
$observedPrimary = Select-MbRecordingPrimaryTarget -Target $observedExplorerSplitButton -PointTarget `
    (New-MbClickPointTargetInfo -X 470 -Y 155 -Window $window) -Window $window
Add-Result ([bool]$observedPrimary.isFallback) '実機で観測した幅24%のExplorer誤候補もクリック点へ寄せる'

$smallButton = [pscustomobject]@{
    name = '保存'; controlType = 'ControlType.Button'; provider = 'UIA-CACHE'
    left = 430.0; top = 340.0; width = 90.0; height = 32.0
}
$smallPoint = New-MbClickPointTargetInfo -X 470 -Y 356 -Window $window
$smallPrimary = Select-MbRecordingPrimaryTarget -Target $smallButton -PointTarget $smallPoint -Window $window
Add-Result (-not [bool]$smallPrimary.isFallback -and [string]$smallPrimary.name -eq '保存') '妥当なUIAボタンはクリック点へ置き換えない'

$nearbySmallIcon = [pscustomobject]@{
    name = '別アイコン'; controlType = 'ControlType.Button'; isActionable = $true
    left = 525.0; top = 348.0; width = 16.0; height = 16.0
}
$containingButton = [pscustomobject]@{
    name = '保存'; controlType = 'ControlType.Button'; isActionable = $true
    left = 430.0; top = 340.0; width = 90.0; height = 32.0
}
$containmentWinner = Select-MbUiaTargetInfo -Candidates @($nearbySmallIcon, $containingButton) -X 520 -Y 356
Add-Result ([string]$containmentWinner.name -eq '保存') '近くの小要素より実クリックを含む操作対象を優先する'
$namedContainmentWinner = Select-MbUiaNamedTargetInfo -Candidates @($nearbySmallIcon, $containingButton) -X 520 -Y 356 -Window $window
Add-Result ([string]$namedContainmentWinner.name -eq '保存') '名前だけで補う経路でも実クリックを含む操作対象を優先する'

$nearbyOnlyTarget = Select-MbUiaTargetInfo -Candidates @($nearbySmallIcon) -X 520 -Y 356
$nearbyPrimary = Select-MbRecordingPrimaryTarget -Target $nearbyOnlyTarget -PointTarget `
    (New-MbClickPointTargetInfo -X 520 -Y 356 -Window $window) -Window $window
Add-Result ([bool]$nearbyPrimary.isFallback -and [string]$nearbyPrimary.name -eq '' -and
    [string](Get-MbRecorderTargetEvidence -Target $nearbyPrimary).confidence -eq 'low') `
    'クリックを含まない近傍候補は名前を移さず低信頼のクリック点にする'

$edgeButton = [pscustomobject]@{
    name = '戻る'; controlType = 'ControlType.Button'; provider = 'UIA'
    left = 100.0; top = 100.0; width = 20.0; height = 20.0
}
$edgePrimary = Select-MbRecordingPrimaryTarget -Target $edgeButton -PointTarget $edgeFallback -Window $window
Add-Result (-not [bool]$edgePrimary.isFallback -and [string]$edgePrimary.name -eq '戻る') `
    'ウィンドウ端の正しい小要素をクランプ後の枠中心で誤降格しない'

$adjacentCell = [pscustomobject]@{
    name = 'F9'; controlType = 'ControlType.DataItem'; provider = 'UIA-CACHE'
    left = 320.0; top = 420.0; width = 52.0; height = 24.0
}
$adjacentPoint = New-MbClickPointTargetInfo -X 346 -Y 405 -Window $window
$adjacentPrimary = Select-MbRecordingPrimaryTarget -Target $adjacentCell -PointTarget $adjacentPoint -Window $window
Add-Result ([bool]$adjacentPrimary.isFallback -and [string]$adjacentPrimary.controlType -eq 'ControlType.ClickPoint') `
    '隣のセルを指すUIA候補より実クリック位置の赤枠を優先する'

$wideEdit = [pscustomobject]@{
    name = '数式バー'; controlType = 'ControlType.Edit'; provider = 'UIA'
    left = 210.0; top = 140.0; width = 680.0; height = 34.0
}
$wideEditPrimary = Select-MbRecordingPrimaryTarget -Target $wideEdit -PointTarget `
    (New-MbClickPointTargetInfo -X 260 -Y 165 -Window $window) -Window $window
Add-Result ([bool]$wideEditPrimary.isFallback -and [string]$wideEditPrimary.controlType -eq 'ControlType.ClickPoint') `
    '数式バーのような過大な横長Editを既定赤枠にしない'

$normalizedPoint = ConvertTo-MbNormalizedClickPoint -Region $capturedRegion -X 500 -Y 400
Add-Result ([Math]::Abs([double]$normalizedPoint.x - 0.5) -lt 0.001 -and
    [Math]::Abs([double]$normalizedPoint.y - 0.5) -lt 0.001) 'クリック座標を記録画像内の位置へ正規化する'

# ---------------------------------------------------------------------
# 記録用EdgeのDOM座標
# ---------------------------------------------------------------------
$domSnapshot = [pscustomobject]@{
    name = '申請する'; role = 'button'; tag = 'button'; type = ''
    clientX = 100.0; clientY = 50.0; dpr = 2.0
    rect = [pscustomobject]@{ left = 80.0; top = 40.0; width = 60.0; height = 24.0 }
}
$domTarget = ConvertFrom-MbDomSnapshotTarget -Snapshot $domSnapshot -X 1000 -Y 500
Add-Result ($null -ne $domTarget -and [string]$domTarget.provider -eq 'DOM') '記録用EdgeのDOM要素を操作対象へ変換する'
Add-Result ([string]$domTarget.controlType -eq 'ControlType.Button' -and [string]$domTarget.name -eq '申請する') 'DOMの役割と名前を保持する'
Add-Result ([Math]::Abs([double]$domTarget.left - 960.0) -lt 0.001 -and
    [Math]::Abs([double]$domTarget.top - 480.0) -lt 0.001 -and
    [Math]::Abs([double]$domTarget.width - 120.0) -lt 0.001) '表示倍率を含むDOM矩形を物理ピクセルへ変換する'
$anchoredDomSnapshot = [pscustomobject]@{
    name = '申請する'; role = 'button'; tag = 'button'; type = ''
    clientX = 100.0; clientY = 50.0; screenX = 980.0; screenY = 490.0; dpr = 2.0
    rect = [pscustomobject]@{ left = 80.0; top = 40.0; width = 60.0; height = 24.0 }
}
$anchoredDomTarget = ConvertFrom-MbDomSnapshotTarget -Snapshot $anchoredDomSnapshot -X 1000 -Y 500
Add-Result ([Math]::Abs([double]$anchoredDomTarget.left - 940.0) -lt 0.001 -and
    [Math]::Abs([double]$anchoredDomTarget.top - 470.0) -lt 0.001) 'クリック後にマウスが動いてもpointerdown時のEdge赤枠を維持する'
$fileDomSnapshot = [pscustomobject]@{
    name = 'ファイルの選択'; role = ''; tag = 'input'; type = 'file'
    clientX = 100.0; clientY = 50.0; dpr = 1.0
    rect = [pscustomobject]@{ left = 80.0; top = 40.0; width = 240.0; height = 24.0 }
}
$fileDomTarget = ConvertFrom-MbDomSnapshotTarget -Snapshot $fileDomSnapshot -X 1000 -Y 500
Add-Result ([string]$fileDomTarget.inputType -eq 'file') 'DOMのファイル選択欄を内側のWindowsボタンへ補正できるよう識別する'
$farDomSnapshot = [pscustomobject]@{
    name = '別の要素'; role = 'link'; tag = 'a'; type = ''
    clientX = 5.0; clientY = 5.0; dpr = 1.0
    rect = [pscustomobject]@{ left = 500.0; top = 500.0; width = 20.0; height = 20.0 }
}
Add-Result ($null -eq (ConvertFrom-MbDomSnapshotTarget -Snapshot $farDomSnapshot -X 1000 -Y 500)) 'クリック点と離れたDOM矩形を誤適用しない'

# ページ遷移直後は、現在ページのタイトルとpointerdown時のタイトルが異なる。
# 遷移前画像には旧対象を使えるが、遷移後画像へ旧矩形を重ねてはならない。
$domCachePath = Join-Path $env:TEMP ('ManualBuilder-DomTarget-' + [guid]::NewGuid().ToString('N') + '.json')
try {
    $epoch = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
    $pointerAt = [long](([DateTime]::UtcNow - $epoch).TotalMilliseconds)
    $transitionPointer = [pscustomobject]@{
        source = 'pointerdown'; at = $pointerAt; pageTitle = '申請入力'; pageUrl = 'https://example.test/input'
        name = '確認へ'; role = 'button'; tag = 'button'; type = ''
        clientX = 100.0; clientY = 50.0; screenX = 1000.0; screenY = 500.0; dpr = 1.0
        rect = [pscustomobject]@{ left = 80.0; top = 40.0; width = 60.0; height = 24.0 }
    }
    $transitionCache = [pscustomobject]@{
        updatedAtUtc = [DateTime]::UtcNow.ToString('o'); cursorX = 1000; cursorY = 500
        page = [pscustomobject]@{ title = '申請完了'; url = 'https://example.test/done'; pointer = $transitionPointer; hover = $null }
    }
    [IO.File]::WriteAllText($domCachePath, ($transitionCache | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))
    $beforeTransitionWindow = [pscustomobject]@{
        handle = 9876; title = '申請入力 - Microsoft Edge'; class = 'Chrome_WidgetWin_1'
        left = 0; top = 0; width = 1200; height = 800
    }
    $afterTransitionWindow = [pscustomobject]@{
        handle = 9876; title = '申請完了 - Microsoft Edge'; class = 'Chrome_WidgetWin_1'
        left = 0; top = 0; width = 1200; height = 800
    }
    Add-Result ($null -ne (Get-MbDomTargetFromCache -Path $domCachePath -X 1000 -Y 500 -Window $beforeTransitionWindow)) `
        '遷移前の画像とタイトルにはpointerdown時のDOM対象を使う'
    Add-Result ($null -eq (Get-MbDomTargetFromCache -Path $domCachePath -X 1000 -Y 500 -Window $afterTransitionWindow)) `
        '遷移後の画像へ遷移前のDOM矩形を誤適用しない'
} finally {
    Remove-Item -LiteralPath $domCachePath -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 入力の検出に使うキー
# ---------------------------------------------------------------------
$typingKeys = @(Get-MbWatchedTypingKeys)
$commitKeys = @(Get-MbWatchedCommitKeys)
Add-Result ($typingKeys.Count -gt 30) '入力の検出に十分な数のキーを見る'
Add-Result ($typingKeys -contains 0x41) '英字を見る'
Add-Result ($typingKeys -contains 0x30) '数字を見る'
Add-Result ($typingKeys -contains 0x08) 'BackSpaceを見る'
# 修飾キーだけの操作を入力とみなすと、Ctrl+Cのたびに手順ができてしまう。
Add-Result ($typingKeys -notcontains 0x11) 'Ctrlだけでは入力とみなさない'
Add-Result ($typingKeys -notcontains 0x10) 'Shiftだけでは入力とみなさない'
Add-Result ($typingKeys -notcontains 0x09) 'Tabだけでは入力とみなさない'
Add-Result ($typingKeys -notcontains 0x0D) 'Enterだけでは入力とみなさない'
Add-Result ($commitKeys -contains 0x0D -and $commitKeys -contains 0x09) 'EnterとTabを入力確定として別に監視する'
Add-Result (-not (Test-MbTextChangingShortcutKey -VirtualKey 0x4C)) 'Ctrl+Lの移動を入力内容の変更とみなさない'
Add-Result (-not (Test-MbTextChangingShortcutKey -VirtualKey 0x46)) 'Ctrl+Fの検索開始を入力内容の変更とみなさない'
Add-Result (Test-MbTextChangingShortcutKey -VirtualKey 0x56) 'Ctrl+Vの貼り付けは入力内容の変更として残す'
Add-Result (Test-MbTextChangingShortcutKey -VirtualKey 0x5A) 'Ctrl+ZのUndoは入力内容の変更として残す'
Add-Result ((Resolve-MbTypingEventElapsedMs -CurrentElapsedMs 3100 -Clicked $true -ClickElapsedMs 3000) -eq 2999) `
    'クリックで確定した入力を後続クリックの直前へ並べる'
Add-Result ((Resolve-MbTypingEventElapsedMs -CurrentElapsedMs 3100 -Clicked $false -ClickElapsedMs 0) -eq 3100) `
    '待機で確定した入力は現在時刻を保つ'
$frequency = 10000000L
$queuedKey1 = [pscustomobject]@{ Kind = 1; WindowHandle = 101L; Timestamp = 10000000L }
$queuedKey2 = [pscustomobject]@{ Kind = 1; WindowHandle = 101L; Timestamp = 11000000L }
$queuedEnter = [pscustomobject]@{ Kind = 2; WindowHandle = 101L; Timestamp = 12000000L }
$queuedOtherWindow = [pscustomobject]@{ Kind = 1; WindowHandle = 202L; Timestamp = 11000000L }
$queuedLateKey = [pscustomobject]@{ Kind = 1; WindowHandle = 101L; Timestamp = 26000000L }
$queuedMouse = [pscustomobject]@{ Timestamp = 10500000L }
Add-Result (Test-MbQueuedKeyboardContinuation -CurrentActivity $queuedKey1 -NextActivity $queuedKey2 `
        -NextMouseClick $null -TypingIdleMs 1200 -TimestampFrequency $frequency) `
    '処理停止中に溜まった同じウィンドウの連続キーを1入力として保持する'
Add-Result (Test-MbQueuedKeyboardContinuation -CurrentActivity $queuedKey2 -NextActivity $queuedEnter `
        -NextMouseClick $null -TypingIdleMs 1200 -TimestampFrequency $frequency) `
    '連続キーの後に溜まったEnterまで入力確定を待つ'
Add-Result (Test-MbQueuedKeyboardContinuation -CurrentActivity $queuedKey1 -NextActivity $queuedOtherWindow `
        -NextMouseClick $null -TypingIdleMs 1200 -TimestampFrequency $frequency) `
    '別ウィンドウの入力が続くとき先に現在入力の境界処理へ渡す'
Add-Result (Test-MbQueuedKeyboardContinuation -CurrentActivity $queuedKey1 -NextActivity $null `
        -NextMouseClick $queuedMouse -TypingIdleMs 1200 -TimestampFrequency $frequency) `
    '入力直後のクリックが滞留してもidle確定で時系列を逆転させない'
Add-Result (-not (Test-MbQueuedKeyboardContinuation -CurrentActivity $queuedKey1 -NextActivity $queuedLateKey `
        -NextMouseClick $null -TypingIdleMs 1200 -TimestampFrequency $frequency)) `
    '十分に間が空いたキーは別入力として扱う'
Add-Result ($recorderSourceText -match '\$lastTypingMs \+ \$TypingIdleMs') `
    '滞留後のidle確定時刻を現在時刻ではなく実入力時刻から決める'

# タッチパッドの短いタップは次の巡回時には離されていることがある。
# その場合も GetAsyncKeyState の下位ビットから押下を拾う。
Add-Result ((Test-MbAsyncKeyStateDown -State 0x8000) -eq $true) '押されているキーを上位ビットで検出する'
Add-Result ((Test-MbAsyncKeyStatePressed -State 0x0001) -eq $true) '巡回の間に終わった短い押下を下位ビットで検出する'
Add-Result ((Test-MbAsyncKeyStateDown -State 0x0001) -eq $false) '離された短い押下を押下中とは扱わない'
Add-Result ((Test-MbAsyncKeyStatePressed -State 0x0000) -eq $false) '操作のない状態を押下とは扱わない'

# スクリーンショットやUI解析中のクリックを失わないよう、押下の受け取りだけは
# 独立スレッドの低レベルフックで行う。重い処理をコールバックへ入れないことも固定する。
Add-Result ($recorderSourceText -match 'ConcurrentQueue<MouseClick>' -and
    $recorderSourceText -match 'MouseClicks\.Enqueue' -and
    $recorderSourceText -match 'DequeueMouseClick') 'クリックを独立キューへ蓄積する'
Add-Result ($recorderSourceText -match 'Timestamp = Stopwatch\.GetTimestamp\(\)' -and
    $recorderSourceText -match '\$clickElapsedMs - \$preClickCaptureAtMs') `
    '実際のクリック時刻で画像バッファの前後関係を判定する'
Add-Result ($recorderSourceText -match 'if \(\$mouseHookActive\)' -and
    $recorderSourceText -match 'Test-MbAsyncKeyStatePressed -State \$leftState') `
    'フックを開始できない環境では従来の押下検出へ戻る'
Add-Result ($recorderSourceText -match 'ConcurrentQueue<KeyboardActivity>' -and
    $recorderSourceText -match 'DequeueKeyboardActivity' -and
    $recorderSourceText -match '\$hookTextActivity') `
    '重い画面取得中の短い入力も内容を保存せず独立キューで検出する'
Add-Result ($recorderSourceText -match 'public sealed class KeyboardActivity\s*\{\s*public int Kind;[\s\S]*?public long Timestamp;' -and
    $recorderSourceText -notmatch 'public sealed class KeyboardActivity[\s\S]*?public (?:int|uint) VirtualKey') `
    '入力検出キューへ実際のキー値を保存しない'
Add-Result ($recorderSourceText -match 'if \(!MouseHookStarted\)[\s\S]*?UnhookWindowsHookEx\(KeyboardHookHandle\)[\s\S]*?KeyboardHookStarted = false;') `
    'マウスフック開始失敗時にキーボードフックを残さない'

# ---------------------------------------------------------------------
# 実行環境で記録できるかどうか
# ---------------------------------------------------------------------
$capability = Get-MbRecorderCapability
Add-Result ($capability.PSObject.Properties.Name -contains 'available') '記録できるかどうかを判定できる'
if (-not $capability.available) {
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$capability.reason)) '記録できない場合は理由を返す'
    Write-Host ("     この環境では記録できません: " + [string]$capability.reason) -ForegroundColor Yellow
} else {
    Add-Result $true '記録に必要な機能がそろっている'
    # DPI認識は座標系を揃えるために欠かせない。呼べること自体を確かめる。
    $dpi = Set-MbProcessDpiAware
    Add-Result ($dpi -in @('per-monitor', 'system', 'none')) "DPI認識の設定を行える（$dpi）"
    $hookStarted = [MbRecorderNative]::StartMouseHook()
    Add-Result $hookStarted '独立したマウス記録スレッドを開始できる'
    [MbRecorderNative]::StopMouseHook()
}

# ---------------------------------------------------------------------
# 話した内容を操作へ振り分ける
# ---------------------------------------------------------------------
Import-Module (Join-Path $srcRoot 'ManualBuilder.RecorderServer.psm1') -Force

$ocrLabel = & (Get-Module ManualBuilder.RecorderServer) { ConvertTo-MbRecorderOcrLabel -Value '詳 細 を 表 示' }
$ocrBackLabel = & (Get-Module ManualBuilder.RecorderServer) { ConvertTo-MbRecorderOcrLabel -Value '- 覧 へ 戻 る' }
Add-Result ($ocrLabel -eq '詳細を表示' -and $ocrBackLabel -eq '一覧へ戻る') `
    'Windows OCRが分割・誤認した日本語ボタン名を整える'
$ocrConfidence = & (Get-Module ManualBuilder.RecorderServer) { Get-MbRecorderOcrLabelConfidence -Label '詳細を表示' }
$mixedOcrConfidence = & (Get-Module ManualBuilder.RecorderServer) { Get-MbRecorderOcrLabelConfidence -Label '自 i 2 ロ n' }
$acronymOcrConfidence = & (Get-Module ManualBuilder.RecorderServer) { Get-MbRecorderOcrLabelConfidence -Label 'CSVを出力' }
Add-Result ($ocrConfidence -eq 'medium' -and $mixedOcrConfidence -eq 'low' -and $acronymOcrConfidence -eq 'medium') `
    '日本語OCRの崩れた日英交互列だけを要確認へ下げる'
$ocrInputValue = & (Get-Module ManualBuilder.RecorderServer) {
    Test-MbRecorderUnreliableOcrLabel -Label '01042' -ActionKind 'input' -Source 'click-point+OCR'
}
$ocrExcelError = & (Get-Module ManualBuilder.RecorderServer) {
    Test-MbRecorderUnreliableOcrLabel -Label '#CALC!' -ActionKind 'click' -Source 'click-point+OCR'
}
$ocrFragment = & (Get-Module ManualBuilder.RecorderServer) {
    Test-MbRecorderUnreliableOcrLabel -Label '数' -ActionKind 'click' -Source 'click-point+OCR'
}
$ocrNormalButton = & (Get-Module ManualBuilder.RecorderServer) {
    Test-MbRecorderUnreliableOcrLabel -Label '詳細を表示' -ActionKind 'click' -Source 'click-point+OCR'
}
Add-Result ($ocrInputValue -and $ocrExcelError -and $ocrFragment -and -not $ocrNormalButton) `
    '入力値・Excelエラー・1文字断片をOCRの操作対象名として断定しない'
$ocrInputSequence = @(
    [pscustomobject]@{
        index = 1; timeMs = 1000; kind = 'click'; targetName = '注文番号'; targetType = 'ControlType.OcrText'
        targetSource = 'click-point+OCR'; confidence = 'medium'; windowTitle = '受注検索 - Edge'
        rect = [pscustomobject]@{ x1 = 0.2; y1 = 0.2; x2 = 0.3; y2 = 0.25 }
    },
    [pscustomobject]@{ index = 2; timeMs = 2600; kind = 'input'; targetName = ''; targetType = ''; windowTitle = '受注検索 - Edge' }
)
$repairedOcrInput = @(& (Get-Module ManualBuilder.RecorderServer) {
    param($Events) Repair-MbRecorderUnlabeledInputAnchors -Events $Events
} $ocrInputSequence)
Add-Result ([string]$repairedOcrInput[1].targetName -eq '注文番号' -and
    [string]$repairedOcrInput[1].targetType -eq 'ControlType.Edit' -and
    $repairedOcrInput[1].rect.x1 -eq 0.2) 'OCRクリック直後の入力へ対象名と赤枠を引き継ぐ'
Add-Result (@(Merge-MbRecordedEditInteractions -Events $repairedOcrInput).Count -eq 1) `
    'OCRで復元した入力欄クリックと入力を1手順へまとめる'
$buttonThenInput = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; targetName = '検索'; targetType = 'ControlType.Button'; windowTitle = '検索 - Edge' },
    [pscustomobject]@{ index = 2; timeMs = 1500; kind = 'input'; targetName = ''; targetType = ''; windowTitle = '検索 - Edge' }
)
$buttonInputResult = @(& (Get-Module ManualBuilder.RecorderServer) {
    param($Events) Repair-MbRecorderUnlabeledInputAnchors -Events $Events
} $buttonThenInput)
Add-Result ([string]::IsNullOrWhiteSpace([string]$buttonInputResult[1].targetName)) `
    '通常ボタンの名前を後続入力へ誤って引き継がない'

$trustedUiaEvent = [pscustomobject]@{
    kind = 'click'; targetName = '検索'; targetType = 'ControlType.Button'; targetSource = 'UIA'; confidence = 'medium'
    rect = [pscustomobject]@{ x1 = 0.2; y1 = 0.3; x2 = 0.32; y2 = 0.38 }
    clickPoint = [pscustomobject]@{ x = 0.25; y = 0.34 }
}
$trustedUia = & (Get-Module ManualBuilder.RecorderServer) {
    param($Event) Test-MbRecorderTrustedLocalTarget -Event $Event -ActionKind 'click'
} $trustedUiaEvent
Add-Result ([bool]$trustedUia) 'クリック点と一致する小さなUIA対象だけを自動採用できる'
$cachedUiaEvent = $trustedUiaEvent.PSObject.Copy(); $cachedUiaEvent.targetSource = 'UIA-CACHE'
$trustedCachedUia = & (Get-Module ManualBuilder.RecorderServer) {
    param($Event) Test-MbRecorderTrustedLocalTarget -Event $Event -ActionKind 'click'
} $cachedUiaEvent
Add-Result ([bool]$trustedCachedUia) '同じ画面・同じウィンドウで保持したクリック前UIAも座標一致時だけ自動採用できる'
$outsideUiaEvent = $trustedUiaEvent.PSObject.Copy()
$outsideUiaEvent.clickPoint = [pscustomobject]@{ x = 0.7; y = 0.7 }
$outsideUia = & (Get-Module ManualBuilder.RecorderServer) {
    param($Event) Test-MbRecorderTrustedLocalTarget -Event $Event -ActionKind 'click'
} $outsideUiaEvent
Add-Result (-not [bool]$outsideUia) 'クリック点を含まないUIA対象は自動採用しない'
$domEvent = $trustedUiaEvent.PSObject.Copy(); $domEvent.targetSource = 'DOM'
$trustedDom = & (Get-Module ManualBuilder.RecorderServer) {
    param($Event) Test-MbRecorderTrustedLocalTarget -Event $Event -ActionKind 'click'
} $domEvent
Add-Result (-not [bool]$trustedDom) 'DOM候補を短い確認のために安易に自動採用しない'
$trustedInput = & (Get-Module ManualBuilder.RecorderServer) {
    param($Event) Test-MbRecorderTrustedLocalTarget -Event $Event -ActionKind 'input'
} $trustedUiaEvent
Add-Result (-not [bool]$trustedInput) '入力内容を記録しない入力手順は引き続き確認対象にする'

# 記録一覧で操作前／操作後を関連付け、取り込み後も別画像として保持する。
$importRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderImport-' + [guid]::NewGuid().ToString('N'))
$beforeBitmap = $null; $afterBitmap = $null
try {
    $eventDirectory = Join-Path $importRoot 'events'
    [void](New-Item -ItemType Directory -Path $eventDirectory -Force)
    $beforePath = Join-Path $eventDirectory 'event-001.jpg'
    $afterPath = Join-Path $eventDirectory 'event-001-result.jpg'
    $excelPath = Join-Path $eventDirectory 'event-002.jpg'
    $beforeBitmap = New-Object Drawing.Bitmap -ArgumentList @(64, 40)
    $afterBitmap = New-Object Drawing.Bitmap -ArgumentList @(64, 40)
    $beforeGraphics = [Drawing.Graphics]::FromImage($beforeBitmap)
    $afterGraphics = [Drawing.Graphics]::FromImage($afterBitmap)
    try {
        $beforeGraphics.Clear([Drawing.Color]::White)
        $afterGraphics.Clear([Drawing.Color]::LightBlue)
        $beforeBitmap.Save($beforePath, [Drawing.Imaging.ImageFormat]::Jpeg)
        $afterBitmap.Save($afterPath, [Drawing.Imaging.ImageFormat]::Jpeg)
        $beforeBitmap.Save($excelPath, [Drawing.Imaging.ImageFormat]::Jpeg)
    } finally { $beforeGraphics.Dispose(); $afterGraphics.Dispose() }
    $eventsPath = Join-Path $importRoot 'events.jsonl'
    $edgeEvidenceId = 'evidence-' + [guid]::NewGuid().ToString('N')
    $excelEvidenceId = 'evidence-' + [guid]::NewGuid().ToString('N')
    $edgeEventJson = [pscustomobject]@{
        index = 1; kind = 'click'; timeMs = 1000; image = 'event-001.jpg'
        evidenceId = $edgeEvidenceId
        windowTitle = '申請画面 - Microsoft Edge'; targetName = '詳細を表示'; targetType = 'ControlType.Button'
        rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.3; y2 = 0.2 }
    } | ConvertTo-Json -Compress -Depth 5
    $excelEventJson = [pscustomobject]@{
        index = 2; kind = 'click'; timeMs = 2200; image = 'event-002.jpg'
        evidenceId = $excelEvidenceId
        windowTitle = 'Book1 - Excel'; targetName = 'F8'; targetType = 'ControlType.Cell'
        rect = [pscustomobject]@{ x1 = 0.4; y1 = 0.4; x2 = 0.5; y2 = 0.5 }
    } | ConvertTo-Json -Compress -Depth 5
    [IO.File]::WriteAllLines($eventsPath, @($edgeEventJson, $excelEventJson), [Text.UTF8Encoding]::new($false))
    $recordingJobId = 'record-' + [guid]::NewGuid().ToString('N')
    $evidenceDirectory = Join-Path $importRoot 'evidence-source'
    [void](New-Item -ItemType Directory -Path $evidenceDirectory -Force)
    Copy-Item -LiteralPath $beforePath -Destination (Join-Path $evidenceDirectory ($edgeEvidenceId + '.jpg'))
    Copy-Item -LiteralPath $excelPath -Destination (Join-Path $evidenceDirectory ($excelEvidenceId + '.jpg'))
    $ledgerPath = Join-Path $importRoot 'evidence-ledger.jsonl'
    [IO.File]::WriteAllLines($ledgerPath, @(
        ([ordered]@{ recordType='capture-start'; formatVersion=2; sessionId=$recordingJobId; mouseHook=$true; keyboardHook=$true; completeness='no-known-gaps' } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$edgeEvidenceId; sessionId=$recordingJobId; kind='click'; timeMs=1000; image=($edgeEvidenceId + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$excelEvidenceId; sessionId=$recordingJobId; kind='click'; timeMs=2200; image=($excelEvidenceId + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='capture-end'; formatVersion=2; sessionId=$recordingJobId; operationCount=2; reason='stopped'; completeness='no-known-gaps'; warning='' } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    $job = [pscustomobject]@{
        JobId = $recordingJobId; EventsPath = $eventsPath; EventsDirectory = $eventDirectory
        EvidenceDirectory = $evidenceDirectory; LedgerPath = $ledgerPath
    }
    $oldLedgerPath = Join-Path $importRoot 'old-evidence-ledger.jsonl'
    [IO.File]::WriteAllLines($oldLedgerPath, @(
        ([ordered]@{ recordType='operation'; id=$edgeEvidenceId; sessionId=$recordingJobId; kind='click'; timeMs=1000; image=($edgeEvidenceId + '.jpg') } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    $oldJob = [pscustomobject]@{
        JobId = $recordingJobId; EventsPath = $eventsPath; EventsDirectory = $eventDirectory
        EvidenceDirectory = $evidenceDirectory; LedgerPath = $oldLedgerPath
    }
    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $oldJob
    $oldFormatRejected = $false
    try {
        $oldProject = New-MbProject
        [void](Import-MbRecordedEvents -Project $oldProject -ProjectPath (Join-Path $importRoot 'old-project.json') -SheetId $oldProject.sheets[0].id)
    } catch {
        $oldFormatRejected = $_.Exception.Message -like '*現在の証拠形式ではありません*'
    }
    Add-Result $oldFormatRejected '旧録画形式を推測で取り込まず新形式での記録を求める'

    $validLedgerLines = @([IO.File]::ReadAllLines($ledgerPath, [Text.Encoding]::UTF8))
    $duplicateLedgerPath = Join-Path $importRoot 'duplicate-start-ledger.jsonl'
    [IO.File]::WriteAllLines($duplicateLedgerPath, @($validLedgerLines[0], $validLedgerLines[0]) + @($validLedgerLines[1..3]), [Text.UTF8Encoding]::new($false))
    $duplicateJob = $job.PSObject.Copy(); $duplicateJob.LedgerPath = $duplicateLedgerPath
    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $duplicateJob
    $duplicateStartRejected = $false
    try {
        [void](& (Get-Module ManualBuilder.RecorderServer) {
            param($Project, $Path) Import-MbRecordedEvidenceSession -Project $Project -ProjectPath $Path
        } (New-MbProject) (Join-Path $importRoot 'duplicate-project.json'))
    } catch { $duplicateStartRejected = $_.Exception.Message -like '*現在の証拠形式ではありません*' }
    Add-Result $duplicateStartRejected '重複したcapture-startを取り込み前に拒否する'

    $mismatchLedgerPath = Join-Path $importRoot 'count-mismatch-ledger.jsonl'
    [IO.File]::WriteAllLines($mismatchLedgerPath, @(
        $validLedgerLines[0], $validLedgerLines[1], $validLedgerLines[2],
        ([ordered]@{ recordType='capture-end'; formatVersion=2; sessionId=$recordingJobId; operationCount=1; reason='stopped'; completeness='no-known-gaps'; warning='' } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    $mismatchJob = $job.PSObject.Copy(); $mismatchJob.LedgerPath = $mismatchLedgerPath
    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $mismatchJob
    $countMismatchRejected = $false
    try {
        [void](& (Get-Module ManualBuilder.RecorderServer) {
            param($Project, $Path) Import-MbRecordedEvidenceSession -Project $Project -ProjectPath $Path
        } (New-MbProject) (Join-Path $importRoot 'mismatch-project.json'))
    } catch { $countMismatchRejected = $_.Exception.Message -like '*件数が終了台帳と一致しません*' }
    Add-Result $countMismatchRejected 'capture-end件数とoperation台帳件数の不一致を拒否する'

    $foreignSessionLedgerPath = Join-Path $importRoot 'foreign-session-ledger.jsonl'
    $foreignOperation = $validLedgerLines[1] | ConvertFrom-Json
    $foreignOperation.sessionId = 'record-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
    [IO.File]::WriteAllLines($foreignSessionLedgerPath, @(
        $validLedgerLines[0], ($foreignOperation | ConvertTo-Json -Compress), $validLedgerLines[2], $validLedgerLines[3]
    ), [Text.UTF8Encoding]::new($false))
    $foreignSessionJob = $job.PSObject.Copy(); $foreignSessionJob.LedgerPath = $foreignSessionLedgerPath
    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $foreignSessionJob
    $foreignSessionRejected = $false
    try {
        [void](& (Get-Module ManualBuilder.RecorderServer) {
            param($Project, $Path) Import-MbRecordedEvidenceSession -Project $Project -ProjectPath $Path
        } (New-MbProject) (Join-Path $importRoot 'foreign-project.json'))
    } catch { $foreignSessionRejected = $_.Exception.Message -like '*セッション情報が一致しません*' }
    Add-Result $foreignSessionRejected '別セッションのoperation混入を拒否する'

    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $job
    $listedEvents = @(Get-MbRecordedEvents)
    Add-Result ($listedEvents.Count -eq 2) 'EdgeからExcelへ移った操作をどちらも確認一覧へ残す'
    Add-Result ([string]$listedEvents[0].windowTitle -like '*Edge' -and [string]$listedEvents[1].windowTitle -like '*Excel') '複数アプリの操作順を保持する'
    Add-Result ([string]$listedEvents[0].resultImage -eq 'event-001-result.jpg') '確認一覧でクリック後画像を同じ操作へ関連付ける'

    $importProject = New-MbProject
    $importProjectPath = Join-Path $importRoot 'project.json'
    $importProject = Save-MbProject -Project $importProject -Path $importProjectPath
    $imported = Import-MbRecordedEvents -Project $importProject -ProjectPath $importProjectPath -SheetId $importProject.sheets[0].id
    $importedStep = @($importProject.sheets[0].steps)[0]
    Add-Result ([int]$imported.added -eq 2) 'EdgeとExcelの操作をまとめて手順へ取り込む'
    Add-Result ([int]$imported.generated -eq 2) '取り込みと同時にローカルで文章を作る'
    Add-Result (@($importProject.sheets[0].steps).Count -eq 2) '複数アプリの手順を取り込み時に欠落させない'
    $importedSteps = @($importProject.sheets[0].steps)
    Add-Result ([string]$importedSteps[0].capture.windowTitle -like '*Edge' -and [string]$importedSteps[1].capture.windowTitle -like '*Excel') '取り込み後もEdgeからExcelへの操作順を保持する'
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$importedStep.resultImageId)) '取り込み後も操作後画像を手順へ保持する'
    Add-Result ([string]$importedStep.imageId -ne [string]$importedStep.resultImageId) '操作前画像と操作後画像を混同しない'
    Add-Result ([string]$importedStep.imageLayout -eq 'before') '操作後画像を保持しても初稿は案内画像1枚で取り込む'
    Add-Result ([string]$importedStep.title -eq '詳細を表示' -and [string]$importedStep.description -eq '［詳細を表示］をクリックします。') '操作対象名から編集可能な初稿を作る'
    Add-Result ([double]$importedStep.crop.x -eq 0 -and [double]$importedStep.crop.y -eq 0 -and
        [double]$importedStep.crop.width -eq 1 -and [double]$importedStep.crop.height -eq 1) '取り込み時は画面全体を残し自動拡大しない'
    Add-Result (@($importProject.evidenceSessions).Count -eq 1 -and
        [string]$importProject.evidenceSessions[0].id -eq $recordingJobId -and
        [int]$importProject.evidenceSessions[0].operationCount -eq 2 -and
        [string]$importProject.evidenceSessions[0].captureCompleteness -eq 'no-known-gaps') `
        '取り込み元の操作証拠セッションと完全性状態をプロジェクトへ保存する'
    Add-Result ([string]$importedStep.capture.sourceSessionId -eq $recordingJobId -and
        @($importedStep.capture.evidenceIds) -contains $edgeEvidenceId -and
        -not [string]::IsNullOrWhiteSpace([string]$importedStep.capture.transformationReason)) `
        '完成手順から元操作と変換理由をたどれる'
    $archiveRoot = Join-Path (Join-Path $importRoot 'evidence') $recordingJobId
    Add-Result ((Test-Path -LiteralPath (Join-Path $archiveRoot 'evidence-ledger.jsonl') -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path (Join-Path $archiveRoot 'images') ($edgeEvidenceId + '.jpg')) -PathType Leaf)) `
        '作業用記録を片付けても元台帳と証拠画像がプロジェクト側に残る'
    $decisionsPath = Join-Path $archiveRoot 'transformations.jsonl'
    $decisions = @([IO.File]::ReadAllLines($decisionsPath) | ForEach-Object { $_ | ConvertFrom-Json })
    Add-Result ($decisions.Count -eq 2 -and @($decisions | Where-Object accepted).Count -eq 2 -and
        @($decisions | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.reason) }).Count -eq 2) `
        '各候補の採否と変換理由を追記履歴へ保存する'
    $importProject = Save-MbProject -Project $importProject -Path $importProjectPath
    Add-Result (@($importProject.evidenceSessions).Count -eq 1) '証拠参照を含むプロジェクトを検証して再保存できる'

    $gapJobId = 'record-' + [guid]::NewGuid().ToString('N')
    $gapLedgerPath = Join-Path $importRoot 'gap-evidence-ledger.jsonl'
    [IO.File]::WriteAllLines($gapLedgerPath, @(
        ([ordered]@{ recordType='capture-start'; formatVersion=2; sessionId=$gapJobId; mouseHook=$true; keyboardHook=$true; completeness='no-known-gaps' } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$edgeEvidenceId; sessionId=$gapJobId; kind='click'; timeMs=1000; image=($edgeEvidenceId + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$excelEvidenceId; sessionId=$gapJobId; kind='click'; timeMs=2200; image=($excelEvidenceId + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='capture-gap'; formatVersion=2; sessionId=$gapJobId; droppedMouseClicks=1; droppedKeyboardActivities=0; reason='capture-queue-overflow' } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='capture-end'; formatVersion=2; sessionId=$gapJobId; operationCount=2; reason='stopped'; completeness='known-gaps'; warning='操作が短時間に集中しました。' } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    $gapJob = [pscustomobject]@{
        JobId = $gapJobId; EventsPath = $eventsPath; EventsDirectory = $eventDirectory
        EvidenceDirectory = $evidenceDirectory; LedgerPath = $gapLedgerPath
    }
    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $gapJob
    $gapProject = New-MbProject
    $gapProjectPath = Join-Path $importRoot 'gap-project.json'
    [void](Import-MbRecordedEvents -Project $gapProject -ProjectPath $gapProjectPath -SheetId $gapProject.sheets[0].id)
    Add-Result ([string]$gapProject.evidenceSessions[0].captureCompleteness -eq 'known-gaps' -and
        [string]$gapProject.evidenceSessions[0].captureWarning -eq '操作が短時間に集中しました。') `
        '既知の取りこぼしを証拠セッションへ保存して確認対象にできる'
} finally {
    & (Get-Module ManualBuilder.RecorderServer) { $script:MbRecordingJob = $null }
    if ($beforeBitmap) { $beforeBitmap.Dispose() }
    if ($afterBitmap) { $afterBitmap.Dispose() }
    Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$practicalEvents = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; targetType = 'ControlType.Edit'; targetName = '検索'; windowTitle = 'Explorer' },
    [pscustomobject]@{ index = 2; timeMs = 2400; kind = 'input'; targetType = 'ControlType.Edit'; targetName = '検索'; windowTitle = 'Explorer' },
    [pscustomobject]@{ index = 3; timeMs = 4000; kind = 'click'; targetType = 'ControlType.Button'; targetName = '並べ替え'; windowTitle = 'Explorer' },
    [pscustomobject]@{ index = 4; timeMs = 4700; kind = 'click'; targetType = 'ControlType.MenuItem'; targetName = '更新日時'; windowTitle = 'Explorer' }
)
$mergedPracticalEvents = @(Merge-MbRecordedEditInteractions -Events $practicalEvents)
Add-Result ($mergedPracticalEvents.Count -eq 3) '入力欄のクリックと直後の入力を1手順へまとめる'
Add-Result ([int]$mergedPracticalEvents[0].index -eq 2) '入力済み画面を持つ入力手順を残す'
Add-Result ([int]$mergedPracticalEvents[1].index -eq 3 -and [int]$mergedPracticalEvents[2].index -eq 4) `
    '並べ替えメニューと選択肢は別々の操作として残す'
$slowEditEvents = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; targetType = 'ControlType.Edit'; targetName = '検索'; windowTitle = 'Explorer' },
    [pscustomobject]@{ index = 2; timeMs = 8000; kind = 'input'; targetType = 'ControlType.Edit'; targetName = '検索'; windowTitle = 'Explorer' }
)
Add-Result (@(Merge-MbRecordedEditInteractions -Events $slowEditEvents).Count -eq 2) '間を置いた入力欄クリックは独立操作として残す'
$excelCellInputEvents = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; targetType = 'ControlType.DataItem'; targetName = 'B2'; windowTitle = 'Book1 - Excel'; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.2; x2 = 0.2; y2 = 0.3 } },
    [pscustomobject]@{ index = 2; timeMs = 2800; kind = 'input'; targetType = 'ControlType.DataItem'; targetName = 'B2'; windowTitle = 'Book1 - Excel'; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.2; x2 = 0.2; y2 = 0.3 } }
)
$mergedExcelCellInputEvents = @(Merge-MbRecordedEditInteractions -Events $excelCellInputEvents)
Add-Result ($mergedExcelCellInputEvents.Count -eq 1 -and [string]$mergedExcelCellInputEvents[0].kind -eq 'input') `
    'Excelセルの選択と直後の入力を1手順へまとめる'
$sameNameDifferentFields = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; targetType = 'ControlType.Edit'; targetName = '値'; windowTitle = '設定'; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.3; y2 = 0.2 } },
    [pscustomobject]@{ index = 2; timeMs = 2000; kind = 'input'; targetType = 'ControlType.Edit'; targetName = '値'; windowTitle = '設定'; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.6; x2 = 0.3; y2 = 0.7 } }
)
Add-Result (@(Merge-MbRecordedEditInteractions -Events $sameNameDifferentFields).Count -eq 2) '同じ名前でも位置が違う入力欄を誤ってまとめない'

$unknownFieldRect = [pscustomobject]@{ x1 = 0.20; y1 = 0.20; x2 = 0.24; y2 = 0.24 }
$unknownFieldEvents = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; targetType = 'ControlType.ClickPoint'; targetName = ''; windowTitle = '業務アプリ'; rect = $unknownFieldRect },
    [pscustomobject]@{ index = 2; timeMs = 2300; kind = 'input'; targetType = 'ControlType.ClickPoint'; targetName = ''; windowTitle = '業務アプリ'; rect = $unknownFieldRect }
)
$mergedUnknownFieldEvents = @(Merge-MbRecordedEditInteractions -Events $unknownFieldEvents)
Add-Result ($mergedUnknownFieldEvents.Count -eq 1 -and [string]$mergedUnknownFieldEvents[0].kind -eq 'input') `
    '対象名を取得できない入力欄も同じクリック位置なら1手順へまとめる'

$repeatedClicks = @(
    [pscustomobject]@{ index = 1; timeMs = 1000; kind = 'click'; windowTitle = 'Book1 - Excel'; targetName = 'F8'; targetType = 'ControlType.DataItem' },
    [pscustomobject]@{ index = 2; timeMs = 1150; kind = 'click'; windowTitle = 'Book1 - Excel'; targetName = 'F8'; targetType = 'ControlType.DataItem' },
    [pscustomobject]@{ index = 3; timeMs = 1300; kind = 'click'; windowTitle = 'Book1 - Excel'; targetName = 'F8'; targetType = 'ControlType.DataItem' }
)
Add-Result (@(Merge-MbRecordedEditInteractions -Events $repeatedClicks).Count -eq 3) `
    '似た連続クリックを復元不能な形で自動削除しない'

$candidateProject = New-MbProject
$candidateStep = Add-MbStep -Project $candidateProject -SheetId $candidateProject.sheets[0].id
[void](Set-MbStepCapture -Project $candidateProject -StepId $candidateStep.id -TargetCandidateId 'missing' `
    -TargetCandidatesJson '[{"id":"dom-1","source":"DOM","confidence":"medium","label":"送信","targetType":"ControlType.Button","rect":{"x1":0.4,"y1":0.4,"x2":0.6,"y2":0.5}},{"id":"dom-1","source":"UIA","confidence":"low","label":"重複","targetType":"ControlType.Text","rect":{"x1":0.4,"y1":0.4,"x2":0.6,"y2":0.5}}]' `
    -ClickPointJson '{"x":0.52,"y":0.48}')
$candidateStep = Get-MbStepById -Project $candidateProject -StepId $candidateStep.id
Add-Result (@($candidateStep.capture.targetCandidates).Count -eq 1) '重複する候補IDを保存しない'
Add-Result ([string]$candidateStep.capture.targetCandidateId -eq 'dom-1') '現在候補が一覧外なら保存済みの先頭候補へ戻す'
Add-Result ([Math]::Abs([double]$candidateStep.capture.clickPoint.x - 0.52) -lt 0.001 -and
    [Math]::Abs([double]$candidateStep.capture.clickPoint.y - 0.48) -lt 0.001) 'クリック位置を手順へ引き継ぐ'

$bulkProject = New-MbProject
$sourceSheet = $bulkProject.sheets[0]
$bulkStep1 = Add-MbStep -Project $bulkProject -SheetId $sourceSheet.id
$bulkStep2 = Add-MbStep -Project $bulkProject -SheetId $sourceSheet.id
$targetSheet = Add-MbSheet -Project $bulkProject
[void](Move-MbStepsToSheet -Project $bulkProject -StepIds @($bulkStep1.id, $bulkStep2.id) -TargetSheetId $targetSheet.id)
Add-Result (@($sourceSheet.steps).Count -eq 0 -and @($targetSheet.steps).Count -eq 2) '複数手順をまとめて別シートへ移動する'
Add-Result ([string]$targetSheet.steps[0].id -eq [string]$bulkStep1.id -and [string]$targetSheet.steps[1].id -eq [string]$bulkStep2.id) '複数手順の順序を保って移動する'
Remove-MbSteps -Project $bulkProject -StepIds @($bulkStep1.id, $bulkStep2.id)
Add-Result (@($targetSheet.steps).Count -eq 0) '複数手順をまとめて削除する'

# ---------------------------------------------------------------------
Write-Host ''
if ($errors.Count -eq 0) {
    Write-Host '操作記録の検査はすべて成功しました。' -ForegroundColor Green
    exit 0
}
Write-Host ("失敗: " + $errors.Count + " 件") -ForegroundColor Red
exit 1
