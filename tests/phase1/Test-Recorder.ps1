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
    $cacheValue.updatedAtUtc = [DateTime]::UtcNow.AddMilliseconds(-800).ToString('o')
    [IO.File]::WriteAllText($uiaCachePath, ($cacheValue | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
    $otherWindow = [pscustomobject]@{ handle = 54321; left = 0; top = 0; width = 1000; height = 700 }
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
$smallPrimary = Select-MbRecordingPrimaryTarget -Target $smallButton -PointTarget $pointFallback -Window $window
Add-Result (-not [bool]$smallPrimary.isFallback -and [string]$smallPrimary.name -eq '保存') '妥当なUIAボタンはクリック点へ置き換えない'

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
Add-Result ($typingKeys.Count -gt 30) '入力の検出に十分な数のキーを見る'
Add-Result ($typingKeys -contains 0x41) '英字を見る'
Add-Result ($typingKeys -contains 0x30) '数字を見る'
Add-Result ($typingKeys -contains 0x08) 'BackSpaceを見る'
# 修飾キーだけの操作を入力とみなすと、Ctrl+Cのたびに手順ができてしまう。
Add-Result ($typingKeys -notcontains 0x11) 'Ctrlだけでは入力とみなさない'
Add-Result ($typingKeys -notcontains 0x10) 'Shiftだけでは入力とみなさない'
Add-Result ($typingKeys -notcontains 0x09) 'Tabだけでは入力とみなさない'
Add-Result (-not (Test-MbTextChangingShortcutKey -VirtualKey 0x4C)) 'Ctrl+Lの移動を入力内容の変更とみなさない'
Add-Result (-not (Test-MbTextChangingShortcutKey -VirtualKey 0x46)) 'Ctrl+Fの検索開始を入力内容の変更とみなさない'
Add-Result (Test-MbTextChangingShortcutKey -VirtualKey 0x56) 'Ctrl+Vの貼り付けは入力内容の変更として残す'
Add-Result (Test-MbTextChangingShortcutKey -VirtualKey 0x5A) 'Ctrl+ZのUndoは入力内容の変更として残す'

# タッチパッドの短いタップは次の巡回時には離されていることがある。
# その場合も GetAsyncKeyState の下位ビットから押下を拾う。
Add-Result ((Test-MbAsyncKeyStateDown -State 0x8000) -eq $true) '押されているキーを上位ビットで検出する'
Add-Result ((Test-MbAsyncKeyStatePressed -State 0x0001) -eq $true) '巡回の間に終わった短い押下を下位ビットで検出する'
Add-Result ((Test-MbAsyncKeyStateDown -State 0x0001) -eq $false) '離された短い押下を押下中とは扱わない'
Add-Result ((Test-MbAsyncKeyStatePressed -State 0x0000) -eq $false) '操作のない状態を押下とは扱わない'

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
}

# ---------------------------------------------------------------------
# 話した内容を操作へ振り分ける
# ---------------------------------------------------------------------
Import-Module (Join-Path $srcRoot 'ManualBuilder.RecorderServer.psm1') -Force

# 記録一覧で操作前／操作後を関連付け、取り込み後も別画像として保持する。
$importRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderImport-' + [guid]::NewGuid().ToString('N'))
$beforeBitmap = $null; $afterBitmap = $null
try {
    $eventDirectory = Join-Path $importRoot 'events'
    [void](New-Item -ItemType Directory -Path $eventDirectory -Force)
    $beforePath = Join-Path $eventDirectory 'event-001.jpg'
    $afterPath = Join-Path $eventDirectory 'event-001-result.jpg'
    $beforeBitmap = New-Object Drawing.Bitmap -ArgumentList @(64, 40)
    $afterBitmap = New-Object Drawing.Bitmap -ArgumentList @(64, 40)
    $beforeGraphics = [Drawing.Graphics]::FromImage($beforeBitmap)
    $afterGraphics = [Drawing.Graphics]::FromImage($afterBitmap)
    try {
        $beforeGraphics.Clear([Drawing.Color]::White)
        $afterGraphics.Clear([Drawing.Color]::LightBlue)
        $beforeBitmap.Save($beforePath, [Drawing.Imaging.ImageFormat]::Jpeg)
        $afterBitmap.Save($afterPath, [Drawing.Imaging.ImageFormat]::Jpeg)
    } finally { $beforeGraphics.Dispose(); $afterGraphics.Dispose() }
    $eventsPath = Join-Path $importRoot 'events.jsonl'
    $eventJson = [pscustomobject]@{
        index = 1; kind = 'click'; timeMs = 1000; image = 'event-001.jpg'
        windowTitle = '詳細'; targetName = '詳細を表示'; targetType = 'ControlType.Button'
        rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.3; y2 = 0.2 }
    } | ConvertTo-Json -Compress -Depth 5
    [IO.File]::WriteAllText($eventsPath, $eventJson + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    $job = [pscustomobject]@{ EventsPath = $eventsPath; EventsDirectory = $eventDirectory; NarrationPath = (Join-Path $importRoot 'narration.jsonl') }
    & (Get-Module ManualBuilder.RecorderServer) { param($Job) $script:MbRecordingJob = $Job } $job
    $listedEvents = @(Get-MbRecordedEvents)
    Add-Result ([string]$listedEvents[0].resultImage -eq 'event-001-result.jpg') '確認一覧でクリック後画像を同じ操作へ関連付ける'

    $importProject = New-MbProject
    $importProjectPath = Join-Path $importRoot 'project.json'
    $importProject = Save-MbProject -Project $importProject -Path $importProjectPath
    $imported = Import-MbRecordedEvents -Project $importProject -ProjectPath $importProjectPath -SheetId $importProject.sheets[0].id
    $importedStep = @($importProject.sheets[0].steps)[0]
    Add-Result ([int]$imported.added -eq 1) '操作後画像つきの記録を手順へ取り込む'
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$importedStep.resultImageId)) '取り込み後も操作後画像を手順へ保持する'
    Add-Result ([string]$importedStep.imageId -ne [string]$importedStep.resultImageId) '操作前画像と操作後画像を混同しない'
    Add-Result ([string]$importedStep.imageLayout -eq 'side-by-side') '操作後画像つきの記録を左右比較で取り込む'
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

$focusRect = [pscustomobject]@{ x1 = 0.7; y1 = 0.7; x2 = 0.8; y2 = 0.75 }
$focusCrop = Get-MbRecorderTargetCrop -Rect $focusRect -TargetType 'ControlType.Button'
Add-Result ($null -ne $focusCrop -and [double]$focusCrop.width -eq 0.55 -and [double]$focusCrop.height -eq 0.55) '特定できた操作対象の周辺を初期表示する'
Add-Result ([double]$focusCrop.x -ge 0 -and ([double]$focusCrop.x + [double]$focusCrop.width) -le 1) '対象周辺の切り抜きを画像内へ収める'
$fallbackCrop = Get-MbRecorderTargetCrop -Rect $focusRect -TargetType 'ControlType.ClickPoint'
Add-Result ($null -eq $fallbackCrop) '対象不明のクリックは自動で切り抜かない'

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

$events = @(
    [pscustomobject]@{ index = 1; timeMs = 5000 },
    [pscustomobject]@{ index = 2; timeMs = 12000 },
    [pscustomobject]@{ index = 3; timeMs = 20000 }
)

# 「ここで申請ボタンを押します」と言ってから押す。次に来る操作の説明になる。
$before = @([pscustomobject]@{ startMs = 2000; endMs = 4000; text = 'ここで申請ボタンを押します' })
$mapped = Merge-MbNarrationIntoEvents -Events $events -Phrases $before
Add-Result ($mapped.ContainsKey(1)) '操作の前に話した内容はその操作へ付く'
Add-Result ([string]$mapped[1] -eq 'ここで申請ボタンを押します') '発話がそのまま入る'
Add-Result (-not $mapped.ContainsKey(2)) '別の操作には付かない'

# 押してからすぐ「これで一覧に出ました」と言う。直前の操作への補足になる。
$after = @([pscustomobject]@{ startMs = 5800; endMs = 7500; text = 'これで一覧に出ました' })
$mapped = Merge-MbNarrationIntoEvents -Events $events -Phrases $after
Add-Result ($mapped.ContainsKey(1)) '操作の直後に話した内容は直前の操作へ付く'
Add-Result (-not $mapped.ContainsKey(2)) '直後の発話が次の操作へ流れない'

# どの操作からも遠い独り言は捨てる。
$stray = @([pscustomobject]@{ startMs = 30000; endMs = 31000; text = 'ええと' })
$mapped = Merge-MbNarrationIntoEvents -Events $events -Phrases $stray
Add-Result ($mapped.Count -eq 0) 'どの操作からも離れた発話は捨てる'

# 1つの操作について複数回話した場合はつなげる。
$multiple = @(
    [pscustomobject]@{ startMs = 9000; endMs = 10000; text = '次に金額を入れます' },
    [pscustomobject]@{ startMs = 10200; endMs = 11500; text = '税込で入力します' }
)
$mapped = Merge-MbNarrationIntoEvents -Events $events -Phrases $multiple
Add-Result ($mapped.ContainsKey(2)) '複数の発話が同じ操作へ付く'
Add-Result ([string]$mapped[2] -eq '次に金額を入れます 税込で入力します') '発話を話した順につなげる'

# 同じ発話が複数の手順に出ると読みにくい。1つの操作にだけ付ける。
$single = @([pscustomobject]@{ startMs = 4000; endMs = 4500; text = '押します' })
$mapped = Merge-MbNarrationIntoEvents -Events $events -Phrases $single
Add-Result ($mapped.Count -eq 1) '1つの発話は1つの操作にしか付かない'

# 空の入力で落ちないこと。
Add-Result ((Merge-MbNarrationIntoEvents -Events @() -Phrases $before).Count -eq 0) '操作が無ければ何も返さない'
Add-Result ((Merge-MbNarrationIntoEvents -Events $events -Phrases @()).Count -eq 0) '発話が無ければ何も返さない'

# ---------------------------------------------------------------------
# 音声入力が使えるかどうか
# ---------------------------------------------------------------------
Import-Module (Join-Path $srcRoot 'ManualBuilder.Dictation.psm1') -Force
$dictation = Get-MbDictationCapability
Add-Result ($dictation.PSObject.Properties.Name -contains 'available') '音声入力の可否を判定できる'
if ($dictation.available) {
    Add-Result ([string]$dictation.language -like 'ja*') '日本語の音声入力を使う'
} else {
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$dictation.reason)) '使えない場合は対処が分かる理由を返す'
    Write-Host ("     この環境では音声入力を使えません: " + [string]$dictation.reason) -ForegroundColor Yellow
}

# 認識結果から記録を作る部分。PhraseStartTime が取れない場合の代用も見る。
$startedAt = [DateTime]::UtcNow
$fake = [pscustomobject]@{
    Text = '申請ボタンを押します'
    Status = 'Success'
    Confidence = 'Medium'
    PhraseStartTime = [DateTimeOffset]::new($startedAt.AddSeconds(3))
    PhraseDuration = [TimeSpan]::FromMilliseconds(1500)
}
$record = ConvertTo-MbDictationRecord -Result $fake -StartedAtUtc $startedAt -ReceivedAtMs 6000
Add-Result ($null -ne $record) '認識結果から記録を作れる'
Add-Result ([Math]::Abs([int]$record.startMs - 3000) -le 50) '発話の開始時刻を使う（受信時刻ではない）'
Add-Result ([int]$record.endMs -eq ([int]$record.startMs + 1500)) '発話の長さから終了時刻を出す'

$empty = [pscustomobject]@{
    Text = '   '; Status = 'Success'; Confidence = 'High'
    PhraseStartTime = [DateTimeOffset]::new($startedAt); PhraseDuration = [TimeSpan]::FromSeconds(1)
}
Add-Result ($null -eq (ConvertTo-MbDictationRecord -Result $empty -StartedAtUtc $startedAt -ReceivedAtMs 1000)) '空の発話は記録しない'

# ---------------------------------------------------------------------
Write-Host ''
if ($errors.Count -eq 0) {
    Write-Host '操作記録の検査はすべて成功しました。' -ForegroundColor Green
    exit 0
}
Write-Host ("失敗: " + $errors.Count + " 件") -ForegroundColor Red
exit 1
