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

$sliver = [pscustomobject]@{ left = 500.0; top = 400.0; width = 0.5; height = 30.0 }
Add-Result ($null -eq (ConvertTo-MbRegionRect -Region $capturedRegion -Target $sliver)) '潰れた矩形は赤枠にしない'

Add-Result ($null -eq (ConvertTo-MbRegionRect -Region $capturedRegion -Target $null)) '操作対象が無ければ矩形は作らない'

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
