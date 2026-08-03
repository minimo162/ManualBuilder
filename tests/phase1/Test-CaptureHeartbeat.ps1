# Phase 1 capture heartbeat state machine test.
# ブラウザーは裏へ回ったタブのタイマーを1分に1回まで間引く。撮影のたびにタブが裏へ回る
# ManualBuilderでは、猶予が短いと正常なタブでも失効し、監視が止まって撮影ぶんを取りこぼす。
# ここではサーバー本体から状態遷移の関数だけを取り出し、間引き・凍結・引き継ぎを再現する。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$serverPath = Join-Path $repoRoot 'src\Start-ManualBuilder.ps1'
$serverText = [IO.File]::ReadAllText($serverPath, [Text.Encoding]::UTF8)
$tokens = $null
$parseErrors = $null
$serverAst = [System.Management.Automation.Language.Parser]::ParseInput($serverText, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw "Start-ManualBuilder.ps1 を解析できません: $($parseErrors[0].Message)" }

foreach ($functionName in @('Get-MbCaptureRole', 'Update-MbCaptureHeartbeatState', 'Set-MbCaptureHeartbeat', 'Test-MbWatchCandidate', 'Invoke-MbWatcherFlush')) {
    $found = @($serverAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    }, $true))
    if ($found.Count -ne 1) { throw "サーバー本体に関数が見つかりません: $functionName" }
    Invoke-Expression $found[0].Extent.Text
}

function Write-MbLog {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host "     $Level : $Message" -ForegroundColor DarkGray
}

# サーバー本体の既定値と同じ条件で確かめる。
$HeartbeatTimeoutSec = 90
$CaptureStandbySec = 900
$script:Watcher = [pscustomobject]@{ Name = 'test-watcher' }
$script:WatchDirectory = 'C:\ManualBuilderTest\Screenshots'
$script:WatcherState = 'suspended'
$script:CaptureOwnerTab = $null
$script:CaptureOwnerLastHeartbeat = $null
$script:CaptureOwnerSheetId = $null
$script:ImportWatermark = Get-Date
$script:PendingImages = New-Object System.Collections.ArrayList
$script:PendingImageLimit = 200
$script:WatcherEventIds = @('ManualBuilder.Test.Created', 'ManualBuilder.Test.Renamed')

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Add-MbTestPending {
    param([string]$Path)
    [void]$script:PendingImages.Add([pscustomobject]@{ Path = $Path; Size = [long]-1; Tries = 0 })
}

function Set-MbTestSilence {
    param([double]$Seconds)
    $script:CaptureOwnerLastHeartbeat = (Get-Date).AddSeconds(-$Seconds)
}

$role = Set-MbCaptureHeartbeat -TabId 'tab-A' -SheetId 'sheet-1'
Assert-Mb ($role -eq 'owner') '最初のタブが撮影対象になる'
Assert-Mb ($script:WatcherState -eq 'active') 'ハートビート受信で監視中になる'
$watermark = $script:ImportWatermark

Set-MbTestSilence 60
Update-MbCaptureHeartbeatState
Assert-Mb ($script:WatcherState -eq 'active') '裏タブの間引き（60秒間隔）では監視が止まらない'

Set-MbTestSilence 120
Update-MbCaptureHeartbeatState
Assert-Mb ($script:WatcherState -eq 'standby') '猶予を超えると保留へ移る'

Add-MbTestPending 'C:\ManualBuilderTest\Screenshots\shot-1.png'
Invoke-MbWatcherFlush
Assert-Mb ($script:PendingImages.Count -eq 1) '保留中の新着を破棄しない'
Assert-Mb ([int]$script:PendingImages[0].Tries -eq 0) '保留中は再試行回数を進めない'
Assert-Mb ($script:ImportWatermark -eq $watermark) '保留中は取り込み基準時刻を動かさない'

$role = Set-MbCaptureHeartbeat -TabId 'tab-A' -SheetId 'sheet-1'
Assert-Mb ($role -eq 'owner') '戻ってきたタブが撮影対象のまま'
Assert-Mb ($script:WatcherState -eq 'active') 'タブが戻ると監視を再開する'
Assert-Mb ($script:PendingImages.Count -eq 1) '保留していた新着が残っている'
Assert-Mb ($script:ImportWatermark -eq $watermark) '再開時に取り込み基準時刻を引き直さない'
Invoke-MbWatcherFlush
Assert-Mb ([int]$script:PendingImages[0].Tries -eq 1) '再開後は取り込み処理が進む'

Set-MbTestSilence 1200
Update-MbCaptureHeartbeatState
Assert-Mb ($script:WatcherState -eq 'suspended') '保留の期限を超えると一時停止する'
Assert-Mb ($script:PendingImages.Count -eq 0) '一時停止で保留を破棄する'

$script:PendingImages.Clear()
[void](Set-MbCaptureHeartbeat -TabId 'tab-A' -SheetId 'sheet-1')
Add-MbTestPending 'C:\ManualBuilderTest\Screenshots\shot-2.png'
$watermark = $script:ImportWatermark
Start-Sleep -Milliseconds 20
Set-MbTestSilence 120
$role = Set-MbCaptureHeartbeat -TabId 'tab-B' -SheetId 'sheet-1'
Assert-Mb ($role -eq 'owner') '失効後は別タブが撮影対象を引き継ぐ'
Assert-Mb ($script:PendingImages.Count -eq 0) '引き継ぎ時は前のタブ向けの保留を捨てる'
Assert-Mb ($script:ImportWatermark -gt $watermark) '引き継ぎ時は取り込み基準時刻を引き直す'

$role = Set-MbCaptureHeartbeat -TabId 'tab-C' -SheetId 'sheet-1'
Assert-Mb ($role -eq 'viewer') '撮影対象が生きている間、別タブは閲覧専用になる'

$script:PendingImages.Clear()
[void](Set-MbCaptureHeartbeat -TabId 'tab-B' -SheetId 'sheet-1')
Set-MbTestSilence 120
Update-MbCaptureHeartbeatState
for ($i = 0; $i -lt ($script:PendingImageLimit + 5); $i++) {
    Add-MbTestPending ("C:\ManualBuilderTest\Screenshots\bulk-$i.png")
}
Invoke-MbWatcherFlush
Assert-Mb ($script:PendingImages.Count -eq $script:PendingImageLimit) '保留は上限件数で頭打ちになる'
Assert-Mb ([string]$script:PendingImages[0].Path -eq 'C:\ManualBuilderTest\Screenshots\bulk-5.png') '上限を超えたぶんは古い保留から捨てる'

Write-Host ''
Write-Host 'Capture heartbeat tests passed.' -ForegroundColor Cyan
