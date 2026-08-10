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

# 一覧画像モジュールは別プロセスでも単独読込されるため、他モジュールの
# System.Drawing初期化へ依存していないことを最初に確認する。
Import-Module (Join-Path $srcRoot 'ManualBuilder.RecorderCopilot.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.RecorderServer.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Project.psm1') -Force
Add-Result ((Get-MbRecorderWindowAppKey -WindowTitle '申請 - Google Chrome') -eq 'chrome') `
    'processNameがない旧録画でもGoogle Chromeのタイトルからアプリを判定する'

$currentRecorderProcess = Get-Process -Id $PID
try {
    $recorderIdentity = & (Get-Module ManualBuilder.RecorderServer) { param($Process) New-MbRecorderProcessIdentity -Process $Process } $currentRecorderProcess
    $recorderIdentityMatches = & (Get-Module ManualBuilder.RecorderServer) { param($Identity) Test-MbRecorderProcessIdentity -Identity $Identity } $recorderIdentity
    $reusedRecorderIdentity = $recorderIdentity | Select-Object *
    $reusedRecorderIdentity.StartTimeUtcTicks = [int64]$reusedRecorderIdentity.StartTimeUtcTicks + 1
    $reusedRecorderIdentityRejected = -not (& (Get-Module ManualBuilder.RecorderServer) { param($Identity) Test-MbRecorderProcessIdentity -Identity $Identity } $reusedRecorderIdentity)
    Add-Result ($recorderIdentityMatches -and $reusedRecorderIdentityRejected) `
        '録画workerはPIDだけでなく名前・開始時刻・実行ファイルが一致する所有プロセスだけを扱う'
} finally { $currentRecorderProcess.Dispose() }

$copilotModuleSource = Get-Content -LiteralPath (Join-Path $srcRoot 'ManualBuilder.RecorderCopilot.psm1') -Raw -Encoding UTF8
$recorderServerSource = Get-Content -LiteralPath (Join-Path $srcRoot 'ManualBuilder.RecorderServer.psm1') -Raw -Encoding UTF8
$recorderSource = Get-Content -LiteralPath (Join-Path $srcRoot 'ManualBuilder.Recorder.psm1') -Raw -Encoding UTF8
Add-Result ($recorderSource -match 'GetWindowThreadProcessId' -and
    $recorderSource -match 'processName\s*=\s*\$\(if' -and $recorderSource -match 'windowClass\s*=\s*\$\(if') `
    '記録フレームとイベントへ前面アプリのプロセス識別子を保存する'

$testRoot = Join-Path $env:TEMP ('ManualBuilder-RecorderCopilot-' + [guid]::NewGuid().ToString('N'))
$framesDirectory = Join-Path $testRoot 'frames'
$contactDirectory = Join-Path $testRoot 'contact-sheets'
$projectPath = Join-Path $testRoot 'project.json'
$framesPath = Join-Path $testRoot 'frames.jsonl'
$eventsPath = Join-Path $testRoot 'events.jsonl'
$bitmap = $null
$graphics = $null
try {
    [void](New-Item -ItemType Directory -Path $framesDirectory -Force)
    $frameLines = New-Object System.Collections.ArrayList
    for ($index = 1; $index -le 25; $index++) {
        $fileName = 'frame-{0:d5}.jpg' -f $index
        $bitmap = New-Object Drawing.Bitmap -ArgumentList @(640, 360)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([Drawing.Color]::FromArgb(20 + ($index * 5), 50 + ($index * 3), 90 + ($index * 2)))
        $graphics.DrawString(('FRAME {0:d5}' -f $index), [Drawing.SystemFonts]::CaptionFont, [Drawing.Brushes]::White, 24, 24)
        $bitmap.Save((Join-Path $framesDirectory $fileName), [Drawing.Imaging.ImageFormat]::Jpeg)
        $graphics.Dispose(); $graphics = $null
        $bitmap.Dispose(); $bitmap = $null
        [void]$frameLines.Add(([pscustomobject]@{
            id = 'F{0:d5}' -f $index
            index = $index
            timeMs = ($index - 1) * 500
            image = $fileName
            windowTitle = $(if ($index -le 12) { '申請画面 - Microsoft Edge' } else { 'Book1 - Excel' })
        } | ConvertTo-Json -Compress))
    }
    [IO.File]::WriteAllLines($framesPath, @($frameLines), [Text.UTF8Encoding]::new($false))
    $testEvidenceIds = @(
        ('evidence-' + [guid]::NewGuid().ToString('N')),
        ('evidence-' + [guid]::NewGuid().ToString('N'))
    )
    $eventLines = @(
        ([pscustomobject]@{ index = 1; evidenceId = $testEvidenceIds[0]; timeMs = 2100; kind = 'click'; targetName = '詳細を表示'; targetType = 'ControlType.Button'; windowTitle = '申請画面 - Microsoft Edge'; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.2; x2 = 0.3; y2 = 0.3 } } | ConvertTo-Json -Compress -Depth 5),
        ([pscustomobject]@{ index = 2; evidenceId = $testEvidenceIds[1]; timeMs = 7100; kind = 'input'; targetName = 'F7'; targetType = 'ControlType.DataItem'; windowTitle = 'Book1 - Excel'; rect = $null } | ConvertTo-Json -Compress -Depth 5)
    )
    [IO.File]::WriteAllLines($eventsPath, $eventLines, [Text.UTF8Encoding]::new($false))

    $frames = @(Read-MbRecorderJsonLines -Path $framesPath)
    $events = @(Read-MbRecorderJsonLines -Path $eventsPath)
    $selected = @(Select-MbRecorderTimelineFrames -Frames $frames -Events $events -Maximum 20)
    Add-Result ($selected.Count -gt 1 -and $selected.Count -le 20) 'AIへ渡すフレームを重要場面だけに絞る'
    $singleSheetSelection = @(Select-MbRecorderTimelineFrames -Frames $frames -Events $events -Maximum 15)
    Add-Result ($singleSheetSelection.Count -gt 1 -and $singleSheetSelection.Count -le 15) '1枚分の上限でもイベント前後を優先して絞る'
    $eventWindowSelection = @(Select-MbRecorderEventWindowFrames -Frames $frames -Events $events)
    Add-Result ($eventWindowSelection.Count -gt 1 -and [int]$eventWindowSelection[0].timeMs -ge 600 -and
        [int]$eventWindowSelection[-1].timeMs -le 9100) '最初と最後の操作から離れた準備・終了後画面を一覧から外す'
    $foreignFrame = [pscustomobject]@{ id = 'FX'; index = 99; timeMs = 7200; image = 'frame-00001.jpg'; windowTitle = 'ChatGPT' }
    $eventWindowWithForeign = @(Select-MbRecorderEventWindowFrames -Frames @($frames + $foreignFrame) -Events $events)
    Add-Result (@($eventWindowWithForeign | Where-Object { $_.id -eq 'FX' }).Count -eq 0) `
        '操作イベントにないChatGPTやManualBuilderへの復帰画面を一覧から外す'
    Add-Result (@($selected | Where-Object { [Math]::Abs([int]$_.timeMs - 2100) -le 600 }).Count -gt 0) 'クリック時刻付近の原本フレームを残す'
    Add-Result (@($selected | Where-Object { [Math]::Abs([int]$_.timeMs - 7100) -le 600 }).Count -gt 0) '入力時刻付近の原本フレームを残す'

    $copilotSourceFrames = @(Select-MbRecorderCopilotSourceFrames -Frames $frames -Events $events -Maximum 30)
    $localOnlyFrames = @(Select-MbRecorderCandidateFrames -Frames $frames -Candidates @(
        [pscustomobject]@{ beforeFrame = 'F00003'; afterFrame = 'F00019' }
    ))
    Add-Result ($copilotSourceFrames.Count -gt $localOnlyFrames.Count -and
        @($copilotSourceFrames | Where-Object { $_.id -eq 'F00010' }).Count -eq 1 -and
        @($localOnlyFrames | Where-Object { $_.id -eq 'F00010' }).Count -eq 0) `
        'ローカル候補から漏れた中間コマもCopilot用の時系列原本へ残す'
    $excelOnlyEvents = @($events | Where-Object { [string]$_.windowTitle -match 'Excel$' })
    $missingAppSource = @(Select-MbRecorderCopilotSourceFrames -Frames $frames -Events $excelOnlyEvents -Maximum 30)
    Add-Result (@($missingAppSource | Where-Object { [string]$_.windowTitle -match 'Edge$' }).Count -gt 0) `
        'イベントを全て取り逃したアプリの周期コマもCopilot用原本から除外しない'
    $manualBuilderNamedPageFrames = @(
        [pscustomobject]@{ id='N00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='受注検索 - ManualBuilder Recorder Smoke - Microsoft Edge'; processName='msedge'; visualChange=0.0 },
        [pscustomobject]@{ id='N00002'; index=2; timeMs=1500; image='frame-00002.jpg'; windowTitle='ManualBuilder Recorder Smoke および他 1 ページ - Microsoft Edge'; processName='msedge'; visualChange=0.1 },
        [pscustomobject]@{ id='N00003'; index=3; timeMs=2000; image='frame-00003.jpg'; windowTitle='ManualBuilder - Google Chrome'; processName='chrome'; visualChange=0.1 }
    )
    $manualBuilderNamedPageSelection = @(Select-MbRecorderCopilotSourceFrames -Frames $manualBuilderNamedPageFrames -Events @() -Maximum 8)
    Add-Result ($manualBuilderNamedPageSelection.Count -eq 2 -and
        @($manualBuilderNamedPageSelection | Where-Object { [string]$_.id -in @('N00001','N00002') }).Count -eq 2) `
        'ManualBuilderを名称に含む業務画面を本体タブと誤認せず、タイトル先頭が本体の画面だけ除外する'

    $longFrames = @($frames) + @($frames | ForEach-Object {
        $copy = $_ | Select-Object *
        $copy.index = [int]$copy.index + 25
        $copy.timeMs = [int]$copy.timeMs + 12500
        $copy.id = 'F{0:d5}' -f [int]$copy.index
        $copy
    })
    $fastSelection = @(Select-MbRecorderTimelineFrames -Frames $longFrames -Events $events -Maximum 30)
    Add-Result ($fastSelection.Count -gt 1 -and $fastSelection.Count -le 30) '長い記録も最大30コマへ絞る'
    $longSourceEvents = @($events) + @(
        [pscustomobject]@{ index = 3; timeMs = 22000; kind = 'click'; windowTitle = 'Book1 - Excel' }
    )
    $longCopilotSource = @(Select-MbRecorderCopilotSourceFrames -Frames $longFrames -Events $longSourceEvents -Maximum 30)
    $longSourceTimes = @($longCopilotSource | ForEach-Object { [int]$_.timeMs })
    $maximumGap = 0
    for ($sourceIndex = 1; $sourceIndex -lt $longSourceTimes.Count; $sourceIndex++) {
        $maximumGap = [Math]::Max($maximumGap, $longSourceTimes[$sourceIndex] - $longSourceTimes[$sourceIndex - 1])
    }
    Add-Result ($longCopilotSource.Count -eq 30 -and $maximumGap -le 1500) `
        '長い記録は原本30コマを時間軸全体へ分散してCopilotへ渡す'

    $priorityFrames = for ($index = 1; $index -le 40; $index++) {
        $frame = [pscustomobject]@{ id=('P{0:d5}' -f $index); index=$index; timeMs=$index * 500;
            image='frame-00001.jpg'; windowTitle='Book1 - Excel' }
        if ($index -in @(7, 21, 37)) { $frame | Add-Member -NotePropertyName role -NotePropertyValue 'input-evidence' }
        $frame
    }
    $priorityEvents = for ($index = 1; $index -le 20; $index++) {
        [pscustomobject]@{ index=$index; timeMs=$index * 900; kind='click'; windowTitle='Book1 - Excel' }
    }
    $prioritySelection = @(Select-MbRecorderTimelineFrames -Frames $priorityFrames -Events $priorityEvents -Maximum 8)
    $priorityIds = @($prioritySelection | ForEach-Object { [string]$_.id })
    Add-Result (@('P00007','P00021','P00037' | Where-Object { $_ -notin $priorityIds }).Count -eq 0) `
        '上限を超える長い記録でも全input-evidenceを等間隔候補より先に保護する'

    # 実機で失敗した時刻関係を固定する。入力途中ではなく、入力完了と
    # 次操作直前の安定画面が残ることを回帰テストにする。
    $realTimes = @(13813,14513,15130,15980,16652,17294,17914,18547,19212,19828,20713,21446,22081,22746,23428,24098,24715,25383,25981,26597,27230,27898,28581,29197,29798,30414,31082,31682)
    $realFrames = for ($index = 0; $index -lt $realTimes.Count; $index++) {
        $number = $index + 25
        [pscustomobject]@{ id = 'F{0:d5}' -f $number; index = $number; timeMs = $realTimes[$index]; image = 'frame-00001.jpg';
            windowTitle = $(if ($number -le 39) { '申請画面 - Microsoft Edge' } else { 'scenario.xlsx - Excel' }) }
    }
    $realEvents = @(
        [pscustomobject]@{ index=1; timeMs=15664; kind='click'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ index=2; timeMs=18046; kind='input'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ index=3; timeMs=18078; kind='click'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ index=4; timeMs=19727; kind='click'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ index=5; timeMs=20617; kind='click'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ index=6; timeMs=26992; kind='input'; windowTitle='scenario.xlsx - Excel' },
        [pscustomobject]@{ index=7; timeMs=27016; kind='click'; windowTitle='scenario.xlsx - Excel' },
        [pscustomobject]@{ index=8; timeMs=29071; kind='click'; windowTitle='scenario.xlsx - Excel' },
        [pscustomobject]@{ index=9; timeMs=31449; kind='input'; windowTitle='scenario.xlsx - Excel' }
    )
    $realSelection = @(Select-MbRecorderTimelineFrames -Frames $realFrames -Events $realEvents -Maximum 30)
    $realIds = @($realSelection | ForEach-Object { [string]$_.id })
    $requiredIds = @('F00027','F00031','F00033','F00034','F00035','F00042','F00044','F00047','F00051')
    Add-Result (@($requiredIds | Where-Object { $_ -notin $realIds }).Count -eq 0) `
        '実機で必要だった入力完了・検索結果・Excel操作前後の全コマを保護する'
    $transientIds = @('F00029','F00030','F00032','F00046')
    Add-Result ($realIds.Count -le 12 -and @($transientIds | Where-Object { $_ -in $realIds }).Count -eq 0) `
        '入力途中・ドロップダウン展開中のコマで上限を埋めない'

    # 数式入力の確定前画面は、単なる不安定コマではなく文章化に必要な証拠。
    # クリック+入力を1手順にし、数式が表示されたコマを操作後へ割り当てる。
    $meaningFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=900; image='frame-00001.jpg'; windowTitle='Book1 - Excel' },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=1300; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; role='input-evidence'; evidenceKind='paste' },
        [pscustomobject]@{ id='F00003'; index=3; timeMs=2500; image='frame-00003.jpg'; windowTitle='Book1 - Excel' },
        [pscustomobject]@{ id='F00004'; index=4; timeMs=3200; image='frame-00004.jpg'; windowTitle='Book1 - Excel' }
    )
    $meaningEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='B4'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} },
        [pscustomobject]@{ index=2; timeMs=2300; kind='input'; targetName='B4'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} },
        [pscustomobject]@{ index=3; timeMs=3000; kind='click'; targetName='通貨表示形式'; targetType='ControlType.Button'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.4;y1=.1;x2=.5;y2=.2} }
    )
    $meaningSelection = @(Select-MbRecorderTimelineFrames -Frames $meaningFrames -Events $meaningEvents -Maximum 8)
    $localCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $meaningFrames -Events $meaningEvents -MaximumFrames 8)
    Add-Result (@($meaningSelection | Where-Object { $_.id -eq 'F00002' }).Count -eq 1) `
        '数式など入力中の最終表示を安定画面より優先して保護する'
    Add-Result ($localCandidates.Count -eq 2 -and @($localCandidates[0].eventIds).Count -eq 2 -and
        [string]$localCandidates[0].beforeFrame -eq 'F00001' -and [string]$localCandidates[0].afterFrame -eq 'F00003') `
        'ローカル候補でもクリックと入力をまとめ、入力途中より完成した通常コマを操作後へ割り当てる'
    Add-Result ([int]$localCandidates[0].timeMs -eq 1000 -and [int]$localCandidates[1].timeMs -eq 3000) `
        '同じ操作前画像を共有してもイベント時刻で手順順序を保つ'

    $duplicateClickFrames = @(
        [pscustomobject]@{ id='D00001'; index=1; timeMs=800; image='frame-00001.jpg'; windowTitle='申請 - Microsoft Edge' },
        [pscustomobject]@{ id='D00002'; index=2; timeMs=1600; image='frame-00002.jpg'; windowTitle='申請 - Microsoft Edge' }
    )
    $duplicateClickEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='保存'; windowTitle='申請 - Microsoft Edge'; rect=[pscustomobject]@{x1=.10;y1=.10;x2=.20;y2=.20}; clickPoint=[pscustomobject]@{x=.15;y=.15} },
        [pscustomobject]@{ index=2; timeMs=1080; kind='click'; targetName=''; windowTitle='申請 - Microsoft Edge'; rect=[pscustomobject]@{x1=.11;y1=.11;x2=.21;y2=.21}; clickPoint=[pscustomobject]@{x=.16;y=.16} },
        [pscustomobject]@{ index=3; timeMs=1200; kind='click'; targetName='保存'; windowTitle='申請 - Microsoft Edge'; rect=$null; clickPoint=[pscustomobject]@{x=.17;y=.16} }
    )
    $duplicateClickCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $duplicateClickFrames `
        -Events $duplicateClickEvents -MaximumFrames 8)
    Add-Result ($duplicateClickCandidates.Count -eq 1 -and
        (@($duplicateClickCandidates[0].eventIds) -join ',') -eq '1,2,3') `
        '同じ対象を450ms以内に複数経路で記録したクリックは全eventIdsを持つ1候補にする'

    $differentTargetEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='保存'; windowTitle='申請 - Microsoft Edge'; rect=[pscustomobject]@{x1=.10;y1=.10;x2=.20;y2=.20} },
        [pscustomobject]@{ index=2; timeMs=1200; kind='click'; targetName='閉じる'; windowTitle='申請 - Microsoft Edge'; rect=[pscustomobject]@{x1=.70;y1=.10;x2=.80;y2=.20} }
    )
    $differentTargetCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $duplicateClickFrames `
        -Events $differentTargetEvents -MaximumFrames 8)
    Add-Result ($differentTargetCandidates.Count -eq 2 -and
        @($differentTargetCandidates[0].eventIds).Count -eq 1 -and
        @($differentTargetCandidates[1].eventIds).Count -eq 1) `
        '短時間のクリックでも対象が異なれば別候補のまま残す'

    $candidateFrames = @(Select-MbRecorderCandidateFrames -Frames $meaningFrames -Candidates $localCandidates)
    Add-Result ($candidateFrames.Count -eq 4 -and @($candidateFrames | Where-Object { $_.id -eq 'F00002' }).Count -eq 1 -and
        @($candidateFrames | Where-Object { $_.id -eq 'F00003' }).Count -eq 1) `
        'ローカル候補の比較用フレームでは貼り付け証拠も保持する'

    $missingEventFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=800; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=1400; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0006 },
        [pscustomobject]@{ id='F00003'; index=3; timeMs=2600; image='frame-00003.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0009 }
    )
    $missingEventEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='B2'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} }
    )
    $missingCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $missingEventFrames -Events $missingEventEvents -MaximumFrames 8)
    Add-Result (@($missingCandidates | Where-Object {
        $_.actionKind -eq 'visual-change' -and $_.afterFrame -eq 'F00003' -and
            [int]$_.targetEventId -eq 0 -and @($_.eventIds).Count -eq 0
    }).Count -eq 1) `
        'クリックイベントが欠けても画面差分から要確認候補を補う'

    $noEventFrames = @(
        [pscustomobject]@{ id='N00001'; index=1; timeMs=500; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='N00002'; index=2; timeMs=1100; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0008 },
        [pscustomobject]@{ id='N00003'; index=3; timeMs=1700; image='frame-00003.jpg'; windowTitle='読み込み中 - Excel'; visualChange=0.0020 }
    )
    $noEventCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $noEventFrames -Events @() -MaximumFrames 8)
    Add-Result ($noEventCandidates.Count -eq 1 -and
        [string]$noEventCandidates[0].beforeFrame -eq 'N00001' -and
        [string]$noEventCandidates[0].afterFrame -eq 'N00002' -and
        [string]$noEventCandidates[0].actionKind -eq 'visual-change' -and
        [int]$noEventCandidates[0].targetEventId -eq 0 -and @($noEventCandidates[0].eventIds).Count -eq 0) `
        '操作イベントが0件でも安定した画像差分を赤枠なしの要確認候補として回収する'

    $visualEpisodeFrames = @(
        [pscustomobject]@{ id='V00001'; index=1; timeMs=500; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='V00002'; index=2; timeMs=1000; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0008 },
        [pscustomobject]@{ id='V00003'; index=3; timeMs=2300; image='frame-00003.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0012 }
    )
    $visualEpisodeCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $visualEpisodeFrames `
        -Events @() -MaximumFrames 8)
    Add-Result ($visualEpisodeCandidates.Count -eq 1 -and
        [string]$visualEpisodeCandidates[0].beforeFrame -eq 'V00001' -and
        [string]$visualEpisodeCandidates[0].afterFrame -eq 'V00003' -and
        @($visualEpisodeCandidates[0].eventIds).Count -eq 0 -and
        [int]$visualEpisodeCandidates[0].targetEventId -eq 0) `
        '同じアプリで1500ms以内に続く画面変化を最初から最後までの1episodeにまとめる'

    $leadingGapFrames = @(
        [pscustomobject]@{ id='L00001'; index=1; timeMs=800; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='L00002'; index=2; timeMs=1400; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0009 },
        [pscustomobject]@{ id='L00003'; index=3; timeMs=3300; image='frame-00003.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='L00004'; index=4; timeMs=4300; image='frame-00004.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 }
    )
    $leadingGapEvents = @(
        [pscustomobject]@{ index=1; timeMs=3500; kind='click'; targetName='保存'; targetType='ControlType.Button'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.7;y1=.1;x2=.8;y2=.2} }
    )
    $leadingGapCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $leadingGapFrames -Events $leadingGapEvents -MaximumFrames 8)
    Add-Result (@($leadingGapCandidates | Where-Object {
        $_.actionKind -eq 'visual-change' -and $_.beforeFrame -eq 'L00001' -and $_.afterFrame -eq 'L00002' -and
            [int]$_.targetEventId -eq 0
    }).Count -eq 1) `
        '最初の操作イベントより前に取り逃した画面変化も要確認候補として回収する'

    $longNoEventFrames = for ($index = 1; $index -le 24; $index++) {
        [pscustomobject]@{
            id=('Z{0:d5}' -f $index); index=$index; timeMs=$index * 500; image='frame-00001.jpg';
            windowTitle='Book1 - Excel'; visualChange=$(if ($index -in @(4, 21)) { 0.001 } else { 0.0 })
        }
    }
    $longNoEventSelection = @(Select-MbRecorderTimelineFrames -Frames $longNoEventFrames -Events @() -Maximum 8)
    $longNoEventIds = @($longNoEventSelection | ForEach-Object { [string]$_.id })
    Add-Result (@('Z00003','Z00004','Z00020','Z00021' | Where-Object { $_ -notin $longNoEventIds }).Count -eq 0) `
        'イベント0件の長い記録でも変化画像と直前画像を等間隔間引きより先に保護する'

    $coveredInputFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=800; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=1400; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0012 },
        [pscustomobject]@{ id='F00003'; index=3; timeMs=1900; image='frame-00003.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0014 },
        [pscustomobject]@{ id='F00004'; index=4; timeMs=2300; image='frame-00004.jpg'; windowTitle='Book1 - Excel'; role='input-evidence'; visualChange=0.0001 },
        [pscustomobject]@{ id='F00005'; index=5; timeMs=2600; image='frame-00005.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0008 }
    )
    $coveredInputEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='B4'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} },
        [pscustomobject]@{ index=2; timeMs=2500; kind='input'; targetName='B4'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} }
    )
    $coveredCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $coveredInputFrames -Events $coveredInputEvents -MaximumFrames 8)
    Add-Result ($coveredCandidates.Count -eq 1 -and [string]$coveredCandidates[0].actionKind -eq 'input' -and
        [string]$coveredCandidates[0].afterFrame -eq 'F00005') `
        '入力イベントが取れた区間は完成した通常コマを使い、途中差分を別手順にしない'

    $inputOnlyCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $coveredInputFrames -Events @(
        [pscustomobject]@{ index=2; timeMs=2500; kind='input'; targetName='B3'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.2;y1=.3;x2=.3;y2=.4} }
    ) -MaximumFrames 8)
    Add-Result ($inputOnlyCandidates.Count -eq 1 -and [int]$inputOnlyCandidates[0].targetEventId -eq 0) `
        'クリックを取り逃した入力へ古いセルの赤枠アンカーを付けない'

    $excelSequenceEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='B2'; targetType='ControlType.DataItem'; processName='EXCEL'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.10;y1=.20;x2=.20;y2=.24} },
        [pscustomobject]@{ index=2; timeMs=2000; kind='input'; targetName=''; targetType=''; processName='EXCEL'; windowTitle='Book1 - Excel'; rect=$null },
        [pscustomobject]@{ index=3; timeMs=3000; kind='click'; targetName='B4'; targetType='ControlType.DataItem'; processName='EXCEL'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.10;y1=.28;x2=.20;y2=.32} }
    )
    $repairedExcelSequenceEvents = @(Repair-MbRecorderExcelInputEventAnchors -Events $excelSequenceEvents)
    $repairedExcelCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $coveredInputFrames -Events $repairedExcelSequenceEvents -MaximumFrames 8)
    Add-Result ([string]$repairedExcelSequenceEvents[1].targetName -eq 'B3' -and
        [Math]::Abs([double]$repairedExcelSequenceEvents[1].rect.y1 - .24) -lt .000001 -and
        @($repairedExcelCandidates | Where-Object { [int]$_.targetEventId -eq 2 }).Count -eq 1) `
        '同じ列の前後セルで一意に決まる取り逃しだけ、実測矩形の中間セルへ復元する'

    $duplicateImageFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; imageSha256='AAA' },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=1200; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; imageSha256='AAA' }
    )
    $deduplicatedFrames = @(Select-MbRecorderCandidateFrames -Frames $duplicateImageFrames -Candidates @(
        [pscustomobject]@{ beforeFrame='F00001'; afterFrame='F00002' }
    ))
    Add-Result ($deduplicatedFrames.Count -eq 1) '短時間に連続した同一内容の別ID画像をCopilot一覧へ重複掲載しない'

    $revisitedFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='Browser'; imageSha256='AAA' },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=5000; image='frame-00002.jpg'; windowTitle='Browser'; imageSha256='AAA' }
    )
    $revisitedSelection = @(Select-MbRecorderCandidateFrames -Frames $revisitedFrames -Candidates @(
        [pscustomobject]@{ beforeFrame='F00001'; afterFrame='F00002' }
    ))
    Add-Result ($revisitedSelection.Count -eq 2) '一覧へ戻るなど時間を置いて再訪した同一画面は操作後として残す'

    $clickEvidenceFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='Browser'; role='click-evidence'; evidenceEventId=1; visualChange=0.0 },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=1500; image='frame-00002.jpg'; windowTitle='Browser'; visualChange=0.0 },
        [pscustomobject]@{ id='F00003'; index=3; timeMs=2000; image='frame-00003.jpg'; windowTitle='Browser'; role='click-evidence'; evidenceEventId=2; visualChange=0.0 },
        [pscustomobject]@{ id='F00004'; index=4; timeMs=2600; image='frame-00004.jpg'; windowTitle='Browser'; visualChange=0.0 }
    )
    $clickEvidenceEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='検索'; targetType='ControlType.Button'; windowTitle='Browser'; rect=[pscustomobject]@{x1=.1;y1=.1;x2=.2;y2=.2} },
        [pscustomobject]@{ index=2; timeMs=2000; kind='click'; targetName='詳細'; targetType='ControlType.Button'; windowTitle='Browser'; rect=[pscustomobject]@{x1=.2;y1=.2;x2=.3;y2=.3} }
    )
    $clickEvidenceCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $clickEvidenceFrames -Events $clickEvidenceEvents -MaximumFrames 8)
    Add-Result ($clickEvidenceCandidates.Count -eq 2 -and
        [string]$clickEvidenceCandidates[0].beforeFrame -eq 'F00001' -and
        [string]$clickEvidenceCandidates[0].afterFrame -eq 'F00002' -and
        [string]$clickEvidenceCandidates[1].beforeFrame -eq 'F00002') `
        '操作後は次操作のclick-evidenceを越えず直前の安定画面にする'

    $inputBoundaryFrames = @(
        [pscustomobject]@{ id='I00001'; index=1; timeMs=800; image='frame-00001.jpg'; windowTitle='Book1 - Excel' },
        [pscustomobject]@{ id='I00002'; index=2; timeMs=1500; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; role='input-evidence' },
        [pscustomobject]@{ id='I00003'; index=3; timeMs=2100; image='frame-00003.jpg'; windowTitle='Book1 - Excel' },
        [pscustomobject]@{ id='I00004'; index=4; timeMs=2600; image='frame-00004.jpg'; windowTitle='Book1 - Excel'; role='click-evidence'; evidenceEventId=3 },
        [pscustomobject]@{ id='I00005'; index=5; timeMs=2800; image='frame-00005.jpg'; windowTitle='Book1 - Excel' }
    )
    $inputBoundaryEvents = @(
        [pscustomobject]@{ index=1; timeMs=900; kind='click'; targetName='B2'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} },
        [pscustomobject]@{ index=2; timeMs=2200; kind='input'; targetName='B2'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} },
        [pscustomobject]@{ index=3; timeMs=2600; kind='click'; targetName='B3'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.3;x2=.2;y2=.4} }
    )
    $inputBoundaryCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $inputBoundaryFrames -Events $inputBoundaryEvents -MaximumFrames 8)
    Add-Result ([string]$inputBoundaryCandidates[0].afterFrame -eq 'I00003' -and
        [int](@($inputBoundaryFrames | Where-Object { $_.id -eq $inputBoundaryCandidates[0].afterFrame })[0].timeMs) -lt 2600) `
        '入力手順の操作後画像も次の操作時刻を越えない'

    $sheets = @(New-MbRecorderContactSheets -Frames $frames -FramesDirectory $framesDirectory -OutputDirectory $contactDirectory)
    Add-Result ($sheets.Count -eq 3) '25フレームを複数の番号付き一覧画像へ分ける'
    $sheetImage = [Drawing.Image]::FromFile([string]$sheets[0].path)
    try { Add-Result ($sheetImage.Width -eq 1920 -and $sheetImage.Height -eq 2870) '一覧画像を小さな値も読める2列の固定寸法で作る' }
    finally { $sheetImage.Dispose() }
    $lastSheetImage = [Drawing.Image]::FromFile([string]$sheets[$sheets.Count - 1].path)
    try { Add-Result ($lastSheetImage.Width -eq 1920 -and $lastSheetImage.Height -eq 1722) '最終ページを空の10コマ分まで水増ししない' }
    finally { $lastSheetImage.Dispose() }

    $prompt = New-MbRecorderCopilotPrompt -Frames $frames -Events $events -PacketNumber 2 -TotalPackets 2 -PreviousFrameId 'F00015' -Marker 'MB_TEST_END'
    Add-Result ($prompt -match '中間的なアニメーション' -and $prompt -match 'ManualBuilderへ戻る操作') '遷移中画像と記録終了操作を除く判断基準を伝える'
    Add-Result ($prompt -match '対象が異なる入力イベントは省略せず' -and $prompt -match '前面アプリが切り替わっただけ') '別セル入力を残しアプリ切替を手順にしない基準を伝える'
    Add-Result ($prompt -match '大文字小文字は変えず') '画面に見える値の表記を変えない基準を伝える'
    Add-Result ($prompt -match '数式バー' -and $prompt -match '=SUM\(B2:B3\)') 'Excelの計算結果だけでなく入力した数式を原寸画像から残すよう伝える'
    Add-Result ($prompt -match '全2枚中2枚目' -and $prompt -match 'F00015 まで' -and
        $prompt -match '同じChrome/Edge内' -and $prompt -match '曖昧な表現を避け') `
        '後続一覧でも同一ブラウザーの遷移と具体的な操作名を最後まで確認させる'
    Add-Result ($prompt.Length -lt 3000 -and $prompt -match '画面タイトルの遷移') `
        'Copilot入力上限へ達しない長さで画面タイトルの遷移を要約する'
    Add-Result ($prompt -match 'beforeFrame' -and $prompt -match 'afterFrame' -and $prompt.EndsWith('MB_TEST_END')) '操作前・操作後をフレームIDで返すJSON形式を指定する'

    $groupPrompt = New-MbRecorderCopilotPrompt -Frames $frames -Events $events -InteractionGroups @(
        [pscustomobject]@{ eventIds=@(1,2); actionKind='input'; targetName='B2'; beforeFrame='F00004'; afterFrame='F00007' }
    ) -Marker 'MB_GROUP_END'
    Add-Result ($groupPrompt -match '各Gをちょうど1手順' -and $groupPrompt -match 'G01 events=E001,E002 action=input target=B2' -and
        $groupPrompt -match 'targetは操作開始位置' -and $groupPrompt -match '初期描画や既存値を操作として追加しない') `
        'クリックと入力の操作境界をAIへ明示し一覧境界の重複と初期描画の誤認を防ぐ'

    $answer = [pscustomobject]@{ steps = @(
        [pscustomobject]@{ beforeFrame = 'F00004'; afterFrame = 'F00007'; eventIds = @(1); targetEventId = 1; title = '詳細を表示'; description = '［詳細を表示］をクリックします。'; confidence = 'high'; reason = '操作後の内容が必要' },
        [pscustomobject]@{ beforeFrame = 'F99999'; afterFrame = ''; eventIds = @(); targetEventId = $null; title = '不正'; description = '不正'; confidence = 'high'; reason = '' },
        [pscustomobject]@{ beforeFrame = 'F00012'; afterFrame = 'F00013'; eventIds = @(); targetEventId = $null; title = 'Excelを開く'; description = 'Excelを開きます。'; confidence = 'medium'; reason = '画面が切り替わった' }
    ) }
    $proposals = @(ConvertFrom-MbRecorderCopilotAnswer -Answer $answer -Frames $frames -Events $events)
    Add-Result ($proposals.Count -eq 1 -and [string]$proposals[0].beforeImage -eq 'frame-00004.jpg') '存在するフレームだけを元画像へ安全に戻す'
    Add-Result (@($proposals | Where-Object { $_.title -eq 'Excelを開く' }).Count -eq 0) 'クリック根拠のないアプリ切替を架空手順として採用しない'
    Add-Result ([string]$proposals[0].afterImage -eq 'frame-00007.jpg' -and [int]$proposals[0].targetEventId -eq 1) '操作後画像と赤枠アンカーを別々に保持する'

    $anchorAnswer = [pscustomobject]@{ steps = @(
        [pscustomobject]@{ beforeFrame = 'F00004'; afterFrame = ''; eventIds = @(1, 2); targetEventId = 2; title = '入力する'; description = '値を入力します。'; confidence = 'medium'; reason = '' }
    ) }
    $anchorProposals = @(ConvertFrom-MbRecorderCopilotAnswer -Answer $anchorAnswer -Frames $frames -Events $events)
    Add-Result ($anchorProposals.Count -eq 1 -and [int]$anchorProposals[0].targetEventId -eq 1) '矩形のない入力イベントを同じ手順のクリック赤枠へ補正する'

    $safetyFrames = @(
        [pscustomobject]@{ id='S00001'; timeMs=2000; image='frame-00004.jpg'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ id='S00002'; timeMs=4000; image='frame-00007.jpg'; windowTitle='申請画面 - Microsoft Edge' },
        [pscustomobject]@{ id='S00003'; timeMs=4000; image='frame-00013.jpg'; windowTitle='Book1 - Excel' }
    )
    $validRect = [pscustomobject]@{ x1=.1; y1=.2; x2=.3; y2=.4 }
    $safetyEvents = @(
        [pscustomobject]@{ index=1; timeMs=2100; kind='click'; windowTitle='申請画面 - Microsoft Edge'; rect=$validRect },
        [pscustomobject]@{ index=2; timeMs=2100; kind='click'; windowTitle='Book1 - Excel'; rect=$validRect },
        [pscustomobject]@{ index=3; timeMs=9000; kind='click'; windowTitle='申請画面 - Microsoft Edge'; rect=$validRect },
        [pscustomobject]@{ index=4; timeMs=2200; kind='click'; windowTitle='申請画面 - Microsoft Edge'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=1.2;y2=.4} },
        [pscustomobject]@{ index=5; timeMs=2200; kind='click'; windowTitle='申請画面 - Microsoft Edge'; rect=[pscustomobject]@{x1=.5;y1=.2;x2=.7;y2=.4} }
    )
    $newSafetyStep = { param($EventIds,$Target,$After='') [pscustomobject]@{
        beforeFrame='S00001'; afterFrame=$After; eventIds=$EventIds; targetEventId=$Target
        title='安全確認'; description='対象をクリックします。'; confidence='high'; reason=''
    } }
    $outsideAnchor = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(1) 2)}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($outsideAnchor.Count -eq 1 -and [int]$outsideAnchor[0].targetEventId -eq 0) 'eventIds外の別手順アンカーを採用しない'
    $foreignAnchor = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(2) 2)}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($foreignAnchor.Count -eq 1 -and [int]$foreignAnchor[0].targetEventId -eq 0) '操作前画像と別アプリのアンカーを採用しない'
    $distantAnchor = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(3) 3)}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($distantAnchor.Count -eq 1 -and [int]$distantAnchor[0].targetEventId -eq 0) '操作前画像から離れた後続イベントをアンカーにしない'
    $invalidRectAnchor = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(4) 4)}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($invalidRectAnchor.Count -eq 1 -and [int]$invalidRectAnchor[0].targetEventId -eq 0) '正規化範囲外の矩形をアンカーにしない'
    $ambiguousAnchor = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(1,5) 0)}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($ambiguousAnchor.Count -eq 1 -and [int]$ambiguousAnchor[0].targetEventId -eq 0) '妥当な赤枠候補が複数あるとき自動選択しない'
    $validAnchor = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(1) 1 'S00002')}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($validAnchor.Count -eq 1 -and [int]$validAnchor[0].targetEventId -eq 1) '同一アプリ・時刻近傍・eventIds内の矩形だけをアンカーにする'
    $crossApp = @(ConvertFrom-MbRecorderCopilotAnswer -Answer ([pscustomobject]@{steps=@(& $newSafetyStep @(1) 1 'S00003')}) -Frames $safetyFrames -Events $safetyEvents)
    Add-Result ($crossApp.Count -eq 0) '有効アンカーがあっても別アプリの操作前後画像を手順にしない'

    $browserFrames = @(
        [pscustomobject]@{ id='B00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='Calculator.net'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.0 },
        [pscustomobject]@{ id='B00002'; index=2; timeMs=2000; image='frame-00002.jpg'; windowTitle='Scientific Calculator'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.02 }
    )
    $browserAnswer = [pscustomobject]@{steps=@([pscustomobject]@{
        beforeFrame='B00001'; afterFrame='B00002'; eventIds=@(); targetEventId=$null
        title='科学電卓を開く'; description='［Scientific Calculator］を選択します。'; confidence='high'; reason='画面遷移'
    })}
    $browserProposal = @(ConvertFrom-MbRecorderCopilotAnswer -Answer $browserAnswer -Frames $browserFrames -Events @())
    Add-Result ($browserProposal.Count -eq 1 -and
        (Get-MbRecorderItemAppKey -Item $browserFrames[0]) -eq 'chrome' -and
        (Test-MbRecorderFrameSetHasMeaningfulChange -Frames $browserFrames)) `
        'ページタイトルが変わっても同じChromeプロセスの操作前後を保持する'

    $transitionFrames = @(
        [pscustomobject]@{ id='T00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='Calculator.net'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.0 },
        [pscustomobject]@{ id='T00002'; index=2; timeMs=10000; image='frame-00002.jpg'; windowTitle='Calculator.net'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.01 },
        [pscustomobject]@{ id='T00003'; index=3; timeMs=11000; image='frame-00003.jpg'; windowTitle='Scientific Calculator'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.35 },
        [pscustomobject]@{ id='T00004'; index=4; timeMs=13000; image='frame-00004.jpg'; windowTitle='Scientific Calculator'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.0 },
        [pscustomobject]@{ id='T00005'; index=5; timeMs=19000; image='frame-00005.jpg'; windowTitle='Scientific Calculator'; processName='chrome'; windowClass='Chrome_WidgetWin_1'; visualChange=0.08 }
    )
    $missingTransitionProposals = @(
        [pscustomobject]@{ id='P1'; beforeFrame='T00001'; afterFrame='T00002'; eventIds=@(); targetEventId=0; title='78 + 9を計算する'; description='87を表示します。'; confidence='high'; reason=''; timeMs=1000; beforeImage='frame-00001.jpg'; afterImage='frame-00002.jpg' },
        [pscustomobject]@{ id='P2'; beforeFrame='T00004'; afterFrame='T00005'; eventIds=@(); targetEventId=0; title='関数を入力する'; description='関数を入力します。'; confidence='medium'; reason=''; timeMs=13000; beforeImage='frame-00004.jpg'; afterImage='frame-00005.jpg' }
    )
    $completedTransitionProposals = @(Add-MbRecorderTitleTransitionProposals -Frames $transitionFrames -Proposals $missingTransitionProposals)
    Add-Result ($completedTransitionProposals.Count -eq 3 -and
        [string]$completedTransitionProposals[1].beforeFrame -eq 'T00002' -and
        [string]$completedTransitionProposals[1].afterFrame -eq 'T00003' -and
        [string]$completedTransitionProposals[1].title -eq 'Scientific Calculatorを開く') `
        'Copilot実回答から欠けたページ遷移を独立した手順として補完する'

    $mergedTransitionProposal = @([pscustomobject]@{
        id='P3'; beforeFrame='T00002'; afterFrame='T00005'; eventIds=@(); targetEventId=0
        title='関数電卓で計算する'; description='関数電卓を開いて計算します。'; confidence='medium'; reason=''
        timeMs=10000; beforeImage='frame-00002.jpg'; afterImage='frame-00005.jpg'
    })
    $splitTransitionProposals = @(Add-MbRecorderTitleTransitionProposals -Frames $transitionFrames -Proposals $mergedTransitionProposal)
    Add-Result ($splitTransitionProposals.Count -eq 2 -and
        @($splitTransitionProposals | Where-Object { $_.title -eq 'Scientific Calculatorを開く' }).Count -eq 1 -and
        @($splitTransitionProposals | Where-Object { $_.id -eq 'P3' -and $_.beforeFrame -eq 'T00004' }).Count -eq 1) `
        'Copilotが結合したページ遷移と後続操作を別々の手順へ分ける'

    $crossPacketDuplicates = @(
        [pscustomobject]@{ id='P4'; beforeFrame='T00002'; afterFrame=''; eventIds=@(); targetEventId=0; title='Scientific Calculator リンクを選択する'; description='「Scientific Calculator」をクリックします。'; confidence='medium'; reason=''; timeMs=10000; beforeImage='frame-00002.jpg'; afterImage='' },
        [pscustomobject]@{ id='P5'; beforeFrame='T00003'; afterFrame='T00004'; eventIds=@(); targetEventId=0; title='Scientific Calculatorを開く'; description='Scientific Calculatorページを表示します。'; confidence='high'; reason=''; timeMs=11000; beforeImage='frame-00003.jpg'; afterImage='frame-00004.jpg' }
    )
    # 実回答と同じ構造にするため、後続候補のbefore/afterが異なるタイトルを持つ
    # フレームへ差し替える。
    $crossPacketDuplicates[1].beforeFrame = 'T00002'
    $crossPacketDuplicates[1].beforeImage = 'frame-00002.jpg'
    $crossPacketDuplicates[1].afterFrame = 'T00003'
    $crossPacketDuplicates[1].afterImage = 'frame-00003.jpg'
    $deduplicatedTransitions = @(Merge-MbRecorderDuplicateTransitionProposals -Frames $transitionFrames -Proposals $crossPacketDuplicates)
    Add-Result ($deduplicatedTransitions.Count -eq 1 -and [string]$deduplicatedTransitions[0].id -eq 'P5' -and
        [string]$deduplicatedTransitions[0].afterFrame -eq 'T00003') `
        '一覧画像の境界で重なったリンク選択とページ遷移を1手順へまとめる'

    $rangeEvents = @([pscustomobject]@{ index=7; timeMs=12000; kind='click'; targetName='B2'; processName='EXCEL'; windowTitle='Book1 - Excel'; rect=$validRect })
    $rangeProposals = @(
        [pscustomobject]@{ id='R1'; beforeFrame='S00001'; afterFrame=''; eventIds=@(7); targetEventId=7; title='セルB2を選択する'; description='セルB2を選択する。'; confidence='high'; reason=''; timeMs=12000 },
        [pscustomobject]@{ id='R2'; beforeFrame='S00002'; afterFrame=''; eventIds=@(); targetEventId=0; title='通貨表示形式を適用する'; description='B2:B4を選択した状態で通貨表示形式を適用する。'; confidence='high'; reason=''; timeMs=14000 }
    )
    $expandedRange = @(Expand-MbRecorderExcelRangeSelectionProposals -Events $rangeEvents -Proposals $rangeProposals)
    Add-Result ($expandedRange.Count -eq 2 -and [string]$expandedRange[0].title -eq 'セル範囲B2:B4を選択する' -and
        [string]$expandedRange[0].description -match 'ドラッグ') `
        'Excelのドラッグ開始セルだけを返した場合も後続画像で確認した選択範囲へ補正する'

    $loadingFrames = @(
        [pscustomobject]@{ id='L00001'; index=1; timeMs=1000; image='frame-00001.jpg'; windowTitle='受注検索 - Microsoft Edge'; processName='msedge' },
        [pscustomobject]@{ id='L00002'; index=2; timeMs=2000; image='frame-00002.jpg'; windowTitle='読み込み中 - Microsoft Edge'; processName='msedge' },
        [pscustomobject]@{ id='L00003'; index=3; timeMs=3000; image='frame-00003.jpg'; windowTitle='受注検索 - Microsoft Edge'; processName='msedge' }
    )
    $loadingEvents = @(
        [pscustomobject]@{ index=1; timeMs=1500; kind='click'; targetName='検索'; processName='msedge' },
        [pscustomobject]@{ index=2; timeMs=4000; kind='click'; targetName='詳細'; processName='msedge' }
    )
    $loadingProposal = @([pscustomobject]@{ id='L1'; beforeFrame='L00001'; afterFrame='L00002'; eventIds=@(1); targetEventId=1; title='検索する'; description='検索する。'; confidence='high'; reason=''; timeMs=1000; beforeImage='frame-00001.jpg'; afterImage='frame-00002.jpg' })
    $repairedLoading = @(Repair-MbRecorderTransientAfterFrames -Frames $loadingFrames -Events $loadingEvents -Proposals $loadingProposal)
    Add-Result ($repairedLoading.Count -eq 1 -and [string]$repairedLoading[0].afterFrame -eq 'L00003') `
        'Copilotが読込中を操作後画像に選んでも次操作前の最初の安定画面へ置き換える'

    $testJobId = 'record-' + [guid]::NewGuid().ToString('N')
    $evidenceDirectory = Join-Path $testRoot 'evidence'
    [void](New-Item -ItemType Directory -Path $evidenceDirectory -Force)
    Copy-Item -LiteralPath (Join-Path $framesDirectory 'frame-00004.jpg') -Destination (Join-Path $evidenceDirectory ($testEvidenceIds[0] + '.jpg'))
    Copy-Item -LiteralPath (Join-Path $framesDirectory 'frame-00013.jpg') -Destination (Join-Path $evidenceDirectory ($testEvidenceIds[1] + '.jpg'))
    $ledgerPath = Join-Path $testRoot 'evidence-ledger.jsonl'
    [IO.File]::WriteAllLines($ledgerPath, @(
        ([ordered]@{ recordType='capture-start'; formatVersion=2; sessionId=$testJobId; mouseHook=$true; keyboardHook=$true; completeness='no-known-gaps' } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$testEvidenceIds[0]; sessionId=$testJobId; kind='click'; timeMs=2100; image=($testEvidenceIds[0] + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$testEvidenceIds[1]; sessionId=$testJobId; kind='input'; timeMs=7100; image=($testEvidenceIds[1] + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='capture-end'; formatVersion=2; sessionId=$testJobId; operationCount=2; reason='stopped'; completeness='no-known-gaps'; warning='' } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    $job = [pscustomobject]@{
        JobId = $testJobId; JobDirectory = $testRoot; ProcessId = 0
        FramesDirectory = $framesDirectory; FramesPath = $framesPath
        EventsDirectory = $framesDirectory; EventsPath = $eventsPath
        EvidenceDirectory = $evidenceDirectory; LedgerPath = $ledgerPath
    }
    & (Get-Module ManualBuilder.RecorderServer) { param($Value) $script:MbRecordingJob = $Value } $job
    $directSafetyProject = New-MbProject
    $directCrossJson = [pscustomobject]@{accept=@([pscustomobject]@{
        beforeFrame='F00004'; afterFrame='F00013'; eventIds=@(1); targetEventId=1
        title='不正な切替'; description='切り替えます。'; confidence='high'; reason=''
    })} | ConvertTo-Json -Depth 8 -Compress
    $directCross = Import-MbRecordedLocalSelections -Project $directSafetyProject -ProjectPath $projectPath `
        -SheetId $directSafetyProject.sheets[0].id -SelectionJson $directCrossJson
    Add-Result ([int]$directCross.added -eq 0 -and [int]$directCross.skipped -eq 1) '取り込み側でも別アプリの操作前後画像を拒否する'
    $directForeignJson = [pscustomobject]@{accept=@([pscustomobject]@{
        beforeFrame='F00004'; afterFrame=''; eventIds=@(2); targetEventId=2
        title='要確認'; description='対象を操作します。'; confidence='high'; reason=''
    })} | ConvertTo-Json -Depth 8 -Compress
    $directForeign = Import-MbRecordedLocalSelections -Project $directSafetyProject -ProjectPath $projectPath `
        -SheetId $directSafetyProject.sheets[0].id -SelectionJson $directForeignJson
    $directStep = @($directSafetyProject.sheets[0].steps)[0]
    Add-Result ([int]$directForeign.added -eq 1 -and [int]$directForeign.needsReview -eq 1 -and
        @($directStep.annotations).Count -eq 0 -and [string]$directStep.review.action -eq 'review') `
        '取り込み側でも不整合アンカーを外し高confidenceを要確認にする'
    $localServerProposals = @(Get-MbRecordedLocalProposals)
    Add-Result ($localServerProposals.Count -ge 1 -and [string]$localServerProposals[0].source -eq 'local') `
        'Copilotを使わず安定フレームから編集可能な候補を返す'
    $sameLocalProposals = @(Get-MbRecordedLocalProposals)
    Add-Result ((@($localServerProposals | ForEach-Object { [string]$_.id }) -join ',') -eq
        (@($sameLocalProposals | ForEach-Object { [string]$_.id }) -join ',')) `
        '確認画面と取り込み時で同じ候補IDを使う'
    $reviewedItem = $localServerProposals[0].PSObject.Copy()
    $reviewedItem.title = '利用者が確認した手順'
    $reviewedItem.description = '候補画面で文章を直して確定します。'
    $reviewedItem | Add-Member -NotePropertyName reviewed -NotePropertyValue $true -Force
    $reviewedProject = New-MbProject
    $reviewDecisionItem = [pscustomobject]@{
        id = [string]$reviewedItem.id; accepted = $true; reviewed = $true
        title = [string]$reviewedItem.title; description = [string]$reviewedItem.description
    }
    $reviewedJson = [pscustomobject]@{ accept = @($reviewedItem); decisions = @($reviewDecisionItem) } | ConvertTo-Json -Depth 10 -Compress
    $reviewedImport = Import-MbRecordedLocalSelections -Project $reviewedProject -ProjectPath $projectPath `
        -SheetId $reviewedProject.sheets[0].id -SelectionJson $reviewedJson
    $reviewedStep = @($reviewedProject.sheets[0].steps)[0]
    Add-Result ([int]$reviewedImport.added -eq 1 -and [int]$reviewedImport.needsReview -eq 0 -and
        [string]$reviewedStep.title -eq '利用者が確認した手順' -and [string]$reviewedStep.review.action -eq '') `
        '候補画面で直して確定した手順を取り込み後の再確認対象にしない'
    $decisionPath = Join-Path (Join-Path (Join-Path $testRoot 'evidence') $testJobId) 'transformations.jsonl'
    $reviewDecision = @([IO.File]::ReadAllLines($decisionPath, [Text.Encoding]::UTF8) | ForEach-Object { $_ | ConvertFrom-Json } |
        Where-Object { [string]$_.proposalId -eq [string]$reviewedItem.id -and [bool]$_.accepted -and [bool]$_.reviewed } | Select-Object -Last 1)
    Add-Result ($reviewDecision.Count -eq 1 -and [string]$reviewDecision[0].decisionSource -eq 'user-review' -and
        [string]$reviewDecision[0].finalTitle -eq '利用者が確認した手順') `
        '候補画面での文章修正と利用者確認を変換履歴へ残す'
    $excludedDecisionItem = [pscustomobject]@{
        id = [string]$reviewedItem.id; accepted = $false; reviewed = $true
        title = [string]$reviewedItem.title; description = [string]$reviewedItem.description
    }
    $excludedProject = New-MbProject
    $excludedJson = [pscustomobject]@{ accept = @(); decisions = @($excludedDecisionItem) } | ConvertTo-Json -Depth 10 -Compress
    [void](Import-MbRecordedLocalSelections -Project $excludedProject -ProjectPath $projectPath `
        -SheetId $excludedProject.sheets[0].id -SelectionJson $excludedJson)
    $excludedDecision = @([IO.File]::ReadAllLines($decisionPath, [Text.Encoding]::UTF8) | ForEach-Object { $_ | ConvertFrom-Json } |
        Where-Object { [string]$_.proposalId -eq [string]$reviewedItem.id -and -not [bool]$_.accepted -and [bool]$_.reviewed } | Select-Object -Last 1)
    Add-Result ($excludedDecision.Count -eq 1 -and [string]$excludedDecision[0].decisionSource -eq 'user-review') `
        '候補画面で利用者が除外した判断も変換履歴へ残す'
    $project = New-MbProject
    $project = Save-MbProject -Project $project -Path $projectPath
    $selectionJson = [pscustomobject]@{ accept = @($proposals) } | ConvertTo-Json -Depth 10 -Compress
    $imported = Import-MbRecordedLocalSelections -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id -SelectionJson $selectionJson
    $step = @($project.sheets[0].steps)[0]
    Add-Result ([int]$imported.added -eq 1 -and @($project.sheets[0].steps).Count -eq 1) 'Copilotが選んだ単位で手順を作る'
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$step.resultImageId) -and [string]$step.imageLayout -eq 'before') '結果画像を保持しつつ初稿は案内画像1枚で取り込む'
    Add-Result (@($step.annotations).Count -eq 1 -and [string]$step.capture.kind -eq 'recorded-local') 'クリックイベントは赤枠候補のアンカーとしてだけ使う'
    Add-Result ([string]$step.title -eq '詳細を表示' -and [string]$step.description -match 'クリック') 'Copilotの手順文を編集可能な初稿へ反映する'
} finally {
    & (Get-Module ManualBuilder.RecorderServer) { $script:MbRecordingJob = $null }
    if ($null -ne $graphics) { $graphics.Dispose() }
    if ($null -ne $bitmap) { $bitmap.Dispose() }
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if ($errors.Count -gt 0) {
    Write-Host "Recorder Copilot tests failed: $($errors.Count)" -ForegroundColor Red
    exit 1
}
Write-Host 'Recorder Copilot tests passed.' -ForegroundColor Green
