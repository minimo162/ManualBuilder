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

$workerSource = Get-Content -LiteralPath (Join-Path $srcRoot 'Invoke-ManualBuilderRecorderCopilot.ps1') -Raw -Encoding UTF8
Add-Result ($workerSource -match "copilot_model\s*=\s*'自動,Automatic,Auto'") '場面選定は高速で安定した自動モデルを使う'
Add-Result ($workerSource -match '\$maximumAttempts\s*=\s*2' -and $workerSource -match '回答を読み取れませんでした（\{1\}/\{2\}）') `
    'Copilotの一時的な通信エラーだけ1回再試行できる'
Add-Result ($workerSource -match "completedBy\s*-eq\s*'service-error'" -and
    $workerSource -match "ErrorCode 'COPILOT_SERVICE_UNAVAILABLE'" -and
    $workerSource -match "Properties.Name -contains 'errorCode'") `
    'M365の確定的なサービスエラーでは同じ画像を再送せずローカル候補へ移る'
Add-Result ($workerSource -match 'New-MbRecorderLocalFrameCandidates\s+-Frames \$eventWindowFrames\s+-Events \$events\s+-MaximumFrames 30' -and
    $workerSource -match 'Select-MbRecorderCandidateFrames\s+-Frames \$eventWindowFrames\s+-Candidates \$localCandidates' -and
    $workerSource -match '\$perPacket\s*=\s*1' -and $workerSource -match 'request_timeout\s*=\s*\[Math\]::Min\(90') `
    'ローカル候補の操作前後だけを残し、一覧画像は1枚ずつ90秒以内で処理する'

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
    $eventLines = @(
        ([pscustomobject]@{ index = 1; timeMs = 2100; kind = 'click'; targetName = '詳細を表示'; targetType = 'ControlType.Button'; windowTitle = '申請画面 - Microsoft Edge'; rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.2; x2 = 0.3; y2 = 0.3 } } | ConvertTo-Json -Compress -Depth 5),
        ([pscustomobject]@{ index = 2; timeMs = 7100; kind = 'input'; targetName = 'F7'; targetType = 'ControlType.DataItem'; windowTitle = 'Book1 - Excel'; rect = $null } | ConvertTo-Json -Compress -Depth 5)
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

    $longFrames = @($frames) + @($frames | ForEach-Object {
        $copy = $_ | Select-Object *
        $copy.index = [int]$copy.index + 25
        $copy.timeMs = [int]$copy.timeMs + 12500
        $copy.id = 'F{0:d5}' -f [int]$copy.index
        $copy
    })
    $fastSelection = @(Select-MbRecorderTimelineFrames -Frames $longFrames -Events $events -Maximum 30)
    Add-Result ($fastSelection.Count -gt 1 -and $fastSelection.Count -le 30) '長い記録も最大30コマへ絞る'

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
    $candidateFrames = @(Select-MbRecorderCandidateFrames -Frames $meaningFrames -Candidates $localCandidates)
    Add-Result ($candidateFrames.Count -eq 4 -and @($candidateFrames | Where-Object { $_.id -eq 'F00002' }).Count -eq 1 -and
        @($candidateFrames | Where-Object { $_.id -eq 'F00003' }).Count -eq 1) `
        'Copilot一覧を候補の操作前後と貼り付け証拠へ絞り、数式を確定後の値だけにしない'

    $missingEventFrames = @(
        [pscustomobject]@{ id='F00001'; index=1; timeMs=800; image='frame-00001.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0 },
        [pscustomobject]@{ id='F00002'; index=2; timeMs=1400; image='frame-00002.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0006 },
        [pscustomobject]@{ id='F00003'; index=3; timeMs=2600; image='frame-00003.jpg'; windowTitle='Book1 - Excel'; visualChange=0.0009 }
    )
    $missingEventEvents = @(
        [pscustomobject]@{ index=1; timeMs=1000; kind='click'; targetName='B2'; targetType='ControlType.DataItem'; windowTitle='Book1 - Excel'; rect=[pscustomobject]@{x1=.1;y1=.2;x2=.2;y2=.3} }
    )
    $missingCandidates = @(New-MbRecorderLocalFrameCandidates -Frames $missingEventFrames -Events $missingEventEvents -MaximumFrames 8)
    Add-Result (@($missingCandidates | Where-Object { $_.actionKind -eq 'visual-change' -and $_.afterFrame -eq 'F00003' }).Count -eq 1) `
        'クリックイベントが欠けても画面差分から要確認候補を補う'

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
        [string]$clickEvidenceCandidates[0].afterFrame -eq 'F00003' -and
        [string]$clickEvidenceCandidates[1].beforeFrame -eq 'F00003') `
        '遅延遷移は読み込み中の周期コマでなく次操作直前の安定画面を結果にする'

    $sheets = @(New-MbRecorderContactSheets -Frames $frames -FramesDirectory $framesDirectory -OutputDirectory $contactDirectory)
    Add-Result ($sheets.Count -eq 2) '25フレームを複数の番号付き一覧画像へ分ける'
    $sheetImage = [Drawing.Image]::FromFile([string]$sheets[0].path)
    try { Add-Result ($sheetImage.Width -eq 1920 -and $sheetImage.Height -eq 1970) '一覧画像を文字も読める3列の固定寸法で作る' }
    finally { $sheetImage.Dispose() }
    $lastSheetImage = [Drawing.Image]::FromFile([string]$sheets[1].path)
    try { Add-Result ($lastSheetImage.Width -eq 1920 -and $lastSheetImage.Height -eq 1576) '最終ページを空の15コマ分まで水増ししない' }
    finally { $lastSheetImage.Dispose() }

    $prompt = New-MbRecorderCopilotPrompt -Frames $frames -Events $events -Marker 'MB_TEST_END'
    Add-Result ($prompt -match '中間的なアニメーション' -and $prompt -match 'ManualBuilderへ戻る操作') '遷移中画像と記録終了操作を除く判断基準を伝える'
    Add-Result ($prompt -match '対象が異なる入力イベントは省略せず' -and $prompt -match '前面アプリが切り替わっただけ') '別セル入力を残しアプリ切替を手順にしない基準を伝える'
    Add-Result ($prompt -match '大文字小文字は変えず') '画面に見える値の表記を変えない基準を伝える'
    Add-Result ($prompt -match 'beforeFrame' -and $prompt -match 'afterFrame' -and $prompt.EndsWith('MB_TEST_END')) '操作前・操作後をフレームIDで返すJSON形式を指定する'

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

    $job = [pscustomobject]@{
        JobId = 'test'; JobDirectory = $testRoot; ProcessId = 0
        FramesDirectory = $framesDirectory; FramesPath = $framesPath
        EventsDirectory = $framesDirectory; EventsPath = $eventsPath
        NarrationPath = (Join-Path $testRoot 'narration.jsonl')
    }
    & (Get-Module ManualBuilder.RecorderServer) { param($Value) $script:MbRecordingJob = $Value } $job
    $directSafetyProject = New-MbProject
    $directCrossJson = [pscustomobject]@{accept=@([pscustomobject]@{
        beforeFrame='F00004'; afterFrame='F00013'; eventIds=@(1); targetEventId=1
        title='不正な切替'; description='切り替えます。'; confidence='high'; reason=''
    })} | ConvertTo-Json -Depth 8 -Compress
    $directCross = Import-MbRecordedCopilotSelections -Project $directSafetyProject -ProjectPath $projectPath `
        -SheetId $directSafetyProject.sheets[0].id -SelectionJson $directCrossJson
    Add-Result ([int]$directCross.added -eq 0 -and [int]$directCross.skipped -eq 1) '取り込み側でも別アプリの操作前後画像を拒否する'
    $directForeignJson = [pscustomobject]@{accept=@([pscustomobject]@{
        beforeFrame='F00004'; afterFrame=''; eventIds=@(2); targetEventId=2
        title='要確認'; description='対象を操作します。'; confidence='high'; reason=''
    })} | ConvertTo-Json -Depth 8 -Compress
    $directForeign = Import-MbRecordedCopilotSelections -Project $directSafetyProject -ProjectPath $projectPath `
        -SheetId $directSafetyProject.sheets[0].id -SelectionJson $directForeignJson
    $directStep = @($directSafetyProject.sheets[0].steps)[0]
    Add-Result ([int]$directForeign.added -eq 1 -and [int]$directForeign.needsReview -eq 1 -and
        @($directStep.annotations).Count -eq 0 -and [string]$directStep.review.action -eq 'review') `
        '取り込み側でも不整合アンカーを外し高confidenceを要確認にする'
    $localServerProposals = @(Get-MbRecordedLocalProposals)
    Add-Result ($localServerProposals.Count -ge 1 -and [string]$localServerProposals[0].source -eq 'local') `
        'Copilotを使わず安定フレームから編集可能な候補を返す'
    $project = New-MbProject
    $project = Save-MbProject -Project $project -Path $projectPath
    $selectionJson = [pscustomobject]@{ accept = @($proposals) } | ConvertTo-Json -Depth 10 -Compress
    $imported = Import-MbRecordedCopilotSelections -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id -SelectionJson $selectionJson
    $step = @($project.sheets[0].steps)[0]
    Add-Result ([int]$imported.added -eq 1 -and @($project.sheets[0].steps).Count -eq 1) 'Copilotが選んだ単位で手順を作る'
    Add-Result (-not [string]::IsNullOrWhiteSpace([string]$step.resultImageId) -and [string]$step.imageLayout -eq 'side-by-side') '選ばれた操作前・操作後を左右比較として取り込む'
    Add-Result (@($step.annotations).Count -eq 1 -and [string]$step.capture.kind -eq 'recorded-ai') 'クリックイベントは赤枠候補のアンカーとしてだけ使う'
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
