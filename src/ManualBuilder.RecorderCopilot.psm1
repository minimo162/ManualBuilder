# ManualBuilder recorder timeline / Copilot selection helpers.

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing -ErrorAction Stop

function Format-MbRecorderTimeCode {
    param([int]$Milliseconds)
    $safe = [Math]::Max(0, $Milliseconds)
    $minutes = [int][Math]::Floor($safe / 60000)
    $seconds = [int][Math]::Floor(($safe % 60000) / 1000)
    $tenths = [int][Math]::Floor(($safe % 1000) / 100)
    return ('{0:d2}:{1:d2}.{2}' -f $minutes, $seconds, $tenths)
}

function Get-MbRecorderWindowAppKey {
    param(
        [AllowEmptyString()][string]$WindowTitle,
        [AllowEmptyString()][string]$ProcessName = '',
        [AllowEmptyString()][string]$WindowClass = ''
    )
    $process = $ProcessName.Trim().ToLowerInvariant()
    if ($process) {
        if ($process -in @('msedge', 'microsoftedge', 'microsoftedgecp')) { return 'edge' }
        if ($process -eq 'chrome') { return 'chrome' }
        if ($process -in @('excel', 'winword', 'powerpnt')) {
            return $(switch ($process) { 'winword' { 'word' } 'powerpnt' { 'powerpoint' } default { $process } })
        }
        return $process
    }
    $title = $WindowTitle.Trim()
    if ($title -match '(?i)Microsoft.?.?Edge$') { return 'edge' }
    if ($title -match '(?i)\s-\sGoogle Chrome$') { return 'chrome' }
    if ($title -match '(?i)\s-\sExcel$') { return 'excel' }
    if ($title -match '(?i)\s-\sWord$') { return 'word' }
    # ChromiumのアプリモードやPWAはブラウザー名をタイトル末尾へ付けない。
    # ページ遷移でタイトルが変わっても、同じブラウザー内の操作として扱う。
    if ($WindowClass -like 'Chrome_WidgetWin*') { return 'chromium' }
    if ([string]::IsNullOrWhiteSpace($title)) { return '' }
    return ($title -replace '\s+', ' ').ToLowerInvariant()
}

function Get-MbRecorderItemAppKey {
    param([AllowNull()]$Item)
    if ($null -eq $Item) { return '' }
    $processName = if ($Item.PSObject.Properties.Name -contains 'processName') { [string]$Item.processName } else { '' }
    $windowClass = if ($Item.PSObject.Properties.Name -contains 'windowClass') { [string]$Item.windowClass } else { '' }
    $windowTitle = if ($Item.PSObject.Properties.Name -contains 'windowTitle') { [string]$Item.windowTitle } else { '' }
    return Get-MbRecorderWindowAppKey -WindowTitle $windowTitle -ProcessName $processName -WindowClass $windowClass
}

function Test-MbRecorderEventHasVisualAnchor {
    param([AllowNull()]$Event)
    if ($null -eq $Event) { return $false }
    if ($Event.PSObject.Properties.Name -notcontains 'rect' -or $null -eq $Event.rect) { return $false }
    try {
        $x1 = [double]$Event.rect.x1; $y1 = [double]$Event.rect.y1
        $x2 = [double]$Event.rect.x2; $y2 = [double]$Event.rect.y2
        return ($x1 -ge 0 -and $y1 -ge 0 -and $x2 -le 1 -and $y2 -le 1 -and
            $x2 -gt $x1 -and $y2 -gt $y1)
    } catch { return $false }
}

function Test-MbRecorderAnchorContext {
    param(
        [AllowNull()]$Event,
        [Parameter(Mandatory = $true)]$BeforeFrame,
        [AllowNull()]$AfterFrame
    )
    if (-not (Test-MbRecorderEventHasVisualAnchor -Event $Event)) { return $false }
    $beforeApp = Get-MbRecorderItemAppKey -Item $BeforeFrame
    $eventApp = Get-MbRecorderItemAppKey -Item $Event
    if ([string]::IsNullOrWhiteSpace($beforeApp) -or $beforeApp -ne $eventApp) { return $false }
    $beforeTime = [int]$BeforeFrame.timeMs; $eventTime = [int]$Event.timeMs
    if ([Math]::Abs($eventTime - $beforeTime) -gt 2500) { return $false }
    if ($null -ne $AfterFrame) {
        $afterApp = Get-MbRecorderItemAppKey -Item $AfterFrame
        $afterTime = [int]$AfterFrame.timeMs
        if ([string]::IsNullOrWhiteSpace($afterApp) -or $afterApp -ne $beforeApp -or
            $afterTime -le $beforeTime -or $eventTime -gt ($afterTime + 1000)) { return $false }
    }
    return $true
}

function Read-MbRecorderJsonLines {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    $items = New-Object System.Collections.ArrayList
    foreach ($line in [IO.File]::ReadAllLines($Path, [Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $item = $line | ConvertFrom-Json
            if ($null -ne $item) { [void]$items.Add($item) }
        } catch { }
    }
    return @($items)
}

function Test-MbRecorderTransientFrameTitle {
    param([AllowEmptyString()][string]$WindowTitle)
    return ([string]$WindowTitle -match '(?i)読み込み中|読込中|loading|please wait|処理中|準備中')
}

# 小さなセル値や選択位置の変化を、クリックイベントが欠けた場合の補助証拠にする。
# OCRやアプリ固有APIは使わず、縮小画像の平均画素差だけをメタデータへ付ける。
function Add-MbRecorderFrameVisualMetrics {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory = $true)][string]$FramesDirectory
    )
    $ordered = @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    $previous = $null
    $previousApp = ''
    foreach ($frame in $ordered) {
        $score = 0.0
        $current = $null
        $app = Get-MbRecorderItemAppKey -Item $frame
        $fileName = [string]$frame.image
        $path = if ($fileName -match '^frame-\d{5}\.jpg$' -and [IO.Path]::GetFileName($fileName) -eq $fileName) {
            Join-Path $FramesDirectory $fileName
        } else { '' }
        if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
            try {
                $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash
                $frame | Add-Member -NotePropertyName imageSha256 -NotePropertyValue ([string]$hash) -Force
            } catch { }
            $source = $null; $thumb = $null; $graphics = $null
            try {
                $source = [Drawing.Image]::FromFile($path)
                $thumb = New-Object Drawing.Bitmap -ArgumentList @(80, 45, [Drawing.Imaging.PixelFormat]::Format24bppRgb)
                $graphics = [Drawing.Graphics]::FromImage($thumb)
                $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear
                $graphics.DrawImage($source, 0, 0, 80, 45)
                $current = New-Object byte[] (80 * 45 * 3)
                $offset = 0
                for ($y = 0; $y -lt 45; $y++) {
                    for ($x = 0; $x -lt 80; $x++) {
                        $pixel = $thumb.GetPixel($x, $y)
                        $current[$offset] = $pixel.R; $current[$offset + 1] = $pixel.G; $current[$offset + 2] = $pixel.B
                        $offset += 3
                    }
                }
            } catch { $current = $null }
            finally {
                foreach ($item in @($graphics, $thumb, $source)) { if ($null -ne $item) { try { $item.Dispose() } catch { } } }
            }
        }
        if ($null -ne $current -and $null -ne $previous -and $app -eq $previousApp -and $current.Length -eq $previous.Length) {
            [long]$difference = 0
            for ($i = 0; $i -lt $current.Length; $i++) { $difference += [Math]::Abs([int]$current[$i] - [int]$previous[$i]) }
            $score = $difference / [double]($current.Length * 255)
        }
        $frame | Add-Member -NotePropertyName visualChange -NotePropertyValue ([Math]::Round($score, 8)) -Force
        if ($null -ne $current) { $previous = $current; $previousApp = $app }
    }
    return @($ordered)
}

function Select-MbRecorderTimelineFrames {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [AllowEmptyCollection()][object[]]$Events = @(),
        [ValidateRange(8, 300)][int]$Maximum = 120
    )
    $ordered = @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    $orderedEvents = @($Events | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    if ($orderedEvents.Count -lt 1) {
        if ($ordered.Count -le $Maximum) { return @($ordered) }

        # イベントを全件取り逃した記録でも、画面差分とその直前画像は手順を
        # 復元できる唯一の根拠になる。等間隔だけで上限を埋める前に保護する。
        $important = New-Object 'System.Collections.Generic.HashSet[int]'
        for ($frameIndex = 0; $frameIndex -lt $ordered.Count; $frameIndex++) {
            $frame = $ordered[$frameIndex]
            $visualChange = if ($frame.PSObject.Properties.Name -contains 'visualChange') { [double]$frame.visualChange } else { 0.0 }
            if ($visualChange -lt 0.00035 -or
                (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$frame.windowTitle))) { continue }
            [void]$important.Add($frameIndex)
            if ($frameIndex -gt 0 -and
                (Get-MbRecorderItemAppKey -Item $ordered[$frameIndex - 1]) -eq
                    (Get-MbRecorderItemAppKey -Item $frame)) {
                [void]$important.Add($frameIndex - 1)
            }
        }
        $importantIndexes = @($important | Sort-Object)
        $chosen = New-Object 'System.Collections.Generic.HashSet[int]'
        if ($importantIndexes.Count -gt $Maximum) {
            for ($slot = 0; $slot -lt $Maximum; $slot++) {
                $position = [int][Math]::Round(($slot / [double]([Math]::Max(1, $Maximum - 1))) * ($importantIndexes.Count - 1))
                [void]$chosen.Add([int]$importantIndexes[$position])
            }
        } else {
            foreach ($importantIndex in $importantIndexes) { [void]$chosen.Add([int]$importantIndex) }
        }
        for ($slot = 0; $slot -lt $Maximum -and $chosen.Count -lt $Maximum; $slot++) {
            $position = [int][Math]::Round(($slot / [double]([Math]::Max(1, $Maximum - 1))) * ($ordered.Count - 1))
            [void]$chosen.Add($position)
        }
        return @($chosen | Sort-Object | ForEach-Object { $ordered[[int]$_] })
    }

    # 等間隔の間引きは、実機で「入力完了」を落として入力途中を残した。
    # 各イベントの直前、次の操作直前に確定した結果、アプリ切替後の安定画面を
    # 保護する。Maximum は目標件数ではなく上限であり、ノイズで埋めない。
    $priorities = @{}
    $addCandidate = {
        param([int]$Index, [int]$Priority)
        if ($Index -lt 0 -or $Index -ge $ordered.Count) { return }
        $key = [string]$Index
        if (-not $priorities.ContainsKey($key) -or [int]$priorities[$key] -lt $Priority) {
            $priorities[$key] = $Priority
        }
    }

    $seenAppSegments = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    # 数式や長い識別子は、確定後の安定画面だけでは元の入力内容を復元できない。
    # 記録ワーカーが明示した「入力中の最終表示」は、通常の安定画面より先に保護する。
    for ($frameIndex = 0; $frameIndex -lt $ordered.Count; $frameIndex++) {
        $frame = $ordered[$frameIndex]
        $role = if ($frame.PSObject.Properties.Name -contains 'role') { [string]$frame.role } else { '' }
        if ($role -eq 'input-evidence') { & $addCandidate $frameIndex 120 }
        $visualChange = if ($frame.PSObject.Properties.Name -contains 'visualChange') { [double]$frame.visualChange } else { 0.0 }
        if ($visualChange -ge 0.00035 -and -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$frame.windowTitle))) {
            # UIAイベントを取り逃したセル移動・値確定も、前後を比較できるよう残す。
            & $addCandidate $frameIndex 92
            if ($frameIndex -gt 0 -and
                (Get-MbRecorderItemAppKey -Item $ordered[$frameIndex - 1]) -eq
                    (Get-MbRecorderItemAppKey -Item $frame)) {
                & $addCandidate ($frameIndex - 1) 84
            }
        }
    }

    for ($eventIndex = 0; $eventIndex -lt $orderedEvents.Count; $eventIndex++) {
        $event = $orderedEvents[$eventIndex]
        $eventTime = [int]$event.timeMs
        $eventApp = Get-MbRecorderItemAppKey -Item $event
        $beforeIndex = -1
        for ($frameIndex = 0; $frameIndex -lt $ordered.Count; $frameIndex++) {
            if ([int]$ordered[$frameIndex].timeMs -gt $eventTime) { break }
            $frameApp = Get-MbRecorderItemAppKey -Item $ordered[$frameIndex]
            if (-not $eventApp -or $frameApp -eq $eventApp) { $beforeIndex = $frameIndex }
        }
        if ($beforeIndex -ge 0) { & $addCandidate $beforeIndex 100 }

        # 新しいアプリへ切り替わった直後の先頭2コマは、描画途中やフォーカス移動に
        # なりやすい。3コマ目を操作前の文脈として残す（存在しなければ直前コマ）。
        if ($eventApp -and $seenAppSegments.Add($eventApp)) {
            $firstAppIndex = -1
            for ($frameIndex = 0; $frameIndex -lt $ordered.Count; $frameIndex++) {
                if ((Get-MbRecorderItemAppKey -Item $ordered[$frameIndex]) -eq $eventApp) {
                    $firstAppIndex = $frameIndex
                    break
                }
            }
            if ($firstAppIndex -ge 0) {
                $contextIndex = [Math]::Min($beforeIndex, $firstAppIndex + 2)
                if ($contextIndex -ge $firstAppIndex) { & $addCandidate $contextIndex 90 }
            }
        }

        if ([string]$event.kind -notin @('click', 'right-click')) { continue }
        $nextEvent = if ($eventIndex + 1 -lt $orderedEvents.Count) { $orderedEvents[$eventIndex + 1] } else { $null }
        $nextApp = if ($null -ne $nextEvent) { Get-MbRecorderItemAppKey -Item $nextEvent } else { '' }
        $resultIndex = -1
        if ($null -ne $nextEvent -and $nextApp -eq $eventApp -and [int]$nextEvent.timeMs -gt $eventTime) {
            # 次の操作直前は、クリック結果が描画し終わった最も安定したコマ。
            for ($frameIndex = $beforeIndex + 1; $frameIndex -lt $ordered.Count; $frameIndex++) {
                $frameTime = [int]$ordered[$frameIndex].timeMs
                if ($frameTime -ge [int]$nextEvent.timeMs) { break }
                if ((Get-MbRecorderItemAppKey -Item $ordered[$frameIndex]) -eq $eventApp) {
                    $resultIndex = $frameIndex
                }
            }
        } else {
            # 次が別アプリ、または最後の操作なら、現在アプリ内の最初の操作後コマを残す。
            for ($frameIndex = $beforeIndex + 1; $frameIndex -lt $ordered.Count; $frameIndex++) {
                $frameTime = [int]$ordered[$frameIndex].timeMs
                if ($frameTime -gt $eventTime + 2000) { break }
                if ((Get-MbRecorderItemAppKey -Item $ordered[$frameIndex]) -eq $eventApp) {
                    $resultIndex = $frameIndex
                    break
                }
            }
        }
        if ($resultIndex -ge 0) { & $addCandidate $resultIndex 95 }
    }

    $candidateIndexes = @($priorities.Keys | ForEach-Object { [int]$_ } | Sort-Object)
    if ($candidateIndexes.Count -le $Maximum) {
        return @($candidateIndexes | ForEach-Object { $ordered[$_] })
    }

    # 長い記録ではinput-evidenceを先に予約し、残りを時間軸へ分散する。
    # 先に等間隔候補で上限を埋めると、後段の優先度ループが一度も動かず、
    # 数式の最終表示まで落ちていた。
    $chosen = New-Object 'System.Collections.Generic.HashSet[int]'
    $requiredIndexes = @($candidateIndexes | Where-Object { [int]$priorities[[string]$_] -ge 110 })
    if ($requiredIndexes.Count -gt $Maximum) {
        for ($slot = 0; $slot -lt $Maximum; $slot++) {
            $position = [int][Math]::Round(($slot / [double]([Math]::Max(1, $Maximum - 1))) * ($requiredIndexes.Count - 1))
            [void]$chosen.Add([int]$requiredIndexes[$position])
        }
    } else {
        foreach ($requiredIndex in $requiredIndexes) { [void]$chosen.Add([int]$requiredIndex) }
    }
    for ($slot = 0; $slot -lt $Maximum -and $chosen.Count -lt $Maximum; $slot++) {
        $position = [int][Math]::Round(($slot / [double]([Math]::Max(1, $Maximum - 1))) * ($candidateIndexes.Count - 1))
        [void]$chosen.Add([int]$candidateIndexes[$position])
    }
    foreach ($candidateIndex in @($candidateIndexes | Sort-Object { -[int]$priorities[[string]$_] }, { [int]$_ })) {
        if ($chosen.Count -ge $Maximum) { break }
        [void]$chosen.Add([int]$candidateIndex)
    }
    return @($chosen | Sort-Object | ForEach-Object { $ordered[[int]$_] })
}

# クリックや入力の件数をそのまま手順数にせず、選定済みフレームから
# 「操作前／意味のある操作後」を組み立てる。Copilotが利用できない場合も、
# 読込中のイベント画像へ戻らず同じ安定フレームを確認できるようにする。
function New-MbRecorderLocalProposalId {
    param(
        [Parameter(Mandatory = $true)][string]$BeforeFrame,
        [AllowEmptyString()][string]$AfterFrame = '',
        [AllowEmptyCollection()][object[]]$EventIds = @(),
        [int]$TargetEventId = 0,
        [Parameter(Mandatory = $true)][string]$ActionKind
    )
    # 候補を表示した時と取り込む時で同じIDになるよう、候補の根拠だけから算出する。
    # 乱数IDでは、利用者の採否・文章修正を証拠台帳の候補へ結び付けられない。
    $identity = @($BeforeFrame, $AfterFrame, (@($EventIds) -join ','), $TargetEventId, $ActionKind) -join '|'
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($identity)
        $hex = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
        return 'local-' + $hex.Substring(0, 32)
    } finally { $sha.Dispose() }
}

function New-MbRecorderLocalFrameCandidates {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Events,
        [ValidateRange(8, 300)][int]$MaximumFrames = 30,
        [ValidateRange(250, 10000)][int]$EditMergeGapMs = 5000
    )
    $orderedEvents = @($Events | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    if (@($Frames).Count -lt 1) { return @() }

    # 入力イベントは確定時に記録されるため、クリックを取り逃した最初の入力では
    # 操作前画像がイベントより2秒以上前になることがある。
    # イベントを全件取り逃した場合は、記録全体の画像差分を安全側の候補にする。
    $scoped = if ($orderedEvents.Count -gt 0) {
        @(Select-MbRecorderEventWindowFrames -Frames @($Frames) -Events $orderedEvents -BeforePaddingMs 3000)
    } else {
        @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    }
    $selected = @(Select-MbRecorderTimelineFrames -Frames $scoped -Events $orderedEvents -Maximum $MaximumFrames)
    if ($selected.Count -lt 1) { return @() }

    function Test-MbRecorderSameSemanticTarget {
        param([AllowNull()]$First, [AllowNull()]$Second)
        if ($null -eq $First -or $null -eq $Second) { return $false }

        $firstName = if ($First.PSObject.Properties.Name -contains 'targetName') { ([string]$First.targetName).Trim() } else { '' }
        $secondName = if ($Second.PSObject.Properties.Name -contains 'targetName') { ([string]$Second.targetName).Trim() } else { '' }
        if ($firstName -and $secondName -and
            [string]::Equals($firstName, $secondName, [StringComparison]::OrdinalIgnoreCase)) { return $true }

        if ($First.PSObject.Properties.Name -contains 'rect' -and $null -ne $First.rect -and
            $Second.PSObject.Properties.Name -contains 'rect' -and $null -ne $Second.rect) {
            try {
                $overlapWidth = [Math]::Min([double]$First.rect.x2, [double]$Second.rect.x2) -
                    [Math]::Max([double]$First.rect.x1, [double]$Second.rect.x1)
                $overlapHeight = [Math]::Min([double]$First.rect.y2, [double]$Second.rect.y2) -
                    [Math]::Max([double]$First.rect.y1, [double]$Second.rect.y1)
                if ($overlapWidth -gt 0.0 -and $overlapHeight -gt 0.0) { return $true }
            } catch { }
        }

        if ($First.PSObject.Properties.Name -contains 'clickPoint' -and $null -ne $First.clickPoint -and
            $Second.PSObject.Properties.Name -contains 'clickPoint' -and $null -ne $Second.clickPoint) {
            try {
                $deltaX = [double]$First.clickPoint.x - [double]$Second.clickPoint.x
                $deltaY = [double]$First.clickPoint.y - [double]$Second.clickPoint.y
                if ([Math]::Sqrt(($deltaX * $deltaX) + ($deltaY * $deltaY)) -le 0.045) { return $true }
            } catch { }
        }
        return $false
    }

    # UIA・押下履歴など複数の経路が同じクリックを報告しても、確認候補は1件にする。
    # raw eventは消さず、同じsemantic groupのeventIdsとして全件を後段へ渡す。
    $semanticGroups = New-Object System.Collections.ArrayList
    foreach ($event in $orderedEvents) {
        $previousGroup = if ($semanticGroups.Count -gt 0) { @($semanticGroups[$semanticGroups.Count - 1]) } else { @() }
        $firstPrevious = if (@($previousGroup).Count -gt 0) { @($previousGroup)[0] } else { $null }
        $lastPrevious = if (@($previousGroup).Count -gt 0) { @($previousGroup)[@($previousGroup).Count - 1] } else { $null }
        $eventKind = [string]$event.kind
        $canMergeClick = $null -ne $lastPrevious -and $eventKind -in @('click', 'right-click') -and
            [string]$lastPrevious.kind -eq $eventKind -and
            (Get-MbRecorderItemAppKey -Item $lastPrevious) -eq (Get-MbRecorderItemAppKey -Item $event) -and
            [string]$lastPrevious.windowTitle -eq [string]$event.windowTitle -and
            ([int]$event.timeMs - [int]$lastPrevious.timeMs) -ge 0 -and
            ([int]$event.timeMs - [int]$lastPrevious.timeMs) -le 450 -and
            ([int]$event.timeMs - [int]$firstPrevious.timeMs) -le 450 -and
            (Test-MbRecorderSameSemanticTarget -First $lastPrevious -Second $event)
        if ($canMergeClick) {
            $semanticGroups[$semanticGroups.Count - 1] = [object[]]@(@($previousGroup) + @($event))
        } else {
            [void]$semanticGroups.Add([object[]]@($event))
        }
    }

    # 編集可能な場所へのクリックと、その直後の入力は利用者から見れば1手順。
    # 対象名が取れないEdgeでも、同じアプリ内で他のクリックを挟まない場合だけ結合する。
    $groups = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $semanticGroups.Count; $i++) {
        $currentGroup = @($semanticGroups[$i])
        if (@($currentGroup).Count -lt 1) { continue }
        $current = @($currentGroup)[0]
        $lastCurrent = @($currentGroup)[@($currentGroup).Count - 1]
        $items = New-Object System.Collections.ArrayList
        foreach ($item in $currentGroup) { [void]$items.Add($item) }
        if ([string]$current.kind -in @('click', 'right-click') -and $i + 1 -lt $semanticGroups.Count) {
            $nextGroup = @($semanticGroups[$i + 1])
            $next = if (@($nextGroup).Count -gt 0) { @($nextGroup)[0] } else { $null }
            if ($null -eq $next) { [void]$groups.Add(@($items)); continue }
            $sameApp = (Get-MbRecorderItemAppKey -Item $current) -eq
                (Get-MbRecorderItemAppKey -Item $next)
            $gap = [int]$next.timeMs - [int]$lastCurrent.timeMs
            $currentType = if ($current.PSObject.Properties.Name -contains 'targetType') { [string]$current.targetType } else { '' }
            $currentName = if ($current.PSObject.Properties.Name -contains 'targetName') { [string]$current.targetName } else { '' }
            $nextName = if ($next.PSObject.Properties.Name -contains 'targetName') { [string]$next.targetName } else { '' }
            $editable = $currentType -in @('ControlType.Edit', 'ControlType.DataItem', 'ControlType.ClickPoint') -or
                [string]::IsNullOrWhiteSpace($currentType)
            $sameTarget = [string]::IsNullOrWhiteSpace($currentName) -or [string]::IsNullOrWhiteSpace($nextName) -or
                [string]::Equals($currentName, $nextName, [StringComparison]::OrdinalIgnoreCase)
            if ([string]$next.kind -eq 'input' -and $sameApp -and $gap -ge 0 -and
                $gap -le $EditMergeGapMs -and $editable -and $sameTarget) {
                foreach ($item in $nextGroup) { [void]$items.Add($item) }
                $i++
            }
        }
        [void]$groups.Add(@($items))
    }

    $result = New-Object System.Collections.ArrayList
    $lastAfterFrameByApp = @{}
    for ($groupIndex = 0; $groupIndex -lt $groups.Count; $groupIndex++) {
        $group = @($groups[$groupIndex])
        if ($group.Count -lt 1) { continue }
        $firstEvent = $group[0]
        $lastEvent = $group[$group.Count - 1]
        $app = Get-MbRecorderItemAppKey -Item $firstEvent
        # ローカル提案ではAI向けに絞った一覧だけでなく、イベント周辺の原本も使う。
        # クリックを取り逃した場合でも、入力前の空欄と確定後の完成状態を復元できる。
        $appFrames = @($scoped | Where-Object {
            (Get-MbRecorderItemAppKey -Item $_) -eq $app
        } | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
        if ($appFrames.Count -lt 1) { continue }

        $hasInput = @($group | Where-Object { [string]$_.kind -eq 'input' }).Count -gt 0
        $inputEvidence = $null
        if ($hasInput) {
            # 入力イベントは無操作時間の後に確定するため、意味フレームはイベント時刻より
            # 最大2秒ほど前にある。クリックと結合済みならクリック以後だけを対象にする。
            $evidenceStart = if ($group.Count -gt 1) { [int]$firstEvent.timeMs - 250 } else { [int]$lastEvent.timeMs - 2200 }
            $inputEvidence = @($appFrames | Where-Object {
                $role = if ($_.PSObject.Properties.Name -contains 'role') { [string]$_.role } else { '' }
                $role -eq 'input-evidence' -and [int]$_.timeMs -ge $evidenceStart -and
                    [int]$_.timeMs -le ([int]$lastEvent.timeMs + 250)
            } | Select-Object -Last 1)
            if (@($inputEvidence).Count -gt 0) { $inputEvidence = @($inputEvidence)[0] } else { $inputEvidence = $null }
        }

        $beforeLimit = if ($null -ne $inputEvidence -and $group.Count -eq 1) {
            [int]$inputEvidence.timeMs - 1
        } else { [int]$firstEvent.timeMs }
        $before = @($appFrames | Where-Object {
            $role = if ($_.PSObject.Properties.Name -contains 'role') { [string]$_.role } else { '' }
            [int]$_.timeMs -le $beforeLimit -and $role -ne 'input-evidence'
        } | Select-Object -Last 1)
        if (@($before).Count -gt 0) { $before = @($before)[0] } else { $before = $appFrames[0] }

        $firstClickEvidence = @($appFrames | Where-Object {
            $_.PSObject.Properties.Name -contains 'evidenceEventId' -and
                [int]$_.evidenceEventId -eq [int]$firstEvent.index
        } | Select-Object -First 1)
        if ($firstClickEvidence.Count -gt 0) { $before = $firstClickEvidence[0] }

        # クリックして入力する手順では、クリック直前に残ったツールチップより、
        # 対象セル／入力欄へフォーカスが移った直後の安定画面の方が操作前として明確。
        # 入力証拠より前にある最初の描画済みコマへ置き換える。
        if ($group.Count -gt 1 -and $null -ne $inputEvidence -and
            [string]$firstEvent.kind -in @('click', 'right-click')) {
            $focusedBefore = @($appFrames | Where-Object {
                $role = if ($_.PSObject.Properties.Name -contains 'role') { [string]$_.role } else { '' }
                $role -ne 'input-evidence' -and
                    [int]$_.timeMs -ge ([int]$firstEvent.timeMs + 180) -and
                    [int]$_.timeMs -lt [int]$inputEvidence.timeMs -and
                    -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$_.windowTitle))
            } | Select-Object -First 1)
            if ($focusedBefore.Count -gt 0) { $before = $focusedBefore[0] }
        }

        $nextEventTime = $null
        for ($nextGroupIndex = $groupIndex + 1; $nextGroupIndex -lt $groups.Count; $nextGroupIndex++) {
            $nextGroup = @($groups[$nextGroupIndex])
            if ($nextGroup.Count -lt 1) { continue }
            $nextApp = Get-MbRecorderItemAppKey -Item $nextGroup[0]
            if ($nextApp -eq $app) { $nextEventTime = [int]$nextGroup[0].timeMs; break }
            # 別アプリへ移った後のフレームを、前の操作結果として結ばない。
            break
        }

        $after = $null
        if ($hasInput) {
            # input-evidenceは数式の保護には有効だが、短い数値では最後の1文字より前に
            # 撮られることがある。証拠画像より後、入力イベント時刻に最も近い通常コマを
            # 完成状態として優先し、無ければ証拠画像へ戻る。
            $completionStart = if ($null -ne $inputEvidence) { [int]$inputEvidence.timeMs + 1 } else { [int]$lastEvent.timeMs - 1800 }
            $completionEnd = [int]$lastEvent.timeMs + 1800
            # 次の操作が始まった後の画面は、見た目が安定していても現在の操作結果ではない。
            # 特にExcelの連続入力では、次セルの選択後を前セルの操作後へ混ぜると
            # before/afterの境界と赤枠が1手順ずつずれて見える。
            if ($null -ne $nextEventTime) { $completionEnd = [Math]::Min($completionEnd, [int]$nextEventTime - 1) }
            $completion = @($appFrames | Where-Object {
                $role = if ($_.PSObject.Properties.Name -contains 'role') { [string]$_.role } else { '' }
                $role -notin @('input-evidence', 'click-evidence') -and [int]$_.timeMs -ge $completionStart -and
                    [int]$_.timeMs -le $completionEnd -and
                    -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$_.windowTitle))
            } | Sort-Object { [Math]::Abs([int]$_.timeMs - [int]$lastEvent.timeMs) }, { [int]$_.timeMs } | Select-Object -First 1)
            if ($completion.Count -gt 0) { $after = $completion[0] } else { $after = $inputEvidence }
        }
        if ($null -eq $after) {
            if ($null -ne $nextEventTime) {
                # 次操作のclick-evidenceは「次の操作前」の証拠であり、現在の操作後ではない。
                # event時刻を越えない最後の通常フレームだけを結果候補にする。
                $after = @($appFrames | Where-Object {
                    $role = if ($_.PSObject.Properties.Name -contains 'role') { [string]$_.role } else { '' }
                    [int]$_.timeMs -gt [int]$firstEvent.timeMs -and [int]$_.timeMs -lt [int]$nextEventTime -and
                        $role -ne 'click-evidence' -and
                        -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$_.windowTitle))
                } | Select-Object -Last 1)
            } else {
                $after = @($appFrames | Where-Object {
                    [int]$_.timeMs -gt [int]$lastEvent.timeMs -and [int]$_.timeMs -le ([int]$lastEvent.timeMs + 2200) -and
                        [string]$_.id -ne [string]$before.id
                } | Select-Object -First 1)
            }
            if (@($after).Count -gt 0) { $after = @($after)[0] } else { $after = $null }
        }

        # 同じアプリの連続手順は、直前の完成状態を次の操作前として共有する。
        # これにより入力途中の通常コマが次手順の「操作前」へ紛れ込まない。
        if ($lastAfterFrameByApp.ContainsKey($app)) {
            $previousAfter = $lastAfterFrameByApp[$app]
            if ($null -ne $previousAfter -and [int]$previousAfter.timeMs -le ([int]$lastEvent.timeMs + 1800)) {
                $before = $previousAfter
            }
        }

        $eventIds = @($group | ForEach-Object { [int]$_.index })
        $targetEventId = 0
        $anchorEvents = if ($hasInput) {
            $clickAnchors = @($group | Where-Object { [string]$_.kind -in @('click', 'right-click') })
            if ($clickAnchors.Count -gt 0) {
                $clickAnchors
            } else {
                # 通常のinput矩形は直前セルの残留である可能性があるため使わない。
                # 前後の実測Excelセルから一意に復元したものだけ例外とする。
                @($group | Where-Object {
                    [string]$_.kind -eq 'input' -and $_.PSObject.Properties.Name -contains 'targetSource' -and
                    [string]$_.targetSource -eq 'Excel-sequence-inferred'
                })
            }
        } else { @($group) }
        foreach ($event in $anchorEvents) {
            if (Test-MbRecorderEventHasVisualAnchor -Event $event) { $targetEventId = [int]$event.index; break }
        }
        $actionKind = $(if ($hasInput) { 'input' } else { [string]$firstEvent.kind })
        $afterFrameId = $(if ($null -ne $after -and [string]$after.id -ne [string]$before.id) { [string]$after.id } else { '' })
        [void]$result.Add([pscustomobject]@{
            id = New-MbRecorderLocalProposalId -BeforeFrame ([string]$before.id) -AfterFrame $afterFrameId `
                -EventIds $eventIds -TargetEventId $targetEventId -ActionKind $actionKind
            beforeFrame = [string]$before.id
            afterFrame = $afterFrameId
            eventIds = $eventIds
            targetEventId = $targetEventId
            actionKind = $actionKind
            timeMs = [int]$firstEvent.timeMs
            beforeImage = [string]$before.image
            afterImage = $(if ($null -ne $after -and [string]$after.id -ne [string]$before.id) { [string]$after.image } else { '' })
            source = 'local'
        })
        if ($null -ne $after) { $lastAfterFrameByApp[$app] = $after }
    }

    # クリック監視を取り逃しても、時系列画像にはセル値などの変化が残っている。
    # イベント候補で既に表せていない大きな変化だけを、要確認の補助候補として足す。
    $existingAfter = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in $result) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate.afterFrame)) { [void]$existingAfter.Add([string]$candidate.afterFrame) }
    }
    $inputRanges = New-Object System.Collections.ArrayList
    foreach ($groupValue in $groups) {
        $group = @($groupValue)
        if ($group.Count -lt 1 -or @($group | Where-Object { [string]$_.kind -eq 'input' }).Count -lt 1) { continue }
        [void]$inputRanges.Add([pscustomobject]@{
            app = Get-MbRecorderItemAppKey -Item $group[0]
            startMs = if ($group.Count -gt 1) { [int]$group[0].timeMs - 250 } else { [int]$group[0].timeMs - 2200 }
            # 入力確定直後の計算結果やセル移動は同じ入力手順の操作後であり、
            # 独立した画面変化手順にはしない。
            endMs = [int]$group[$group.Count - 1].timeMs + 1800
        })
    }
    $eligibleVisualFrames = New-Object System.Collections.ArrayList
    foreach ($frame in $selected) {
        $change = if ($frame.PSObject.Properties.Name -contains 'visualChange') { [double]$frame.visualChange } else { 0.0 }
        if ($change -lt 0.00035 -or $existingAfter.Contains([string]$frame.id) -or
            (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$frame.windowTitle))) { continue }
        $frameApp = Get-MbRecorderItemAppKey -Item $frame
        $frameTime = [int]$frame.timeMs

        # 入力イベントが取れている区間は、そのグループの最新input-evidenceで十分。
        # 途中の数文字やカーソル点滅を別の「画面変化」手順として追加しない。
        $insideRecordedInput = @($inputRanges | Where-Object {
            [string]$_.app -eq $frameApp -and $frameTime -ge [int]$_.startMs -and $frameTime -le [int]$_.endMs
        }).Count -gt 0
        if ($insideRecordedInput) { continue }

        # 入力中の最終証拠が直後にあるなら、途中の式ではなく最終証拠を使う。
        $betterInputEvidence = @($selected | Where-Object {
            $role = if ($_.PSObject.Properties.Name -contains 'role') { [string]$_.role } else { '' }
            $role -eq 'input-evidence' -and
                (Get-MbRecorderItemAppKey -Item $_) -eq $frameApp -and
                [int]$_.timeMs -gt $frameTime -and [int]$_.timeMs -le ($frameTime + 1600)
        }).Count -gt 0
        if ($betterInputEvidence) { continue }

        # 既にイベント候補の操作後とほぼ同じ時刻なら重複させない。
        $nearExistingResult = $false
        foreach ($candidate in $result) {
            if ([string]::IsNullOrWhiteSpace([string]$candidate.afterFrame)) { continue }
            $candidateFrame = @($selected | Where-Object { [string]$_.id -eq [string]$candidate.afterFrame } | Select-Object -First 1)
            $resultLag = if ($candidateFrame.Count -gt 0) { [int]$candidateFrame[0].timeMs - $frameTime } else { -1 }
            if ($candidateFrame.Count -gt 0 -and $resultLag -ge 0 -and $resultLag -le 2200) {
                $nearExistingResult = $true; break
            }
        }
        if ($nearExistingResult) { continue }

        # クリック直後の単なる選択枠変化は、そのクリック候補が既にある。
        $nearClick = @($orderedEvents | Where-Object {
            [string]$_.kind -in @('click', 'right-click') -and
                (Get-MbRecorderItemAppKey -Item $_) -eq $frameApp -and
                [Math]::Abs([int]$_.timeMs - $frameTime) -le 900
        }).Count -gt 0
        if ($nearClick) { continue }

        [void]$eligibleVisualFrames.Add($frame)
    }

    # 同じアプリで短時間に続く差分は、描画途中を複数手順にせず1つの変化episodeにする。
    # 先頭変化の直前をbefore、episode末尾の変化をafterとして扱う。
    $visualEpisodes = New-Object System.Collections.ArrayList
    foreach ($frame in @($eligibleVisualFrames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })) {
        $previousEpisode = if ($visualEpisodes.Count -gt 0) { @($visualEpisodes[$visualEpisodes.Count - 1]) } else { @() }
        $lastFrame = if (@($previousEpisode).Count -gt 0) { @($previousEpisode)[@($previousEpisode).Count - 1] } else { $null }
        $sameEpisode = $null -ne $lastFrame -and
            (Get-MbRecorderItemAppKey -Item $lastFrame) -eq (Get-MbRecorderItemAppKey -Item $frame) -and
            ([int]$frame.timeMs - [int]$lastFrame.timeMs) -ge 0 -and
            ([int]$frame.timeMs - [int]$lastFrame.timeMs) -le 1500
        if ($sameEpisode) {
            $visualEpisodes[$visualEpisodes.Count - 1] = [object[]]@(@($previousEpisode) + @($frame))
        } else {
            [void]$visualEpisodes.Add([object[]]@($frame))
        }
    }

    foreach ($episodeValue in $visualEpisodes) {
        $episode = @($episodeValue)
        if (@($episode).Count -lt 1) { continue }
        $firstFrame = @($episode)[0]
        $afterFrame = @($episode)[@($episode).Count - 1]
        $frameApp = Get-MbRecorderItemAppKey -Item $firstFrame
        $before = @($selected | Where-Object {
            [int]$_.timeMs -lt [int]$firstFrame.timeMs -and
                (Get-MbRecorderItemAppKey -Item $_) -eq $frameApp -and
                -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$_.windowTitle))
        } | Sort-Object { [int]$_.timeMs }, { [int]$_.index } | Select-Object -Last 1)
        if ($before.Count -lt 1 -or [string]$before[0].id -eq [string]$afterFrame.id) { continue }
        # 画像差分だけでは、近くのイベントがこの変化を起こしたとは断定できない。
        # 誤った赤枠を付けないため、必ず根拠イベントなし・要確認として残す。
        [void]$result.Add([pscustomobject]@{
            id = New-MbRecorderLocalProposalId -BeforeFrame ([string]$before[0].id) -AfterFrame ([string]$afterFrame.id) `
                -EventIds @() -TargetEventId 0 -ActionKind 'visual-change'
            beforeFrame = [string]$before[0].id
            afterFrame = [string]$afterFrame.id
            eventIds = @()
            targetEventId = 0
            actionKind = 'visual-change'
            timeMs = [int]$before[0].timeMs
            beforeImage = [string]$before[0].image
            afterImage = [string]$afterFrame.image
            source = 'local'
        })
        [void]$existingAfter.Add([string]$afterFrame.id)
    }
    return @($result | Sort-Object { [int]$_.timeMs })
}

# ローカル候補の前後画像だけを取り出す補助関数。診断・比較用に残すが、
# Copilotの原本選定には使わない。同じ境界画像は1枚へまとめる。
function Select-MbRecorderCandidateFrames {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Candidates
    )
    if (@($Frames).Count -lt 1 -or @($Candidates).Count -lt 1) { return @() }
    $frameMap = @{}
    foreach ($frame in @($Frames)) { $frameMap[[string]$frame.id] = $frame }
    $selectedIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in @($Candidates)) {
        foreach ($propertyName in @('beforeFrame', 'afterFrame')) {
            if ($candidate.PSObject.Properties.Name -notcontains $propertyName) { continue }
            $id = ([string]$candidate.$propertyName).Trim()
            if ($id -and $frameMap.ContainsKey($id)) { [void]$selectedIds.Add($id) }
        }
    }
    # 数式を貼り付けた場合、確定後の操作結果だけでは式が1550などの値に変わる。
    # 通常の途中入力は増やさず、貼り付けと明示された確定前証拠だけを追加する。
    foreach ($frame in @($Frames)) {
        $evidenceKind = if ($frame.PSObject.Properties.Name -contains 'evidenceKind') { [string]$frame.evidenceKind } else { '' }
        if ($evidenceKind -eq 'paste') { [void]$selectedIds.Add([string]$frame.id) }
    }
    $ordered = @($selectedIds | ForEach-Object { $frameMap[[string]$_] } |
        Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    $seenImages = @{}
    $deduplicated = New-Object System.Collections.ArrayList
    foreach ($frame in $ordered) {
        $hash = if ($frame.PSObject.Properties.Name -contains 'imageSha256') { [string]$frame.imageSha256 } else { '' }
        $app = Get-MbRecorderItemAppKey -Item $frame
        $key = if ($hash) { 'sha256:' + $hash + '|app:' + $app } else { 'id:' + [string]$frame.id }
        $isNearDuplicate = $seenImages.ContainsKey($key) -and
            ([int]$frame.timeMs - [int]$seenImages[$key]) -ge 0 -and
            ([int]$frame.timeMs - [int]$seenImages[$key]) -le 1500
        if (-not $isNearDuplicate) { [void]$deduplicated.Add($frame) }
        $seenImages[$key] = [int]$frame.timeMs
    }
    return @($deduplicated)
}

function Select-MbRecorderEventWindowFrames {
    param(
        [Parameter(Mandatory = $true)][object[]]$Frames,
        [Parameter(Mandatory = $true)][object[]]$Events,
        [ValidateRange(0, 10000)][int]$BeforePaddingMs = 1500,
        [ValidateRange(0, 10000)][int]$AfterPaddingMs = 2000
    )
    if ($Frames.Count -lt 2 -or $Events.Count -lt 1) { return @($Frames) }
    $times = @($Events | ForEach-Object { try { [int]$_.timeMs } catch { $null } } | Where-Object { $null -ne $_ })
    if ($times.Count -lt 1) { return @($Frames) }
    $allowedApps = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($event in $Events) {
        $app = Get-MbRecorderItemAppKey -Item $event
        if (-not [string]::IsNullOrWhiteSpace($app)) { [void]$allowedApps.Add($app) }
    }
    $first = [int](($times | Measure-Object -Minimum).Minimum) - $BeforePaddingMs
    $last = [int](($times | Measure-Object -Maximum).Maximum) + $AfterPaddingMs
    $scoped = @($Frames | Where-Object {
        if ([int]$_.timeMs -lt $first -or [int]$_.timeMs -gt $last) { return $false }
        if ($allowedApps.Count -lt 1) { return $true }
        $frameApp = Get-MbRecorderItemAppKey -Item $_
        return $allowedApps.Contains($frameApp)
    })
    if ($scoped.Count -lt 2) { return @($Frames) }
    return $scoped
}

function Repair-MbRecorderExcelInputEventAnchors {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Events
    )

    $ordered = @($Events | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    for ($index = 1; $index -lt ($ordered.Count - 1); $index++) {
        $current = $ordered[$index]
        if ([string]$current.kind -ne 'input' -or (Test-MbRecorderEventHasVisualAnchor -Event $current) -or
            (Get-MbRecorderItemAppKey -Item $current) -ne 'excel') { continue }

        # ExcelでEnter確定直後に次セルを素早く押すと、画面保存中の短いクリックを
        # Windowsの押下履歴から拾えないことがある。前後に同じ列の隣接セルがあり、
        # 欠けた行が一意に決まる場合だけ、2つの実測矩形を線形補間する。
        # 名前や座標が曖昧な一般画面では決して推測しない。
        $previous = $null
        for ($scan = $index - 1; $scan -ge 0; $scan--) {
            if ([string]$ordered[$scan].kind -in @('click', 'right-click')) { $previous = $ordered[$scan]; break }
        }
        $next = $null
        for ($scan = $index + 1; $scan -lt $ordered.Count; $scan++) {
            if ([string]$ordered[$scan].kind -in @('click', 'right-click')) { $next = $ordered[$scan]; break }
        }
        if ($null -eq $previous -or $null -eq $next -or
            (Get-MbRecorderItemAppKey -Item $previous) -ne 'excel' -or
            (Get-MbRecorderItemAppKey -Item $next) -ne 'excel' -or
            ([int]$current.timeMs - [int]$previous.timeMs) -gt 6000 -or
            ([int]$next.timeMs - [int]$current.timeMs) -gt 6000 -or
            -not (Test-MbRecorderEventHasVisualAnchor -Event $previous) -or
            -not (Test-MbRecorderEventHasVisualAnchor -Event $next)) { continue }

        $previousMatch = [regex]::Match(([string]$previous.targetName).Trim().ToUpperInvariant(), '^([A-Z]{1,3})([1-9][0-9]*)$')
        $nextMatch = [regex]::Match(([string]$next.targetName).Trim().ToUpperInvariant(), '^([A-Z]{1,3})([1-9][0-9]*)$')
        if (-not $previousMatch.Success -or -not $nextMatch.Success -or
            $previousMatch.Groups[1].Value -ne $nextMatch.Groups[1].Value) { continue }
        $previousRow = [int]$previousMatch.Groups[2].Value
        $nextRow = [int]$nextMatch.Groups[2].Value
        if ($nextRow -ne ($previousRow + 2)) { continue }

        $previousRect = $previous.rect; $nextRect = $next.rect
        $xTolerance = 0.01
        if ([Math]::Abs([double]$previousRect.x1 - [double]$nextRect.x1) -gt $xTolerance -or
            [Math]::Abs([double]$previousRect.x2 - [double]$nextRect.x2) -gt $xTolerance) { continue }
        $rect = [pscustomobject]@{
            x1 = [Math]::Round(([double]$previousRect.x1 + [double]$nextRect.x1) / 2.0, 6)
            y1 = [Math]::Round(([double]$previousRect.y1 + [double]$nextRect.y1) / 2.0, 6)
            x2 = [Math]::Round(([double]$previousRect.x2 + [double]$nextRect.x2) / 2.0, 6)
            y2 = [Math]::Round(([double]$previousRect.y2 + [double]$nextRect.y2) / 2.0, 6)
        }
        $cellName = $previousMatch.Groups[1].Value + ($previousRow + 1)
        $candidate = [pscustomobject]@{
            id = 'excel-sequence-1'
            source = 'Excel-sequence-inferred'
            confidence = 'medium'
            label = $cellName
            targetType = 'ControlType.DataItem'
            rect = $rect
        }
        foreach ($property in @{
            targetName = $cellName
            targetType = 'ControlType.DataItem'
            rect = $rect
            targetSource = 'Excel-sequence-inferred'
            confidence = 'medium'
            targetCandidateId = 'excel-sequence-1'
            targetCandidates = @($candidate)
        }.GetEnumerator()) {
            $current | Add-Member -NotePropertyName $property.Key -NotePropertyValue $property.Value -Force
        }
    }
    return @($ordered)
}

# Copilotにはローカル候補で選んだ操作前／操作後だけでなく、記録区間の原本に
# 近い時系列を渡す。上限以下は全て残し、長い記録だけ操作前後と入力証拠を
# 保護しながら時間軸全体へ分散する。これにより、イベント取得に失敗した操作も
# 一覧画像を見たCopilotが拾える。
function Select-MbRecorderCopilotSourceFrames {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Events = @(),
        [ValidateRange(8, 300)][int]$Maximum = 30
    )
    # イベントの時刻・アプリで原本を切らない。UIA/DOMがあるアプリのイベントだけ
    # 取れた場合でも、イベントが欠けたEdgeやExcelの周期コマをAIへ渡すため。
    $scoped = @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index } | Where-Object {
        $title = ([string]$_.windowTitle).Trim()
        $process = if ($_.PSObject.Properties.Name -contains 'processName') { ([string]$_.processName).Trim() } else { '' }
        # 記録開始・停止のためにManualBuilderへ戻った画面は、AIへ見せても手順では
        # なくノイズになる。プロセス名を固定せず、製品名が明確な画面だけを外す。
        return ($title -notmatch '(?i)^ManualBuilder(?:\s*[-–—|]\s*.*)?$|操作を記録して手順書を作る' -and
            $process -notmatch '(?i)^ManualBuilder(?:\.exe)?$')
    })

    # Computer-useによる実機試験などでは、停止ボタンを押した制御アプリ自体が
    # 最後の低信頼click-pointとして記録される。末尾12秒以内に初登場し、名前の
    # ない低信頼クリックしか持たないアプリは記録終了UIとみなして外す。
    if ($scoped.Count -ge 3 -and @($Events).Count -gt 0) {
        $lastApp = Get-MbRecorderItemAppKey -Item $scoped[$scoped.Count - 1]
        $tailStart = $scoped.Count - 1
        while ($tailStart -gt 0 -and (Get-MbRecorderItemAppKey -Item $scoped[$tailStart - 1]) -eq $lastApp) { $tailStart-- }
        $seenEarlier = $false
        for ($i = 0; $i -lt $tailStart; $i++) {
            if ((Get-MbRecorderItemAppKey -Item $scoped[$i]) -eq $lastApp) { $seenEarlier = $true; break }
        }
        $tailDuration = [int]$scoped[$scoped.Count - 1].timeMs - [int]$scoped[$tailStart].timeMs
        $tailEvents = @($Events | Where-Object {
            (Get-MbRecorderItemAppKey -Item $_) -eq $lastApp -and
            [int]$_.timeMs -ge ([int]$scoped[$tailStart].timeMs - 500)
        })
        $onlyUnnamedLowConfidenceClicks = $tailEvents.Count -gt 0
        foreach ($event in $tailEvents) {
            $target = if ($event.PSObject.Properties.Name -contains 'targetName') { ([string]$event.targetName).Trim() } else { '' }
            $source = if ($event.PSObject.Properties.Name -contains 'targetSource') { ([string]$event.targetSource).Trim() } else { '' }
            $confidence = if ($event.PSObject.Properties.Name -contains 'confidence') { ([string]$event.confidence).Trim() } else { '' }
            if ($target -or $source -ne 'click-point' -or $confidence -ne 'low') { $onlyUnnamedLowConfidenceClicks = $false; break }
        }
        if ($tailStart -gt 0 -and -not $seenEarlier -and $tailDuration -le 12000 -and $onlyUnnamedLowConfidenceClicks) {
            $scoped = @($scoped | Select-Object -First $tailStart)
        }
    }
    if ($scoped.Count -le $Maximum) { return @($scoped) }

    $selected = New-Object 'System.Collections.Generic.HashSet[int]'
    [void]$selected.Add(0)
    [void]$selected.Add($scoped.Count - 1)

    # イベントが欠けても、値の確定やページ遷移が起きた付近は一覧に必ず残す。
    # 差分が出た瞬間だけでなく、0.8秒後の安定画面も保護することで、入力途中や
    # 読み込み中のコマを代表画像にしにくくする。
    for ($frameIndex = 1; $frameIndex -lt $scoped.Count; $frameIndex++) {
        $change = if ($scoped[$frameIndex].PSObject.Properties.Name -contains 'visualChange') { [double]$scoped[$frameIndex].visualChange } else { 0.0 }
        $app = Get-MbRecorderItemAppKey -Item $scoped[$frameIndex]
        $previousApp = Get-MbRecorderItemAppKey -Item $scoped[$frameIndex - 1]
        $titleChanged = $app -eq $previousApp -and ([string]$scoped[$frameIndex].windowTitle) -ne ([string]$scoped[$frameIndex - 1].windowTitle)
        if ($change -lt 0.00035 -and -not $titleChanged) { continue }
        [void]$selected.Add([Math]::Max(0, $frameIndex - 1))
        [void]$selected.Add($frameIndex)
        $changedAt = [int]$scoped[$frameIndex].timeMs
        for ($afterIndex = $frameIndex + 1; $afterIndex -lt $scoped.Count; $afterIndex++) {
            if ((Get-MbRecorderItemAppKey -Item $scoped[$afterIndex]) -ne $app) { break }
            $elapsed = [int]$scoped[$afterIndex].timeMs - $changedAt
            if ($elapsed -ge 800) { [void]$selected.Add($afterIndex); break }
            if ($elapsed -gt 3000) { break }
        }
    }

    # 数式や貼り付け値など、確定後の画面から復元できない入力証拠は必ず保護する。
    for ($frameIndex = 0; $frameIndex -lt $scoped.Count; $frameIndex++) {
        $role = if ($scoped[$frameIndex].PSObject.Properties.Name -contains 'role') { [string]$scoped[$frameIndex].role } else { '' }
        if ($role -eq 'input-evidence') { [void]$selected.Add($frameIndex) }
    }

    # 各操作の直前と直後をアンカーにする。対象取得の精度には依存せず、時刻と
    # アプリだけを使うため、UI Automationが欠けても周辺の原本コマは残る。
    foreach ($event in @($Events)) {
        $eventTime = [int]$event.timeMs
        $eventApp = Get-MbRecorderItemAppKey -Item $event
        $beforeIndex = -1
        $afterIndex = -1
        for ($frameIndex = 0; $frameIndex -lt $scoped.Count; $frameIndex++) {
            $frameApp = Get-MbRecorderItemAppKey -Item $scoped[$frameIndex]
            if ($eventApp -and $frameApp -ne $eventApp) { continue }
            $frameTime = [int]$scoped[$frameIndex].timeMs
            if ($frameTime -le $eventTime) { $beforeIndex = $frameIndex }
            elseif ($afterIndex -lt 0) { $afterIndex = $frameIndex }
        }
        if ($beforeIndex -ge 0) { [void]$selected.Add($beforeIndex) }
        if ($afterIndex -ge 0) { [void]$selected.Add($afterIndex) }
    }

    # アンカーだけで上限を超える場合は、既存の優先選定で入力証拠を保護する。
    if ($selected.Count -gt $Maximum) {
        return @(Select-MbRecorderTimelineFrames -Frames $scoped -Events @($Events) -Maximum $Maximum)
    }

    # 残りは記録全体から均等に加える。丸めの重複で不足した場合は、既に選んだ
    # コマから最も離れた時刻を順に足して、可能な限りMaximum枚まで使う。
    $uniformSlots = $Maximum - $selected.Count
    for ($slot = 1; $slot -le $uniformSlots; $slot++) {
        $position = [int][Math]::Round(($slot / [double]($uniformSlots + 1)) * ($scoped.Count - 1))
        [void]$selected.Add($position)
    }
    while ($selected.Count -lt $Maximum) {
        $bestIndex = -1
        [long]$bestDistance = -1
        for ($frameIndex = 0; $frameIndex -lt $scoped.Count; $frameIndex++) {
            if ($selected.Contains($frameIndex)) { continue }
            [long]$nearest = [long]::MaxValue
            foreach ($chosenIndex in $selected) {
                $distance = [Math]::Abs([long]$scoped[$frameIndex].timeMs - [long]$scoped[[int]$chosenIndex].timeMs)
                if ($distance -lt $nearest) { $nearest = $distance }
            }
            if ($nearest -gt $bestDistance) { $bestDistance = $nearest; $bestIndex = $frameIndex }
        }
        if ($bestIndex -lt 0) { break }
        [void]$selected.Add($bestIndex)
    }
    return @($selected | Sort-Object | ForEach-Object { $scoped[[int]$_] })
}

function Save-MbRecorderContactSheetJpeg {
    param([Parameter(Mandatory = $true)]$Bitmap, [Parameter(Mandatory = $true)][string]$Path, [long]$Quality = 90)
    $encoder = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
    $parameters = New-Object Drawing.Imaging.EncoderParameters -ArgumentList 1
    $parameters.Param[0] = New-Object Drawing.Imaging.EncoderParameter -ArgumentList @([Drawing.Imaging.Encoder]::Quality, $Quality)
    try { $Bitmap.Save($Path, $encoder, $parameters) } finally { $parameters.Dispose() }
}

function New-MbRecorderContactSheets {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory = $true)][string]$FramesDirectory,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [ValidateRange(6, 24)][int]$FramesPerSheet = 10,
        [ValidateRange(2, 6)][int]$Columns = 2
    )
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }
    # 3列640pxでは実機のM365 Copilotが電卓の値を読めず、曖昧な手順になった。
    # 横1920pxは維持し、既定2列で1コマを960x540まで大きくする。
    $sheetWidth = 1920
    $tileWidth = [int][Math]::Floor($sheetWidth / $Columns)
    $tileHeight = [int][Math]::Round($tileWidth * 9.0 / 16.0)
    $labelHeight = 34
    $sheetWidth = $tileWidth * $Columns
    $sheets = New-Object System.Collections.ArrayList
    $ordered = @($Frames)

    for ($offset = 0; $offset -lt $ordered.Count; $offset += $FramesPerSheet) {
        $number = [int]($offset / $FramesPerSheet) + 1
        $pageFrames = @($ordered | Select-Object -Skip $offset -First $FramesPerSheet)
        # 最終ページを15コマ分の黒い余白で水増ししない。添付サイズを抑え、
        # 少数コマでは画像そのものがCopilot画面上で大きく見えるようにする。
        $rows = [Math]::Max(1, [int][Math]::Ceiling($pageFrames.Count / [double]$Columns))
        $sheetHeight = ($tileHeight + $labelHeight) * $rows
        $bitmap = New-Object Drawing.Bitmap -ArgumentList @($sheetWidth, $sheetHeight, [Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $graphics = $null
        $font = $null
        $smallFont = $null
        try {
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([Drawing.Color]::FromArgb(30, 34, 40))
            $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $font = New-Object Drawing.Font -ArgumentList @('Segoe UI', 12, [Drawing.FontStyle]::Bold, [Drawing.GraphicsUnit]::Pixel)
            $smallFont = New-Object Drawing.Font -ArgumentList @('Yu Gothic UI', 11, [Drawing.FontStyle]::Regular, [Drawing.GraphicsUnit]::Pixel)
            for ($i = 0; $i -lt $pageFrames.Count; $i++) {
                $frame = $pageFrames[$i]
                $column = $i % $Columns
                $row = [Math]::Floor($i / $Columns)
                $left = $column * $tileWidth
                $top = $row * ($tileHeight + $labelHeight)
                $graphics.FillRectangle([Drawing.Brushes]::Black, $left, $top, $tileWidth, $tileHeight)
                $fileName = [string]$frame.image
                if ($fileName -notmatch '^frame-\d{5}\.jpg$' -or [IO.Path]::GetFileName($fileName) -ne $fileName) { continue }
                $sourcePath = Join-Path $FramesDirectory $fileName
                if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { continue }
                $image = $null
                try {
                    $image = [Drawing.Image]::FromFile($sourcePath)
                    $scale = [Math]::Min($tileWidth / [double]$image.Width, $tileHeight / [double]$image.Height)
                    $width = [int][Math]::Round($image.Width * $scale)
                    $height = [int][Math]::Round($image.Height * $scale)
                    $x = $left + [int](($tileWidth - $width) / 2)
                    $y = $top + [int](($tileHeight - $height) / 2)
                    $graphics.DrawImage($image, $x, $y, $width, $height)
                } finally { if ($null -ne $image) { $image.Dispose() } }

                $labelTop = $top + $tileHeight
                $labelBrush = $null
                try {
                    $labelBrush = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(238, 242, 247))
                    $graphics.FillRectangle($labelBrush, $left, $labelTop, $tileWidth, $labelHeight)
                } finally { if ($null -ne $labelBrush) { $labelBrush.Dispose() } }
                $primary = ('{0}  {1}' -f [string]$frame.id, (Format-MbRecorderTimeCode -Milliseconds ([int]$frame.timeMs)))
                $graphics.DrawString($primary, $font, [Drawing.Brushes]::Black, $left + 8, $labelTop + 8)
                $title = ([string]$frame.windowTitle) -replace '\s+', ' '
                if ($title.Length -gt 38) { $title = $title.Substring(0, 37) + '…' }
                $graphics.DrawString($title, $smallFont, [Drawing.Brushes]::DimGray, $left + 138, $labelTop + 9)
            }
            $path = Join-Path $OutputDirectory ('timeline-{0:d3}.jpg' -f $number)
            Save-MbRecorderContactSheetJpeg -Bitmap $bitmap -Path $path
            [void]$sheets.Add([pscustomobject]@{
                number = $number
                path = $path
                fileName = [IO.Path]::GetFileName($path)
                frames = @($pageFrames)
                firstTimeMs = $(if ($pageFrames.Count -gt 0) { [int]$pageFrames[0].timeMs } else { 0 })
                lastTimeMs = $(if ($pageFrames.Count -gt 0) { [int]$pageFrames[$pageFrames.Count - 1].timeMs } else { 0 })
            })
        } finally {
            foreach ($disposable in @($smallFont, $font, $graphics, $bitmap)) {
                if ($null -ne $disposable) { try { $disposable.Dispose() } catch { } }
            }
        }
    }
    return @($sheets)
}

function New-MbRecorderCopilotPrompt {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [AllowEmptyCollection()][object[]]$Events = @(),
        [AllowEmptyCollection()][object[]]$InteractionGroups = @(),
        [ValidateRange(1, 100)][int]$PacketNumber = 1,
        [ValidateRange(1, 100)][int]$TotalPackets = 1,
        [AllowEmptyString()][string]$PreviousFrameId = '',
        [string]$Marker = 'MB_END'
    )
    $builder = New-Object Text.StringBuilder
    [void]$builder.AppendLine('あなたは操作マニュアルの編集者です。添付は記録画面を時系列に並べたコンタクトシートです。')
    [void]$builder.AppendLine(("これは全{0}枚中{1}枚目です。画像内のフレームID順に、この区間だけを判定してください。" -f $TotalPackets, $PacketNumber))
    if ($PacketNumber -gt 1 -and -not [string]::IsNullOrWhiteSpace($PreviousFrameId)) {
        [void]$builder.AppendLine(("前の一覧は {0} までです。この一覧はその続きであり、先頭の状態との差分も確認してください。" -f $PreviousFrameId))
    }
    [void]$builder.AppendLine('プログラムの画像差分だけで手順を決めず、前後の流れから「利用者の目的を達成するために必要な手順」を作ってください。')
    [void]$builder.AppendLine('規則:')
    if (@($InteractionGroups).Count -gt 0) {
        # グループ境界がある実運用では重複した一般規則を送らない。M365は長い指示と
        # 大画像を同時に処理すると汎用エラーになりやすいため、判定に必要な制約だけに絞る。
        [void]$builder.AppendLine('- 下記の各Gをちょうど1手順にする。G同士を結合せず、G内のクリック＋入力も分割しない。各GにはそのGの操作だけを書く。')
        [void]$builder.AppendLine('- clickは押した名称を書く。「表示を確認」へ言い換えない。Gにない操作、初期描画や既存値を操作として追加しない。')
        [void]$builder.AppendLine('- targetは操作開始位置の手がかり。画面と矛盾すれば画面を優先し、afterFrameに範囲選択が見える場合はその範囲を書く。')
        [void]$builder.AppendLine('- 読込中・暗転・スピナーは選ばず、beforeFrameは操作対象が見える安定コマ、afterFrameは結果が最初に安定したコマにする。')
        [void]$builder.AppendLine('- ManualBuilderへ戻る操作、記録開始・停止、前面アプリが切り替わっただけの境界は除外する。同じEdge内のページ遷移は除外しない。')
        [void]$builder.AppendLine('- 画面に見える名称・値・英字の大文字小文字を保つ。Excelの数式バーが見える場合は「=SUM(B2:B3)」のように数式をそのまま書く。')
        [void]$builder.AppendLine('- 不明点は推測せず confidence=low にする。')
    } else {
        [void]$builder.AppendLine('- 同じ対象への連続クリックと入力は、1つの意味のある手順にまとめる。')
        [void]$builder.AppendLine('- 電卓で数値・演算子・関数・実行を続けて1つの式を完成させる操作は、ボタンごとに分けず1つの計算手順にまとめる。')
        [void]$builder.AppendLine('- スクロールや表示位置の移動は独立した手順にしない。移動後に押したリンクやボタンを、その操作前後の1手順として残す。')
        [void]$builder.AppendLine('- ある手順のafterFrameが次のbeforeFrameと同じで、一連のボタン操作が同じ式や入力を完成させる場合は、必ず1手順へ統合して操作順を説明に列挙する。')
        [void]$builder.AppendLine('- リンクやボタンで別ページ・ダイアログを開く操作と、開いた画面内で行う入力・計算は目的が異なるため、隣接していても別手順にする。')
        [void]$builder.AppendLine('- B2、B3、B4のように対象が異なる入力イベントは省略せず、対象ごとに別の手順にする。')
        [void]$builder.AppendLine('- 読込中、暗転、スピナー、中間的なアニメーションのコマは代表画像に選ばない。')
        [void]$builder.AppendLine('- ManualBuilderへ戻る操作、記録開始・停止、無意味なフォーカス移動は除外する。')
        [void]$builder.AppendLine('- EdgeからExcelなど、前面アプリが切り替わっただけの境界は手順にしない。操作対象が操作前画面に明確に見えない場合、「開く」「移動する」などの操作を推測しない。')
        [void]$builder.AppendLine('- 同じChrome/Edge内でページ名が変わる遷移は別アプリ切替ではない。クリック対象または遷移結果が画像で確認できる場合は、必要な操作として残す。')
        [void]$builder.AppendLine('- 参考イベントの対象名が画面と矛盾する場合、その対象名を使わず画面を優先する。')
        [void]$builder.AppendLine('- beforeFrame には操作対象と周辺文脈が見える安定コマを選ぶ。結果画面が手順の理解に必要な場合だけ afterFrame を指定し、後続のスクロールや別操作より前に結果が最初に安定したコマを選ぶ。')
        [void]$builder.AppendLine('- クリック時刻と座標は手がかりであり、画面と矛盾する場合は画面を優先する。')
        [void]$builder.AppendLine('- 参考イベントがある区間は、対応する意味のある手順を原則残す。入力イベントが欠けても、別の入力欄やセルの値が変化していれば独立した入力手順にする。')
        [void]$builder.AppendLine('- 参考イベントも画面変化もない区間だけ steps を空配列にできる。')
        [void]$builder.AppendLine('- 画面に見える名称・セル値・英字の大文字小文字は変えず、そのまま記載する。')
        [void]$builder.AppendLine('- Excelの数式バーが原寸画像に見える場合、計算結果だけに言い換えず「=SUM(B2:B3)」のような数式をそのまま記載する。')
        [void]$builder.AppendLine('- 「前半を入力」「画面を操作」のような曖昧な表現を避け、画像で読めるボタン名・値・結果をタイトルまたは説明へ含める。')
        [void]$builder.AppendLine('- 安定した画面変化が複数ある場合、最後の変化まで確認して途中で打ち切らない。')
        [void]$builder.AppendLine('- 操作を推測で追加しない。不明な手順は confidence=low にする。')
    }
    [void]$builder.AppendLine('')
    $titleSequence = New-Object System.Collections.ArrayList
    foreach ($frame in @($Frames)) {
        $title = (([string]$frame.windowTitle) -replace '\s+', ' ').Trim()
        if ($title.Length -gt 45) { $title = $title.Substring(0, 44) + '…' }
        if ($title -and ($titleSequence.Count -eq 0 -or [string]$titleSequence[$titleSequence.Count - 1] -ne $title)) {
            [void]$titleSequence.Add($title)
        }
    }
    if ($titleSequence.Count -gt 0) {
        [void]$builder.AppendLine(('画面タイトルの遷移: ' + (@($titleSequence) -join ' → ')))
    }
    [void]$builder.AppendLine('選択可能なフレーム:')
    foreach ($frame in @($Frames)) {
        $app = Get-MbRecorderItemAppKey -Item $frame
        if ($app.Length -gt 24) { $app = $app.Substring(0, 23) + '…' }
        $change = if ($frame.PSObject.Properties.Name -contains 'visualChange') { [double]$frame.visualChange } else { 0.0 }
        [void]$builder.AppendLine(('- {0} {1} app={2} change={3:F5}' -f [string]$frame.id,
            (Format-MbRecorderTimeCode -Milliseconds ([int]$frame.timeMs)), $app, $change))
    }
    if (@($InteractionGroups).Count -eq 0) {
        [void]$builder.AppendLine('')
        [void]$builder.AppendLine('参考イベント:')
        foreach ($event in @($Events)) {
            $target = ([string]$event.targetName) -replace '\s+', ' '
            if ($target.Length -gt 100) { $target = $target.Substring(0, 99) + '…' }
            $app = Get-MbRecorderItemAppKey -Item $event
            if ($app.Length -gt 24) { $app = $app.Substring(0, 23) + '…' }
            [void]$builder.AppendLine(('- E{0:d3} {1} kind={2} target={3} app={4}' -f [int]$event.index,
                (Format-MbRecorderTimeCode -Milliseconds ([int]$event.timeMs)), [string]$event.kind, $target,
                $app))
        }
    }
    if (@($InteractionGroups).Count -gt 0) {
        [void]$builder.AppendLine('')
        [void]$builder.AppendLine('この一覧で完成する操作グループ:')
        $groupNumber = 0
        foreach ($group in @($InteractionGroups)) {
            $groupNumber++
            $ids = @($group.eventIds | ForEach-Object { 'E{0:d3}' -f [int]$_ }) -join ','
            $target = ([string]$group.targetName -replace '\s+', ' ').Trim()
            if ($target.Length -gt 70) { $target = $target.Substring(0, 69) + '…' }
            $groupBefore = if ($group.PSObject.Properties.Name -contains 'beforeFrame') { ([string]$group.beforeFrame).Trim() } else { '' }
            $groupAfter = if ($group.PSObject.Properties.Name -contains 'afterFrame') { ([string]$group.afterFrame).Trim() } else { '' }
            $framePair = $groupBefore + '→' + $groupAfter
            [void]$builder.AppendLine(('- G{0:d2} events={1} action={2} target={3} frames={4}' -f $groupNumber, $ids,
                [string]$group.actionKind, $target, $framePair))
        }
    }
    [void]$builder.AppendLine('')
    [void]$builder.AppendLine('回答は次のJSONオブジェクトだけにしてください。')
    [void]$builder.AppendLine('{"steps":[{"beforeFrame":"F00001","afterFrame":"F00002または空文字","eventIds":[1],"targetEventId":1,"title":"短い手順名","description":"利用者が行う操作を1文で記載","confidence":"high|medium|low","reason":"選択理由"}]}')
    [void]$builder.AppendLine('手順は beforeFrame の時刻順に並べてください。targetEventId は赤枠の根拠に使える場合だけ指定し、なければ null にします。')
    [void]$builder.Append($Marker)
    return $builder.ToString()
}

function Test-MbRecorderFrameSetHasMeaningfulChange {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames)
    $ordered = @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    if ($ordered.Count -lt 2) { return $false }
    foreach ($frame in $ordered) {
        if ($frame.PSObject.Properties.Name -contains 'visualChange' -and [double]$frame.visualChange -ge 0.00035) {
            return $true
        }
    }
    $first = $ordered[0]
    $firstApp = Get-MbRecorderItemAppKey -Item $first
    $firstTitle = ([string]$first.windowTitle).Trim()
    foreach ($frame in $ordered | Select-Object -Skip 1) {
        if ((Get-MbRecorderItemAppKey -Item $frame) -eq $firstApp -and
            -not [string]::Equals(([string]$frame.windowTitle).Trim(), $firstTitle, [StringComparison]::Ordinal)) {
            return $true
        }
        if ($first.PSObject.Properties.Name -contains 'imageSha256' -and
            $frame.PSObject.Properties.Name -contains 'imageSha256' -and
            [string]$first.imageSha256 -ne [string]$frame.imageSha256) { return $true }
    }
    return $false
}

function ConvertFrom-MbRecorderCopilotAnswer {
    param(
        [AllowNull()]$Answer,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [AllowEmptyCollection()][object[]]$Events = @()
    )
    $result = New-Object System.Collections.ArrayList
    if ($null -eq $Answer -or $Answer.PSObject.Properties.Name -notcontains 'steps') { return @() }
    $frameMap = @{}
    foreach ($frame in @($Frames)) { $frameMap[[string]$frame.id] = $frame }
    $eventMap = @{}
    foreach ($event in @($Events)) { $eventMap[[int]$event.index] = $event }
    foreach ($step in @($Answer.steps)) {
        if ($null -eq $step) { continue }
        $before = ([string]$step.beforeFrame).Trim().ToUpperInvariant()
        if (-not $frameMap.ContainsKey($before)) { continue }
        $after = ''
        if ($step.PSObject.Properties.Name -contains 'afterFrame') { $after = ([string]$step.afterFrame).Trim().ToUpperInvariant() }
        if ($after -eq $before -or -not $frameMap.ContainsKey($after)) { $after = '' }
        $beforeFrame = $frameMap[$before]
        $afterFrame = $(if ($after) { $frameMap[$after] } else { $null })
        if ($null -ne $afterFrame) {
            $beforeApp = Get-MbRecorderItemAppKey -Item $beforeFrame
            $afterApp = Get-MbRecorderItemAppKey -Item $afterFrame
            # 前後画像が別アプリなら、アンカーの有無に関係なく単なるアプリ切替を
            # 操作として作らない。時刻が逆転した回答も同様に破棄する。
            if ([string]::IsNullOrWhiteSpace($beforeApp) -or $beforeApp -ne $afterApp -or
                [int]$afterFrame.timeMs -le [int]$beforeFrame.timeMs) { continue }
        }
        $eventIds = New-Object System.Collections.ArrayList
        if ($step.PSObject.Properties.Name -contains 'eventIds') {
            foreach ($value in @($step.eventIds)) {
                $id = 0
                try { $id = [int]$value } catch { $id = 0 }
                if ($eventMap.ContainsKey($id) -and -not $eventIds.Contains($id)) { [void]$eventIds.Add($id) }
            }
        }
        $targetEventId = 0
        if ($step.PSObject.Properties.Name -contains 'targetEventId') {
            try { $targetEventId = [int]$step.targetEventId } catch { $targetEventId = 0 }
        }
        $validAnchorIds = @($eventIds | Where-Object {
            $eventMap.ContainsKey([int]$_) -and
            (Test-MbRecorderAnchorContext -Event $eventMap[[int]$_] -BeforeFrame $beforeFrame -AfterFrame $afterFrame)
        })
        if ($targetEventId -gt 0 -and -not $eventIds.Contains($targetEventId)) {
            # 別手順のイベントIDを差し込まれた場合は補完せず、アンカーなしにする。
            $targetEventId = 0
        } elseif (-not $eventIds.Contains($targetEventId) -or $targetEventId -notin $validAnchorIds) {
            # 明示アンカーが不正なとき、複数の候補から勝手に別の枠を選ばない。
            # 同じ操作群に妥当な候補が1つだけなら、入力イベントから直前クリックへ安全に補完する。
            $targetEventId = if ($validAnchorIds.Count -eq 1) { [int]$validAnchorIds[0] } else { 0 }
        }
        $title = ([string]$step.title).Trim()
        if ($title.Length -gt 100) { $title = $title.Substring(0, 100) }
        $description = ([string]$step.description).Trim()
        if ($description.Length -gt 500) { $description = $description.Substring(0, 500) }
        if ([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($description)) { continue }
        $confidence = ([string]$step.confidence).Trim().ToLowerInvariant()
        if ($confidence -notin @('high', 'medium', 'low')) { $confidence = 'low' }
        [void]$result.Add([pscustomobject]@{
            id = 'proposal-' + [guid]::NewGuid().ToString('N')
            beforeFrame = $before
            afterFrame = $after
            eventIds = @($eventIds)
            targetEventId = $targetEventId
            title = $title
            description = $description
            confidence = $confidence
            reason = $(if ($step.PSObject.Properties.Name -contains 'reason') { ([string]$step.reason).Trim() } else { '' })
            timeMs = [int]$beforeFrame.timeMs
            beforeImage = [string]$beforeFrame.image
            afterImage = $(if ($null -ne $afterFrame) { [string]$afterFrame.image } else { '' })
        })
    }
    return @($result | Sort-Object { [int]$_.timeMs })
}

function Add-MbRecorderTitleTransitionProposals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Proposals
    )

    $orderedFrames = @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    $result = New-Object System.Collections.ArrayList
    foreach ($proposal in @($Proposals)) { [void]$result.Add($proposal) }

    # Copilotが具体的な入力は読めても、その直前のページ遷移だけを落とすことがある。
    # 同じアプリ内でタイトルと画面が大きく変わった境界は記録上の事実なので、
    # 既存提案と重ならない場合だけ遷移手順を補完する。
    for ($frameIndex = 1; $frameIndex -lt $orderedFrames.Count; $frameIndex++) {
        $before = $orderedFrames[$frameIndex - 1]
        $after = $orderedFrames[$frameIndex]
        $beforeTitle = ([string]$before.windowTitle).Trim()
        $afterTitle = ([string]$after.windowTitle).Trim()
        $afterChange = if ($after.PSObject.Properties.Name -contains 'visualChange') { [double]$after.visualChange } else { 0.0 }
        if ((Get-MbRecorderItemAppKey -Item $before) -ne (Get-MbRecorderItemAppKey -Item $after) -or
            [string]::Equals($beforeTitle, $afterTitle, [StringComparison]::Ordinal) -or
            [string]::IsNullOrWhiteSpace($afterTitle) -or $afterChange -lt 0.02) { continue }

        $beforeTime = [int]$before.timeMs
        $afterTime = [int]$after.timeMs
        $nearTransition = @($result | Where-Object {
            $candidateProposal = $_
            $proposalBefore = [int]$candidateProposal.timeMs
            $proposalAfter = $proposalBefore
            if ([string]$candidateProposal.afterFrame) {
                $matchedAfter = @($orderedFrames | Where-Object { [string]$_.id -eq [string]$candidateProposal.afterFrame } | Select-Object -First 1)
                if ($matchedAfter.Count -gt 0) { $proposalAfter = [int]$matchedAfter[0].timeMs }
            }
            [Math]::Abs($proposalBefore - $beforeTime) -le 2500 -and $proposalAfter -le ($afterTime + 4000)
        })
        if ($nearTransition.Count -gt 0) { continue }

        # 遷移と後続操作を1件へ合体した提案は、遷移後の最初の安定画面から
        # 始まるよう画像を分けてから、遷移手順を追加する。
        $spanning = @($result | Where-Object {
            $candidateProposal = $_
            $proposalBefore = [int]$candidateProposal.timeMs
            $proposalAfter = $proposalBefore
            $matchedAfter = @($orderedFrames | Where-Object { [string]$_.id -eq [string]$candidateProposal.afterFrame } | Select-Object -First 1)
            if ($matchedAfter.Count -gt 0) { $proposalAfter = [int]$matchedAfter[0].timeMs }
            $proposalBefore -le ($beforeTime + 2500) -and $proposalAfter -ge ($afterTime + 5000)
        })
        foreach ($proposal in $spanning) {
            $stableAfter = @($orderedFrames | Where-Object {
                [int]$_.timeMs -ge ($afterTime + 1500) -and
                (Get-MbRecorderItemAppKey -Item $_) -eq (Get-MbRecorderItemAppKey -Item $after) -and
                ([string]$_.windowTitle).Trim() -eq $afterTitle
            } | Select-Object -First 1)
            if ($stableAfter.Count -gt 0) {
                $proposal.beforeFrame = [string]$stableAfter[0].id
                $proposal.beforeImage = [string]$stableAfter[0].image
                $proposal.timeMs = [int]$stableAfter[0].timeMs
            }
        }

        $title = $afterTitle
        if ($title.Length -gt 80) { $title = $title.Substring(0, 80) }
        [void]$result.Add([pscustomobject]@{
            id = 'proposal-' + [guid]::NewGuid().ToString('N')
            beforeFrame = [string]$before.id
            afterFrame = [string]$after.id
            eventIds = @()
            targetEventId = 0
            title = $title + 'を開く'
            description = $title + 'を開きます。'
            confidence = 'medium'
            reason = '同じブラウザー内で画面タイトルと表示内容が大きく変化したため。'
            timeMs = $beforeTime
            beforeImage = [string]$before.image
            afterImage = [string]$after.image
        })
    }
    return @($result | Sort-Object { [int]$_.timeMs })
}

function Merge-MbRecorderDuplicateTransitionProposals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Proposals
    )

    $frameMap = @{}
    foreach ($frame in @($Frames)) { $frameMap[[string]$frame.id] = $frame }
    $ordered = @($Proposals | Sort-Object { [int]$_.timeMs })
    $removeIds = New-Object 'System.Collections.Generic.HashSet[string]'

    # 一覧を1枚ずつ別チャットで処理すると、前半の末尾で「リンクを選択」し、
    # 後半の先頭で同じリンクによる「ページを開く」をもう一度返すことがある。
    # 後続候補が実際のタイトル遷移と操作後画像を持ち、直前候補が遷移先名を
    # 明示している場合だけ、前後画像のない直前候補を重複として除く。
    for ($index = 0; $index -lt ($ordered.Count - 1); $index++) {
        $current = $ordered[$index]
        $next = $ordered[$index + 1]
        if (-not [string]::IsNullOrWhiteSpace([string]$current.afterFrame)) { continue }
        if (-not $frameMap.ContainsKey([string]$current.beforeFrame) -or
            -not $frameMap.ContainsKey([string]$next.beforeFrame) -or
            -not $frameMap.ContainsKey([string]$next.afterFrame)) { continue }

        $currentBefore = $frameMap[[string]$current.beforeFrame]
        $nextBefore = $frameMap[[string]$next.beforeFrame]
        $nextAfter = $frameMap[[string]$next.afterFrame]
        $nextBeforeTitle = ([string]$nextBefore.windowTitle).Trim()
        $nextAfterTitle = ([string]$nextAfter.windowTitle).Trim()
        if ((Get-MbRecorderItemAppKey -Item $currentBefore) -ne (Get-MbRecorderItemAppKey -Item $nextBefore) -or
            (Get-MbRecorderItemAppKey -Item $nextBefore) -ne (Get-MbRecorderItemAppKey -Item $nextAfter) -or
            [string]::Equals($nextBeforeTitle, $nextAfterTitle, [StringComparison]::Ordinal)) { continue }

        $gap = [int]$nextBefore.timeMs - [int]$currentBefore.timeMs
        if ($gap -lt 0 -or $gap -gt 2500) { continue }
        $destination = ($nextAfterTitle -replace '(?i)\s*[-–—]\s*(Microsoft\s*Edge|Google\s*Chrome)\s*$', '').Trim()
        if ($destination.Length -lt 4) { continue }
        $currentText = (([string]$current.title) + ' ' + ([string]$current.description)).Trim()
        if ($currentText.IndexOf($destination, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            [void]$removeIds.Add([string]$current.id)
        }
    }

    return @($ordered | Where-Object { -not $removeIds.Contains([string]$_.id) })
}

function Expand-MbRecorderExcelRangeSelectionProposals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Events,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Proposals
    )

    $eventMap = @{}
    foreach ($event in @($Events)) { $eventMap[[int]$event.index] = $event }
    $ordered = @($Proposals | Sort-Object { [int]$_.timeMs })
    for ($index = 0; $index -lt ($ordered.Count - 1); $index++) {
        $current = $ordered[$index]
        $next = $ordered[$index + 1]
        $targetId = [int]$current.targetEventId
        if ($targetId -le 0 -or -not $eventMap.ContainsKey($targetId)) { continue }
        $event = $eventMap[$targetId]
        if ((Get-MbRecorderItemAppKey -Item $event) -ne 'excel') { continue }
        $startCell = ([string]$event.targetName).Trim().ToUpperInvariant()
        if ($startCell -notmatch '^[A-Z]{1,3}[1-9][0-9]*$') { continue }

        $currentText = (([string]$current.title) + ' ' + ([string]$current.description)).Trim()
        $nextText = (([string]$next.title) + ' ' + ([string]$next.description)).Trim()
        if ($currentText -notmatch '選択|(?i)select' -or $currentText -match '[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*') { continue }
        if ($nextText -notmatch '選択|(?i)select') { continue }
        $rangeMatch = [regex]::Match($nextText, '(?i)([A-Z]{1,3}[1-9][0-9]*):([A-Z]{1,3}[1-9][0-9]*)')
        if (-not $rangeMatch.Success -or
            -not [string]::Equals($rangeMatch.Groups[1].Value, $startCell, [StringComparison]::OrdinalIgnoreCase)) { continue }

        $range = ($rangeMatch.Groups[1].Value + ':' + $rangeMatch.Groups[2].Value).ToUpperInvariant()
        $current.title = 'セル範囲' + $range + 'を選択する'
        $current.description = 'セル' + $startCell + 'から' + $rangeMatch.Groups[2].Value.ToUpperInvariant() +
            'までドラッグし、セル範囲' + $range + 'を選択する。'
        $current.reason = (([string]$current.reason).TrimEnd('。') +
            '。後続の操作前画像で選択範囲' + $range + 'を確認できるため。').TrimStart('。')
    }
    return @($ordered)
}

function Repair-MbRecorderTransientAfterFrames {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Events,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Proposals
    )

    $frameMap = @{}
    foreach ($frame in @($Frames)) { $frameMap[[string]$frame.id] = $frame }
    $orderedFrames = @($Frames | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    $orderedEvents = @($Events | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    foreach ($proposal in @($Proposals)) {
        $afterId = ([string]$proposal.afterFrame).Trim()
        if (-not $afterId -or -not $frameMap.ContainsKey($afterId)) { continue }
        $after = $frameMap[$afterId]
        if (-not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$after.windowTitle))) { continue }
        if (-not $frameMap.ContainsKey([string]$proposal.beforeFrame)) { continue }
        $before = $frameMap[[string]$proposal.beforeFrame]
        $app = Get-MbRecorderItemAppKey -Item $before
        $lastEventTime = [int]$proposal.timeMs
        foreach ($eventId in @($proposal.eventIds)) {
            $matched = @($orderedEvents | Where-Object { [int]$_.index -eq [int]$eventId } | Select-Object -First 1)
            if ($matched.Count -gt 0) { $lastEventTime = [Math]::Max($lastEventTime, [int]$matched[0].timeMs) }
        }
        $nextEvent = @($orderedEvents | Where-Object {
            [int]$_.timeMs -gt $lastEventTime -and (Get-MbRecorderItemAppKey -Item $_) -eq $app
        } | Select-Object -First 1)
        $limit = if ($nextEvent.Count -gt 0) { [int]$nextEvent[0].timeMs } else { [int]$after.timeMs + 5000 }
        $stable = @($orderedFrames | Where-Object {
            [int]$_.timeMs -gt [int]$after.timeMs -and [int]$_.timeMs -lt $limit -and
            (Get-MbRecorderItemAppKey -Item $_) -eq $app -and
            -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$_.windowTitle))
        } | Select-Object -First 1)
        if ($stable.Count -lt 1) { continue }
        $proposal.afterFrame = [string]$stable[0].id
        $proposal.afterImage = [string]$stable[0].image
        $proposal.reason = (([string]$proposal.reason).TrimEnd('。') +
            '。操作後画像は読み込み完了後の最初の安定画面を使用。').TrimStart('。')
    }
    return @($Proposals | Sort-Object { [int]$_.timeMs })
}

Export-ModuleMember -Function @(
    'Format-MbRecorderTimeCode',
    'Get-MbRecorderWindowAppKey',
    'Get-MbRecorderItemAppKey',
    'Repair-MbRecorderExcelInputEventAnchors',
    'Read-MbRecorderJsonLines',
    'Add-MbRecorderFrameVisualMetrics',
    'Select-MbRecorderTimelineFrames',
    'Select-MbRecorderEventWindowFrames',
    'Select-MbRecorderCopilotSourceFrames',
    'New-MbRecorderLocalFrameCandidates',
    'Select-MbRecorderCandidateFrames',
    'New-MbRecorderContactSheets',
    'New-MbRecorderCopilotPrompt',
    'Test-MbRecorderFrameSetHasMeaningfulChange',
    'ConvertFrom-MbRecorderCopilotAnswer',
    'Add-MbRecorderTitleTransitionProposals',
    'Merge-MbRecorderDuplicateTransitionProposals',
    'Expand-MbRecorderExcelRangeSelectionProposals',
    'Repair-MbRecorderTransientAfterFrames'
)
