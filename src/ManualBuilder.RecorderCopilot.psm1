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
    param([AllowEmptyString()][string]$WindowTitle)
    $title = $WindowTitle.Trim()
    if ($title -match '(?i)Microsoft.?.?Edge$') { return 'edge' }
    if ($title -match '(?i)\s-\sExcel$') { return 'excel' }
    if ($title -match '(?i)\s-\sWord$') { return 'word' }
    if ([string]::IsNullOrWhiteSpace($title)) { return '' }
    return ($title -replace '\s+', ' ').ToLowerInvariant()
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
    $beforeApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$BeforeFrame.windowTitle)
    $eventApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$Event.windowTitle)
    if ([string]::IsNullOrWhiteSpace($beforeApp) -or $beforeApp -ne $eventApp) { return $false }
    $beforeTime = [int]$BeforeFrame.timeMs; $eventTime = [int]$Event.timeMs
    if ([Math]::Abs($eventTime - $beforeTime) -gt 2500) { return $false }
    if ($null -ne $AfterFrame) {
        $afterApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$AfterFrame.windowTitle)
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
        $app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$frame.windowTitle)
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
        $uniform = New-Object 'System.Collections.Generic.HashSet[int]'
        for ($slot = 0; $slot -lt $Maximum; $slot++) {
            $position = [int][Math]::Round(($slot / [double]([Math]::Max(1, $Maximum - 1))) * ($ordered.Count - 1))
            [void]$uniform.Add($position)
        }
        return @($uniform | Sort-Object | ForEach-Object { $ordered[[int]$_] })
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
                (Get-MbRecorderWindowAppKey -WindowTitle ([string]$ordered[$frameIndex - 1].windowTitle)) -eq
                    (Get-MbRecorderWindowAppKey -WindowTitle ([string]$frame.windowTitle))) {
                & $addCandidate ($frameIndex - 1) 84
            }
        }
    }

    for ($eventIndex = 0; $eventIndex -lt $orderedEvents.Count; $eventIndex++) {
        $event = $orderedEvents[$eventIndex]
        $eventTime = [int]$event.timeMs
        $eventApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$event.windowTitle)
        $beforeIndex = -1
        for ($frameIndex = 0; $frameIndex -lt $ordered.Count; $frameIndex++) {
            if ([int]$ordered[$frameIndex].timeMs -gt $eventTime) { break }
            $frameApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$ordered[$frameIndex].windowTitle)
            if (-not $eventApp -or $frameApp -eq $eventApp) { $beforeIndex = $frameIndex }
        }
        if ($beforeIndex -ge 0) { & $addCandidate $beforeIndex 100 }

        # 新しいアプリへ切り替わった直後の先頭2コマは、描画途中やフォーカス移動に
        # なりやすい。3コマ目を操作前の文脈として残す（存在しなければ直前コマ）。
        if ($eventApp -and $seenAppSegments.Add($eventApp)) {
            $firstAppIndex = -1
            for ($frameIndex = 0; $frameIndex -lt $ordered.Count; $frameIndex++) {
                if ((Get-MbRecorderWindowAppKey -WindowTitle ([string]$ordered[$frameIndex].windowTitle)) -eq $eventApp) {
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
        $nextApp = if ($null -ne $nextEvent) { Get-MbRecorderWindowAppKey -WindowTitle ([string]$nextEvent.windowTitle) } else { '' }
        $resultIndex = -1
        if ($null -ne $nextEvent -and $nextApp -eq $eventApp -and [int]$nextEvent.timeMs -gt $eventTime) {
            # 次の操作直前は、クリック結果が描画し終わった最も安定したコマ。
            for ($frameIndex = $beforeIndex + 1; $frameIndex -lt $ordered.Count; $frameIndex++) {
                $frameTime = [int]$ordered[$frameIndex].timeMs
                if ($frameTime -ge [int]$nextEvent.timeMs) { break }
                if ((Get-MbRecorderWindowAppKey -WindowTitle ([string]$ordered[$frameIndex].windowTitle)) -eq $eventApp) {
                    $resultIndex = $frameIndex
                }
            }
        } else {
            # 次が別アプリ、または最後の操作なら、現在アプリ内の最初の操作後コマを残す。
            for ($frameIndex = $beforeIndex + 1; $frameIndex -lt $ordered.Count; $frameIndex++) {
                $frameTime = [int]$ordered[$frameIndex].timeMs
                if ($frameTime -gt $eventTime + 2000) { break }
                if ((Get-MbRecorderWindowAppKey -WindowTitle ([string]$ordered[$frameIndex].windowTitle)) -eq $eventApp) {
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
function New-MbRecorderLocalFrameCandidates {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Frames,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Events,
        [ValidateRange(8, 300)][int]$MaximumFrames = 30,
        [ValidateRange(250, 10000)][int]$EditMergeGapMs = 5000
    )
    $orderedEvents = @($Events | Sort-Object { [int]$_.timeMs }, { [int]$_.index })
    if ($orderedEvents.Count -lt 1 -or @($Frames).Count -lt 1) { return @() }

    # 入力イベントは確定時に記録されるため、クリックを取り逃した最初の入力では
    # 操作前画像がイベントより2秒以上前になることがある。
    $scoped = @(Select-MbRecorderEventWindowFrames -Frames @($Frames) -Events $orderedEvents -BeforePaddingMs 3000)
    $selected = @(Select-MbRecorderTimelineFrames -Frames $scoped -Events $orderedEvents -Maximum $MaximumFrames)
    if ($selected.Count -lt 1) { return @() }

    # 編集可能な場所へのクリックと、その直後の入力は利用者から見れば1手順。
    # 対象名が取れないEdgeでも、同じアプリ内で他のクリックを挟まない場合だけ結合する。
    $groups = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $orderedEvents.Count; $i++) {
        $current = $orderedEvents[$i]
        $items = New-Object System.Collections.ArrayList
        [void]$items.Add($current)
        if ([string]$current.kind -in @('click', 'right-click') -and $i + 1 -lt $orderedEvents.Count) {
            $next = $orderedEvents[$i + 1]
            $sameApp = (Get-MbRecorderWindowAppKey -WindowTitle ([string]$current.windowTitle)) -eq
                (Get-MbRecorderWindowAppKey -WindowTitle ([string]$next.windowTitle))
            $gap = [int]$next.timeMs - [int]$current.timeMs
            $currentType = if ($current.PSObject.Properties.Name -contains 'targetType') { [string]$current.targetType } else { '' }
            $currentName = if ($current.PSObject.Properties.Name -contains 'targetName') { [string]$current.targetName } else { '' }
            $nextName = if ($next.PSObject.Properties.Name -contains 'targetName') { [string]$next.targetName } else { '' }
            $editable = $currentType -in @('ControlType.Edit', 'ControlType.DataItem', 'ControlType.ClickPoint') -or
                [string]::IsNullOrWhiteSpace($currentType)
            $sameTarget = [string]::IsNullOrWhiteSpace($currentName) -or [string]::IsNullOrWhiteSpace($nextName) -or
                [string]::Equals($currentName, $nextName, [StringComparison]::OrdinalIgnoreCase)
            if ([string]$next.kind -eq 'input' -and $sameApp -and $gap -ge 0 -and
                $gap -le $EditMergeGapMs -and $editable -and $sameTarget) {
                [void]$items.Add($next)
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
        $app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$firstEvent.windowTitle)
        # ローカル提案ではAI向けに絞った一覧だけでなく、イベント周辺の原本も使う。
        # クリックを取り逃した場合でも、入力前の空欄と確定後の完成状態を復元できる。
        $appFrames = @($scoped | Where-Object {
            (Get-MbRecorderWindowAppKey -WindowTitle ([string]$_.windowTitle)) -eq $app
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
            $nextApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$nextGroup[0].windowTitle)
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
            if ($null -ne $nextEventTime) { $completionEnd = [Math]::Min($completionEnd, [int]$nextEventTime + 1800) }
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
                $nextClickEvidence = @($appFrames | Where-Object {
                    $_.PSObject.Properties.Name -contains 'evidenceEventId' -and
                        [int]$_.evidenceEventId -gt 0 -and [int]$_.timeMs -ge ([int]$nextEventTime - 250) -and
                        [int]$_.timeMs -le ([int]$nextEventTime + 250)
                } | Select-Object -First 1)
                if ($nextClickEvidence.Count -gt 0) {
                    $after = $nextClickEvidence
                } else {
                    $after = @($appFrames | Where-Object {
                        # 入力イベントは無操作待ちの後に記録されるため、結果画面が
                        # event.timeMs より前にある。グループ先頭の操作以後を対象にする。
                        [int]$_.timeMs -gt [int]$firstEvent.timeMs -and [int]$_.timeMs -lt [int]$nextEventTime
                    } | Select-Object -Last 1)
                }
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
            @($group | Where-Object { [string]$_.kind -in @('click', 'right-click') })
        } else { @($group) }
        foreach ($event in $anchorEvents) {
            if (Test-MbRecorderEventHasVisualAnchor -Event $event) { $targetEventId = [int]$event.index; break }
        }
        [void]$result.Add([pscustomobject]@{
            id = 'local-' + [guid]::NewGuid().ToString('N')
            beforeFrame = [string]$before.id
            afterFrame = $(if ($null -ne $after -and [string]$after.id -ne [string]$before.id) { [string]$after.id } else { '' })
            eventIds = $eventIds
            targetEventId = $targetEventId
            actionKind = $(if ($hasInput) { 'input' } else { [string]$firstEvent.kind })
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
            app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$group[0].windowTitle)
            startMs = if ($group.Count -gt 1) { [int]$group[0].timeMs - 250 } else { [int]$group[0].timeMs - 2200 }
            # 入力確定直後の計算結果やセル移動は同じ入力手順の操作後であり、
            # 独立した画面変化手順にはしない。
            endMs = [int]$group[$group.Count - 1].timeMs + 1800
        })
    }
    $firstCandidateStartMs = [int]::MaxValue
    foreach ($candidate in $result) {
        $candidateBefore = @($scoped | Where-Object { [string]$_.id -eq [string]$candidate.beforeFrame } | Select-Object -First 1)
        if ($candidateBefore.Count -gt 0) { $firstCandidateStartMs = [Math]::Min($firstCandidateStartMs, [int]$candidateBefore[0].timeMs) }
    }
    foreach ($frame in $selected) {
        $change = if ($frame.PSObject.Properties.Name -contains 'visualChange') { [double]$frame.visualChange } else { 0.0 }
        if ($change -lt 0.00035 -or $existingAfter.Contains([string]$frame.id) -or
            (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$frame.windowTitle))) { continue }
        $frameApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$frame.windowTitle)
        $frameTime = [int]$frame.timeMs
        if ($frameTime -le $firstCandidateStartMs) { continue }

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
                (Get-MbRecorderWindowAppKey -WindowTitle ([string]$_.windowTitle)) -eq $frameApp -and
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
                (Get-MbRecorderWindowAppKey -WindowTitle ([string]$_.windowTitle)) -eq $frameApp -and
                [Math]::Abs([int]$_.timeMs - $frameTime) -le 900
        }).Count -gt 0
        if ($nearClick) { continue }

        $before = @($selected | Where-Object {
            [int]$_.timeMs -lt $frameTime -and
                (Get-MbRecorderWindowAppKey -WindowTitle ([string]$_.windowTitle)) -eq $frameApp -and
                -not (Test-MbRecorderTransientFrameTitle -WindowTitle ([string]$_.windowTitle))
        } | Sort-Object { [int]$_.timeMs }, { [int]$_.index } | Select-Object -Last 1)
        if ($before.Count -lt 1 -or [string]$before[0].id -eq [string]$frame.id) { continue }
        $anchor = @($orderedEvents | Where-Object {
            (Get-MbRecorderWindowAppKey -WindowTitle ([string]$_.windowTitle)) -eq $frameApp -and
                [int]$_.timeMs -le ($frameTime + 1200) -and [int]$_.timeMs -ge ($frameTime - 2200) -and
                (Test-MbRecorderEventHasVisualAnchor -Event $_)
        } | Sort-Object { [Math]::Abs([int]$_.timeMs - $frameTime) } | Select-Object -First 1)
        $anchorId = if ($anchor.Count -gt 0) { [int]$anchor[0].index } else { 0 }
        [void]$result.Add([pscustomobject]@{
            id = 'local-' + [guid]::NewGuid().ToString('N')
            beforeFrame = [string]$before[0].id
            afterFrame = [string]$frame.id
            eventIds = $(if ($anchorId -gt 0) { @($anchorId) } else { @() })
            targetEventId = $anchorId
            actionKind = 'visual-change'
            timeMs = [int]$before[0].timeMs
            beforeImage = [string]$before[0].image
            afterImage = [string]$frame.image
            source = 'local'
        })
        [void]$existingAfter.Add([string]$frame.id)
    }
    return @($result | Sort-Object { [int]$_.timeMs })
}

# Copilotへは候補生成前の最大30コマをそのまま渡さず、ローカルで組み立てた
# 各手順の操作前／操作後だけを時刻順に渡す。同じ境界画像は1枚へまとめるため、
# 短い記録で入力途中や同一結果のコマが一覧を占有しない。
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
        $app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$frame.windowTitle)
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
        $app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$event.windowTitle)
        if (-not [string]::IsNullOrWhiteSpace($app)) { [void]$allowedApps.Add($app) }
    }
    $first = [int](($times | Measure-Object -Minimum).Minimum) - $BeforePaddingMs
    $last = [int](($times | Measure-Object -Maximum).Maximum) + $AfterPaddingMs
    $scoped = @($Frames | Where-Object {
        if ([int]$_.timeMs -lt $first -or [int]$_.timeMs -gt $last) { return $false }
        if ($allowedApps.Count -lt 1) { return $true }
        $frameApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$_.windowTitle)
        return $allowedApps.Contains($frameApp)
    })
    if ($scoped.Count -lt 2) { return @($Frames) }
    return $scoped
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
        [ValidateRange(8, 24)][int]$FramesPerSheet = 15,
        [ValidateRange(2, 6)][int]$Columns = 3
    )
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }
    # 4列ではExcelのセル値やブラウザーの小さなラベルをAIが読み落とした。
    # 横1920pxは維持し、既定3列で1コマを640x360まで大きくする。
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
        [string]$Marker = 'MB_END'
    )
    $builder = New-Object Text.StringBuilder
    [void]$builder.AppendLine('あなたは操作マニュアルの編集者です。添付は記録画面を時系列に並べたコンタクトシートです。')
    [void]$builder.AppendLine('プログラムの候補分割は信用せず、前後の流れから「利用者の目的を達成するために必要な手順」を作ってください。')
    [void]$builder.AppendLine('規則:')
    [void]$builder.AppendLine('- 同じ対象への連続クリックと入力は、1つの意味のある手順にまとめる。')
    [void]$builder.AppendLine('- B2、B3、B4のように対象が異なる入力イベントは省略せず、対象ごとに別の手順にする。')
    [void]$builder.AppendLine('- 読込中、暗転、スピナー、中間的なアニメーションのコマは代表画像に選ばない。')
    [void]$builder.AppendLine('- ManualBuilderへ戻る操作、記録開始・停止、無意味なフォーカス移動は除外する。')
    [void]$builder.AppendLine('- EdgeからExcelなど、前面アプリが切り替わっただけの境界は手順にしない。操作対象が操作前画面に明確に見えない場合、「開く」「移動する」などの操作を推測しない。')
    [void]$builder.AppendLine('- 参考イベントの対象名が画面と矛盾する場合、その対象名を使わず画面を優先する。')
    [void]$builder.AppendLine('- beforeFrame には操作対象と周辺文脈が見える安定コマを選ぶ。結果画面が手順の理解に必要な場合だけ afterFrame を指定する。')
    [void]$builder.AppendLine('- クリック時刻と座標は手がかりであり、画面と矛盾する場合は画面を優先する。')
    [void]$builder.AppendLine('- 参考イベントがある区間は、対応する意味のある手順を原則残す。入力イベントが欠けても、別の入力欄やセルの値が変化していれば独立した入力手順にする。')
    [void]$builder.AppendLine('- 参考イベントも画面変化もない区間だけ steps を空配列にできる。')
    [void]$builder.AppendLine('- 画面に見える名称・セル値・英字の大文字小文字は変えず、そのまま記載する。')
    [void]$builder.AppendLine('- 操作を推測で追加しない。不明な手順は confidence=low にする。')
    [void]$builder.AppendLine('')
    [void]$builder.AppendLine('選択可能なフレーム:')
    foreach ($frame in @($Frames)) {
        $app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$frame.windowTitle)
        if ($app.Length -gt 24) { $app = $app.Substring(0, 23) + '…' }
        [void]$builder.AppendLine(('- {0} {1} app={2}' -f [string]$frame.id,
            (Format-MbRecorderTimeCode -Milliseconds ([int]$frame.timeMs)), $app))
    }
    [void]$builder.AppendLine('')
    [void]$builder.AppendLine('参考イベント:')
    foreach ($event in @($Events)) {
        $target = ([string]$event.targetName) -replace '\s+', ' '
        if ($target.Length -gt 100) { $target = $target.Substring(0, 99) + '…' }
        $app = Get-MbRecorderWindowAppKey -WindowTitle ([string]$event.windowTitle)
        if ($app.Length -gt 24) { $app = $app.Substring(0, 23) + '…' }
        [void]$builder.AppendLine(('- E{0:d3} {1} kind={2} target={3} app={4}' -f [int]$event.index,
            (Format-MbRecorderTimeCode -Milliseconds ([int]$event.timeMs)), [string]$event.kind, $target,
            $app))
    }
    [void]$builder.AppendLine('')
    [void]$builder.AppendLine('回答は次のJSONオブジェクトだけにしてください。')
    [void]$builder.AppendLine('{"steps":[{"beforeFrame":"F00001","afterFrame":"F00002または空文字","eventIds":[1],"targetEventId":1,"title":"短い手順名","description":"利用者が行う操作を1文で記載","confidence":"high|medium|low","reason":"選択理由"}]}')
    [void]$builder.AppendLine('手順は beforeFrame の時刻順に並べてください。targetEventId は赤枠の根拠に使える場合だけ指定し、なければ null にします。')
    [void]$builder.Append($Marker)
    return $builder.ToString()
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
            $beforeApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$beforeFrame.windowTitle)
            $afterApp = Get-MbRecorderWindowAppKey -WindowTitle ([string]$afterFrame.windowTitle)
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

Export-ModuleMember -Function @(
    'Format-MbRecorderTimeCode',
    'Get-MbRecorderWindowAppKey',
    'Read-MbRecorderJsonLines',
    'Add-MbRecorderFrameVisualMetrics',
    'Select-MbRecorderTimelineFrames',
    'Select-MbRecorderEventWindowFrames',
    'New-MbRecorderLocalFrameCandidates',
    'Select-MbRecorderCandidateFrames',
    'New-MbRecorderContactSheets',
    'New-MbRecorderCopilotPrompt',
    'ConvertFrom-MbRecorderCopilotAnswer'
)
