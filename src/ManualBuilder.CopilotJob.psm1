# 手順の下書きをCopilotへ依頼するジョブ。
#
# 依頼文の考え方:
#   画面を見せて「説明を書いて」と頼んでも、当たり障りのない文しか返らない。
#   ManualBuilderは録画から「どの画面で」「どこが操作されたか」を機械的に割り出し、
#   OCRでその場所の文字まで読んである。つまりCopilotへ渡すのは推測の材料ではなく、
#   確定した事実である。Copilotの仕事は事実を日本語の手順文へ整えることだけになる。
#
#   加えて、すでに人が書いた手順を見本として渡す。文体と粒度の基準は人が決め、
#   Copilotはそれに合わせる。ここを渡さないと部内の書き方から外れた文が返る。

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Excel.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1')

$script:MbDraftTitleMaxLength = 40
$script:MbDraftDescriptionMaxLength = 400
$script:MbDraftNoteMaxLength = 300

function Format-MbTimeCode {
    param([int]$Milliseconds)
    if ($Milliseconds -le 0) { return '' }
    $total = [int][Math]::Floor($Milliseconds / 1000)
    return ('{0}:{1:d2}' -f [int][Math]::Floor($total / 60), ($total % 60))
}

# 項目が欠けている手順でも落ちないようにする。
# 数値として使う項目に空文字を返すと [int] 変換で例外になるため、既定値を受け取る。
function Get-MbStepCaptureValue {
    param(
        [Parameter(Mandatory = $true)]$Step,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = ''
    )
    if ($Step.PSObject.Properties.Name -notcontains 'capture') { return $Default }
    $capture = $Step.capture
    if ($null -eq $capture) { return $Default }
    if ($capture.PSObject.Properties.Name -notcontains $Name) { return $Default }
    $value = $capture.$Name
    if ($null -eq $value) { return $Default }
    return $value
}

# プロジェクト全体の手順を、シートの文脈つきの平らな一覧にする。
function Get-MbCopilotStepList {
    param([Parameter(Mandatory = $true)]$Project)

    $list = New-Object System.Collections.ArrayList
    $order = 0
    foreach ($sheet in @($Project.sheets)) {
        $indexInSheet = 0
        foreach ($step in @($sheet.steps)) {
            $order++
            $indexInSheet++
            $imageEntry = @($Project.images | Where-Object { [string]$_.id -eq [string]$step.imageId }) | Select-Object -First 1
            [void]$list.Add([pscustomobject]@{
                id           = [string]$step.id
                sheetId      = [string]$sheet.id
                sheetName    = [string]$sheet.name
                order        = $order
                indexInSheet = $indexInSheet
                title        = [string]$step.title
                description  = [string]$step.description
                note         = [string]$step.note
                imageId      = [string]$step.imageId
                imageWidth   = $(if ($null -ne $imageEntry) { [int]$imageEntry.width } else { 0 })
                imageHeight  = $(if ($null -ne $imageEntry) { [int]$imageEntry.height } else { 0 })
                annotations  = @($step.annotations)
                crop         = $step.crop
                clickLabel   = [string](Get-MbStepCaptureValue -Step $step -Name 'clickLabel')
                windowTitle  = [string](Get-MbStepCaptureValue -Step $step -Name 'windowTitle')
                screenText   = [string](Get-MbStepCaptureValue -Step $step -Name 'screenText')
                narration    = [string](Get-MbStepCaptureValue -Step $step -Name 'narration')
                videoTimeMs  = [int](Get-MbStepCaptureValue -Step $step -Name 'videoTimeMs' -Default 0)
                targetType       = [string](Get-MbStepCaptureValue -Step $step -Name 'targetType')
                targetSource     = [string](Get-MbStepCaptureValue -Step $step -Name 'targetSource')
                targetConfidence = [string](Get-MbStepCaptureValue -Step $step -Name 'targetConfidence')
                targetCandidateId = [string](Get-MbStepCaptureValue -Step $step -Name 'targetCandidateId')
                targetCandidates = @((Get-MbStepCaptureValue -Step $step -Name 'targetCandidates' -Default @()))
            })
        }
    }
    return @($list)
}

# 文章が書かれている手順を見本として選ぶ。文体と粒度の基準になる。
function Get-MbCopilotStyleSamples {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Steps, [int]$Maximum = 3)

    $written = @($Steps | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.title) -and -not [string]::IsNullOrWhiteSpace($_.description)
    })
    if ($written.Count -eq 0) { return @() }
    # 説明が長すぎず短すぎないものを見本にする。極端な例を基準にしない。
    $sorted = @($written | Sort-Object @{ Expression = { [Math]::Abs($_.description.Length - 60) } })
    return @($sorted | Select-Object -First $Maximum)
}

# 下書きが必要な手順だけを、指定の数ずつのまとまりに分ける。
function Get-MbCopilotPackets {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Steps,
        [int]$StepsPerPacket = 6,
        [switch]$IncludeWritten,
        [ValidateSet('draft', 'review')][string]$Mode = 'draft'
    )

    if ($StepsPerPacket -lt 1) { $StepsPerPacket = 1 }
    if ($Mode -eq 'review') {
        # 校正は書かれている文章が対象。画像は見ないので、文字だけの手順も含める。
        $targets = @($Steps | Where-Object {
            (-not [string]::IsNullOrWhiteSpace($_.title)) -or
            (-not [string]::IsNullOrWhiteSpace($_.description)) -or
            (-not [string]::IsNullOrWhiteSpace($_.note))
        })
    } else {
        $targets = @($Steps | Where-Object {
            $_.imageId -and ($IncludeWritten -or [string]::IsNullOrWhiteSpace($_.title) -or [string]::IsNullOrWhiteSpace($_.description))
        })
    }
    $packets = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $targets.Count; $i += $StepsPerPacket) {
        $count = [Math]::Min($StepsPerPacket, $targets.Count - $i)
        [void]$packets.Add(@($targets[$i..($i + $count - 1)]))
    }
    # パケットが1件だけでも、その中の手順をPowerShellのパイプラインで平坦化しない。
    # 呼び出し側は常に「パケットの配列」として件数と添付単位を扱う。
    Write-Output -NoEnumerate ($packets.ToArray())
    return
}

function New-MbCopilotStepPrompt {
    param(
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)][object[]]$PacketSteps,
        [Parameter(Mandatory = $true)][hashtable]$AttachmentNames,
        [object[]]$StyleSamples = @(),
        [int]$TotalSteps = 0,
        [string]$Marker = 'MB_END'
    )

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('あなたは部内向け操作マニュアルの編集者です。操作を録画した画面から、手順の文章を書いてください。')
    [void]$builder.AppendLine()
    [void]$builder.AppendLine('## このマニュアルについて')
    [void]$builder.AppendLine(('マニュアル名: ' + [string]$Project.title))
    if ($TotalSteps -gt 0) { [void]$builder.AppendLine(('全体の手順数: ' + $TotalSteps)) }
    [void]$builder.AppendLine()

    [void]$builder.AppendLine('## 今回書いてほしい手順')
    [void]$builder.AppendLine('添付画像の番号付き枠は、DOM、UI Automation、OCR、または動画差分から得た操作対象の候補です。いずれも誤る可能性があり、確定した事実ではありません。')
    [void]$builder.AppendLine('候補の意味が画面と操作内容に一致するかを確認し、列挙された候補IDか none を選んでください。座標は生成しないでください。')
    [void]$builder.AppendLine()
    foreach ($step in $PacketSteps) {
        $name = [string]$AttachmentNames[[string]$step.id]
        [void]$builder.AppendLine(('### ' + [string]$step.id))
        [void]$builder.AppendLine(('添付画像: ' + $name))
        [void]$builder.AppendLine(('シート: ' + [string]$step.sheetName + '（このシートの ' + [string]$step.indexInSheet + ' 番目）'))
        $timeCode = Format-MbTimeCode -Milliseconds ([int]$step.videoTimeMs)
        if (-not [string]::IsNullOrWhiteSpace($timeCode)) {
            [void]$builder.AppendLine(('録画内の時刻: ' + $timeCode))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$step.windowTitle)) {
            [void]$builder.AppendLine(('操作していた画面: ' + [string]$step.windowTitle))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$step.clickLabel)) {
            [void]$builder.AppendLine(('アプリが暫定選択した操作対象: ' + [string]$step.clickLabel))
        }
        $visualCandidates = @($step.targetCandidates)
        if ($visualCandidates.Count -gt 0) {
            [void]$builder.AppendLine('操作対象候補（添付画像の番号と同じ順）:')
            for ($candidateIndex = 0; $candidateIndex -lt $visualCandidates.Count; $candidateIndex++) {
                $candidate = $visualCandidates[$candidateIndex]
                $candidateLabel = if ($candidate.PSObject.Properties.Name -contains 'label') { [string]$candidate.label } else { '' }
                $candidateSource = if ($candidate.PSObject.Properties.Name -contains 'source') { [string]$candidate.source } else { '' }
                $candidateConfidence = if ($candidate.PSObject.Properties.Name -contains 'confidence') { [string]$candidate.confidence } else { '' }
                [void]$builder.AppendLine(('- {0}: id={1}, 取得元={2}, ローカル信頼度={3}, 名前={4}' -f
                    ($candidateIndex + 1), [string]$candidate.id, $candidateSource, $candidateConfidence, $candidateLabel))
            }
            [void]$builder.AppendLine('- none: 操作対象を示す枠が不要な結果画面、または正しい候補がない')
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$step.narration)) {
            [void]$builder.AppendLine(('操作しながら話した内容: ' + [string]$step.narration))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$step.screenText)) {
            $screenText = ([string]$step.screenText) -replace '[\r\n]+', ' / '
            if ($screenText.Length -gt 600) { $screenText = $screenText.Substring(0, 600) }
            [void]$builder.AppendLine(('画面に出ていた文字: ' + $screenText))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$step.title) -or -not [string]::IsNullOrWhiteSpace([string]$step.description)) {
            [void]$builder.AppendLine(('すでに人が書いた内容: 手順名「' + [string]$step.title + '」説明「' + [string]$step.description + '」'))
        }
        [void]$builder.AppendLine()
    }

    if (@($StyleSamples).Count -gt 0) {
        [void]$builder.AppendLine('## 文体と粒度の見本（このマニュアルで人が書いた手順）')
        [void]$builder.AppendLine('この書き方に合わせてください。')
        foreach ($sample in $StyleSamples) {
            [void]$builder.AppendLine(('- 手順名「' + [string]$sample.title + '」／説明「' + [string]$sample.description + '」'))
        }
        [void]$builder.AppendLine()
    }

    [void]$builder.AppendLine('## 書き方の決まり')
    [void]$builder.AppendLine('- 手順名は体言止めで20文字以内。')
    [void]$builder.AppendLine('- 説明は「〜します。」の敬体。操作の文と、その結果どうなるかの文で、2文までにする。')
    [void]$builder.AppendLine('- 補足は、間違えやすい点や前提がある場合だけ書く。無ければ空文字にする。')
    [void]$builder.AppendLine('- 番号付き候補を盲信しない。画像上の意味と一致する候補だけを選ぶ。正しい候補がなければ targetCandidateId を none、visualConfident を false にする。')
    [void]$builder.AppendLine('- 結果確認の画面やスクロール場面では、無理に操作対象を選ばず none を使う。')
    [void]$builder.AppendLine('- 操作対象が小さく周辺文脈を残した拡大が有効なら zoom=focus、画面全体の確認なら zoom=full、判断できなければ zoom=keep にする。')
    [void]$builder.AppendLine('- 画像と与えられた情報から読み取れないことは書かない。想像で補わない。')
    [void]$builder.AppendLine('- 判断できない手順は confident を false にし、reason に理由を短く書く。')
    [void]$builder.AppendLine('- 直前の手順と同じ画面である、操作されていない、といった理由で手順として不要な場合は keep を false にする。')
    [void]$builder.AppendLine('- すでに人が書いた内容がある手順は、それを尊重して整えるだけにする。')
    [void]$builder.AppendLine()

    [void]$builder.AppendLine('## 出力の形')
    [void]$builder.AppendLine('次の形のJSONだけを出力してください。説明文や前置きは書かないでください。')
    [void]$builder.AppendLine('{"steps":[{"id":"手順のid","keep":true,"targetCandidateId":"列挙された候補IDまたはnone","zoom":"keepまたはfocusまたはfull","visualConfident":true,"visualReason":"候補を選んだ根拠","title":"手順名","description":"説明","note":"補足","confident":true,"reason":""}]}')
    [void]$builder.AppendLine(('id は上に並べた ' + (@($PacketSteps | ForEach-Object { [string]$_.id }) -join ', ') + ' をそのまま使ってください。'))
    [void]$builder.AppendLine(('JSONを出力し終えたら、最後の行に ' + $Marker + ' とだけ書いてください。'))
    # この一文が回答の始まりを見つける目印になる。必ず依頼文の最後に置く。
    [void]$builder.Append((Get-MbCopilotPromptTailAnchor))
    return $builder.ToString()
}

# 文章を整えてもらうための依頼文。
#
# 下書きと違い、画像は渡さない。表記ゆれや用語の不統一は、文章を一度にまとめて
# 見ないと分からないため、1回で渡す手順の数を多くする。画像が無いぶん軽い。
function New-MbCopilotReviewPrompt {
    param(
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$PacketSteps,
        [int]$TotalSteps = 0,
        [string]$Marker = 'MB_END'
    )

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('あなたは部内向け操作マニュアルの校正者です。次の手順の文章を読み、直したほうがよい箇所だけを挙げてください。')
    [void]$builder.AppendLine()
    [void]$builder.AppendLine('## このマニュアルについて')
    [void]$builder.AppendLine(('マニュアル名: ' + [string]$Project.title))
    if ($TotalSteps -gt 0) { [void]$builder.AppendLine(('全体の手順数: ' + $TotalSteps)) }
    [void]$builder.AppendLine()

    [void]$builder.AppendLine('## 手順の文章')
    foreach ($step in $PacketSteps) {
        [void]$builder.AppendLine(('### ' + [string]$step.id))
        [void]$builder.AppendLine(('シート: ' + [string]$step.sheetName + '（このシートの ' + [string]$step.indexInSheet + ' 番目）'))
        [void]$builder.AppendLine(('手順名: ' + [string]$step.title))
        [void]$builder.AppendLine(('説明: ' + [string]$step.description))
        [void]$builder.AppendLine(('補足: ' + [string]$step.note))
        [void]$builder.AppendLine()
    }

    [void]$builder.AppendLine('## 見るところ')
    [void]$builder.AppendLine('- 敬体の統一。「〜します。」に揃える。')
    [void]$builder.AppendLine('- 表記ゆれ。同じものが別の書き方になっていないか（送り仮名、全角と半角、カタカナの長音）。')
    [void]$builder.AppendLine('- 用語の不統一。同じ画面や操作を、手順によって別の名前で呼んでいないか。')
    [void]$builder.AppendLine('- 誤字脱字。')
    [void]$builder.AppendLine('- 一文が長すぎて読みにくい箇所。')
    [void]$builder.AppendLine()

    [void]$builder.AppendLine('## 決まり')
    [void]$builder.AppendLine('- 直す必要がない手順は挙げないでください。')
    [void]$builder.AppendLine('- 意味を変えないでください。書き方だけを整えます。')
    [void]$builder.AppendLine('- 手順の順序や、操作そのものの是非は指摘しないでください。校正の範囲を超えます。')
    [void]$builder.AppendLine('- 直す項目だけを書き、直さない項目は空文字にしてください。')
    [void]$builder.AppendLine('- 表記ゆれや用語の不統一を直すときは、このマニュアルの中で多いほうへ揃えてください。')
    [void]$builder.AppendLine()

    [void]$builder.AppendLine('## 出力の形')
    [void]$builder.AppendLine('次の形のJSONだけを出力してください。説明文や前置きは書かないでください。')
    [void]$builder.AppendLine('{"steps":[{"id":"手順のid","title":"直した手順名","description":"直した説明","note":"直した補足","kind":"敬体","reason":"直した理由を一行で"}]}')
    [void]$builder.AppendLine('kind は 敬体 / 表記ゆれ / 用語 / 誤字 / 長文 のいずれかにしてください。')
    [void]$builder.AppendLine(('JSONを出力し終えたら、最後の行に ' + $Marker + ' とだけ書いてください。'))
    [void]$builder.Append((Get-MbCopilotPromptTailAnchor))
    return $builder.ToString()
}

function Get-MbTrimmedText {
    param([AllowNull()]$Value, [int]$MaxLength)
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value).Trim()
    if ($MaxLength -gt 0 -and $text.Length -gt $MaxLength) { $text = $text.Substring(0, $MaxLength) }
    return $text
}

function Get-MbBooleanOrDefault {
    param($Container, [string]$Name, [bool]$Default)
    if ($null -eq $Container) { return $Default }
    if ($Container.PSObject.Properties.Name -notcontains $Name) { return $Default }
    $value = $Container.$Name
    if ($value -is [bool]) { return $value }
    return [bool]([string]$value -match '^(?i:true|1|yes)$')
}

# Copilotの回答を、手順idごとの下書きへ整える。
# 依頼していないidや、形の合わない項目は捨てる。
function ConvertFrom-MbCopilotStepAnswer {
    param(
        [AllowNull()]$Answer,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$PacketSteps,
        [ValidateSet('draft', 'review')][string]$Mode = 'draft'
    )

    $known = @{}
    foreach ($step in $PacketSteps) { $known[[string]$step.id] = $step }
    $drafts = New-Object System.Collections.ArrayList
    if ($null -eq $Answer -or $Answer.PSObject.Properties.Name -notcontains 'steps') { return @($drafts) }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($item in @($Answer.steps)) {
        if ($null -eq $item) { continue }
        if ($item.PSObject.Properties.Name -notcontains 'id') { continue }
        $id = [string]$item.id
        if (-not $known.ContainsKey($id)) { continue }
        if (-not $seen.Add($id)) { continue }

        $source = $known[$id]
        $title = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'title') { $item.title } else { '' }) -MaxLength $script:MbDraftTitleMaxLength
        $description = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'description') { $item.description } else { '' }) -MaxLength $script:MbDraftDescriptionMaxLength
        $note = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'note') { $item.note } else { '' }) -MaxLength $script:MbDraftNoteMaxLength
        $reason = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'reason') { $item.reason } else { '' }) -MaxLength 200
        $kind = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'kind') { $item.kind } else { '' }) -MaxLength 40
        $visualReason = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'visualReason') { $item.visualReason } else { '' }) -MaxLength 200
        $targetCandidateId = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'targetCandidateId') { $item.targetCandidateId } else { '' }) -MaxLength 80
        $zoom = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'zoom') { $item.zoom } else { 'keep' }) -MaxLength 10
        if ($zoom -notin @('keep', 'focus', 'full')) { $zoom = 'keep' }
        $visualConfident = Get-MbBooleanOrDefault -Container $item -Name 'visualConfident' -Default $false
        $allowedCandidateIds = @(@($source.targetCandidates) | ForEach-Object { [string]$_.id })
        if ($allowedCandidateIds.Count -eq 0) {
            # 視覚候補がない手順では、Copilotが none や zoom を補っても画像編集へ使わない。
            $targetCandidateId = ''
            $zoom = 'keep'
            $visualConfident = $false
        } elseif ([string]::IsNullOrWhiteSpace($targetCandidateId)) {
            # 候補IDがないのに拡大だけを適用すると、現在の自動赤枠を誤って外し得る。
            $zoom = 'keep'
            $visualConfident = $false
        } elseif ($targetCandidateId -ne 'none' -and $allowedCandidateIds -notcontains $targetCandidateId) {
            $targetCandidateId = ''
            $zoom = 'keep'
            $visualConfident = $false
            if ([string]::IsNullOrWhiteSpace($visualReason)) { $visualReason = 'Copilotが一覧にない候補を返しました。' }
        }
        $selectedCandidate = $null
        if (-not [string]::IsNullOrWhiteSpace($targetCandidateId) -and $targetCandidateId -ne 'none') {
            $selectedCandidate = @($source.targetCandidates | Where-Object { [string]$_.id -eq $targetCandidateId }) | Select-Object -First 1
        }

        # 校正で3項目とも空なら、直すところが無いという意味。確認画面へ出さない。
        if ($Mode -eq 'review' -and
            [string]::IsNullOrWhiteSpace($title) -and
            [string]::IsNullOrWhiteSpace($description) -and
            [string]::IsNullOrWhiteSpace($note)) { continue }

        [void]$drafts.Add([pscustomobject]@{
            id              = $id
            sheetId         = [string]$source.sheetId
            sheetName       = [string]$source.sheetName
            order           = [int]$source.order
            keep            = Get-MbBooleanOrDefault -Container $item -Name 'keep' -Default $true
            confident       = Get-MbBooleanOrDefault -Container $item -Name 'confident' -Default $true
            reason          = $reason
            targetCandidateId = $targetCandidateId
            visualConfident = $visualConfident
            visualReason    = $visualReason
            zoom            = $zoom
            kind            = $kind
            title           = $title
            description     = $description
            note            = $note
            currentTitle    = [string]$source.title
            currentDescription = [string]$source.description
            currentNote     = [string]$source.note
            imageId         = [string]$source.imageId
            imageWidth      = $(if ($source.PSObject.Properties.Name -contains 'imageWidth') { [int]$source.imageWidth } else { 0 })
            imageHeight     = $(if ($source.PSObject.Properties.Name -contains 'imageHeight') { [int]$source.imageHeight } else { 0 })
            targetRect      = $(if ($null -ne $selectedCandidate) { $selectedCandidate.rect } else { $null })
            clickLabel      = $(if ($null -ne $selectedCandidate -and
                -not [string]::IsNullOrWhiteSpace([string]$selectedCandidate.label)) {
                    [string]$selectedCandidate.label
                } elseif ($targetCandidateId -eq 'none') { '' } else { [string]$source.clickLabel })
        })
    }
    return @($drafts)
}

function Get-MbCopilotCandidateBadgePoint {
    param(
        [Parameter(Mandatory = $true)]$Rect,
        [ValidateRange(0, 3)][int]$CandidateIndex
    )
    $badgeStep = 0.04
    $badgeOffset = $CandidateIndex * $badgeStep
    $badgeY = if (([double]$Rect.y1 + (3 * $badgeStep)) -le 0.98) {
        [double]$Rect.y1 + $badgeOffset
    } else {
        [double]$Rect.y1 - $badgeOffset
    }
    return [pscustomobject]@{
        x = [Math]::Max(0.02, [Math]::Min(0.98, [double]$Rect.x1))
        y = [Math]::Max(0.02, [Math]::Min(0.98, $badgeY))
    }
}

# 添付用に、赤枠を焼き込んだ画像を作る。
# 注釈も切り抜きも無い場合、New-MbAnnotatedImage は元のパスを返す。
# その戻り値を捨てると画像を見失うので、必ず受け取って使う。
function New-MbCopilotAttachment {
    param(
        [Parameter(Mandatory = $true)]$Step,
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$WorkDirectory,
        [Parameter(Mandatory = $true)][string]$FileName
    )

    if (-not (Test-Path -LiteralPath $WorkDirectory)) {
        [void](New-Item -ItemType Directory -Path $WorkDirectory -Force)
    }
    $burnedPath = Join-Path $WorkDirectory ('burn-' + $FileName)
    $attachmentAnnotations = @($Step.annotations)
    $attachmentCrop = $Step.crop
    $visualCandidates = @($Step.targetCandidates)
    if ($visualCandidates.Count -gt 0) {
        # 現在の赤枠を正解として刷り込まず、全候補を同じ画像へ番号付きで示す。
        # 黒塗りなど利用者が加えた非矩形注釈は残す。
        # 現在の自動切り抜き自体が誤っていても比較できるよう、候補画像は全画面にする。
        $attachmentCrop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
        $candidateAnnotations = New-Object System.Collections.ArrayList
        foreach ($annotation in @($Step.annotations)) {
            if ([string]$annotation.type -ne 'rect') { [void]$candidateAnnotations.Add($annotation) }
        }
        for ($candidateIndex = 0; $candidateIndex -lt $visualCandidates.Count; $candidateIndex++) {
            $candidate = $visualCandidates[$candidateIndex]
            if ($null -eq $candidate -or $candidate.PSObject.Properties.Name -notcontains 'rect' -or
                -not (Test-MbNormalizedRect -Rect $candidate.rect)) { continue }
            $rect = $candidate.rect
            [void]$candidateAnnotations.Add([pscustomobject]@{
                type = 'rect'; x1 = $rect.x1; y1 = $rect.y1; x2 = $rect.x2; y2 = $rect.y2; label = 0
            })
            # 複数候補の左上が同じでも番号が重ならないよう、最大4件を縦へずらす。
            $badge = Get-MbCopilotCandidateBadgePoint -Rect $rect -CandidateIndex $candidateIndex
            [void]$candidateAnnotations.Add([pscustomobject]@{
                type = 'number'
                x1 = $badge.x; y1 = $badge.y; x2 = $badge.x; y2 = $badge.y
                label = ($candidateIndex + 1)
            })
        }
        $attachmentAnnotations = @($candidateAnnotations)
    }
    $rendered = New-MbAnnotatedImage -SourcePath $SourcePath -Annotations $attachmentAnnotations -Crop $attachmentCrop -DestinationPath $burnedPath
    if ([string]::IsNullOrWhiteSpace($rendered)) { $rendered = $SourcePath }

    # Copilotの画面では添付名で照合するため、手順が分かる名前へ揃える。
    $destination = Join-Path $WorkDirectory $FileName
    [IO.File]::Copy($rendered, $destination, $true)
    return $destination
}

Export-ModuleMember -Function @(
    'Format-MbTimeCode',
    'Get-MbCopilotStepList',
    'Get-MbCopilotStyleSamples',
    'Get-MbCopilotPackets',
    'New-MbCopilotStepPrompt',
    'New-MbCopilotReviewPrompt',
    'ConvertFrom-MbCopilotStepAnswer',
    'New-MbCopilotAttachment',
    'Get-MbTrimmedText'
)
