# 手順の下書きをCopilotへ依頼するジョブ。
#
# 依頼文の考え方:
#   画面を見せて「説明を書いて」と頼んでも、当たり障りのない文しか返らない。
#   ManualBuilderはクリック座標と前後画像を事実として保持する。UIAやOCRの矩形は
#   あくまで候補であり、操作対象・意味・文章はCopilotの結果を利用者が確認して確定する。
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
                annotations  = @($step.annotations)
                crop         = $step.crop
                clickLabel   = [string](Get-MbStepCaptureValue -Step $step -Name 'clickLabel')
                windowTitle  = [string](Get-MbStepCaptureValue -Step $step -Name 'windowTitle')
                screenText   = [string](Get-MbStepCaptureValue -Step $step -Name 'screenText')
                narration    = [string](Get-MbStepCaptureValue -Step $step -Name 'narration')
                kind         = [string](Get-MbStepCaptureValue -Step $step -Name 'kind')
                videoTimeMs  = [int](Get-MbStepCaptureValue -Step $step -Name 'videoTimeMs' -Default 0)
                targetType   = [string](Get-MbStepCaptureValue -Step $step -Name 'targetType')
                clickX       = [double](Get-MbStepCaptureValue -Step $step -Name 'clickX' -Default -1.0)
                clickY       = [double](Get-MbStepCaptureValue -Step $step -Name 'clickY' -Default -1.0)
                afterImageId = [string](Get-MbStepCaptureValue -Step $step -Name 'afterImageId')
                targetCandidates = @((Get-MbStepCaptureValue -Step $step -Name 'targetCandidates' -Default @()))
                analysisState = [string](Get-MbStepCaptureValue -Step $step -Name 'analysisState')
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
        [ValidateSet('draft', 'review', 'operation')][string]$Mode = 'draft'
    )

    if ($StepsPerPacket -lt 1) { $StepsPerPacket = 1 }
    if ($Mode -eq 'review') {
        # 校正は書かれている文章が対象。画像は見ないので、文字だけの手順も含める。
        $targets = @($Steps | Where-Object {
            (-not [string]::IsNullOrWhiteSpace($_.title)) -or
            (-not [string]::IsNullOrWhiteSpace($_.description)) -or
            (-not [string]::IsNullOrWhiteSpace($_.note))
        })
    } elseif ($Mode -eq 'operation') {
        # 記録由来で未確定の操作はすべてCopilotへ送る。UIAの信頼度では間引かない。
        $targets = @($Steps | Where-Object {
            $_.imageId -and [string]$_.analysisState -in @('pending', 'needs-review')
        })
        $StepsPerPacket = 1
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
    [void]$builder.AppendLine('添付画像に赤枠がある場合、その枠は利用者が確定した注釈または以前の解析結果です。画像全体の文脈も確認してください。')
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
            [void]$builder.AppendLine(('赤枠の位置にあった操作対象: ' + [string]$step.clickLabel))
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
    [void]$builder.AppendLine('- 赤枠がある場合は操作対象の手掛かりにする。赤枠が無い画面は、その画面が何を表しているかを書く。')
    [void]$builder.AppendLine('- 画像と与えられた情報から読み取れないことは書かない。想像で補わない。')
    [void]$builder.AppendLine('- 判断できない手順は confident を false にし、reason に理由を短く書く。')
    [void]$builder.AppendLine('- 直前の手順と同じ画面である、操作されていない、といった理由で手順として不要な場合は keep を false にする。')
    [void]$builder.AppendLine('- すでに人が書いた内容がある手順は、それを尊重して整えるだけにする。')
    [void]$builder.AppendLine()

    [void]$builder.AppendLine('## 出力の形')
    [void]$builder.AppendLine('次の形のJSONだけを出力してください。説明文や前置きは書かないでください。')
    [void]$builder.AppendLine('{"steps":[{"id":"手順のid","keep":true,"title":"手順名","description":"説明","note":"補足","confident":true,"reason":""}]}')
    [void]$builder.AppendLine(('id は上に並べた ' + (@($PacketSteps | ForEach-Object { [string]$_.id }) -join ', ') + ' をそのまま使ってください。'))
    [void]$builder.AppendLine(('JSONを出力し終えたら、最後の行に ' + $Marker + ' とだけ書いてください。'))
    # この一文が回答の始まりを見つける目印になる。必ず依頼文の最後に置く。
    [void]$builder.Append((Get-MbCopilotPromptTailAnchor))
    return $builder.ToString()
}

# 操作記録は「1操作=1チャット」とし、決定論的な検出結果を正解扱いしない。
# クリック座標と前後画像が事実、UIA/MSAA/OCRは候補にすぎないことを明記する。
function New-MbCopilotOperationPrompt {
    param(
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)]$Step,
        [Parameter(Mandatory = $true)][hashtable]$EvidenceNames,
        [object[]]$StyleSamples = @(),
        [string]$Marker = 'MB_END'
    )

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('あなたは操作マニュアルの記録解析者です。1件の操作について、操作対象、操作の意味、手順文を同時に判断してください。')
    [void]$builder.AppendLine('クリック座標と撮影時刻は記録された事実です。候補枠、UIA名、コントロール種類は誤ることがあるため、正解として扱わないでください。')
    [void]$builder.AppendLine('操作前と操作後の差、クリック点の周囲、画面全体の文脈を合わせて判断してください。ページ遷移後の空白や別画面を操作対象にしないでください。')
    [void]$builder.AppendLine()
    [void]$builder.AppendLine(('マニュアル名: ' + [string]$Project.title))
    [void]$builder.AppendLine(('手順ID: ' + [string]$Step.id))
    [void]$builder.AppendLine(('操作種別: ' + [string]$Step.kind))
    if (-not [string]::IsNullOrWhiteSpace([string]$Step.windowTitle)) {
        [void]$builder.AppendLine(('操作時のウィンドウ: ' + [string]$Step.windowTitle))
    }
    if ([double]$Step.clickX -ge 0 -and [double]$Step.clickY -ge 0) {
        [void]$builder.AppendLine(('操作点（画像左上を0,0・右下を1,1）: ' +
            ([double]$Step.clickX).ToString('0.######', [Globalization.CultureInfo]::InvariantCulture) + ', ' +
            ([double]$Step.clickY).ToString('0.######', [Globalization.CultureInfo]::InvariantCulture)))
    } else {
        [void]$builder.AppendLine('操作点: 取得できませんでした。入力操作ではフォーカス候補と前後差を優先してください。')
    }
    [void]$builder.AppendLine(('操作前の全体画像（クリック点マーカー付き）: ' + [string]$EvidenceNames.before))
    [void]$builder.AppendLine(('操作点周辺の候補画像: ' + [string]$EvidenceNames.detail))
    if ($EvidenceNames.ContainsKey('after')) {
        [void]$builder.AppendLine(('操作後の全体画像: ' + [string]$EvidenceNames.after))
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Step.narration)) {
        [void]$builder.AppendLine(('利用者が話した内容: ' + [string]$Step.narration))
    }
    [void]$builder.AppendLine()
    [void]$builder.AppendLine('候補（候補番号は周辺画像の番号と対応）:')
    $candidateNumber = 0
    foreach ($candidate in @($Step.targetCandidates)) {
        $candidateNumber++
        $candidateId = if ($candidate.PSObject.Properties.Name -contains 'id') { [string]$candidate.id } else { [string]$candidateNumber }
        $candidateName = if ($candidate.PSObject.Properties.Name -contains 'name') { [string]$candidate.name } else { '' }
        $candidateType = if ($candidate.PSObject.Properties.Name -contains 'type') { [string]$candidate.type } else { '' }
        $candidateSource = if ($candidate.PSObject.Properties.Name -contains 'source') { [string]$candidate.source } else { '' }
        [void]$builder.AppendLine(('- ' + $candidateId + '（画像内番号 ' + $candidateNumber + '）: source=' + $candidateSource + '; type=' + $candidateType + '; name=' + $candidateName))
    }
    [void]$builder.AppendLine('候補がどれも違う場合は candidateId を空にし、bboxへ正しい矩形を指定してください。対象を確定できない場合も操作は捨てず、needsReview=trueにしてください。')
    [void]$builder.AppendLine('意味のない空クリックだと判断した場合だけ keep=falseにします。画面遷移や結果表示があれば、候補名が空でもkeep=trueにしてください。')
    [void]$builder.AppendLine()
    if (@($StyleSamples).Count -gt 0) {
        [void]$builder.AppendLine('文体の見本:')
        foreach ($sample in $StyleSamples) {
            [void]$builder.AppendLine(('- 手順名「' + [string]$sample.title + '」／説明「' + [string]$sample.description + '」'))
        }
        [void]$builder.AppendLine()
    }
    [void]$builder.AppendLine('JSONだけを次の形で返してください。bboxは原画像の正規化座標で、対象を確定できない場合はnullです。')
    [void]$builder.AppendLine('{"steps":[{"id":"手順ID","keep":true,"candidateId":"A","bbox":{"x1":0.1,"y1":0.1,"x2":0.2,"y2":0.2},"targetLabel":"対象名","targetType":"button","title":"手順名","description":"説明","note":"補足","needsReview":false,"reason":"判断根拠"}]}')
    [void]$builder.AppendLine('手順名は20文字程度の体言止め、説明は「〜します。」の敬体で2文までにしてください。画像から分からない結果を想像で補わないでください。')
    [void]$builder.AppendLine(('JSONの後、最後の行に ' + $Marker + ' とだけ書いてください。'))
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
        [ValidateSet('draft', 'review', 'operation')][string]$Mode = 'draft'
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
        $candidateId = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'candidateId') { $item.candidateId } else { '' }) -MaxLength 8
        $targetLabel = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'targetLabel') { $item.targetLabel } else { '' }) -MaxLength 300
        $targetType = Get-MbTrimmedText -Value $(if ($item.PSObject.Properties.Name -contains 'targetType') { $item.targetType } else { '' }) -MaxLength 80
        $targetRect = $null
        if ($Mode -eq 'operation') {
            # candidateIdが有効ならローカル候補の座標を使い、Copilotに座標を転記させない。
            foreach ($candidate in @($source.targetCandidates)) {
                if ($candidate.PSObject.Properties.Name -contains 'id' -and [string]$candidate.id -eq $candidateId -and
                    $candidate.PSObject.Properties.Name -contains 'rect' -and (Test-MbNormalizedRect -Rect $candidate.rect)) {
                    $targetRect = $candidate.rect
                    break
                }
            }
            if ($null -eq $targetRect -and $item.PSObject.Properties.Name -contains 'bbox' -and
                (Test-MbNormalizedRect -Rect $item.bbox)) {
                $area = ([double]$item.bbox.x2 - [double]$item.bbox.x1) * ([double]$item.bbox.y2 - [double]$item.bbox.y1)
                if ($area -lt 0.72) { $targetRect = $item.bbox }
            }
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
            kind            = $kind
            title           = $title
            description     = $description
            note            = $note
            currentTitle    = [string]$source.title
            currentDescription = [string]$source.description
            currentNote     = [string]$source.note
            clickLabel      = [string]$source.clickLabel
            candidateId     = $candidateId
            targetLabel     = $targetLabel
            targetType      = $targetType
            targetRect      = $targetRect
            needsReview     = Get-MbBooleanOrDefault -Container $item -Name 'needsReview' -Default ($null -eq $targetRect)
        })
    }
    return @($drafts)
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
    $rendered = New-MbAnnotatedImage -SourcePath $SourcePath -Annotations @($Step.annotations) -Crop $Step.crop -DestinationPath $burnedPath
    if ([string]::IsNullOrWhiteSpace($rendered)) { $rendered = $SourcePath }

    # Copilotの画面では添付名で照合するため、手順が分かる名前へ揃える。
    $destination = Join-Path $WorkDirectory $FileName
    [IO.File]::Copy($rendered, $destination, $true)
    return $destination
}

# 操作解析用の2〜3枚を作る。プロジェクト内の原画像は一切変更せず、
# 手動の黒塗りだけを引き継いだ一時コピーへクリック点・候補番号を重ねる。
function New-MbCopilotOperationEvidence {
    param(
        [Parameter(Mandatory = $true)]$Step,
        [Parameter(Mandatory = $true)][string]$BeforeSourcePath,
        [AllowEmptyString()][string]$AfterSourcePath = '',
        [Parameter(Mandatory = $true)][string]$WorkDirectory
    )

    if (-not (Test-Path -LiteralPath $WorkDirectory)) {
        [void](New-Item -ItemType Directory -Path $WorkDirectory -Force)
    }
    $fullCrop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
    $blackouts = @($Step.annotations | Where-Object { [string]$_.type -eq 'blackout' })
    $clickX = [double]$Step.clickX
    $clickY = [double]$Step.clickY
    if ($clickX -lt 0 -or $clickX -gt 1 -or $clickY -lt 0 -or $clickY -gt 1) {
        $clickX = 0.5; $clickY = 0.5
    }
    $markerWidth = 0.018
    $markerHeight = 0.024
    $clickMarker = [pscustomobject]@{
        id = 'annotation-' + [guid]::NewGuid().ToString('N'); type = 'rect'; label = 0
        x1 = [Math]::Max(0.0, $clickX - $markerWidth); y1 = [Math]::Max(0.0, $clickY - $markerHeight)
        x2 = [Math]::Min(1.0, $clickX + $markerWidth); y2 = [Math]::Min(1.0, $clickY + $markerHeight)
    }

    $beforeName = 'operation-before.jpg'
    $beforePath = Join-Path $WorkDirectory $beforeName
    $beforeRendered = New-MbAnnotatedImage -SourcePath $BeforeSourcePath -Annotations (@($blackouts) + @($clickMarker)) `
        -Crop $fullCrop -DestinationPath $beforePath
    if ([string]$beforeRendered -ne $beforePath) { [IO.File]::Copy([string]$beforeRendered, $beforePath, $true) }

    $candidateAnnotations = New-Object System.Collections.ArrayList
    $candidateNumber = 0
    foreach ($candidate in @($Step.targetCandidates | Select-Object -First 8)) {
        if ($candidate.PSObject.Properties.Name -notcontains 'rect' -or -not (Test-MbNormalizedRect -Rect $candidate.rect)) { continue }
        $candidateNumber++
        [void]$candidateAnnotations.Add([pscustomobject]@{
            id = 'annotation-' + [guid]::NewGuid().ToString('N'); type = 'rect'; label = 0
            x1 = [double]$candidate.rect.x1; y1 = [double]$candidate.rect.y1
            x2 = [double]$candidate.rect.x2; y2 = [double]$candidate.rect.y2
        })
        [void]$candidateAnnotations.Add([pscustomobject]@{
            id = 'annotation-' + [guid]::NewGuid().ToString('N'); type = 'number'; label = $candidateNumber
            x1 = [double]$candidate.rect.x1; y1 = [double]$candidate.rect.y1
            x2 = [Math]::Min(1.0, [double]$candidate.rect.x1 + 0.02)
            y2 = [Math]::Min(1.0, [double]$candidate.rect.y1 + 0.02)
        })
    }
    if ($candidateNumber -eq 0) { [void]$candidateAnnotations.Add($clickMarker) }
    $detailWidth = 0.52
    $detailHeight = 0.52
    $detailCrop = [pscustomobject]@{
        x = [Math]::Round([Math]::Max(0.0, [Math]::Min(1.0 - $detailWidth, $clickX - ($detailWidth / 2.0))), 6)
        y = [Math]::Round([Math]::Max(0.0, [Math]::Min(1.0 - $detailHeight, $clickY - ($detailHeight / 2.0))), 6)
        width = $detailWidth; height = $detailHeight
    }
    $detailName = 'operation-detail.jpg'
    $detailPath = Join-Path $WorkDirectory $detailName
    $detailRendered = New-MbAnnotatedImage -SourcePath $BeforeSourcePath `
        -Annotations (@($blackouts) + @($candidateAnnotations)) -Crop $detailCrop -DestinationPath $detailPath
    if ([string]$detailRendered -ne $detailPath) { [IO.File]::Copy([string]$detailRendered, $detailPath, $true) }

    $attachments = New-Object System.Collections.ArrayList
    [void]$attachments.Add($beforePath)
    [void]$attachments.Add($detailPath)
    $names = @{ before = $beforeName; detail = $detailName }
    if (-not [string]::IsNullOrWhiteSpace($AfterSourcePath) -and (Test-Path -LiteralPath $AfterSourcePath -PathType Leaf)) {
        $afterName = 'operation-after.jpg'
        $afterPath = Join-Path $WorkDirectory $afterName
        $afterRendered = New-MbAnnotatedImage -SourcePath $AfterSourcePath -Annotations $blackouts -Crop $fullCrop -DestinationPath $afterPath
        if ([string]$afterRendered -ne $afterPath) { [IO.File]::Copy([string]$afterRendered, $afterPath, $true) }
        [void]$attachments.Add($afterPath)
        $names.after = $afterName
    }
    return [pscustomobject]@{ attachments = @($attachments); names = $names; detailCrop = $detailCrop }
}

function Test-MbCopilotOperationPreflight {
    param([Parameter(Mandatory = $true)][string]$WorkDirectory)

    $directory = Join-Path $WorkDirectory 'preflight'
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $beforePath = Join-Path $directory 'synthetic-before.png'
    $afterPath = Join-Path $directory 'synthetic-after.png'
    foreach ($item in @(@{ path = $beforePath; color = [Drawing.Color]::SteelBlue }, @{ path = $afterPath; color = [Drawing.Color]::SeaGreen })) {
        $bitmap = New-Object Drawing.Bitmap -ArgumentList @(320, 180)
        $graphics = $null
        try {
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            $graphics.Clear([Drawing.Color]::White)
            $brush = New-Object Drawing.SolidBrush -ArgumentList $item.color
            try { $graphics.FillRectangle($brush, 112, 72, 96, 36) } finally { $brush.Dispose() }
            $bitmap.Save([string]$item.path, [Drawing.Imaging.ImageFormat]::Png)
        } finally {
            if ($null -ne $graphics) { $graphics.Dispose() }
            $bitmap.Dispose()
        }
    }
    $candidateRect = [pscustomobject]@{ x1 = 0.35; y1 = 0.4; x2 = 0.65; y2 = 0.6 }
    $step = [pscustomobject]@{
        annotations = @(); clickX = 0.5; clickY = 0.5
        targetCandidates = @([pscustomobject]@{ id = 'A'; source = 'preflight'; name = 'テスト'; type = 'button'; rect = $candidateRect })
    }
    $evidence = New-MbCopilotOperationEvidence -Step $step -BeforeSourcePath $beforePath `
        -AfterSourcePath $afterPath -WorkDirectory (Join-Path $directory 'evidence')
    if (@($evidence.attachments).Count -ne 3) { throw 'Copilot事前確認で3枚の証跡を作れませんでした。' }
    foreach ($path in @($evidence.attachments)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Item -LiteralPath $path).Length -lt 256) {
            throw 'Copilot事前確認で証跡画像を検証できませんでした。'
        }
    }
    return $evidence
}

Export-ModuleMember -Function @(
    'Format-MbTimeCode',
    'Get-MbCopilotStepList',
    'Get-MbCopilotStyleSamples',
    'Get-MbCopilotPackets',
    'New-MbCopilotStepPrompt',
    'New-MbCopilotOperationPrompt',
    'New-MbCopilotReviewPrompt',
    'ConvertFrom-MbCopilotStepAnswer',
    'New-MbCopilotAttachment',
    'New-MbCopilotOperationEvidence',
    'Test-MbCopilotOperationPreflight',
    'Get-MbTrimmedText'
)
