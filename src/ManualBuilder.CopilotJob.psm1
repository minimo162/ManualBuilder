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

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Excel.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1') -Force

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
                videoTimeMs  = [int](Get-MbStepCaptureValue -Step $step -Name 'videoTimeMs' -Default 0)
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
        [switch]$IncludeWritten
    )

    if ($StepsPerPacket -lt 1) { $StepsPerPacket = 1 }
    $targets = @($Steps | Where-Object {
        $_.imageId -and ($IncludeWritten -or [string]::IsNullOrWhiteSpace($_.title) -or [string]::IsNullOrWhiteSpace($_.description))
    })
    $packets = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $targets.Count; $i += $StepsPerPacket) {
        $count = [Math]::Min($StepsPerPacket, $targets.Count - $i)
        [void]$packets.Add(@($targets[$i..($i + $count - 1)]))
    }
    return @($packets)
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
    [void]$builder.AppendLine('添付した画像は、操作された場所を赤枠で囲んであります。赤枠はアプリが録画の変化から機械的に求めたもので、推測ではありません。')
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
            [void]$builder.AppendLine(('録画の音声: ' + [string]$step.narration))
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
    [void]$builder.AppendLine('- 赤枠と読み取れた操作対象を必ず主語にする。赤枠が無い画面は、その画面が何を表しているかを書く。')
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
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$PacketSteps
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

        [void]$drafts.Add([pscustomobject]@{
            id              = $id
            sheetId         = [string]$source.sheetId
            sheetName       = [string]$source.sheetName
            order           = [int]$source.order
            keep            = Get-MbBooleanOrDefault -Container $item -Name 'keep' -Default $true
            confident       = Get-MbBooleanOrDefault -Container $item -Name 'confident' -Default $true
            reason          = $reason
            title           = $title
            description     = $description
            note            = $note
            currentTitle    = [string]$source.title
            currentDescription = [string]$source.description
            currentNote     = [string]$source.note
            clickLabel      = [string]$source.clickLabel
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

Export-ModuleMember -Function @(
    'Format-MbTimeCode',
    'Get-MbCopilotStepList',
    'Get-MbCopilotStyleSamples',
    'Get-MbCopilotPackets',
    'New-MbCopilotStepPrompt',
    'ConvertFrom-MbCopilotStepAnswer',
    'New-MbCopilotAttachment',
    'Get-MbTrimmedText'
)
