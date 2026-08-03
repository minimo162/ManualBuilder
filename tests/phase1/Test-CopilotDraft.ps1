# Copilot下書きの組み立てと取り込みを検査する。
# Copilotへ実際に接続はしない。依頼文の作り方と、回答の読み取り方だけを見る。

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

Import-Module (Join-Path $srcRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Copilot.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.CopilotJob.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Ocr.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.CopilotServer.psm1') -Force

# ---------------------------------------------------------------------
# 手順の一覧とパケット分割
# ---------------------------------------------------------------------
$project = New-MbProject
$project.title = '経費精算システム操作手順'
$sheet = $project.sheets[0]
$sheet.name = '申請を出す'

$stepIds = @()
for ($i = 1; $i -le 5; $i++) {
    $step = Add-MbStep -Project $project -SheetId $sheet.id
    $step.imageId = ('image-' + ('{0:d32}' -f $i))
    $stepIds += [string]$step.id
}
# 1件目だけ人が書き終えている。
$first = Get-MbStepById -Project $project -StepId $stepIds[0]
$first.title = 'ログイン'
$first.description = '社員番号とパスワードを入力してログインします。一覧画面が表示されます。'
# 4件目は画像が無い。下書きの対象から外れる。
$fourth = Get-MbStepById -Project $project -StepId $stepIds[3]
$fourth.imageId = $null

$steps = Get-MbCopilotStepList -Project $project
Add-Result (@($steps).Count -eq 5) '手順の一覧が5件になる'
Add-Result ([string]$steps[0].sheetName -eq '申請を出す') '一覧にシート名が入る'
Add-Result ([int]$steps[2].order -eq 3) '一覧に通し番号が入る'

$packets = Get-MbCopilotPackets -Steps $steps -StepsPerPacket 2
$flat = @($packets | ForEach-Object { $_ })
Add-Result (@($flat).Count -eq 3) '文章が無く画像がある手順だけが対象になる（3件）'
Add-Result (@($flat | Where-Object { $_.id -eq $stepIds[0] }).Count -eq 0) '書き終えた手順は対象外になる'
Add-Result (@($flat | Where-Object { $_.id -eq $stepIds[3] }).Count -eq 0) '画像の無い手順は対象外になる'
Add-Result (@($packets).Count -eq 2) '2件ずつのまとまりに分かれる'
Add-Result (@($packets[0]).Count -eq 2 -and @($packets[1]).Count -eq 1) '端数のまとまりが残る'

$allPackets = Get-MbCopilotPackets -Steps $steps -StepsPerPacket 10 -IncludeWritten
Add-Result (@($allPackets[0]).Count -eq 4) '書き直しを選ぶと書き終えた手順も対象になる'

$samples = Get-MbCopilotStyleSamples -Steps $steps
Add-Result (@($samples).Count -eq 1 -and [string]$samples[0].title -eq 'ログイン') '文章のある手順が見本になる'

# ---------------------------------------------------------------------
# 依頼文
# ---------------------------------------------------------------------
$packet = @($packets[0])
$names = @{}
foreach ($step in $packet) { $names[[string]$step.id] = ('step-{0:d3}.jpg' -f [int]$step.order) }
$packet[0].clickLabel = '申請'
$packet[0].videoTimeMs = 72000
$packet[0].narration = 'ここで申請ボタンを押します'

$prompt = New-MbCopilotStepPrompt -Project $project -PacketSteps $packet -AttachmentNames $names `
    -StyleSamples $samples -TotalSteps 5 -Marker 'MB_END'

Add-Result ($prompt.Contains('経費精算システム操作手順')) '依頼文にマニュアル名が入る'
Add-Result ($prompt.Contains('申請を出す')) '依頼文にシート名が入る'
Add-Result ($prompt.Contains([string]$packet[0].id)) '依頼文に手順のidが入る'
Add-Result ($prompt.Contains('step-002.jpg')) '依頼文に添付画像の名前が入る'
Add-Result ($prompt.Contains('赤枠の位置から読み取れた操作対象: 申請')) '読み取った操作対象が依頼文に入る'
Add-Result ($prompt.Contains('1:12')) '録画内の時刻が入る'
Add-Result ($prompt.Contains('ここで申請ボタンを押します')) '録画の音声が入る'
Add-Result ($prompt.Contains('ログイン')) '文体の見本が入る'
Add-Result ($prompt.Contains('MB_END')) '終了の合図が入る'
# 回答の始まりを見つける目印。末尾に無いと依頼文自体を回答と読み違える。
Add-Result ($prompt.EndsWith((Get-MbCopilotPromptTailAnchor))) '依頼文が目印で終わる'

Add-Result ((Format-MbTimeCode -Milliseconds 0) -eq '') '時刻0は空文字になる'
Add-Result ((Format-MbTimeCode -Milliseconds 5000) -eq '0:05') '秒が2桁で並ぶ'

# ---------------------------------------------------------------------
# 回答からJSONを取り出す
# ---------------------------------------------------------------------
$targetId = [string]$packet[0].id
$body = '{"steps":[{"id":"' + $targetId + '","keep":true,"title":"申請の作成","description":"「申請」を選択します。入力画面が表示されます。","note":"","confident":true,"reason":""}]}'

$answer = Get-MbStepAnswerJson -Text ("承知しました。以下が結果です。`n" + $body + "`nMB_END")
Add-Result ($null -ne $answer) '前後に説明文があってもJSONを取り出せる'
Add-Result (@($answer.steps).Count -eq 1) '取り出したJSONに手順が入る'

$withTrailingComma = '{"steps":[{"id":"' + $targetId + '","title":"申請の作成","description":"説明です。",}]}'
Add-Result ($null -ne (Get-MbStepAnswerJson -Text $withTrailingComma)) '末尾のカンマがあっても読み取れる'

Add-Result ($null -eq (Get-MbStepAnswerJson -Text 'お手伝いできません。')) 'JSONが無ければnullになる'
Add-Result ($null -eq (Get-MbStepAnswerJson -Text '{"result":"ok"}')) 'stepsが無いJSONは採用しない'
Add-Result ($null -eq (Get-MbStepAnswerJson -Text '{"steps":[]}')) '空のstepsは採用しない'

$candidates = @(Get-MbJsonObjectCandidates -Text 'a {"x":"}"} b {"y":1}')
Add-Result (@($candidates).Count -eq 2) '文字列の中の波括弧に惑わされない'

# ---------------------------------------------------------------------
# 回答の取り込み
# ---------------------------------------------------------------------
$parsed = Get-MbStepAnswerJson -Text $body
$drafts = ConvertFrom-MbCopilotStepAnswer -Answer $parsed -PacketSteps $packet
Add-Result (@($drafts).Count -eq 1) '依頼した手順の下書きを受け取る'
Add-Result ([string]$drafts[0].title -eq '申請の作成') '手順名が入る'
Add-Result ($drafts[0].keep -eq $true -and $drafts[0].confident -eq $true) '既定で採用・自信ありになる'
Add-Result ([string]$drafts[0].clickLabel -eq '申請') '読み取った操作対象が下書きに残る'

$strayBody = '{"steps":[{"id":"step-does-not-exist","title":"別のもの","description":"説明"}]}'
$strayDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $strayBody) -PacketSteps $packet
Add-Result (@($strayDrafts).Count -eq 0) '依頼していないidは捨てる'

$duplicateBody = '{"steps":[{"id":"' + $targetId + '","title":"1つ目","description":"説明"},{"id":"' + $targetId + '","title":"2つ目","description":"説明"}]}'
$duplicateDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $duplicateBody) -PacketSteps $packet
Add-Result (@($duplicateDrafts).Count -eq 1 -and [string]$duplicateDrafts[0].title -eq '1つ目') '同じidが重なったら最初だけを使う'

$longTitle = 'あ' * 200
$longBody = '{"steps":[{"id":"' + $targetId + '","title":"' + $longTitle + '","description":"説明"}]}'
$longDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $longBody) -PacketSteps $packet
Add-Result ([string]$longDrafts[0].title.Length -le 40) '長すぎる手順名は切り詰める'

$dropBody = '{"steps":[{"id":"' + $targetId + '","keep":false,"confident":false,"reason":"直前と同じ画面です","title":"","description":""}]}'
$dropDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $dropBody) -PacketSteps $packet
Add-Result ($dropDrafts[0].keep -eq $false -and $dropDrafts[0].confident -eq $false) '不要・自信なしの判断が残る'
Add-Result ([string]$dropDrafts[0].reason -eq '直前と同じ画面です') '判断の理由が残る'

Add-Result (@(ConvertFrom-MbCopilotStepAnswer -Answer $null -PacketSteps $packet).Count -eq 0) '回答がnullでも落ちない'

# ---------------------------------------------------------------------
# 採用した下書きの書き込み
# ---------------------------------------------------------------------
$applyTarget = $stepIds[1]
$selection = '{"accept":[{"id":"' + $applyTarget + '","title":"申請の作成","description":"「申請」を選択します。","note":""}]}'
$applied = Set-MbCopilotDraftSelection -Project $project -SelectionJson $selection
$appliedStep = Get-MbStepById -Project $project -StepId $applyTarget
Add-Result ($applied -eq 1) '採用した件数を返す'
Add-Result ([string]$appliedStep.title -eq '申請の作成') '採用した手順名が入る'

# 空文字は「変えない」の意味。既に書いた文章を消してはいけない。
$keepSelection = '{"accept":[{"id":"' + $applyTarget + '","title":"","description":"","note":""}]}'
[void](Set-MbCopilotDraftSelection -Project $project -SelectionJson $keepSelection)
$appliedStep = Get-MbStepById -Project $project -StepId $applyTarget
Add-Result ([string]$appliedStep.title -eq '申請の作成') '空の下書きは既存の文章を消さない'

$missingSelection = '{"accept":[{"id":"step-missing","title":"あ","description":"い"}]}'
Add-Result ((Set-MbCopilotDraftSelection -Project $project -SelectionJson $missingSelection) -eq 0) '存在しない手順は飛ばす'

# ---------------------------------------------------------------------
# 手順に付く録画の情報
# ---------------------------------------------------------------------
[void](Set-MbStepCapture -Project $project -StepId $applyTarget -Kind 'video-scene' -VideoTimeMs 4500 `
    -ClickLabel '登録' -ScreenText '経費申請 / 登録 / 取消' -Narration '登録を押します')
$captured = Get-MbStepById -Project $project -StepId $applyTarget
Add-Result ([int]$captured.capture.videoTimeMs -eq 4500) '録画の時刻を保存する'
Add-Result ([string]$captured.capture.clickLabel -eq '登録') '操作対象を保存する'

# capture を持たない古いプロジェクトを保存して読み直しても壊れないこと。
# v0.27.3以前で作ったマニュアルがこの形になっている。
$legacyRoot = Join-Path ([IO.Path]::GetTempPath()) ('mb-copilot-test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $legacyRoot -Force)
try {
    $legacyPath = Join-Path $legacyRoot 'project.json'
    $legacy = New-MbProject
    $legacyStep = Add-MbStep -Project $legacy -SheetId $legacy.sheets[0].id
    $legacyStep.PSObject.Properties.Remove('capture')
    [void](Save-MbProject -Project $legacy -Path $legacyPath)
    $reloaded = Get-MbProject -Path $legacyPath
    $repaired = Get-MbStepById -Project $reloaded -StepId ([string]$legacyStep.id)
    Add-Result ($null -ne $repaired) '古いプロジェクトを読み直せる'
    Add-Result ($repaired.PSObject.Properties.Name -contains 'capture') '古い手順にも capture が補われる'
    Add-Result ([int]$repaired.capture.videoTimeMs -eq 0) '補われた capture が既定値になる'
} finally {
    Remove-Item -LiteralPath $legacyRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 操作位置の矩形
# ---------------------------------------------------------------------
Add-Result ((Test-MbNormalizedRect -Rect ([pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.3; y2 = 0.2 })) -eq $true) '正しい矩形を受け入れる'
Add-Result ((Test-MbNormalizedRect -Rect ([pscustomobject]@{ x1 = -0.1; y1 = 0.1; x2 = 0.3; y2 = 0.2 })) -eq $false) '範囲外の矩形をはねる'
Add-Result ((Test-MbNormalizedRect -Rect ([pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.1005; y2 = 0.2 })) -eq $false) '潰れた矩形をはねる'
Add-Result ((Test-MbNormalizedRect -Rect $null) -eq $false) 'nullをはねる'

# OCRの結果を作って、赤枠の寄せ方だけを見る（実際の文字認識は呼ばない）。
$snapshot = [pscustomobject]@{
    available = $true
    words = @(
        [pscustomobject]@{ text = '申請'; x1 = 0.50; y1 = 0.40; x2 = 0.56; y2 = 0.43 },
        [pscustomobject]@{ text = '取消'; x1 = 0.60; y1 = 0.40; x2 = 0.66; y2 = 0.43 }
    )
}
$inside = Resolve-MbOperationRect -Rect ([pscustomobject]@{ x1 = 0.48; y1 = 0.38; x2 = 0.58; y2 = 0.45 }) -Snapshot $snapshot
Add-Result ([string]$inside.matched -eq 'inside') '変化領域の中の文字を操作対象とする'
Add-Result ([string]$inside.label -eq '申請') '操作対象の名前を取り出す'
Add-Result ([double]$inside.rect.x1 -eq 0.48) '中に文字があるときは枠を変えない'

$nearest = Resolve-MbOperationRect -Rect ([pscustomobject]@{ x1 = 0.470; y1 = 0.405; x2 = 0.492; y2 = 0.425 }) -Snapshot $snapshot
Add-Result ([string]$nearest.matched -eq 'nearest') '近くの文字へ寄せる'
Add-Result ([double]$nearest.rect.x2 -ge 0.56) '寄せた枠が文字を含む'
Add-Result ([string]$nearest.label -eq '申請') '寄せた文字の名前を使う'

$far = Resolve-MbOperationRect -Rect ([pscustomobject]@{ x1 = 0.02; y1 = 0.90; x2 = 0.06; y2 = 0.94 }) -Snapshot $snapshot
Add-Result ([string]$far.matched -eq 'none') '近くに文字が無ければ枠を変えない'
Add-Result ([string]$far.label -eq '') '名前は空のままにする'

$noOcr = Resolve-MbOperationRect -Rect ([pscustomobject]@{ x1 = 0.1; y1 = 0.1; x2 = 0.2; y2 = 0.2 }) -Snapshot ([pscustomobject]@{ available = $false; words = @() })
Add-Result ([string]$noOcr.matched -eq 'none' -and [double]$noOcr.rect.x2 -eq 0.2) '文字認識が使えなくても枠はそのまま使える'

# ---------------------------------------------------------------------
Write-Host ''
if ($errors.Count -eq 0) {
    Write-Host 'Copilot下書きの検査はすべて成功しました。' -ForegroundColor Green
    exit 0
}
Write-Host ("失敗: " + $errors.Count + " 件") -ForegroundColor Red
exit 1
