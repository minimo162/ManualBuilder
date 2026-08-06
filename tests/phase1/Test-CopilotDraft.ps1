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
Import-Module (Join-Path $srcRoot 'ManualBuilder.Capture.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Copilot.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.CopilotJob.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Ocr.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.CopilotServer.psm1') -Force

$copilotWorkerText = [IO.File]::ReadAllText((Join-Path $srcRoot 'Invoke-ManualBuilderCopilotJob.ps1'), [Text.Encoding]::UTF8)
Add-Result ($copilotWorkerText -match '\$maximumAttempts\s*=\s*2' -and
    $copilotWorkerText -match '回答形式を読み取れないため再試行') `
    'Copilotの一時的な空回答を同じパケットで1回だけ再試行する'

# ---------------------------------------------------------------------
# Copilotの画像利用確認
# ---------------------------------------------------------------------
$copilotModule = Get-Module ManualBuilder.Copilot
$japaneseConsent = '開始する前に… 画像の分析や編集を Copilot に手伝ってもらうことができます。 確認して続行 今はしない'
$englishConsent = 'Before you begin Images may be processed by Copilot. Confirm and continue Not now'
$ordinaryDialog = '画像を添付します。続行しますか。'
Add-Result (& $copilotModule { param($Text) Test-MbCopilotImageConsentText -Text $Text } $japaneseConsent) `
    '日本語の画像利用確認を検出する'
Add-Result (& $copilotModule { param($Text) Test-MbCopilotImageConsentText -Text $Text } $englishConsent) `
    '英語の画像利用確認を検出する'
Add-Result (-not (& $copilotModule { param($Text) Test-MbCopilotImageConsentText -Text $Text } $ordinaryDialog)) `
    '通常の画像確認を初回同意画面と誤認しない'
$pageCandidates = @(
    [pscustomobject]@{ page = [pscustomobject]@{ id = 'blocked' }; ready = $true; consentRequired = $true },
    [pscustomobject]@{ page = [pscustomobject]@{ id = 'usable' }; ready = $true; consentRequired = $false }
)
$usablePage = & $copilotModule { param($Items) Select-MbCopilotPageCandidate -Candidates $Items } $pageCandidates
$consentPage = & $copilotModule { param($Items) Select-MbCopilotPageCandidate -Candidates $Items -PreferConsent } $pageCandidates
Add-Result ([string]$usablePage.page.id -eq 'usable') '通常処理では画像同意に塞がれていないCopilotタブを選ぶ'
Add-Result ([string]$consentPage.page.id -eq 'blocked') 'Copilot画面を開く操作では画像同意のあるタブを選ぶ'

# ---------------------------------------------------------------------
# 下書きジョブのスナップショット
# ---------------------------------------------------------------------
# Copilotを起動せず、ワーカー起動だけを差し替えて、ジョブが参照する画像を検査する。
$snapshotRoot = Join-Path ([IO.Path]::GetTempPath()) ('mb-copilot-snapshot-test-' + [guid]::NewGuid().ToString('N'))
$snapshotProjectRoot = Join-Path $snapshotRoot 'project'
$snapshotJobsRoot = Join-Path $snapshotRoot 'jobs'
$snapshotProfileRoot = Join-Path $snapshotRoot 'profile'
[void](New-Item -ItemType Directory -Path (Join-Path $snapshotProjectRoot 'images') -Force)
[void](New-Item -ItemType Directory -Path $snapshotJobsRoot -Force)
[void](New-Item -ItemType Directory -Path $snapshotProfileRoot -Force)
try {
    $snapshotProjectPath = Join-Path $snapshotProjectRoot 'project.json'
    $snapshotProject = New-MbProject
    $snapshotSheet = $snapshotProject.sheets[0]
    $targetImageId = 'image-' + ('1' * 32)
    $writtenImageId = 'image-' + ('2' * 32)
    $unusedImageId = 'image-' + ('3' * 32)
    $resultImageId = 'image-' + ('4' * 32)
    $snapshotProject.images = @(
        [pscustomobject]@{ id = $targetImageId; fileName = "$targetImageId.png"; sha256 = ('A' * 64); width = 1; height = 1; byteLength = 4; mimeType = 'image/png'; source = 'recorder'; createdAt = [DateTime]::UtcNow.ToString('o') },
        [pscustomobject]@{ id = $writtenImageId; fileName = "$writtenImageId.png"; sha256 = ('B' * 64); width = 1; height = 1; byteLength = 4; mimeType = 'image/png'; source = 'recorder'; createdAt = [DateTime]::UtcNow.ToString('o') },
        [pscustomobject]@{ id = $unusedImageId; fileName = "$unusedImageId.png"; sha256 = ('C' * 64); width = 1; height = 1; byteLength = 4; mimeType = 'image/png'; source = 'recorder'; createdAt = [DateTime]::UtcNow.ToString('o') },
        [pscustomobject]@{ id = $resultImageId; fileName = "$resultImageId.png"; sha256 = ('D' * 64); width = 1; height = 1; byteLength = 4; mimeType = 'image/png'; source = 'recorder'; createdAt = [DateTime]::UtcNow.ToString('o') }
    )
    $targetStep = Add-MbStep -Project $snapshotProject -SheetId $snapshotSheet.id
    $targetStep.imageId = $targetImageId
    $targetStep.resultImageId = $resultImageId
    $writtenStep = Add-MbStep -Project $snapshotProject -SheetId $snapshotSheet.id
    $writtenStep.imageId = $writtenImageId
    $writtenStep.title = '記入済み'
    $writtenStep.description = 'この手順は下書き対象外です。'
    foreach ($image in @($snapshotProject.images)) {
        [IO.File]::WriteAllBytes((Join-Path (Join-Path $snapshotProjectRoot 'images') ([string]$image.fileName)), [byte[]](1, 2, 3, 4))
    }
    [void](Save-MbProject -Project $snapshotProject -Path $snapshotProjectPath)

    Initialize-MbCopilotServer -JobsRoot $snapshotJobsRoot -ScriptRoot $srcRoot -ProfileRoot $snapshotProfileRoot
    $snapshotTestStartProcessCalls = 0
    $workerStarter = {
        param([string]$FilePath, [object[]]$ArgumentList)
        $script:snapshotTestStartProcessCalls++
        return [pscustomobject]@{ Id = $PID }
    }

    $jobStatus = Start-MbCopilotDraftJob -ProjectPath $snapshotProjectPath -WorkerStarter $workerStarter
    $jobDirectories = @(Get-ChildItem -LiteralPath $snapshotJobsRoot -Directory)
    Add-Result ([string]$jobStatus.state -eq 'queued') '実Copilotへ接続せず下書きジョブを開始できる'
    Add-Result ($jobDirectories.Count -eq 1) '下書きジョブのスナップショットを1件作る'
    if ($jobDirectories.Count -eq 1) {
        $jobProjectPath = Join-Path $jobDirectories[0].FullName 'project.json'
        $jobProject = Get-MbProject -Path $jobProjectPath
        $resolvedTargetPath = Get-MbImageFilePath -Project $jobProject -ProjectPath $jobProjectPath -ImageId $targetImageId
        $resolvedResultPath = Get-MbImageFilePath -Project $jobProject -ProjectPath $jobProjectPath -ImageId $resultImageId
        Add-Result (Test-Path -LiteralPath $resolvedTargetPath -PathType Leaf) 'ワーカーが下書き対象画像をスナップショットから解決できる'
        Add-Result (Test-Path -LiteralPath $resolvedResultPath -PathType Leaf) 'ワーカーが操作後の結果画像もスナップショットから解決できる'
        Add-Result (-not (Test-Path -LiteralPath (Join-Path $jobDirectories[0].FullName "images\$writtenImageId.png") -PathType Leaf)) '記入済みで対象外の画像は複製しない'
        Add-Result (-not (Test-Path -LiteralPath (Join-Path $jobDirectories[0].FullName "images\$unusedImageId.png") -PathType Leaf)) '手順から参照されない画像は複製しない'
    }
    Add-Result ([int]$snapshotTestStartProcessCalls -eq 1) 'テストではワーカー起動を差し替え、Copilotへ接続しない'

    # 必要な画像が欠けている場合は、不完全なジョブを起動・放置しない。
    Remove-MbCopilotDraftJob
    Remove-Item -LiteralPath (Join-Path (Join-Path $snapshotProjectRoot 'images') "$targetImageId.png") -Force
    $missingImageRejected = $false
    try {
        [void](Start-MbCopilotDraftJob -ProjectPath $snapshotProjectPath -WorkerStarter $workerStarter)
    } catch {
        $missingImageRejected = $_.Exception.Message -like '*画像ファイルが見つかりません*'
    }
    Add-Result $missingImageRejected '下書き対象画像が欠けていればワーカー起動前に知らせる'
    Add-Result ([int]$snapshotTestStartProcessCalls -eq 1) '画像欠損時はワーカーを起動しない'
    Add-Result (@(Get-ChildItem -LiteralPath $snapshotJobsRoot -Directory).Count -eq 0) '画像欠損時は不完全なジョブを残さない'
} finally {
    Remove-MbCopilotDraftJob
    Remove-Item -LiteralPath $snapshotRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
# 録画で同じ画面へ戻った場面の取り込み
# ---------------------------------------------------------------------
$sceneRoot = Join-Path ([IO.Path]::GetTempPath()) ('mb-video-scene-test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $sceneRoot -Force)
try {
    $sceneProjectPath = Join-Path $sceneRoot 'project.json'
    $sceneProject = New-MbProject
    $sceneSheet = $sceneProject.sheets[0]
    $sameFrame = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==')

    $firstVisit = Import-MbVideoScene -Project $sceneProject -ProjectPath $sceneProjectPath `
        -SheetId ([string]$sceneSheet.id) -Bytes $sameFrame -TimeMs 1000 -SkipOcr
    $returnVisit = Import-MbVideoScene -Project $sceneProject -ProjectPath $sceneProjectPath `
        -SheetId ([string]$sceneSheet.id) -Bytes $sameFrame -TimeMs 5000 -SkipOcr

    $sceneSteps = @($sceneProject.sheets[0].steps)
    Add-Result ($firstVisit.status -eq 'added' -and $returnVisit.status -eq 'added') '同じ画面へ戻った場面も別の手順として取り込む'
    Add-Result ($sceneSteps.Count -eq 2 -and [string]$sceneSteps[0].id -ne [string]$sceneSteps[1].id) '再訪した場面が別の手順IDを持つ'
    Add-Result (@($sceneProject.images).Count -eq 1 -and [string]$sceneSteps[0].imageId -eq [string]$sceneSteps[1].imageId) '再訪した場面は画像実体を共有する'
    Add-Result ([int]$sceneSteps[0].capture.videoTimeMs -eq 1000 -and [int]$sceneSteps[1].capture.videoTimeMs -eq 5000) '再訪した各手順に録画時刻を残す'
    $retryVisit = Import-MbVideoScene -Project $sceneProject -ProjectPath $sceneProjectPath `
        -SheetId ([string]$sceneSheet.id) -Bytes $sameFrame -TimeMs 5000 -SkipOcr
    Add-Result ($retryVisit.status -eq 'duplicate' -and @($sceneProject.sheets[0].steps).Count -eq 2) '同じ場面の再実行では手順を二重にしない'
} finally {
    Remove-Item -LiteralPath $sceneRoot -Recurse -Force -ErrorAction SilentlyContinue
}

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

$steps[1].resultImageId = 'image-result-000000000000000000000001'
$imageLimitedPackets = Get-MbCopilotPackets -Steps $steps -StepsPerPacket 6 -MaxAttachmentsPerPacket 3
$imageLimitedCounts = @($imageLimitedPackets | ForEach-Object {
    $count = 0
    foreach ($item in @($_)) { $count += 1 + $(if ($item.resultImageId) { 1 } else { 0 }) }
    $count
})
Add-Result (@($imageLimitedPackets).Count -eq 2) '操作後画像を含めてもCopilotの画像上限ごとに分割する'
Add-Result (@($imageLimitedCounts | Where-Object { $_ -gt 3 }).Count -eq 0) '各Copilot依頼の画像を3枚以下にする'
Add-Result (Test-MbCopilotImageLimitText -Text '追加しようとしている画像の数が上限を超えています。Copilot では現在、一度に最大 3 個の画像を追加できます。') '画像上限の警告を即時検出する'
Add-Result ((Get-MbCopilotDefaultSettings).max_images_per_request -eq 2) '実機で安定した2枚をCopilot依頼の安全上限にする'
Add-Result (Test-MbAttachmentNameMatch -Actual 'mb-de09314…' -Expected 'mb-de09314a0b-p01-step-001-result.jpg') `
    'Copilot画面で省略された固有添付名を照合する'

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
$resultNames = @{ ([string]$packet[0].id) = 'step-002-result.jpg' }
$packet[0].clickLabel = '申請'
$packet[0].videoTimeMs = 72000
$packet[0].narration = 'ここで申請ボタンを押します'
$packet[0].targetCandidates = @(
    [pscustomobject]@{
        id = 'video-diff-1'; source = 'video-diff'; confidence = 'low'; label = '申請'; targetType = ''
        rect = [pscustomobject]@{ x1 = 0.1; y1 = 0.2; x2 = 0.3; y2 = 0.4 }
    },
    [pscustomobject]@{
        id = 'video-diff-2'; source = 'video-diff'; confidence = 'low'; label = '取消'; targetType = ''
        rect = [pscustomobject]@{ x1 = 0.5; y1 = 0.6; x2 = 0.7; y2 = 0.8 }
    }
)

$prompt = New-MbCopilotStepPrompt -Project $project -PacketSteps $packet -AttachmentNames $names `
    -ResultAttachmentNames $resultNames -StyleSamples $samples -TotalSteps 5 -Marker 'MB_END'

Add-Result ($prompt.Contains('経費精算システム操作手順')) '依頼文にマニュアル名が入る'
Add-Result ($prompt.Contains('申請を出す')) '依頼文にシート名が入る'
Add-Result ($prompt.Contains([string]$packet[0].id)) '依頼文に手順のidが入る'
Add-Result ($prompt.Contains('step-002.jpg')) '依頼文に添付画像の名前が入る'
Add-Result ($prompt.Contains('step-002-result.jpg') -and $prompt.Contains('操作後の結果')) '依頼文でクリック前と操作後の結果を比較できる'
Add-Result ($prompt.Contains('アプリが暫定選択した操作対象: 申請')) '暫定の操作対象が依頼文に入る'
Add-Result ($prompt.Contains('id=video-diff-1')) '選択できる候補IDが依頼文に入る'
Add-Result ($prompt.Contains('確定した事実ではありません')) 'DOM・UIA・動画差分を確定扱いしない'
Add-Result ($prompt.Contains('UIA候補が低信頼') -and $prompt.Contains('click-pointを優先')) '横長UIAよりクリック座標を優先する条件を伝える'
Add-Result (-not $prompt.Contains('推測ではありません')) '誤った確定表現を依頼文へ入れない'
Add-Result ($prompt.Contains('1:12')) '録画内の時刻が入る'
Add-Result ($prompt.Contains('ここで申請ボタンを押します')) '録画の音声が入る'
Add-Result ($prompt.Contains('ログイン')) '文体の見本が入る'
Add-Result ($prompt.Contains('原則として操作を表す1文')) '説明を簡潔な1文にするよう頼む'
Add-Result ($prompt.Contains('自明な結果は繰り返さない')) '操作と同義の結果を重ねないよう頼む'
Add-Result ($prompt.Contains('MB_END')) '終了の合図が入る'
# 回答の始まりを見つける目印。末尾に無いと依頼文自体を回答と読み違える。
Add-Result ($prompt.EndsWith((Get-MbCopilotPromptTailAnchor))) '依頼文が目印で終わる'

$badgeRect = [pscustomobject]@{ x1 = 0.2; y1 = 0.3; x2 = 0.5; y2 = 0.6 }
$badgePoints = & (Get-Module ManualBuilder.CopilotJob) {
    param($Rect)
    @(0..3 | ForEach-Object { Get-MbCopilotCandidateBadgePoint -Rect $Rect -CandidateIndex $_ })
} $badgeRect
$badgeYs = @($badgePoints | ForEach-Object { [Math]::Round([double]$_.y, 3) } | Sort-Object -Unique)
Add-Result ($badgeYs.Count -eq 4 -and ([double]$badgeYs[1] - [double]$badgeYs[0]) -ge 0.039) `
    '同じ左上の候補番号を識別できる間隔でずらす'
Add-Result ([double]$badgePoints[0].y -lt [double]$badgeRect.y1 -and
    [double]$badgePoints[1].y -gt [double]$badgeRect.y1) '候補番号を枠の外側へ上下に逃がして対象を隠さない'

$attachmentRoot = Join-Path ([IO.Path]::GetTempPath()) ('mb-copilot-attachment-test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $attachmentRoot -Force)
try {
    Add-Type -AssemblyName System.Drawing
    $attachmentSource = Join-Path $attachmentRoot 'source.png'
    $attachmentBitmap = New-Object Drawing.Bitmap 100, 80
    try { $attachmentBitmap.Save($attachmentSource, [Drawing.Imaging.ImageFormat]::Png) } finally { $attachmentBitmap.Dispose() }
    $attachmentStep = [pscustomobject]@{
        annotations = @()
        crop = [pscustomobject]@{ x = 0.25; y = 0.25; width = 0.5; height = 0.5 }
        targetCandidates = @(0..3 | ForEach-Object {
            [pscustomobject]@{
                id = 'candidate-' + ($_ + 1); rect = $badgeRect
                source = 'video-diff'; confidence = 'low'; label = ''; targetType = ''
            }
        })
    }
    $attachmentPath = New-MbCopilotAttachment -Step $attachmentStep -SourcePath $attachmentSource `
        -WorkDirectory $attachmentRoot -FileName 'candidate.jpg'
    $attachmentImage = [Drawing.Image]::FromFile($attachmentPath)
    try {
        Add-Result ($attachmentImage.Width -eq 100 -and $attachmentImage.Height -eq 80) `
            '候補添付では現在のcropを使わず全画面を保持する'
    } finally { $attachmentImage.Dispose() }
} finally {
    Remove-Item -LiteralPath $attachmentRoot -Recurse -Force -ErrorAction SilentlyContinue
}

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

$visualBody = '{"steps":[{"id":"' + $targetId + '","targetCandidateId":"video-diff-2","zoom":"focus","visualConfident":true,"visualReason":"取消ボタンと一致","title":"取消","description":"取消を選択します。"}]}'
$visualDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $visualBody) -PacketSteps $packet
Add-Result ([string]$visualDrafts[0].targetCandidateId -eq 'video-diff-2') '列挙した視覚候補IDを受け取る'
Add-Result ($visualDrafts[0].visualConfident -eq $true -and [string]$visualDrafts[0].zoom -eq 'focus') '枠の確信度と拡大方針を受け取る'
Add-Result ([string]$visualDrafts[0].clickLabel -eq '取消') 'Copilotが選んだ候補の名前を確認画面へ返す'
Add-Result ([Math]::Abs([double]$visualDrafts[0].targetRect.x1 - 0.5) -lt 0.001 -and
    -not [string]::IsNullOrWhiteSpace([string]$visualDrafts[0].imageId)) '選んだ赤枠と元画像を確認画面へ返す'
$ordinalVisualBody = '{"steps":[{"id":"' + $targetId + '","targetCandidateId":"候補2","zoom":"focus","visualConfident":true,"visualReason":"候補2が取消ボタンと一致","title":"取消","description":"取消を選択します。"}]}'
$ordinalVisualDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $ordinalVisualBody) -PacketSteps $packet
Add-Result ([string]$ordinalVisualDrafts[0].targetCandidateId -eq 'video-diff-2' -and $ordinalVisualDrafts[0].visualConfident) `
    'Copilotが返した候補番号を列挙済みIDへ正規化する'
$reasonOnlyVisualBody = '{"steps":[{"id":"' + $targetId + '","targetCandidateId":"","zoom":"focus","visualConfident":true,"visualReason":"候補1は申請ボタンと一致します","title":"申請","description":"申請を選択します。"}]}'
$reasonOnlyVisualDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $reasonOnlyVisualBody) -PacketSteps $packet
Add-Result ([string]$reasonOnlyVisualDrafts[0].targetCandidateId -eq 'video-diff-1' -and $reasonOnlyVisualDrafts[0].visualConfident) `
    '候補を明示した根拠と確信があれば欠けた候補IDを補う'
$unknownVisualBody = '{"steps":[{"id":"' + $targetId + '","targetCandidateId":"made-up","zoom":"focus","visualConfident":true,"title":"申請","description":"申請を選択します。"}]}'
$unknownVisualDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $unknownVisualBody) -PacketSteps $packet
Add-Result ([string]::IsNullOrWhiteSpace([string]$unknownVisualDrafts[0].targetCandidateId) -and -not $unknownVisualDrafts[0].visualConfident) '一覧にない候補IDを拒否する'
Add-Result ([string]$unknownVisualDrafts[0].zoom -eq 'keep') '一覧にない候補の拡大指示を拒否する'

$withoutCandidates = @($packet | ForEach-Object { $_.PSObject.Copy() })
$withoutCandidates[0].targetCandidates = @()
$noneWithoutCandidatesBody = '{"steps":[{"id":"' + $targetId + '","targetCandidateId":"none","zoom":"full","visualConfident":true,"title":"申請","description":"申請を選択します。"}]}'
$noneWithoutCandidatesDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $noneWithoutCandidatesBody) -PacketSteps $withoutCandidates
Add-Result ([string]::IsNullOrWhiteSpace([string]$noneWithoutCandidatesDrafts[0].targetCandidateId) -and
    [string]$noneWithoutCandidatesDrafts[0].zoom -eq 'keep' -and -not $noneWithoutCandidatesDrafts[0].visualConfident) `
    '候補がない手順のnoneと拡大指示を拒否する'

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

$attentionTarget = $stepIds[2]
$attentionSelection = '{"accept":[],"attention":[{"id":"' + $attentionTarget + '","action":"review","reason":"赤枠を特定できませんでした。"}]}'
[void](Set-MbCopilotDraftSelection -Project $project -SelectionJson $attentionSelection)
$attentionStep = Get-MbStepById -Project $project -StepId $attentionTarget
Add-Result ([bool]$attentionStep.review.required -and [string]$attentionStep.review.action -eq 'review') 'Copilotの自信なし候補を要確認として保存する'
Add-Result ([string]$attentionStep.review.reason -eq '赤枠を特定できませんでした。') '要確認の理由を保存する'
$resolveByAccepting = '{"accept":[{"id":"' + $attentionTarget + '","title":"","description":"","note":""}],"attention":[]}'
[void](Set-MbCopilotDraftSelection -Project $project -SelectionJson $resolveByAccepting)
Add-Result (-not [bool]$attentionStep.review.required) '採用した候補は要確認を解除する'

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
    -ClickLabel '登録' -ScreenText '経費申請 / 登録 / 取消' -Narration '登録を押します' `
    -TargetSource 'video-diff' -TargetConfidence 'low' -TargetCandidateId 'video-diff-1' `
    -TargetCandidatesJson '[{"id":"video-diff-1","source":"video-diff","confidence":"low","label":"登録","targetType":"","rect":{"x1":0.70,"y1":0.70,"x2":0.90,"y2":0.82}},{"id":"video-diff-2","source":"video-diff","confidence":"low","label":"取消","targetType":"","rect":{"x1":0.45,"y1":0.70,"x2":0.65,"y2":0.82}}]')
$captured = Get-MbStepById -Project $project -StepId $applyTarget
Add-Result ([int]$captured.capture.videoTimeMs -eq 4500) '録画の時刻を保存する'
Add-Result ([string]$captured.capture.clickLabel -eq '登録') '操作対象を保存する'
Add-Result (@($captured.capture.targetCandidates).Count -eq 2) '操作対象の複数候補を保存する'

# 空の候補IDと拡大指示を組み合わせても、現在の自動赤枠を外さない。
$existingAutoRect = [pscustomobject]@{
    id = 'annotation-' + [guid]::NewGuid().ToString('N'); type = 'rect'; label = 0
    x1 = 0.70; y1 = 0.70; x2 = 0.90; y2 = 0.82
}
[void](Set-MbStepAnnotations -Project $project -StepId $applyTarget `
    -AnnotationsJson (ConvertTo-Json -InputObject @($existingAutoRect) -Depth 5))
$emptyCandidateSelection = '{"accept":[{"id":"' + $applyTarget + '","title":"","description":"","note":"","targetCandidateId":"","zoom":"focus"}]}'
Add-Result ((Set-MbCopilotDraftSelection -Project $project -SelectionJson $emptyCandidateSelection) -eq 0 -and
    @((Get-MbStepById -Project $project -StepId $applyTarget).annotations).Count -eq 1) `
    '空の候補IDでは拡大も赤枠変更も適用しない'

$visualSelection = '{"accept":[{"id":"' + $applyTarget + '","title":"","description":"","note":"","targetCandidateId":"video-diff-2","zoom":"focus"}]}'
[void](Set-MbCopilotDraftSelection -Project $project -SelectionJson $visualSelection)
$captured = Get-MbStepById -Project $project -StepId $applyTarget
Add-Result ([string]$captured.capture.targetCandidateId -eq 'video-diff-2' -and [string]$captured.capture.clickLabel -eq '取消') '採用した候補へ赤枠の根拠を切り替える'
Add-Result ([double]$captured.crop.width -lt 1.0 -and [double]$captured.crop.height -lt 1.0) '選んだ候補の周辺へ拡大する'

# 候補がない手順では、noneとfullが来ても手動cropと操作対象名を保持する。
$noCandidateTarget = $stepIds[2]
$noCandidateStep = Get-MbStepById -Project $project -StepId $noCandidateTarget
[void](Set-MbStepCapture -Project $project -StepId $noCandidateTarget -ClickLabel '利用者の対象名')
[void](Set-MbStepImageEdits -Project $project -StepId $noCandidateTarget -AnnotationsJson '[]' `
    -CropJson '{"x":0.1,"y":0.1,"width":0.6,"height":0.6}')
$noCandidateSelection = '{"accept":[{"id":"' + $noCandidateTarget + '","title":"","description":"","note":"","targetCandidateId":"none","zoom":"full"}]}'
Add-Result ((Set-MbCopilotDraftSelection -Project $project -SelectionJson $noCandidateSelection) -eq 0 -and
    [double]$noCandidateStep.crop.width -eq 0.6 -and [string]$noCandidateStep.capture.clickLabel -eq '利用者の対象名') `
    '候補がない手順のnoneと拡大指示では既存編集を変えない'

# capture を持たない古いプロジェクトを保存して読み直しても壊れないこと。
# v0.27.3以前で作ったマニュアルがこの形になっている。
$legacyRoot = Join-Path ([IO.Path]::GetTempPath()) ('mb-copilot-test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $legacyRoot -Force)
try {
    $legacyPath = Join-Path $legacyRoot 'project.json'
    $legacy = New-MbProject
    $legacyStep = Add-MbStep -Project $legacy -SheetId $legacy.sheets[0].id
    $legacyStep.PSObject.Properties.Remove('capture')
    $legacyStep.PSObject.Properties.Remove('review')
    [void](Save-MbProject -Project $legacy -Path $legacyPath)
    $reloaded = Get-MbProject -Path $legacyPath
    $repaired = Get-MbStepById -Project $reloaded -StepId ([string]$legacyStep.id)
    Add-Result ($null -ne $repaired) '古いプロジェクトを読み直せる'
    Add-Result ($repaired.PSObject.Properties.Name -contains 'capture') '古い手順にも capture が補われる'
    Add-Result ([int]$repaired.capture.videoTimeMs -eq 0) '補われた capture が既定値になる'
    Add-Result ($repaired.PSObject.Properties.Name -contains 'review' -and -not [bool]$repaired.review.required) '古い手順にも要確認の既定値が補われる'
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
# 文章を整える（校正）
# ---------------------------------------------------------------------
$reviewProject = New-MbProject
$reviewProject.title = '経費精算システム操作手順'
$reviewSheet = $reviewProject.sheets[0]
$reviewSheet.name = '申請を出す'

$written = Add-MbStep -Project $reviewProject -SheetId $reviewSheet.id
$written.imageId = 'image-' + ('{0:d32}' -f 1)
$written.title = '申請の作成'
$written.description = '「申請」ボタンをクリックする。'

# 画像が無く文字だけの手順も、校正の対象になる。
$textOnly = Add-MbStep -Project $reviewProject -SheetId $reviewSheet.id
$textOnly.title = '注意'
$textOnly.description = '金額は税込で入力します。'

# 文章がまったく無い手順は対象にしない。
$emptyStep = Add-MbStep -Project $reviewProject -SheetId $reviewSheet.id
$emptyStep.imageId = 'image-' + ('{0:d32}' -f 2)

$reviewSteps = Get-MbCopilotStepList -Project $reviewProject
$reviewPackets = Get-MbCopilotPackets -Steps $reviewSteps -StepsPerPacket 25 -Mode 'review'
$reviewFlat = @($reviewPackets | ForEach-Object { $_ })
Add-Result (@($reviewFlat).Count -eq 2) '文章のある手順だけが校正の対象になる'
Add-Result (@($reviewFlat | Where-Object { $_.id -eq $textOnly.id }).Count -eq 1) '画像の無い手順も校正の対象になる'
Add-Result (@($reviewFlat | Where-Object { $_.id -eq $emptyStep.id }).Count -eq 0) '文章の無い手順は校正しない'

$reviewPrompt = New-MbCopilotReviewPrompt -Project $reviewProject -PacketSteps $reviewFlat -TotalSteps 3 -Marker 'MB_END'
Add-Result ($reviewPrompt.Contains('「申請」ボタンをクリックする。')) '今の文章を依頼文へ入れる'
Add-Result ($reviewPrompt.Contains('敬体')) '敬体の統一を見るよう頼む'
Add-Result ($reviewPrompt.Contains('表記ゆれ')) '表記ゆれを見るよう頼む'
Add-Result ($reviewPrompt.Contains('用語')) '用語の不統一を見るよう頼む'
Add-Result ($reviewPrompt.Contains('意味を変えない')) '意味を変えないよう釘を刺す'
Add-Result ($reviewPrompt.EndsWith((Get-MbCopilotPromptTailAnchor))) '依頼文が目印で終わる'
# 校正では画像を渡さない。依頼文に添付の話が出ないこと。
Add-Result (-not $reviewPrompt.Contains('添付画像')) '校正では画像を渡さない'

$reviewId = [string]$reviewFlat[0].id
$reviewBody = '{"steps":[{"id":"' + $reviewId + '","title":"","description":"「申請」を選択します。","note":"","kind":"敬体","reason":"常体を敬体へ揃えました"}]}'
$reviewDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $reviewBody) -PacketSteps $reviewFlat -Mode 'review'
Add-Result (@($reviewDrafts).Count -eq 1) '直した箇所を受け取る'
Add-Result ([string]$reviewDrafts[0].kind -eq '敬体') '指摘の種類が残る'
Add-Result ([string]$reviewDrafts[0].description -eq '「申請」を選択します。') '直した文章が入る'
Add-Result ([string]$reviewDrafts[0].title -eq '') '直さない項目は空のまま'
Add-Result ([string]$reviewDrafts[0].currentDescription -eq '「申請」ボタンをクリックする。') '直す前の文章も持つ'

# 3項目とも空の指摘は、直すところが無いという意味。確認画面へ出さない。
$noChange = '{"steps":[{"id":"' + $reviewId + '","title":"","description":"","note":"","reason":"問題ありません"}]}'
$noChangeDrafts = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $noChange) -PacketSteps $reviewFlat -Mode 'review'
Add-Result (@($noChangeDrafts).Count -eq 0) '直すところが無い指摘は捨てる'

# 下書きでは空の項目も「変更しない」の意味で残す。校正と扱いが違う。
$draftNoChange = ConvertFrom-MbCopilotStepAnswer -Answer (Get-MbStepAnswerJson -Text $noChange) -PacketSteps $reviewFlat -Mode 'draft'
Add-Result (@($draftNoChange).Count -eq 1) '下書きでは空の項目でも受け取る'

# ---------------------------------------------------------------------
Write-Host ''
if ($errors.Count -eq 0) {
    Write-Host 'Copilot下書きの検査はすべて成功しました。' -ForegroundColor Green
    exit 0
}
Write-Host ("失敗: " + $errors.Count + " 件") -ForegroundColor Red
exit 1
