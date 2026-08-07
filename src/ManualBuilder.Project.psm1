# ManualBuilder project model and atomic JSON persistence.

Set-StrictMode -Version 2.0

function New-MbId {
    param([Parameter(Mandatory = $true)][string]$Prefix)
    return ($Prefix + '-' + [guid]::NewGuid().ToString('N'))
}

function Get-MbUtcTimestamp {
    return [DateTime]::UtcNow.ToString('o')
}

function Get-MbText {
    param(
        [AllowNull()][object]$Value,
        [int]$MaxLength,
        [string]$FieldName,
        [switch]$Required
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n")
    if ($Required -and [string]::IsNullOrWhiteSpace($text)) {
        throw "$FieldName を入力してください。"
    }
    if ($text.Length -gt $MaxLength) {
        throw "$FieldName は $MaxLength 文字以内で入力してください。"
    }
    return $text
}

# 録画から取り込んだ手順に付く情報。手で作った手順では空のまま残る。
# Copilotへ渡す材料であり、出力（Excel・Word・HTML）には出さない。
function New-MbStepCapture {
    return [pscustomobject]@{
        kind        = ''   # 'video-scene' なら録画の場面から取り込んだ手順
        videoTimeMs = 0    # 録画のどの時点か
        clickLabel  = ''   # 操作された場所から読み取れた文字、または押したコントロールの名前
        windowTitle = ''   # 操作していたウィンドウの題名
        screenText  = ''   # 画面に出ていた文字
        narration   = ''   # 操作しながら話した内容（操作記録モードでのみ入る）
        targetType       = ''   # DOM/UIA等が返したコントロール種別
        targetSource     = ''   # DOM / UIA / UIA-CACHE / MSAA / click-point / video-diff
        targetConfidence = ''   # high / medium / low。空は古いデータ
        targetCandidateId = ''  # 現在採用している候補
        targetCandidates = @()  # Copilotが選び直せる、正規化矩形つきの候補
        clickPoint       = $null # 画像内のクリック位置（x, y は0〜1）。場面と赤枠のアンカー
    }
}

function New-MbStepReview {
    return [pscustomobject]@{
        required = $false
        action   = ''       # review / delete。空は確認済み
        reason   = ''       # Copilotが要確認とした理由
    }
}

function New-MbStep {
    $now = Get-MbUtcTimestamp
    return [pscustomobject]@{
        id          = New-MbId -Prefix 'step'
        title       = ''
        description = ''
        note        = ''
        imageId     = $null
        resultImageId = $null
        imageLayout = 'before'
        imageOrder  = 'before-after'
        videoId     = $null
        annotations = @()
        crop        = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
        resultAnnotations = @()
        resultCrop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
        capture     = New-MbStepCapture
        review      = New-MbStepReview
        createdAt   = $now
        updatedAt   = $now
    }
}

function New-MbSheet {
    param([Parameter(Mandatory = $true)][string]$Name)

    $safeName = Get-MbText -Value $Name -MaxLength 50 -FieldName 'シート名' -Required
    return [pscustomobject]@{
        id        = New-MbId -Prefix 'sheet'
        name      = $safeName
        summary   = ''
        steps     = @()
        createdAt = Get-MbUtcTimestamp
        updatedAt = Get-MbUtcTimestamp
    }
}

function New-MbProject {
    $sheet = New-MbSheet -Name '手順1'
    return [pscustomobject]@{
        schemaVersion   = 1
        revision        = 0
        id              = New-MbId -Prefix 'project'
        title           = '新しいマニュアル'
        selectedSheetId = $sheet.id
        sheets          = @($sheet)
        images          = @()
        videos          = @()
        createdAt       = Get-MbUtcTimestamp
        updatedAt       = Get-MbUtcTimestamp
    }
}

function Add-MbPropertyIfMissing {
    param(
        [Parameter(Mandatory = $true)][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()][object]$Value
    )

    if ($InputObject.PSObject.Properties.Name -notcontains $Name) {
        $InputObject | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Repair-MbProject {
    param([Parameter(Mandatory = $true)][object]$Project)

    Add-MbPropertyIfMissing $Project 'schemaVersion' 1
    Add-MbPropertyIfMissing $Project 'revision' 0
    Add-MbPropertyIfMissing $Project 'id' (New-MbId -Prefix 'project')
    Add-MbPropertyIfMissing $Project 'title' '新しいマニュアル'
    Add-MbPropertyIfMissing $Project 'selectedSheetId' $null
    Add-MbPropertyIfMissing $Project 'sheets' @()
    Add-MbPropertyIfMissing $Project 'images' @()
    Add-MbPropertyIfMissing $Project 'videos' @()
    Add-MbPropertyIfMissing $Project 'createdAt' (Get-MbUtcTimestamp)
    Add-MbPropertyIfMissing $Project 'updatedAt' (Get-MbUtcTimestamp)

    $Project.sheets = @($Project.sheets)
    $Project.images = @($Project.images)
    $Project.videos = @($Project.videos)
    if ($Project.sheets.Count -eq 0) {
        $Project.sheets = @(New-MbSheet -Name '手順1')
    }

    foreach ($sheet in $Project.sheets) {
        Add-MbPropertyIfMissing $sheet 'id' (New-MbId -Prefix 'sheet')
        Add-MbPropertyIfMissing $sheet 'name' '名称未設定'
        Add-MbPropertyIfMissing $sheet 'summary' ''
        Add-MbPropertyIfMissing $sheet 'steps' @()
        Add-MbPropertyIfMissing $sheet 'createdAt' (Get-MbUtcTimestamp)
        Add-MbPropertyIfMissing $sheet 'updatedAt' (Get-MbUtcTimestamp)
        $sheet.steps = @($sheet.steps)

        foreach ($step in $sheet.steps) {
            Add-MbPropertyIfMissing $step 'id' (New-MbId -Prefix 'step')
            Add-MbPropertyIfMissing $step 'title' ''
            Add-MbPropertyIfMissing $step 'description' ''
            Add-MbPropertyIfMissing $step 'note' ''
            Add-MbPropertyIfMissing $step 'imageId' $null
            Add-MbPropertyIfMissing $step 'resultImageId' $null
            $defaultImageLayout = if (-not [string]::IsNullOrWhiteSpace([string]$step.resultImageId)) { 'side-by-side' } else { 'before' }
            Add-MbPropertyIfMissing $step 'imageLayout' $defaultImageLayout
            Add-MbPropertyIfMissing $step 'imageOrder' 'before-after'
            Add-MbPropertyIfMissing $step 'videoId' $null
            Add-MbPropertyIfMissing $step 'annotations' @()
            Add-MbPropertyIfMissing $step 'crop' ([pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 })
            Add-MbPropertyIfMissing $step 'resultAnnotations' @()
            Add-MbPropertyIfMissing $step 'resultCrop' ([pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 })
            Add-MbPropertyIfMissing $step 'capture' (New-MbStepCapture)
            Add-MbPropertyIfMissing $step 'review' (New-MbStepReview)
            Add-MbPropertyIfMissing $step 'createdAt' (Get-MbUtcTimestamp)
            Add-MbPropertyIfMissing $step 'updatedAt' (Get-MbUtcTimestamp)
            $step.annotations = @($step.annotations)
            $step.resultAnnotations = @($step.resultAnnotations)
            if ($null -eq $step.crop) {
                $step.crop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
            }
            if ($null -eq $step.resultCrop) {
                $step.resultCrop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
            }
            if ($null -eq $step.capture) {
                $step.capture = New-MbStepCapture
            } else {
                # 古いプロジェクトには項目が欠けていることがある。読み出し側で毎回確かめずに済むよう補う。
                Add-MbPropertyIfMissing $step.capture 'kind' ''
                Add-MbPropertyIfMissing $step.capture 'videoTimeMs' 0
                Add-MbPropertyIfMissing $step.capture 'clickLabel' ''
                Add-MbPropertyIfMissing $step.capture 'windowTitle' ''
                Add-MbPropertyIfMissing $step.capture 'screenText' ''
                Add-MbPropertyIfMissing $step.capture 'narration' ''
                Add-MbPropertyIfMissing $step.capture 'targetType' ''
                Add-MbPropertyIfMissing $step.capture 'targetSource' ''
                Add-MbPropertyIfMissing $step.capture 'targetConfidence' ''
                Add-MbPropertyIfMissing $step.capture 'targetCandidateId' ''
                Add-MbPropertyIfMissing $step.capture 'targetCandidates' @()
                Add-MbPropertyIfMissing $step.capture 'clickPoint' $null
                $step.capture.targetCandidates = @($step.capture.targetCandidates)
            }
            if ($null -eq $step.review) {
                $step.review = New-MbStepReview
            } else {
                Add-MbPropertyIfMissing $step.review 'required' $false
                Add-MbPropertyIfMissing $step.review 'action' ''
                Add-MbPropertyIfMissing $step.review 'reason' ''
            }
        }
    }

    $selectedExists = @($Project.sheets | Where-Object { $_.id -eq $Project.selectedSheetId }).Count -gt 0
    if (-not $selectedExists) {
        $Project.selectedSheetId = $Project.sheets[0].id
    }
    return $Project
}

function Test-MbProject {
    param([Parameter(Mandatory = $true)][object]$Project)

    if ([int]$Project.schemaVersion -ne 1) {
        throw "未対応のプロジェクト形式です: schemaVersion=$($Project.schemaVersion)"
    }
    if ([string]$Project.id -notmatch '^project-[a-f0-9]{32}$') {
        throw 'プロジェクトIDの形式が不正です。'
    }
    [void](Get-MbText -Value $Project.title -MaxLength 100 -FieldName '文書タイトル' -Required)
    if (@($Project.sheets).Count -lt 1) {
        throw 'プロジェクトには1件以上のシートが必要です。'
    }
    if (@($Project.sheets).Count -gt 50) {
        throw 'シートは50件までです。'
    }

    $ids = New-Object 'System.Collections.Generic.HashSet[string]'
    $referencedImageIds = New-Object 'System.Collections.Generic.List[string]'
    $referencedVideoIds = New-Object 'System.Collections.Generic.List[string]'
    foreach ($sheet in @($Project.sheets)) {
        if ([string]$sheet.id -notmatch '^sheet-[a-f0-9]{32}$') { throw 'シートIDの形式が不正です。' }
        if (-not $ids.Add([string]$sheet.id)) { throw 'シートIDが重複しています。' }
        [void](Get-MbText -Value $sheet.name -MaxLength 50 -FieldName 'シート名' -Required)
        if (@($sheet.steps).Count -gt 500) { throw '1シートの手順は500件までです。' }
        foreach ($step in @($sheet.steps)) {
            if ([string]$step.id -notmatch '^step-[a-f0-9]{32}$') { throw '手順IDの形式が不正です。' }
            if (-not $ids.Add([string]$step.id)) { throw '手順IDが重複しています。' }
            [void](Get-MbText -Value $step.title -MaxLength 100 -FieldName '手順タイトル')
            [void](Get-MbText -Value $step.description -MaxLength 4000 -FieldName '説明')
            [void](Get-MbText -Value $step.note -MaxLength 2000 -FieldName '補足')
            if ([string]$step.imageLayout -notin @('before', 'after', 'side-by-side', 'stacked')) {
                throw '画像の見せ方が不正です。'
            }
            if ([string]$step.imageOrder -notin @('before-after', 'after-before')) {
                throw '画像の並び順が不正です。'
            }
            if ([bool]$step.review.required) {
                if ([string]$step.review.action -notin @('review', 'delete')) { throw '要確認の操作が不正です。' }
                [void](Get-MbText -Value $step.review.reason -MaxLength 500 -FieldName '要確認の理由')
            } elseif (-not [string]::IsNullOrWhiteSpace([string]$step.review.action)) {
                throw '確認済み手順に要確認の操作が残っています。'
            }
            foreach ($imageEdits in @(
                [pscustomobject]@{ crop = $step.crop; annotations = @($step.annotations) },
                [pscustomobject]@{ crop = $step.resultCrop; annotations = @($step.resultAnnotations) }
            )) {
                foreach ($cropProperty in @('x', 'y', 'width', 'height')) {
                    if ($imageEdits.crop.PSObject.Properties.Name -notcontains $cropProperty) { throw '切り抜き範囲が不足しています。' }
                    $cropValue = [double]$imageEdits.crop.$cropProperty
                    if ([double]::IsNaN($cropValue) -or [double]::IsInfinity($cropValue)) { throw '切り抜き範囲が不正です。' }
                }
                $cropX = [double]$imageEdits.crop.x
                $cropY = [double]$imageEdits.crop.y
                $cropWidth = [double]$imageEdits.crop.width
                $cropHeight = [double]$imageEdits.crop.height
                if ($cropX -lt 0 -or $cropY -lt 0 -or $cropWidth -lt 0.05 -or $cropHeight -lt 0.05 -or
                    $cropX -gt 0.95 -or $cropY -gt 0.95 -or ($cropX + $cropWidth) -gt 1.000001 -or ($cropY + $cropHeight) -gt 1.000001) {
                    throw '切り抜き範囲が画像の外です。'
                }
                if (@($imageEdits.annotations).Count -gt 100) { throw '1枚の画像の注釈は100件までです。' }
                $annotationIds = New-Object 'System.Collections.Generic.HashSet[string]'
                foreach ($annotation in @($imageEdits.annotations)) {
                    # StrictMode下で生の英語例外にならないよう、必須プロパティの存在を先に確かめる。
                    $annotationProperties = @()
                    if ($null -ne $annotation) { $annotationProperties = @($annotation.PSObject.Properties.Name) }
                    if (($annotationProperties -notcontains 'id') -or ($annotationProperties -notcontains 'type')) {
                        throw '注釈データの形式が不正です。'
                    }
                    if ([string]$annotation.id -notmatch '^annotation-[a-f0-9]{32}$') { throw '注釈IDの形式が不正です。' }
                    if (-not $annotationIds.Add([string]$annotation.id)) { throw '注釈IDが重複しています。' }
                    if ([string]$annotation.type -notin @('rect', 'arrow', 'number', 'blackout')) { throw '注釈種類が不正です。' }
                    foreach ($coordinate in @('x1', 'y1', 'x2', 'y2')) {
                        if ($annotation.PSObject.Properties.Name -notcontains $coordinate) { throw '注釈座標が不足しています。' }
                        $value = [double]$annotation.$coordinate
                        if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0 -or $value -gt 1) {
                            throw '注釈座標が範囲外です。'
                        }
                    }
                    $label = if ($annotationProperties -contains 'label') { [int]$annotation.label } else { 0 }
                    if ([string]$annotation.type -eq 'number' -and ($label -lt 1 -or $label -gt 99)) { throw '番号注釈は1〜99です。' }
                    if ([string]$annotation.type -ne 'number' -and $label -ne 0) { throw '番号以外の注釈ラベルが不正です。' }
                }
            }
            if ($step.imageId) { [void]$referencedImageIds.Add([string]$step.imageId) }
            if ($step.resultImageId) { [void]$referencedImageIds.Add([string]$step.resultImageId) }
            if ($step.videoId) { [void]$referencedVideoIds.Add([string]$step.videoId) }
        }
    }

    if (@($Project.images).Count -gt 25000) { throw '画像は25000件までです。' }
    $imageIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($image in @($Project.images)) {
        if ([string]$image.id -notmatch '^image-[a-f0-9]{32}$') { throw '画像IDの形式が不正です。' }
        if (-not $ids.Add([string]$image.id)) { throw '画像IDが重複しています。' }
        [void]$imageIds.Add([string]$image.id)
        if ([string]$image.fileName -notmatch '^image-[a-f0-9]{32}\.(png|jpg|bmp)$') { throw '画像ファイル名の形式が不正です。' }
        if ([string]$image.sha256 -notmatch '^[A-F0-9]{64}$') { throw '画像ハッシュの形式が不正です。' }
        if ([int]$image.width -lt 1 -or [int]$image.width -gt 12000) { throw '画像幅が範囲外です。' }
        if ([int]$image.height -lt 1 -or [int]$image.height -gt 12000) { throw '画像高さが範囲外です。' }
        if ([long]$image.byteLength -lt 1 -or [long]$image.byteLength -gt (20 * 1024 * 1024)) { throw '画像ファイルサイズが範囲外です。' }
        if ([string]$image.mimeType -notin @('image/png', 'image/jpeg', 'image/bmp')) { throw '画像MIMEタイプが不正です。' }
    }
    foreach ($imageId in $referencedImageIds) {
        if (-not $imageIds.Contains($imageId)) { throw "手順が参照する画像が見つかりません: $imageId" }
    }

    # 動画はExcel出力から再生する。扱いやすい大きさに収めるため、
    # 1本30MB・1マニュアル50本までとする。
    if (@($Project.videos).Count -gt 50) { throw '動画は1マニュアル50本までです。' }
    $videoIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($video in @($Project.videos)) {
        if ([string]$video.id -notmatch '^video-[a-f0-9]{32}$') { throw '動画IDの形式が不正です。' }
        if (-not $ids.Add([string]$video.id)) { throw '動画IDが重複しています。' }
        [void]$videoIds.Add([string]$video.id)
        if ([string]$video.fileName -notmatch '^video-[a-f0-9]{32}\.(mp4|webm)$') { throw '動画ファイル名の形式が不正です。' }
        if ([string]$video.sha256 -notmatch '^[A-F0-9]{64}$') { throw '動画ハッシュの形式が不正です。' }
        if ([long]$video.byteLength -lt 1 -or [long]$video.byteLength -gt (30 * 1024 * 1024)) { throw '動画ファイルサイズが範囲外です。' }
        if ([string]$video.mimeType -notin @('video/mp4', 'video/webm')) { throw '動画MIMEタイプが不正です。' }
        $duration = [double]$video.durationSec
        if ([double]::IsNaN($duration) -or [double]::IsInfinity($duration) -or $duration -lt 0 -or $duration -gt 3600) {
            throw '動画の長さが範囲外です。'
        }
    }
    foreach ($videoId in $referencedVideoIds) {
        if (-not $videoIds.Contains($videoId)) { throw "手順が参照する動画が見つかりません: $videoId" }
    }
}

function Save-MbProject {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $Project = Repair-MbProject -Project $Project
    Test-MbProject -Project $Project
    # revisionとupdatedAtはJSONへ載せるため書込み前に更新するが、
    # 書込みに失敗した場合はメモリとディスクがずれないよう元へ戻す。
    $previousRevision = [int]$Project.revision
    $previousUpdatedAt = [string]$Project.updatedAt
    $directory = Split-Path -Parent $Path
    $tempPath = Join-Path $directory ('.project-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = "$Path.bak"
    $projectFileName = [IO.Path]::GetFileName($Path)
    $recoveryPrefix = ".$projectFileName.recovery-"
    $utf8 = New-Object System.Text.UTF8Encoding($false)

    $Project.revision = $previousRevision + 1
    $Project.updatedAt = Get-MbUtcTimestamp
    try {
        # 保存先の作成やJSON変換も、revisionを進めた後に失敗し得る。
        # 書込みだけでなく保存準備を含む全工程をロールバック対象にする。
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
            [void](New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop)
        }
        $json = $Project | ConvertTo-Json -Depth 12 -ErrorAction Stop
        [IO.File]::WriteAllText($tempPath, $json, $utf8)
        if (Test-Path -LiteralPath $Path) {
            # 固定の.bakをFile.Replaceへ直接渡すと、ウイルス対策などが一瞬開いただけで
            # 本体の保存まで失敗する。置換ごとに固有の退避先を使い、短いロックは再試行する。
            $saved = $false
            $successfulRecoveryPath = ''
            $lastWriteError = $null
            foreach ($delay in @(0, 25, 50, 100, 200, 400)) {
                if ([int]$delay -gt 0) { Start-Sleep -Milliseconds ([int]$delay) }
                $attemptRecoveryPath = Join-Path $directory ($recoveryPrefix +
                    [DateTime]::UtcNow.Ticks.ToString('D19') + '-' + [guid]::NewGuid().ToString('N') + '.bak')
                try {
                    [IO.File]::Replace($tempPath, $Path, $attemptRecoveryPath, $true)
                    $saved = $true
                    $successfulRecoveryPath = $attemptRecoveryPath
                    break
                } catch {
                    $lastWriteError = $_
                    $retryable = $_.Exception -is [IO.IOException] -or $_.Exception -is [UnauthorizedAccessException]
                    if (-not $retryable -or -not (Test-Path -LiteralPath $tempPath -PathType Leaf)) { throw }
                }
            }
            if (-not $saved) { throw $lastWriteError }

            # 本体は保存済みなので、従来名の.bak更新が失敗しても保存失敗にはしない。
            # 固有名の退避を復旧候補として残す。
            $backupPublished = $false
            $staleBackupPath = Join-Path $directory ('.project-stale-' + [guid]::NewGuid().ToString('N') + '.tmp')
            try {
                if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                    [IO.File]::Replace($successfulRecoveryPath, $backupPath, $staleBackupPath, $true)
                } else {
                    [IO.File]::Move($successfulRecoveryPath, $backupPath)
                }
                $backupPublished = $true
            } catch {
                # $successfulRecoveryPathを消さずに残す。
            } finally {
                Remove-Item -LiteralPath $staleBackupPath -Force -ErrorAction SilentlyContinue
            }
            if ($backupPublished) {
                # 正式な.bakを更新できたら、過去の固有名退避は不要。
                foreach ($candidate in @([IO.Directory]::GetFiles($directory))) {
                    $candidateName = [IO.Path]::GetFileName($candidate)
                    if ($candidateName.StartsWith($recoveryPrefix, [StringComparison]::OrdinalIgnoreCase) -and
                        $candidateName.EndsWith('.bak', [StringComparison]::OrdinalIgnoreCase)) {
                        Remove-Item -LiteralPath $candidate -Force -ErrorAction SilentlyContinue
                    }
                }
            } else {
                # 長いロックが続いても退避が増え続けないよう、最新3世代だけ残す。
                $recoveryCandidates = @([IO.Directory]::GetFiles($directory) | Where-Object {
                    $name = [IO.Path]::GetFileName($_)
                    $name.StartsWith($recoveryPrefix, [StringComparison]::OrdinalIgnoreCase) -and
                    $name.EndsWith('.bak', [StringComparison]::OrdinalIgnoreCase)
                } | Sort-Object { [IO.Path]::GetFileName($_) } -Descending)
                foreach ($oldRecovery in @($recoveryCandidates | Select-Object -Skip 3)) {
                    Remove-Item -LiteralPath $oldRecovery -Force -ErrorAction SilentlyContinue
                }
            }
        } else {
            [IO.File]::Move($tempPath, $Path)
        }
    } catch {
        $Project.revision = $previousRevision
        $Project.updatedAt = $previousUpdatedAt
        throw
    } finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf -ErrorAction SilentlyContinue) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
    return $Project
}

function Get-MbProject {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return Save-MbProject -Project (New-MbProject) -Path $Path
    }

    $primaryError = ''
    try {
        $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
        $project = $raw | ConvertFrom-Json
        $project = Repair-MbProject -Project $project
        Test-MbProject -Project $project
        return $project
    } catch {
        $primaryError = $_.Exception.Message
    }

    # 本体が壊れている場合だけ、Save-MbProjectが残した直前のバックアップから復旧を試みる。
    # 検証に通ったものだけを採用し、通らなければ従来どおり安全停止する。
    $backupPath = "$Path.bak"
    $backupCandidates = New-Object System.Collections.ArrayList
    $backupDirectory = Split-Path -Parent $Path
    $projectFileName = [IO.Path]::GetFileName($Path)
    $recoveryPrefix = ".$projectFileName.recovery-"
    $recoveryFiles = New-Object System.Collections.ArrayList
    if (Test-Path -LiteralPath $backupDirectory -PathType Container) {
        foreach ($candidatePath in @([IO.Directory]::GetFiles($backupDirectory))) {
            $candidateName = [IO.Path]::GetFileName($candidatePath)
            if ($candidateName.StartsWith($recoveryPrefix, [StringComparison]::OrdinalIgnoreCase) -and
                $candidateName.EndsWith('.bak', [StringComparison]::OrdinalIgnoreCase)) {
                [void]$recoveryFiles.Add((Get-Item -LiteralPath $candidatePath))
            }
        }
    }
    foreach ($recoveryFile in @($recoveryFiles | Sort-Object Name -Descending)) {
        [void]$backupCandidates.Add($recoveryFile)
    }
    if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
        [void]$backupCandidates.Add((Get-Item -LiteralPath $backupPath))
    }
    foreach ($backupFile in @($backupCandidates)) {
        try {
            $backupRaw = [IO.File]::ReadAllText([string]$backupFile.FullName, [Text.Encoding]::UTF8)
            $backupProject = $backupRaw | ConvertFrom-Json
            $backupProject = Repair-MbProject -Project $backupProject
            Test-MbProject -Project $backupProject
            try {
                # 壊れた本体をそのままReplaceのバックアップへ昇格させない。
                # いったん同じフォルダーへ退避し、新規ファイルとして復元する。
                $damagedPath = Join-Path $backupDirectory ('.project-damaged-' + [guid]::NewGuid().ToString('N') + '.tmp')
                [IO.File]::Move($Path, $damagedPath)
                try {
                    $restoredProject = Save-MbProject -Project $backupProject -Path $Path
                    Remove-Item -LiteralPath $damagedPath -Force -ErrorAction SilentlyContinue
                    return $restoredProject
                } catch {
                    if (-not (Test-Path -LiteralPath $Path -PathType Leaf) -and
                        (Test-Path -LiteralPath $damagedPath -PathType Leaf)) {
                        [IO.File]::Move($damagedPath, $Path)
                    }
                    throw
                }
            } catch {
                return $backupProject
            }
        } catch { }
    }

    throw "プロジェクトを読み込めません。破損の可能性があります。$primaryError"
}

function Get-MbSelectedSheet {
    param([Parameter(Mandatory = $true)][object]$Project)
    return @($Project.sheets | Where-Object { $_.id -eq $Project.selectedSheetId })[0]
}

function Set-MbProjectTitle {
    param([object]$Project, [AllowEmptyString()][string]$Title)
    $Project.title = Get-MbText -Value $Title -MaxLength 100 -FieldName '文書タイトル' -Required
}

function Add-MbSheet {
    param([Parameter(Mandatory = $true)][object]$Project)

    if (@($Project.sheets).Count -ge 50) { throw 'シートは50件までです。' }
    $base = 'シート ' + (@($Project.sheets).Count + 1)
    $name = $base
    $suffix = 2
    $names = @($Project.sheets | ForEach-Object { $_.name })
    while ($names -contains $name) {
        $name = "$base ($suffix)"
        $suffix++
    }
    $sheet = New-MbSheet -Name $name
    $Project.sheets = @($Project.sheets) + @($sheet)
    $Project.selectedSheetId = $sheet.id
    return $sheet
}

function Copy-MbSheet {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$SheetId
    )

    if (@($Project.sheets).Count -ge 50) { throw 'シートは50件までです。' }
    $source = @($Project.sheets | Where-Object { [string]$_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $source) { throw '対象シートが見つかりません。' }

    $existingNames = @($Project.sheets | ForEach-Object { [string]$_.name })
    $copyNumber = 1
    do {
        $suffix = if ($copyNumber -eq 1) { ' のコピー' } else { " のコピー ($copyNumber)" }
        $baseLength = [Math]::Max(1, 50 - $suffix.Length)
        $baseName = [string]$source.name
        if ($baseName.Length -gt $baseLength) { $baseName = $baseName.Substring(0, $baseLength) }
        $candidateName = $baseName.TrimEnd() + $suffix
        $copyNumber++
    } while ($existingNames -contains $candidateName)

    # JSONを介して、注釈・切り抜き・記録情報を含む入れ子の値を独立したオブジェクトにする。
    # 画像・動画IDは同じ素材を参照し、シートと手順のIDだけを新しく発行する。
    $copy = $source | ConvertTo-Json -Depth 100 -Compress | ConvertFrom-Json
    $now = Get-MbUtcTimestamp
    $copy.id = New-MbId -Prefix 'sheet'
    $copy.name = $candidateName
    $copy.createdAt = $now
    $copy.updatedAt = $now
    foreach ($step in @($copy.steps)) {
        $step.id = New-MbId -Prefix 'step'
        $step.createdAt = $now
        $step.updatedAt = $now
    }

    $sourceIndex = [array]::IndexOf(@($Project.sheets), $source)
    $sheets = New-Object System.Collections.ArrayList
    foreach ($sheet in @($Project.sheets)) { [void]$sheets.Add($sheet) }
    $sheets.Insert($sourceIndex + 1, $copy)
    $Project.sheets = @($sheets)
    $Project.selectedSheetId = $copy.id
    return $copy
}

function Select-MbSheet {
    param([object]$Project, [string]$SheetId)
    $sheet = @($Project.sheets | Where-Object { $_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $sheet) { throw '対象シートが見つかりません。' }
    $Project.selectedSheetId = $sheet.id
}

function Rename-MbSheet {
    param([object]$Project, [string]$SheetId, [AllowEmptyString()][string]$Name)
    $sheet = @($Project.sheets | Where-Object { $_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $sheet) { throw '対象シートが見つかりません。' }
    $sheet.name = Get-MbText -Value $Name -MaxLength 50 -FieldName 'シート名' -Required
    $sheet.updatedAt = Get-MbUtcTimestamp
}

function Remove-MbSheet {
    param([object]$Project, [string]$SheetId)
    if (@($Project.sheets).Count -le 1) { throw '最後のシートは削除できません。' }
    $before = @($Project.sheets).Count
    $Project.sheets = @($Project.sheets | Where-Object { $_.id -ne $SheetId })
    if (@($Project.sheets).Count -eq $before) { throw '対象シートが見つかりません。' }
    if ($Project.selectedSheetId -eq $SheetId) {
        $Project.selectedSheetId = $Project.sheets[0].id
    }
}

function Restore-MbSheet {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][int]$Index
    )

    $sheetId = [string]$Sheet.id
    if ([string]::IsNullOrWhiteSpace($sheetId)) { throw '復元するシートが不正です。' }
    if (@($Project.sheets | Where-Object { [string]$_.id -eq $sheetId }).Count -gt 0) {
        throw '同じシートがすでにあります。'
    }
    if (@($Project.sheets).Count -ge 50) { throw 'シートは50件までです。' }

    $sheets = New-Object System.Collections.ArrayList
    foreach ($current in @($Project.sheets)) { [void]$sheets.Add($current) }
    $insertAt = [Math]::Max(0, [Math]::Min($Index, $sheets.Count))
    $sheets.Insert($insertAt, $Sheet)
    $Project.sheets = @($sheets)
    $Project.selectedSheetId = $sheetId
    return $Sheet
}

function Set-MbSheetOrder {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string[]]$SheetIds
    )

    $currentSheets = @($Project.sheets)
    $requestedIds = @($SheetIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($requestedIds.Count -ne $currentSheets.Count) { throw 'シートの件数が一致しません。' }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $ordered = New-Object System.Collections.ArrayList
    foreach ($sheetId in $requestedIds) {
        if (-not $seen.Add([string]$sheetId)) { throw 'シートIDが重複しています。' }
        $sheet = @($currentSheets | Where-Object { $_.id -eq $sheetId }) | Select-Object -First 1
        if (-not $sheet) { throw '対象シートが見つかりません。' }
        [void]$ordered.Add($sheet)
    }
    $Project.sheets = @($ordered)
}

function Add-MbStep {
    param(
        [object]$Project,
        [string]$SheetId,
        [AllowEmptyString()][string]$AfterStepId = ''
    )
    $sheet = @($Project.sheets | Where-Object { $_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $sheet) { throw '対象シートが見つかりません。' }
    if (@($sheet.steps).Count -ge 500) { throw '1シートの手順は500件までです。' }
    $step = New-MbStep
    $currentSteps = @($sheet.steps)
    if ([string]::IsNullOrWhiteSpace($AfterStepId)) {
        $sheet.steps = $currentSteps + @($step)
    } else {
        $afterIndex = -1
        for ($i = 0; $i -lt $currentSteps.Count; $i++) {
            if ([string]$currentSteps[$i].id -eq $AfterStepId) { $afterIndex = $i; break }
        }
        if ($afterIndex -lt 0) { throw '追加位置の手順が見つかりません。' }
        $before = if ($afterIndex -ge 0) { @($currentSteps | Select-Object -First ($afterIndex + 1)) } else { @() }
        $after = @($currentSteps | Select-Object -Skip ($afterIndex + 1))
        $sheet.steps = @($before) + @($step) + @($after)
    }
    $sheet.updatedAt = Get-MbUtcTimestamp
    return $step
}

function Update-MbStep {
    param(
        [object]$Project,
        [string]$StepId,
        [AllowEmptyString()][string]$Title,
        [AllowEmptyString()][string]$Description,
        [AllowEmptyString()][string]$Note
    )

    $target = $null
    foreach ($sheet in @($Project.sheets)) {
        $target = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($target) { break }
    }
    if (-not $target) { throw '対象手順が見つかりません。' }
    $target.title = Get-MbText -Value $Title -MaxLength 100 -FieldName '手順タイトル'
    $target.description = Get-MbText -Value $Description -MaxLength 4000 -FieldName '説明'
    $target.note = Get-MbText -Value $Note -MaxLength 2000 -FieldName '補足'
    $target.updatedAt = Get-MbUtcTimestamp
}

# 注釈の矩形として使える形かどうか。0〜1の正規化座標で、潰れていないこと。
# 録画からの取り込みと操作記録の両方が使うため、ここに1つだけ置く。
function Test-MbNormalizedRect {
    param([AllowNull()]$Rect)
    if ($null -eq $Rect) { return $false }
    foreach ($name in @('x1', 'y1', 'x2', 'y2')) {
        if ($Rect.PSObject.Properties.Name -notcontains $name) { return $false }
        # 画面から来た値は数値とは限らない。変換に失敗したら不正として扱う。
        $value = 0.0
        try { $value = [double]$Rect.$name } catch { return $false }
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { return $false }
        if ($value -lt 0 -or $value -gt 1) { return $false }
    }
    if (([double]$Rect.x2 - [double]$Rect.x1) -lt 0.004) { return $false }
    if (([double]$Rect.y2 - [double]$Rect.y1) -lt 0.004) { return $false }
    return $true
}

function Get-MbStepById {
    param([Parameter(Mandatory = $true)][object]$Project, [Parameter(Mandatory = $true)][string]$StepId)
    foreach ($sheet in @($Project.sheets)) {
        $found = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($found) { return $found }
    }
    return $null
}

function Set-MbStepImageLayout {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [ValidateSet('before', 'after', 'side-by-side', 'stacked')][string]$Layout,
        [ValidateSet('before-after', 'after-before')][string]$Order = 'before-after'
    )
    $step = Get-MbStepById -Project $Project -StepId $StepId
    if (-not $step) { throw '対象手順が見つかりません。' }
    $hasResult = -not [string]::IsNullOrWhiteSpace([string]$step.resultImageId)
    if (-not $hasResult -and $Layout -ne 'before') {
        throw '2枚目の画像を追加してから見せ方を選んでください。'
    }
    $step.imageLayout = $Layout
    $step.imageOrder = $Order
    $step.updatedAt = Get-MbUtcTimestamp
    return $step
}

# 録画から取り込んだ情報を手順へ書き込む。文章は触らない。
function Set-MbStepCapture {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [string]$Kind = 'video-scene',
        [int]$VideoTimeMs = 0,
        [AllowEmptyString()][string]$ClickLabel = '',
        [AllowEmptyString()][string]$WindowTitle = '',
        [AllowEmptyString()][string]$ScreenText = '',
        [AllowEmptyString()][string]$Narration = '',
        [AllowEmptyString()][string]$TargetType = '',
        [AllowEmptyString()][string]$TargetSource = '',
        [ValidateSet('', 'high', 'medium', 'low')][string]$TargetConfidence = '',
        [AllowEmptyString()][string]$TargetCandidateId = '',
        [AllowEmptyString()][string]$TargetCandidatesJson = '',
        [AllowEmptyString()][string]$ClickPointJson = ''
    )

    $target = Get-MbStepById -Project $Project -StepId $StepId
    if (-not $target) { throw '対象手順が見つかりません。' }
    if ($target.PSObject.Properties.Name -notcontains 'capture' -or $null -eq $target.capture) {
        $target | Add-Member -NotePropertyName 'capture' -NotePropertyValue (New-MbStepCapture) -Force
    }
    $target.capture.kind = Get-MbText -Value $Kind -MaxLength 40 -FieldName '取り込み種別'
    $target.capture.videoTimeMs = [Math]::Max(0, $VideoTimeMs)
    $target.capture.clickLabel = Get-MbText -Value $ClickLabel -MaxLength 200 -FieldName '操作対象'
    $target.capture.windowTitle = Get-MbText -Value $WindowTitle -MaxLength 300 -FieldName 'ウィンドウの題名'
    $target.capture.screenText = Get-MbText -Value $ScreenText -MaxLength 4000 -FieldName '画面の文字'
    $target.capture.narration = Get-MbText -Value $Narration -MaxLength 2000 -FieldName '話した内容'
    $target.capture.targetType = Get-MbText -Value $TargetType -MaxLength 100 -FieldName '操作対象の種類'
    $target.capture.targetSource = Get-MbText -Value $TargetSource -MaxLength 40 -FieldName '操作対象の取得元'
    $target.capture.targetConfidence = $TargetConfidence
    $target.capture.targetCandidateId = Get-MbText -Value $TargetCandidateId -MaxLength 80 -FieldName '操作対象候補'
    if (-not [string]::IsNullOrWhiteSpace($ClickPointJson)) {
        $parsedPoint = $null
        try { $parsedPoint = $ClickPointJson | ConvertFrom-Json } catch { throw 'クリック位置を読み取れません。' }
        if ($null -eq $parsedPoint -or $parsedPoint.PSObject.Properties.Name -notcontains 'x' -or
            $parsedPoint.PSObject.Properties.Name -notcontains 'y') { throw 'クリック位置が正しくありません。' }
        $pointX = [double]$parsedPoint.x
        $pointY = [double]$parsedPoint.y
        if ([double]::IsNaN($pointX) -or [double]::IsInfinity($pointX) -or
            [double]::IsNaN($pointY) -or [double]::IsInfinity($pointY) -or
            $pointX -lt 0 -or $pointX -gt 1 -or $pointY -lt 0 -or $pointY -gt 1) {
            throw 'クリック位置が画像の範囲外です。'
        }
        $target.capture.clickPoint = [pscustomobject]@{
            x = [Math]::Round($pointX, 6); y = [Math]::Round($pointY, 6)
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($TargetCandidatesJson)) {
        $parsedCandidates = $null
        try { $parsedCandidates = $TargetCandidatesJson | ConvertFrom-Json } catch { throw '操作対象候補を読み取れません。' }
        $safeCandidates = New-Object System.Collections.ArrayList
        $candidateIds = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($candidate in @(@($parsedCandidates) | Select-Object -First 8)) {
            if ($null -eq $candidate -or $candidate.PSObject.Properties.Name -notcontains 'id' -or
                $candidate.PSObject.Properties.Name -notcontains 'rect') { continue }
            if (-not (Test-MbNormalizedRect -Rect $candidate.rect)) { continue }
            $candidateId = Get-MbText -Value $candidate.id -MaxLength 80 -FieldName '操作対象候補ID'
            if ([string]::IsNullOrWhiteSpace($candidateId) -or -not $candidateIds.Add($candidateId)) { continue }
            [void]$safeCandidates.Add([pscustomobject]@{
                id = $candidateId
                source = Get-MbText -Value $(if ($candidate.PSObject.Properties.Name -contains 'source') { $candidate.source } else { '' }) -MaxLength 40 -FieldName '候補の取得元'
                confidence = Get-MbText -Value $(if ($candidate.PSObject.Properties.Name -contains 'confidence') { $candidate.confidence } else { '' }) -MaxLength 20 -FieldName '候補の信頼度'
                label = Get-MbText -Value $(if ($candidate.PSObject.Properties.Name -contains 'label') { $candidate.label } else { '' }) -MaxLength 200 -FieldName '候補名'
                targetType = Get-MbText -Value $(if ($candidate.PSObject.Properties.Name -contains 'targetType') { $candidate.targetType } else { '' }) -MaxLength 100 -FieldName '候補の種類'
                rect = [pscustomobject]@{
                    x1 = [Math]::Round([double]$candidate.rect.x1, 6); y1 = [Math]::Round([double]$candidate.rect.y1, 6)
                    x2 = [Math]::Round([double]$candidate.rect.x2, 6); y2 = [Math]::Round([double]$candidate.rect.y2, 6)
                }
            })
        }
        $target.capture.targetCandidates = @($safeCandidates)
        $savedCandidateIds = @($target.capture.targetCandidates | ForEach-Object { [string]$_.id })
        if ($savedCandidateIds -notcontains [string]$target.capture.targetCandidateId) {
            $target.capture.targetCandidateId = if ($savedCandidateIds.Count -gt 0) { [string]$savedCandidateIds[0] } else { '' }
        }
    }
    $target.updatedAt = Get-MbUtcTimestamp
    return $target
}

# Copilotの下書きのうち、利用者が採用した手順だけを書き込む。
# 空文字の項目は「変更しない」を意味する。誤って既存の文章を消さないため。
function Set-MbStepDraft {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [AllowEmptyString()][string]$Title = '',
        [AllowEmptyString()][string]$Description = '',
        [AllowEmptyString()][string]$Note = ''
    )

    $target = Get-MbStepById -Project $Project -StepId $StepId
    if (-not $target) { throw '対象手順が見つかりません。' }
    $changed = $false
    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        $target.title = Get-MbText -Value $Title -MaxLength 100 -FieldName '手順タイトル'
        $changed = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($Description)) {
        $target.description = Get-MbText -Value $Description -MaxLength 4000 -FieldName '説明'
        $changed = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($Note)) {
        $target.note = Get-MbText -Value $Note -MaxLength 2000 -FieldName '補足'
        $changed = $true
    }
    if ($changed) { $target.updatedAt = Get-MbUtcTimestamp }
    return $changed
}

function Set-MbStepOrder {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [Parameter(Mandatory = $true)][string[]]$StepIds
    )

    $sheet = @($Project.sheets | Where-Object { $_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $sheet) { throw '対象シートが見つかりません。' }
    $currentSteps = @($sheet.steps)
    $requestedIds = @($StepIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($requestedIds.Count -ne $currentSteps.Count) { throw '手順の件数が一致しません。' }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $ordered = New-Object System.Collections.ArrayList
    foreach ($stepId in $requestedIds) {
        if (-not $seen.Add([string]$stepId)) { throw '手順IDが重複しています。' }
        $step = @($currentSteps | Where-Object { $_.id -eq $stepId }) | Select-Object -First 1
        if (-not $step) { throw '対象手順が見つかりません。' }
        [void]$ordered.Add($step)
    }
    $sheet.steps = @($ordered)
    $sheet.updatedAt = Get-MbUtcTimestamp
}

function Move-MbStepToSheet {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [Parameter(Mandatory = $true)][string]$TargetSheetId
    )

    $sourceSheet = $null
    $step = $null
    foreach ($sheet in @($Project.sheets)) {
        $candidate = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($candidate) {
            $sourceSheet = $sheet
            $step = $candidate
            break
        }
    }
    if (-not $step) { throw '対象手順が見つかりません。' }

    $targetSheet = @($Project.sheets | Where-Object { $_.id -eq $TargetSheetId }) | Select-Object -First 1
    if (-not $targetSheet) { throw '移動先シートが見つかりません。' }
    if ($sourceSheet.id -eq $targetSheet.id) {
        $Project.selectedSheetId = $targetSheet.id
        return $step
    }
    if (@($targetSheet.steps).Count -ge 500) { throw '移動先シートの手順は500件までです。' }

    $sourceSheet.steps = @($sourceSheet.steps | Where-Object { $_.id -ne $StepId })
    $targetSheet.steps = @($targetSheet.steps) + @($step)
    $now = Get-MbUtcTimestamp
    $sourceSheet.updatedAt = $now
    $targetSheet.updatedAt = $now
    $step.updatedAt = $now
    $Project.selectedSheetId = $targetSheet.id
    return $step
}

function Move-MbStepsToSheet {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string[]]$StepIds,
        [Parameter(Mandatory = $true)][string]$TargetSheetId
    )

    $requestedIds = @($StepIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($requestedIds.Count -lt 1) { throw '移動する手順を選んでください。' }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $moving = New-Object System.Collections.ArrayList
    $sources = @{}
    foreach ($stepId in $requestedIds) {
        if (-not $seen.Add([string]$stepId)) { throw '手順IDが重複しています。' }
        $source = $null
        $step = $null
        foreach ($sheet in @($Project.sheets)) {
            $step = @($sheet.steps | Where-Object { $_.id -eq $stepId }) | Select-Object -First 1
            if ($step) { $source = $sheet; break }
        }
        if (-not $step) { throw '対象手順が見つかりません。' }
        [void]$moving.Add($step)
        $sources[[string]$stepId] = $source
    }

    $targetSheet = @($Project.sheets | Where-Object { $_.id -eq $TargetSheetId }) | Select-Object -First 1
    if (-not $targetSheet) { throw '移動先シートが見つかりません。' }
    $newCount = 0
    foreach ($movingStep in @($moving)) {
        $sourceForStep = $sources[[string]$movingStep.id]
        if ([string]$sourceForStep.id -ne [string]$targetSheet.id) { $newCount++ }
    }
    if (@($targetSheet.steps).Count + $newCount -gt 500) { throw '移動先シートの手順は500件までです。' }

    $now = Get-MbUtcTimestamp
    foreach ($sheet in @($Project.sheets)) {
        $before = @($sheet.steps).Count
        $sheet.steps = @($sheet.steps | Where-Object { -not $seen.Contains([string]$_.id) })
        if (@($sheet.steps).Count -ne $before) { $sheet.updatedAt = $now }
    }
    $targetSheet.steps = @($targetSheet.steps) + @($moving)
    $targetSheet.updatedAt = $now
    foreach ($step in @($moving)) { $step.updatedAt = $now }
    $Project.selectedSheetId = $targetSheet.id
    return @($moving)
}

function Set-MbStepAnnotations {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [AllowEmptyString()][string]$AnnotationsJson,
        [ValidateSet('before', 'result')][string]$Target = 'before'
    )

    if ($AnnotationsJson.Length -gt 100000) { throw '注釈データが大きすぎます。' }
    $targetStep = $null
    foreach ($sheet in @($Project.sheets)) {
        $targetStep = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($targetStep) { break }
    }
    if (-not $targetStep) { throw '対象手順が見つかりません。' }
    $targetImageId = if ($Target -eq 'result') { [string]$targetStep.resultImageId } else { [string]$targetStep.imageId }
    if ([string]::IsNullOrWhiteSpace($targetImageId)) { throw '画像のない手順には注釈を保存できません。' }

    try {
        $parsed = if ([string]::IsNullOrWhiteSpace($AnnotationsJson)) { @() } else { @($AnnotationsJson | ConvertFrom-Json) }
    } catch {
        throw '注釈データを読み込めません。'
    }
    if ($parsed.Count -gt 100) { throw '1手順の注釈は100件までです。' }

    $normalized = New-Object System.Collections.ArrayList
    $annotationIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($annotation in $parsed) {
        # 受信JSONは配列要素がオブジェクトとは限らないため、プロパティ参照前に形式を確かめる。
        $annotationProperties = @()
        if ($null -ne $annotation) { $annotationProperties = @($annotation.PSObject.Properties.Name) }
        if (($annotationProperties -notcontains 'id') -or ($annotationProperties -notcontains 'type')) {
            throw '注釈データの形式が不正です。'
        }
        if ([string]$annotation.id -notmatch '^annotation-[a-f0-9]{32}$') { throw '注釈IDの形式が不正です。' }
        if (-not $annotationIds.Add([string]$annotation.id)) { throw '注釈IDが重複しています。' }
        $type = [string]$annotation.type
        if ($type -notin @('rect', 'arrow', 'number', 'blackout')) { throw '注釈種類が不正です。' }
        $coordinates = @{}
        foreach ($coordinate in @('x1', 'y1', 'x2', 'y2')) {
            if ($annotation.PSObject.Properties.Name -notcontains $coordinate) { throw '注釈座標が不足しています。' }
            $value = [double]$annotation.$coordinate
            if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0 -or $value -gt 1) {
                throw '注釈座標が範囲外です。'
            }
            $coordinates[$coordinate] = [Math]::Round($value, 6)
        }
        $label = if ($annotation.PSObject.Properties.Name -contains 'label') { [int]$annotation.label } else { 0 }
        if ($type -eq 'number' -and ($label -lt 1 -or $label -gt 99)) { throw '番号注釈は1〜99です。' }
        if ($type -ne 'number') { $label = 0 }
        [void]$normalized.Add([pscustomobject]@{
            id    = [string]$annotation.id
            type  = $type
            x1    = $coordinates.x1
            y1    = $coordinates.y1
            x2    = $coordinates.x2
            y2    = $coordinates.y2
            label = $label
        })
    }
    if ($Target -eq 'result') { $targetStep.resultAnnotations = @($normalized) }
    else { $targetStep.annotations = @($normalized) }
    $targetStep.updatedAt = Get-MbUtcTimestamp
}

function Set-MbStepImageEdits {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [AllowEmptyString()][string]$AnnotationsJson,
        [AllowEmptyString()][string]$CropJson,
        [ValidateSet('before', 'result')][string]$Target = 'before'
    )

    if ($CropJson.Length -gt 1000) { throw '切り抜きデータが大きすぎます。' }
    try {
        $crop = if ([string]::IsNullOrWhiteSpace($CropJson)) {
            [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
        } else {
            $CropJson | ConvertFrom-Json
        }
    } catch {
        throw '切り抜きデータを読み込めません。'
    }

    $normalizedCrop = @{}
    foreach ($property in @('x', 'y', 'width', 'height')) {
        if ($crop.PSObject.Properties.Name -notcontains $property) { throw '切り抜き範囲が不足しています。' }
        $value = [double]$crop.$property
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { throw '切り抜き範囲が不正です。' }
        $normalizedCrop[$property] = [Math]::Round($value, 6)
    }
    if ($normalizedCrop.x -lt 0 -or $normalizedCrop.y -lt 0 -or
        $normalizedCrop.width -lt 0.05 -or $normalizedCrop.height -lt 0.05 -or
        ($normalizedCrop.x + $normalizedCrop.width) -gt 1.000001 -or
        ($normalizedCrop.y + $normalizedCrop.height) -gt 1.000001) {
        throw '切り抜き範囲が画像の外です。'
    }

    Set-MbStepAnnotations -Project $Project -StepId $StepId -AnnotationsJson $AnnotationsJson -Target $Target
    $targetStep = $null
    foreach ($sheet in @($Project.sheets)) {
        $targetStep = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($targetStep) { break }
    }
    if (-not $targetStep) { throw '対象手順が見つかりません。' }
    $normalizedValue = [pscustomobject]@{
        x      = [double]$normalizedCrop.x
        y      = [double]$normalizedCrop.y
        width  = [double]$normalizedCrop.width
        height = [double]$normalizedCrop.height
    }
    if ($Target -eq 'result') { $targetStep.resultCrop = $normalizedValue }
    else { $targetStep.crop = $normalizedValue }
    $targetStep.updatedAt = Get-MbUtcTimestamp
}

function Remove-MbStep {
    param([object]$Project, [string]$StepId)
    foreach ($sheet in @($Project.sheets)) {
        $before = @($sheet.steps).Count
        $sheet.steps = @($sheet.steps | Where-Object { $_.id -ne $StepId })
        if (@($sheet.steps).Count -lt $before) {
            $sheet.updatedAt = Get-MbUtcTimestamp
            return
        }
    }
    throw '対象手順が見つかりません。'
}

function Set-MbStepReview {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [ValidateSet('', 'review', 'delete')][string]$Action = '',
        [AllowEmptyString()][string]$Reason = ''
    )

    $target = Get-MbStepById -Project $Project -StepId $StepId
    if (-not $target) { throw '対象手順が見つかりません。' }
    $safeReason = Get-MbText -Value $Reason -MaxLength 500 -FieldName '要確認の理由'
    $target.review = [pscustomobject]@{
        required = -not [string]::IsNullOrWhiteSpace($Action)
        action   = $Action
        reason   = if ([string]::IsNullOrWhiteSpace($Action)) { '' } else { $safeReason }
    }
    $target.updatedAt = Get-MbUtcTimestamp
    return $target
}

function Remove-MbSteps {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string[]]$StepIds
    )

    $requestedIds = @($StepIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($requestedIds.Count -lt 1) { throw '削除する手順を選んでください。' }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($stepId in $requestedIds) {
        if (-not $seen.Add([string]$stepId)) { throw '手順IDが重複しています。' }
        if (-not (Get-MbStepById -Project $Project -StepId $stepId)) { throw '対象手順が見つかりません。' }
    }

    $now = Get-MbUtcTimestamp
    foreach ($sheet in @($Project.sheets)) {
        $before = @($sheet.steps).Count
        $sheet.steps = @($sheet.steps | Where-Object { -not $seen.Contains([string]$_.id) })
        if (@($sheet.steps).Count -ne $before) { $sheet.updatedAt = $now }
    }
}

function Restore-MbSteps {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][object[]]$Items
    )

    $restoreItems = @($Items)
    if ($restoreItems.Count -lt 1) { throw '復元する手順がありません。' }
    $existingIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) { [void]$existingIds.Add([string]$step.id) }
    }
    $restoringIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $countsBySheet = @{}
    foreach ($item in $restoreItems) {
        $sheetId = [string]$item.sheetId
        $stepId = [string]$item.step.id
        if ([string]::IsNullOrWhiteSpace($sheetId) -or [string]::IsNullOrWhiteSpace($stepId)) {
            throw '復元する手順が不正です。'
        }
        $sheet = @($Project.sheets | Where-Object { [string]$_.id -eq $sheetId }) | Select-Object -First 1
        if (-not $sheet) { throw '復元先のシートが見つかりません。' }
        if ($existingIds.Contains($stepId) -or -not $restoringIds.Add($stepId)) {
            throw '同じ手順がすでにあります。'
        }
        if (-not $countsBySheet.ContainsKey($sheetId)) { $countsBySheet[$sheetId] = 0 }
        $countsBySheet[$sheetId]++
    }
    foreach ($sheetId in @($countsBySheet.Keys)) {
        $sheet = @($Project.sheets | Where-Object { [string]$_.id -eq [string]$sheetId }) | Select-Object -First 1
        if (@($sheet.steps).Count + [int]$countsBySheet[$sheetId] -gt 500) { throw '1シートの手順は500件までです。' }
    }

    $now = Get-MbUtcTimestamp
    foreach ($sheetId in @($countsBySheet.Keys)) {
        $sheet = @($Project.sheets | Where-Object { [string]$_.id -eq [string]$sheetId }) | Select-Object -First 1
        $steps = New-Object System.Collections.ArrayList
        foreach ($step in @($sheet.steps)) { [void]$steps.Add($step) }
        $forSheet = @($restoreItems | Where-Object { [string]$_.sheetId -eq [string]$sheetId } | Sort-Object { [int]$_.index })
        foreach ($item in $forSheet) {
            $insertAt = [Math]::Max(0, [Math]::Min([int]$item.index, $steps.Count))
            $steps.Insert($insertAt, $item.step)
        }
        $sheet.steps = @($steps)
        $sheet.updatedAt = $now
    }
    return @($restoreItems | ForEach-Object { $_.step })
}

Export-ModuleMember -Function @(
    'New-MbProject',
    'New-MbSheet',
    'New-MbStep',
    'New-MbStepCapture',
    'New-MbStepReview',
    'Get-MbStepById',
    'Set-MbStepImageLayout',
    'Test-MbNormalizedRect',
    'Set-MbStepCapture',
    'Set-MbStepDraft',
    'Get-MbProject',
    'Save-MbProject',
    'Test-MbProject',
    'Get-MbSelectedSheet',
    'Set-MbProjectTitle',
    'Add-MbSheet',
    'Copy-MbSheet',
    'Select-MbSheet',
    'Rename-MbSheet',
    'Remove-MbSheet',
    'Restore-MbSheet',
    'Set-MbSheetOrder',
    'Add-MbStep',
    'Update-MbStep',
    'Set-MbStepOrder',
    'Move-MbStepToSheet',
    'Move-MbStepsToSheet',
    'Set-MbStepAnnotations',
    'Set-MbStepImageEdits',
    'Set-MbStepReview',
    'Remove-MbStep',
    'Remove-MbSteps',
    'Restore-MbSteps'
)
