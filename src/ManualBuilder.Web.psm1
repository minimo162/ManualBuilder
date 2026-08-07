# ManualBuilder server-side HTML fragments.

Set-StrictMode -Version 2.0

function ConvertTo-MbHtml {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-MbSaveStatusHtml {
    param(
        [string]$Message = '保存済み',
        [ValidateSet('saved', 'saving', 'error')][string]$State = 'saved'
    )

    $icon = switch ($State) {
        'saved' { '●' }
        'saving' { '◌' }
        'error' { '!' }
    }
    $time = if ($State -eq 'saved') { ' ' + (Get-Date).ToString('HH:mm') } else { '' }
    return '<span id="save-status" class="save-status save-status--' + $State + '" role="status"><span aria-hidden="true">' + $icon + '</span> ' + (ConvertTo-MbHtml ($Message + $time)) + '</span>'
}

function Render-MbSheetNavigation {
    param([Parameter(Mandatory = $true)][object]$Project)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<nav class="sheet-nav" aria-label="シート一覧">')
    [void]$sb.AppendLine('<div class="sheet-nav__heading"><span>シート</span><span class="count-badge">' + @($Project.sheets).Count + '</span></div>')
    [void]$sb.AppendLine('<div class="sheet-nav__guide" data-sheet-sort-guide aria-live="polite">シートをつかんで並べ替え。Alt＋↑↓でも移動</div>')
    [void]$sb.AppendLine('<div class="sheet-nav__list">')
    foreach ($sheet in @($Project.sheets)) {
        $active = if ($sheet.id -eq $Project.selectedSheetId) { ' sheet-nav__item--active' } else { '' }
        $current = if ($sheet.id -eq $Project.selectedSheetId) { ' aria-current="page"' } else { '' }
        $name = ConvertTo-MbHtml $sheet.name
        $stepCount = @($sheet.steps).Count
        [void]$sb.AppendLine('<div class="sheet-nav__item' + $active + '" data-sheet-direct-drag data-sheet-nav-item data-sheet-id="' + $sheet.id + '" data-sheet-drop-target>')
        [void]$sb.AppendLine('<button type="button" class="sheet-nav__main" hx-post="/api/sheets/select" hx-vals=''{"sheetId":"' + $sheet.id + '"}'' hx-target="#workspace" hx-swap="outerHTML"' + $current + '><span class="sheet-nav__name">' + $name + '</span><span class="sheet-nav__count">' + $stepCount + '</span></button>')
        [void]$sb.AppendLine('</div>')
    }
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<button type="button" class="button button--sidebar" hx-post="/api/sheets/add" hx-target="#workspace" hx-swap="outerHTML">＋ シートを追加</button>')
    [void]$sb.AppendLine('</nav>')
    return $sb.ToString()
}

function Render-MbStepNavigation {
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][object]$Project
    )

    $steps = @($Sheet.steps)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<nav class="step-nav" aria-label="選択中シートの手順">')
    [void]$sb.AppendLine('<div class="step-nav__heading"><span>手順 <span class="count-badge">' + $steps.Count + '</span></span></div>')
    [void]$sb.AppendLine('<div class="step-nav__guide" data-step-sort-guide aria-live="polite">カードをつかんで並べ替え。Ctrl・Shiftで複数選択。Alt＋↑↓でも移動</div>')
    [void]$sb.AppendLine('<div id="step-nav-list" class="step-nav__list">')
    for ($i = 0; $i -lt $steps.Count; $i++) {
        $step = $steps[$i]
        $stepId = ConvertTo-MbHtml $step.id
        $titleClass = 'step-nav__title'
        if (-not [string]::IsNullOrWhiteSpace([string]$step.title)) {
            $title = [string]$step.title
        } elseif (-not [string]::IsNullOrWhiteSpace([string]$step.description)) {
            $title = ([string]$step.description).Trim() -replace '[\r\n\t]+', ' '
            if ($title.Length -gt 24) { $title = $title.Substring(0, 24) + '…' }
            $titleClass += ' step-nav__title--fallback'
        } else {
            $title = '（未入力）'
            $titleClass += ' step-nav__title--fallback'
        }
        $status = if (-not $step.imageId) { 'empty' } elseif ([string]::IsNullOrWhiteSpace([string]$step.description)) { 'incomplete' } else { 'complete' }
        $statusLabel = if ($status -eq 'complete') { '入力済み' } elseif ($status -eq 'incomplete') { '説明未入力' } else { '画像なし' }
        # 未完了の目印は role="img" を付けて読み上げ対象にする。付けない span の aria-label は
        # 支援技術へ届かず、色と形だけの区別になってしまう（UX-10）。
        $statusMark = if ($status -eq 'complete') {
            ''
        } else {
            '<span class="step-nav__status" role="img" title="' + $statusLabel + '" aria-label="' + $statusLabel + '"></span>'
        }
        $active = if ($i -eq 0) { ' step-nav__item--active' } else { '' }
        [void]$sb.AppendLine('<div class="step-nav__item step-nav__item--' + $status + $active + '" data-step-direct-drag data-step-nav-item data-step-id="' + $stepId + '">')
        [void]$sb.AppendLine('<button type="button" class="step-nav__main" data-step-jump="' + $stepId + '"><span class="step-nav__number">' + ($i + 1) + '</span><span class="' + $titleClass + '">' + (ConvertTo-MbHtml $title) + '</span>' + $statusMark + '</button>')
        [void]$sb.AppendLine('<details class="step-nav__actions action-menu"><summary class="step-nav__more" aria-label="手順 ' + ($i + 1) + ' の操作">…</summary><div class="action-menu__panel action-menu__panel--right step-nav__menu-panel"><button type="button" class="menu-command" data-step-nav-add-after>この下に手順を追加</button><button type="button" class="menu-command" data-step-nav-order="up"' + $(if ($i -eq 0) { ' disabled' } else { '' }) + '>1つ上へ移動</button><button type="button" class="menu-command" data-step-nav-order="down"' + $(if ($i -ge ($steps.Count - 1)) { ' disabled' } else { '' }) + '>1つ下へ移動</button><button type="button" class="menu-command menu-command--danger" data-step-nav-delete>削除</button></div></details></div>')
    }
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<div class="step-nav__bulk" data-step-bulk-actions hidden>')
    [void]$sb.AppendLine('<div class="step-nav__bulk-heading"><strong data-step-selection-count aria-live="polite">0件選択</strong><div><button type="button" data-step-selection-all>すべて選択</button><button type="button" data-step-selection-clear>選択解除</button></div></div>')
    [void]$sb.AppendLine('<div class="step-nav__bulk-order" aria-label="選択した手順の一括並べ替え"><button type="button" data-step-bulk-order="top">先頭へ</button><button type="button" data-step-bulk-order="up">↑ 上へ</button><button type="button" data-step-bulk-order="down">↓ 下へ</button><button type="button" data-step-bulk-order="bottom">末尾へ</button></div>')
    $targetSheets = @($Project.sheets | Where-Object { $_.id -ne $Sheet.id })
    if ($targetSheets.Count -gt 0) {
        [void]$sb.AppendLine('<select data-step-bulk-target aria-label="選択した手順の移動先"><option value="">移動先のシートを選択</option>')
        foreach ($targetSheet in $targetSheets) {
            [void]$sb.AppendLine('<option value="' + (ConvertTo-MbHtml $targetSheet.id) + '">' + (ConvertTo-MbHtml $targetSheet.name) + '</option>')
        }
        [void]$sb.AppendLine('</select>')
        [void]$sb.AppendLine('<div class="step-nav__bulk-buttons"><button type="button" class="button button--secondary" data-step-bulk-move disabled>まとめて移動</button><button type="button" class="button button--ghost step-nav__bulk-delete" data-step-bulk-delete>まとめて削除</button></div>')
    } else {
        [void]$sb.AppendLine('<div class="step-nav__bulk-buttons step-nav__bulk-buttons--single"><button type="button" class="button button--ghost step-nav__bulk-delete" data-step-bulk-delete>まとめて削除</button></div>')
    }
    [void]$sb.AppendLine('</div></nav>')
    return $sb.ToString()
}

function Get-MbStepVideoEntry {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][object]$Step
    )
    if ($Step.PSObject.Properties.Name -notcontains 'videoId') { return $null }
    $videoId = [string]$Step.videoId
    if ([string]::IsNullOrWhiteSpace($videoId)) { return $null }
    if ($Project.PSObject.Properties.Name -notcontains 'videos') { return $null }
    return (@($Project.videos | Where-Object { $_.id -eq $videoId }) | Select-Object -First 1)
}

function ConvertTo-MbStepCardHtml {
    param(
        [Parameter(Mandatory = $true)][object]$Step,
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter(Mandatory = $true)][int]$Total,
        [string]$Token = '',
        [bool]$CanUndoImageReplacement = $false,
        [AllowNull()][object]$Video = $null
    )

    $title = ConvertTo-MbHtml $Step.title
    $description = ConvertTo-MbHtml $Step.description
    $note = ConvertTo-MbHtml $Step.note
    $stepId = ConvertTo-MbHtml $Step.id
    $reviewRequired = $null -ne $Step.review -and [bool]$Step.review.required
    $reviewAction = if ($reviewRequired) { [string]$Step.review.action } else { '' }
    $reviewReason = if ($reviewRequired) { ConvertTo-MbHtml ([string]$Step.review.reason) } else { '' }
    $annotations = @($Step.annotations)
    # 注釈は必ずJSON配列で渡す。パイプへ流すと1件のときだけ配列でなくオブジェクトへ変換され、
    # 画面側の Array.isArray 判定で空扱いになり、注釈が消えたうえ画像編集の保存で失われる。
    $annotationsJson = if ($annotations.Count -eq 0) {
        '[]'
    } else {
        $json = ConvertTo-Json -InputObject @($annotations) -Compress -Depth 5
        if ($json.StartsWith('[')) { $json } else { '[' + $json + ']' }
    }
    $crop = if ($Step.PSObject.Properties.Name -contains 'crop' -and $null -ne $Step.crop) { $Step.crop } else { [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 } }
    $cropJson = $crop | ConvertTo-Json -Compress -Depth 3
    $isCropped = ([double]$crop.x -gt 0.000001 -or [double]$crop.y -gt 0.000001 -or [double]$crop.width -lt 0.999999 -or [double]$crop.height -lt 0.999999)
    $resultAnnotations = @(if ($Step.PSObject.Properties.Name -contains 'resultAnnotations') { @($Step.resultAnnotations) })
    $resultAnnotationsJson = if ($resultAnnotations.Count -eq 0) {
        '[]'
    } else {
        $json = ConvertTo-Json -InputObject @($resultAnnotations) -Compress -Depth 5
        if ($json.StartsWith('[')) { $json } else { '[' + $json + ']' }
    }
    $resultCrop = if ($Step.PSObject.Properties.Name -contains 'resultCrop' -and $null -ne $Step.resultCrop) { $Step.resultCrop } else { [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 } }
    $resultCropJson = $resultCrop | ConvertTo-Json -Compress -Depth 3
    $resultIsCropped = ([double]$resultCrop.x -gt 0.000001 -or [double]$resultCrop.y -gt 0.000001 -or [double]$resultCrop.width -lt 0.999999 -or [double]$resultCrop.height -lt 0.999999)
    $sb = New-Object System.Text.StringBuilder

    # 動画はExcel出力に同梱する。画面では添付されていることだけを示す。
    $videoRow = ''
    if ($Video) {
        $videoSizeText = if ([long]$Video.byteLength -ge (1024 * 1024)) {
            [string][Math]::Round([long]$Video.byteLength / 1MB, 1) + 'MB'
        } else {
            [string][Math]::Max(1, [Math]::Round([long]$Video.byteLength / 1KB)) + 'KB'
        }
        $videoLengthText = [string][Math]::Round([double]$Video.durationSec, 1) + '秒'
        $videoRow = '<div class="step-video" data-step-video><span class="step-video__mark" aria-hidden="true">▶</span><span class="step-video__text">動画つき ' + (ConvertTo-MbHtml $videoLengthText) + ' ・ ' + (ConvertTo-MbHtml $videoSizeText) + '<span class="step-video__hint">Excelで作成すると再生できます</span></span><button type="button" class="image-secondary-button" data-detach-video>動画を外す</button></div>'
    }

    $imageStateClass = if ($Step.imageId) { '' } else { ' step-card--no-image' }
    [void]$sb.AppendLine('<article class="step-card' + $imageStateClass + '" id="step-' + $stepId + '" data-step-id="' + $stepId + '">')
    [void]$sb.AppendLine('<input type="hidden" name="stepId" value="' + $stepId + '">')
    [void]$sb.AppendLine('<div class="step-card__bar">')
    [void]$sb.AppendLine('<span class="step-number">' + $Number + '</span>')
    [void]$sb.AppendLine('<span class="step-card__label">手順 ' + $Number + '</span>')
    [void]$sb.AppendLine('</div>')

    if ($reviewRequired) {
        $reviewTitle = if ($reviewAction -eq 'delete') { '不要な手順の可能性があります' } else { '操作対象を確認してください' }
        $reviewCopy = if ([string]::IsNullOrWhiteSpace($reviewReason)) { '内容と赤枠・番号を確認してください。' } else { $reviewReason }
        [void]$sb.AppendLine('<div class="step-review-notice" data-step-review-notice data-review-action="' + (ConvertTo-MbHtml $reviewAction) + '"><div><strong>' + $reviewTitle + '</strong><span>' + $reviewCopy + '</span></div><button type="button" class="button button--secondary" data-step-review-resolve>確認済みにする</button></div>')
    }

    [void]$sb.AppendLine('<div class="step-card__body">')
    [void]$sb.AppendLine('<section class="image-panel" aria-label="スクリーンショット">')
    if ($Step.imageId) {
        $imageId = ConvertTo-MbHtml $Step.imageId
        $imageUrl = '/images/' + $imageId + '?token=' + (ConvertTo-MbHtml $Token)
        $hasResultImage = ($Step.PSObject.Properties.Name -contains 'resultImageId' -and
            -not [string]::IsNullOrWhiteSpace([string]$Step.resultImageId))
        $imageLayout = if ($Step.PSObject.Properties.Name -contains 'imageLayout') { [string]$Step.imageLayout } else { 'before' }
        if ($imageLayout -notin @('before', 'after', 'side-by-side', 'stacked') -or (-not $hasResultImage -and $imageLayout -ne 'before')) { $imageLayout = 'before' }
        $imageOrder = if ($Step.PSObject.Properties.Name -contains 'imageOrder') { [string]$Step.imageOrder } else { 'before-after' }
        if ($imageOrder -notin @('before-after', 'after-before')) { $imageOrder = 'before-after' }
        $reverseClass = if ($imageOrder -eq 'after-before') { ' step-visual-layout--reverse' } else { '' }
        $cropBadge = if ($isCropped) { '<span class="image-edit-button__badge crop-badge">切り抜き済み</span>' } else { '' }
        $undoButton = if ($CanUndoImageReplacement) { '<button type="button" class="image-secondary-button image-secondary-button--undo" data-undo-image-replace>元の画像へ戻す</button>' } else { '' }
        $addResultButton = if (-not $hasResultImage) { '<button type="button" class="image-secondary-button image-secondary-button--compare" data-add-result-image>比較画像を追加</button>' } else { '' }
        $annotationBadge = '<span class="image-edit-button__badge annotation-badge">注釈 <span class="annotation-count">' + $annotations.Count + '</span></span>'
        $focusRectCount = @($annotations | Where-Object { $_.type -eq 'rect' }).Count
        $annotationActionTitle = if ($focusRectCount -gt 0) { '赤枠を確認・修正' } else { '赤枠・番号を追加' }
        $annotationActionHint = if ($focusRectCount -gt 0) { '合わない枠は移動・削除できます' } else { '拡大・切り抜きもここで' }
        $removeFocusRectButton = if ($focusRectCount -gt 0) { '<button type="button" class="image-secondary-button image-secondary-button--remove-rect" data-remove-focus-rect>赤枠を外す</button>' } else { '' }
        if ($hasResultImage) {
            $layoutLabels = [ordered]@{ before = '操作前だけ'; after = '操作後だけ'; 'side-by-side' = '左右に並べる'; stacked = '上下に並べる' }
            [void]$sb.AppendLine('<div class="image-layout-editor"><div class="image-layout-editor__heading"><strong>画像の見せ方</strong><span>選んだ配置でマニュアルに出力します</span></div><div class="image-layout-options" role="radiogroup" aria-label="画像の見せ方">')
            foreach ($layoutName in $layoutLabels.Keys) {
                $pressed = if ($imageLayout -eq $layoutName) { 'true' } else { 'false' }
                [void]$sb.AppendLine('<button type="button" class="image-layout-option" data-image-layout-option="' + $layoutName + '" aria-pressed="' + $pressed + '">' + $layoutLabels[$layoutName] + '</button>')
            }
            $resultAnnotationBadge = if ($resultAnnotations.Count -gt 0) { ' <span class="annotation-count" aria-label="注釈 ' + $resultAnnotations.Count + '件">' + $resultAnnotations.Count + '</span>' } else { '' }
            $resultCropBadge = if ($resultIsCropped) { ' <span class="image-edit-button__badge crop-badge">切り抜き済み</span>' } else { '' }
            [void]$sb.AppendLine('</div><div class="image-layout-editor__actions"><button type="button" class="image-secondary-button" data-open-annotation data-image-edit-target="result">操作後を編集' + $resultAnnotationBadge + $resultCropBadge + '</button><button type="button" class="image-secondary-button" data-swap-image-order>前後を入れ替える</button><button type="button" class="image-secondary-button" data-replace-result-image>操作後を差し替え</button><button type="button" class="image-secondary-button image-secondary-button--danger" data-remove-result-image>操作後を外す</button></div></div>')
        }
        [void]$sb.AppendLine('<div class="step-visual-layout step-visual-layout--' + $imageLayout + $reverseClass + '" data-step-visual data-image-layout="' + $imageLayout + '" data-image-order="' + $imageOrder + '">')
        [void]$sb.AppendLine('<div class="step-visual-item step-visual-item--before"><span class="step-visual-item__label">操作前</span><div class="step-image-frame"><button type="button" class="step-image-button" data-image-preview="' + $imageUrl + '" data-image-preview-kind="before" aria-label="手順 ' + $Number + ' の操作前画面を拡大"><span class="step-image-viewport"><img class="step-image" src="' + $imageUrl + '" alt="手順 ' + $Number + ' の操作前画面" loading="lazy"><svg class="step-annotation-overlay" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true"></svg></span></button></div></div>')
        if ($hasResultImage) {
            $resultImageId = ConvertTo-MbHtml ([string]$Step.resultImageId)
            $resultImageUrl = '/images/' + $resultImageId + '?token=' + (ConvertTo-MbHtml $Token)
            [void]$sb.AppendLine('<div class="step-visual-item step-visual-item--after"><span class="step-visual-item__label">操作後</span><div class="step-result-image"><button type="button" class="step-result-image__button" data-image-preview="' + $resultImageUrl + '" data-image-preview-kind="result" aria-label="手順 ' + $Number + ' の操作後画面を拡大"><span class="step-result-image__viewport"><img class="step-result-image__image" src="' + $resultImageUrl + '" alt="手順 ' + $Number + ' の操作後画面" loading="lazy"><svg class="step-annotation-overlay step-result-annotation-overlay" data-image-edit-target="result" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true"></svg></span></button></div></div>')
        }
        [void]$sb.AppendLine('</div>')
        [void]$sb.AppendLine('<div class="image-edit-actions" aria-label="画像の操作"><button type="button" class="image-edit-button" data-open-annotation data-image-edit-target="before" title="赤枠・番号・切り抜き・矢印・黒塗り"><span class="image-edit-button__icon" aria-hidden="true">＋</span><span class="image-edit-button__copy"><strong>' + $annotationActionTitle + '</strong><span>' + $annotationActionHint + '</span></span>' + $annotationBadge + $cropBadge + '</button><div class="image-secondary-actions">' + $removeFocusRectButton + '<button type="button" class="image-secondary-button" data-replace-image>操作前を差し替え</button>' + $undoButton + $addResultButton + '</div>' + $videoRow + '</div>')
        [void]$sb.AppendLine('<textarea class="step-annotations-data" hidden>' + (ConvertTo-MbHtml $annotationsJson) + '</textarea>')
        [void]$sb.AppendLine('<textarea class="step-crop-data" hidden>' + (ConvertTo-MbHtml $cropJson) + '</textarea>')
        [void]$sb.AppendLine('<textarea class="step-result-annotations-data" hidden>' + (ConvertTo-MbHtml $resultAnnotationsJson) + '</textarea>')
        [void]$sb.AppendLine('<textarea class="step-result-crop-data" hidden>' + (ConvertTo-MbHtml $resultCropJson) + '</textarea>')
    } else {
        [void]$sb.AppendLine('<div class="image-placeholder"><span class="image-placeholder__icon" aria-hidden="true">▧</span><strong>画像がありません</strong><button type="button" class="button button--primary image-placeholder__button" data-add-image-to-step>画像を追加</button><span>貼り付け・ドロップもできます</span></div>')
    }
    [void]$sb.AppendLine('</section>')

    [void]$sb.AppendLine('<section class="step-fields">')
    [void]$sb.AppendLine('<label class="field field--name name-field"><span class="sr-only">手順名</span>')
    [void]$sb.AppendLine('<input type="text" name="title" maxlength="100" value="' + $title + '" placeholder="この手順の名前" aria-label="手順名。入力して変更" hx-post="/api/steps/update" hx-trigger="input changed delay:700ms, change" hx-include="closest .step-card" hx-target="#save-status" hx-swap="outerHTML"></label>')
    [void]$sb.AppendLine('<label class="field field--grow"><span class="field__label">説明</span>')
    [void]$sb.AppendLine('<textarea name="description" maxlength="4000" rows="6" placeholder="操作内容を入力" hx-post="/api/steps/update" hx-trigger="input changed delay:700ms, change" hx-include="closest .step-card" hx-target="#save-status" hx-swap="outerHTML">' + $description + '</textarea></label>')
    [void]$sb.AppendLine('<label class="field field--note"><span class="field__label">補足 <span class="field__optional">任意</span></span>')
    [void]$sb.AppendLine('<textarea name="note" maxlength="2000" rows="2" placeholder="注意点や前提条件" hx-post="/api/steps/update" hx-trigger="input changed delay:700ms, change" hx-include="closest .step-card" hx-target="#save-status" hx-swap="outerHTML">' + $note + '</textarea></label>')
    [void]$sb.AppendLine('</section></div></article>')
    return $sb.ToString()
}

function ConvertTo-MbWatchStatusHtml {
    param(
        [ValidateSet('active', 'standby', 'suspended', 'disabled')][string]$State = 'suspended',
        [ValidateSet('owner', 'viewer', 'available')][string]$Role = 'available',
        [string]$Directory = ''
    )

    $label = if ($State -eq 'disabled') {
        '自動監視なし・貼り付け利用可'
    } elseif ($State -eq 'active' -and $Role -eq 'owner') {
        '監視中・このタブに追加'
    } elseif ($State -eq 'active' -and $Role -eq 'viewer') {
        '閲覧専用・別タブが撮影対象'
    } elseif ($State -eq 'standby' -and $Role -eq 'owner') {
        '保留中・このタブに戻ると取り込み'
    } elseif ($Role -eq 'viewer') {
        '閲覧専用・撮影タブ待ち'
    } else {
        '監視準備中'
    }
    $icon = if ($State -eq 'active' -and $Role -eq 'owner') { '●' } elseif ($State -eq 'disabled') { '!' } elseif ($State -eq 'standby') { '◐' } else { '○' }
    $detail = if ($Directory) { 'スクリーンショット保存先: ' + $Directory } else { '保存先を検出できないため、貼り付け・ドロップ・画像選択を利用できます。' }
    # 狭い画面では文言を隠してアイコンだけにする。以前は要素ごと幅で切っていたため、
    # 文言が途中で欠けたまま読めず、撮影状態が判別できなくなっていた（UX-04・UX-09）。
    # 隠したときも状態が分かるよう、状態名は必ずツールチップの先頭へ入れる。
    $title = $label + ' / ' + $detail
    return '<span id="watch-status" class="watch-status watch-status--' + $State + '" role="status" title="' + (ConvertTo-MbHtml $title) + '"><span class="watch-status__icon" aria-hidden="true">' + $icon + '</span><span class="watch-status__label">' + (ConvertTo-MbHtml $label) + '</span></span>'
}

function ConvertTo-MbCaptureSnapshotHtml {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [string]$Token = '',
        [int]$Version = 0,
        [ValidateSet('poll', 'added', 'duplicate')][string]$Status = 'poll',
        [string[]]$UndoImageStepIds = @()
    )

    $sheet = @($Project.sheets | Where-Object { $_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $sheet) { throw '対象シートが見つかりません。' }
    $steps = @($sheet.steps)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div class="capture-snapshot" data-capture-version="' + $Version + '" data-step-count="' + $steps.Count + '" data-import-status="' + $Status + '">')
    for ($i = 0; $i -lt $steps.Count; $i++) {
        $canUndo = $UndoImageStepIds -contains [string]$steps[$i].id
        $stepVideo = Get-MbStepVideoEntry -Project $Project -Step $steps[$i]
        [void]$sb.AppendLine((ConvertTo-MbStepCardHtml -Step $steps[$i] -Number ($i + 1) -Total $steps.Count -Token $Token -CanUndoImageReplacement $canUndo -Video $stepVideo))
    }
    [void]$sb.AppendLine('</div>')
    return $sb.ToString()
}

function ConvertTo-MbProjectLibraryHtml {
    param(
        [object[]]$Projects = @(),
        [object[]]$ArchivedProjects = @(),
        [AllowEmptyString()][string]$LastOpenedProjectKey = ''
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div id="workspace" class="workspace project-library" data-app-version="0.46.0">')
    [void]$sb.AppendLine('<header class="topbar project-library__topbar"><button type="button" class="brand brand--home" data-project-home hx-post="/api/projects/home" hx-target="#workspace" hx-swap="outerHTML" title="マニュアル一覧" aria-label="マニュアル一覧" aria-current="page"><span class="brand__mark" aria-hidden="true">M</span><span>ManualBuilder</span></button><div class="project-library__topbar-title">マニュアル一覧</div><div></div><div class="topbar__actions"><details class="action-menu topbar-menu"><summary class="icon-button" title="その他" aria-label="その他の操作">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command menu-command--danger" hx-post="/api/shutdown" hx-target="body" hx-swap="none" hx-confirm="ManualBuilderを終了しますか？">ManualBuilderを終了</button></div></details></div></header>')
    [void]$sb.AppendLine('<main class="project-library__main">')
    [void]$sb.AppendLine('<section class="project-library__intro"><div><p class="eyebrow">作成したマニュアル</p><h1>マニュアルを選ぶ</h1></div><div class="project-library__actions"><input id="project-package-input" type="file" accept=".zip,application/zip" hidden><button type="button" class="button button--ghost" data-import-project-package>ZIPを取り込む</button><form class="project-create" hx-post="/api/projects/create" hx-target="#workspace" hx-swap="outerHTML"><label><span class="sr-only">新しいマニュアルの名前</span><input type="text" name="title" maxlength="100" placeholder="新しいマニュアルの名前"></label><button type="submit" class="button button--primary">＋ 新規作成</button></form></div></section>')

    [void]$sb.AppendLine('<section class="project-library__section" aria-labelledby="active-projects-heading"><div class="project-library__section-heading"><div><h2 id="active-projects-heading">マニュアル</h2><span>' + @($Projects).Count + ' 件</span></div><label class="project-search"><span aria-hidden="true">⌕</span><span class="sr-only">マニュアルを検索</span><input type="search" placeholder="名前で検索" data-project-search></label></div>')
    if (@($Projects).Count -eq 0) {
        [void]$sb.AppendLine('<div class="project-library__empty"><strong>マニュアルはまだありません</strong><span>右上の「新規作成」から始められます。</span></div>')
    } else {
        [void]$sb.AppendLine('<div class="project-grid" data-project-grid>')
        foreach ($entry in @($Projects)) {
            $key = ConvertTo-MbHtml $entry.key
            $title = ConvertTo-MbHtml $entry.title
            $searchText = ConvertTo-MbHtml ([string]$entry.title).ToLowerInvariant()
            $updated = try { ([DateTime]$entry.updatedAt).ToLocalTime().ToString('yyyy/MM/dd HH:mm') } catch { '日時不明' }
            $lastBadge = if ([string]$entry.key -eq $LastOpenedProjectKey) { '<span class="project-card__badge">前回開いた項目</span>' } else { '' }
            $disabled = if ([bool]$entry.readable) { '' } else { ' disabled' }
            [void]$sb.AppendLine('<article class="project-card" data-project-card data-project-search-text="' + $searchText + '"><div class="project-card__header"><span class="project-card__mark" aria-hidden="true">▤</span>' + $lastBadge + '<details class="action-menu project-card__menu"><summary class="icon-button" aria-label="' + $title + ' の操作">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command" data-project-export data-project-key="' + $key + '"' + $disabled + '>ZIPに書き出す</button><button type="button" class="menu-command" hx-post="/api/projects/duplicate" hx-vals=''{"projectKey":"' + $key + '"}'' hx-target="#workspace" hx-swap="outerHTML"' + $disabled + '>複製</button><button type="button" class="menu-command menu-command--danger" hx-post="/api/projects/delete" hx-vals=''{"projectKey":"' + $key + '"}'' hx-confirm="このマニュアルを完全に削除しますか？ 元に戻せません。" hx-target="#workspace" hx-swap="outerHTML"' + $disabled + '>削除</button></div></details></div><div class="project-card__body"><h3>' + $title + '</h3><p>' + [int]$entry.sheetCount + ' シート ・ ' + [int]$entry.stepCount + ' 手順</p><span>更新 ' + $updated + '</span></div>')
            if ([bool]$entry.readable) {
                [void]$sb.AppendLine('<button type="button" class="project-card__open" hx-post="/api/projects/open" hx-vals=''{"projectKey":"' + $key + '"}'' hx-target="#workspace" hx-swap="outerHTML"><span>開く</span><span aria-hidden="true">→</span></button></article>')
            } else {
                [void]$sb.AppendLine('<div class="project-card__error" title="' + (ConvertTo-MbHtml $entry.error) + '">読み込みエラー。project.jsonを確認してください。</div></article>')
            }
        }
        [void]$sb.AppendLine('</div><p class="project-search-empty" data-project-search-empty hidden>一致するマニュアルがありません。</p>')
    }
    [void]$sb.AppendLine('</section>')

    [void]$sb.AppendLine('<footer class="project-library__footer"><span>v0.46.0</span></footer></main></div>')
    return $sb.ToString()
}

function ConvertTo-MbWorkspaceHtml {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [string]$Token = '',
        [ValidateSet('active', 'standby', 'suspended', 'disabled')][string]$CaptureState = 'suspended',
        [ValidateSet('owner', 'viewer', 'available')][string]$CaptureRole = 'available',
        [string]$CaptureDirectory = '',
        [int]$CaptureVersion = 0,
        [string[]]$UndoImageStepIds = @()
    )

    $sheet = Get-MbSelectedSheet -Project $Project
    $title = ConvertTo-MbHtml $Project.title
    $sheetName = ConvertTo-MbHtml $sheet.name
    $sheetId = ConvertTo-MbHtml $sheet.id
    $steps = @($sheet.steps)
    $sb = New-Object System.Text.StringBuilder

    # 出力は全シートを対象にするため、仕上げ状況も選択中シートだけではなく
    # プロジェクト全体を基準にする。現在シートの入力値はブラウザー側で上書きする。
    $finishItems = @(
        foreach ($projectSheet in @($Project.sheets)) {
            foreach ($projectStep in @($projectSheet.steps)) {
                $layout = if ($projectStep.PSObject.Properties.Name -contains 'imageLayout') { [string]$projectStep.imageLayout } else { 'before' }
                $beforeAnnotations = if ($projectStep.PSObject.Properties.Name -contains 'annotations') { @($projectStep.annotations) } else { @() }
                $afterAnnotations = if ($projectStep.PSObject.Properties.Name -contains 'resultAnnotations') { @($projectStep.resultAnnotations) } else { @() }
                $visibleAnnotations = if ($layout -eq 'after') {
                    @($afterAnnotations)
                } elseif ($layout -in @('side-by-side', 'stacked')) {
                    @($beforeAnnotations) + @($afterAnnotations)
                } else {
                    @($beforeAnnotations)
                }
                $hasFocusAnnotation = @($visibleAnnotations | Where-Object { $_.type -in @('rect', 'number') }).Count -gt 0
                [ordered]@{
                    sheetId = [string]$projectSheet.id
                    sheetName = [string]$projectSheet.name
                    stepId = [string]$projectStep.id
                    missingText = [string]::IsNullOrWhiteSpace([string]$projectStep.description)
                    missingImage = [string]::IsNullOrWhiteSpace([string]$projectStep.imageId)
                    hasFocusAnnotation = $hasFocusAnnotation
                    reviewRequired = [bool]$projectStep.review.required
                    reviewAction = [string]$projectStep.review.action
                    reviewReason = [string]$projectStep.review.reason
                }
            }
        }
    )
    $finishJson = ConvertTo-MbHtml (ConvertTo-Json -InputObject @($finishItems) -Compress -Depth 5)
    $projectStepCount = @($finishItems).Count

    [void]$sb.AppendLine('<div id="workspace" class="workspace step-view--review" data-app-version="0.46.0" data-revision="' + [int]$Project.revision + '" data-capture-version="' + $CaptureVersion + '">')
    [void]$sb.AppendLine('<textarea hidden data-project-finish-data>' + $finishJson + '</textarea>')
    [void]$sb.AppendLine('<header class="topbar">')
    [void]$sb.AppendLine('<button type="button" class="brand brand--home" data-project-home hx-post="/api/projects/home" hx-target="#workspace" hx-swap="outerHTML" title="マニュアル一覧へ戻る" aria-label="マニュアル一覧へ戻る"><span class="brand__mark" aria-hidden="true">M</span><span>ManualBuilder</span></button>')
    [void]$sb.AppendLine('<label class="project-title editable-name name-field" data-editable-name title="マニュアル名を編集"><span class="sr-only">マニュアル名</span><input type="text" name="title" maxlength="100" value="' + $title + '" aria-label="マニュアル名。入力して変更" hx-post="/api/project/title" hx-trigger="input changed delay:700ms, change" hx-target="#save-status" hx-swap="outerHTML"></label>')
    [void]$sb.AppendLine('<div class="topbar__state">' + (ConvertTo-MbSaveStatusHtml) + (ConvertTo-MbWatchStatusHtml -State $CaptureState -Role $CaptureRole -Directory $CaptureDirectory) + '</div>')
    $topbarRecordAction = if ($steps.Count -gt 0) { '<button type="button" class="button button--primary topbar__main-action topbar__record-action" data-record-operations>操作を記録</button>' } else { '' }
    $topbarVideoAction = if ($steps.Count -gt 0) { '<button type="button" class="button button--secondary topbar__video-action" data-open-video-picker>録画を取り込む</button>' } else { '' }
    $topbarExportAction = if ($projectStepCount -gt 0) { '<button type="button" class="button button--secondary topbar-export" data-open-export-dialog>手順書を出力</button>' } else { '' }
    [void]$sb.AppendLine('<div class="topbar__actions">' + $topbarRecordAction + $topbarVideoAction + $topbarExportAction + '<details class="action-menu topbar-menu"><summary class="icon-button" title="その他" aria-label="その他の操作">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command" data-copilot-review>文章をまとめて整える（任意）</button><button type="button" class="menu-command menu-command--danger" hx-post="/api/shutdown" hx-target="body" hx-swap="none" hx-confirm="ManualBuilderを終了しますか？">ManualBuilderを終了</button></div></details></div>')
    [void]$sb.AppendLine('</header>')

    [void]$sb.AppendLine('<div class="app-layout">')
    [void]$sb.AppendLine('<aside class="sidebar">' + (Render-MbSheetNavigation -Project $Project) + (Render-MbStepNavigation -Sheet $sheet -Project $Project) + '<div class="sidebar__footer"><span class="sidebar__version">v0.46.0</span></div></aside>')
    [void]$sb.AppendLine('<main class="editor">')
    # 編集画面の見出しは入力欄なので、文書構造としての見出しが無い。読み上げの目次から
    # 何を編集中か分かるよう、画面には出さない h1 を置く。
    [void]$sb.AppendLine('<h1 class="sr-only">' + $title + ' の編集</h1>')
    [void]$sb.AppendLine('<div class="editor__heading">')
    [void]$sb.AppendLine('<div class="sheet-heading"><input type="hidden" name="sheetId" value="' + $sheetId + '"><label class="editable-name editable-name--sheet name-field" data-editable-name title="シート名を編集"><span class="sr-only">シート名</span><input class="sheet-name-input" type="text" name="name" maxlength="50" value="' + $sheetName + '" aria-label="シート名。入力して変更" hx-post="/api/sheets/rename" hx-trigger="input changed delay:700ms, change" hx-include="closest .sheet-heading" hx-target="#save-status" hx-swap="outerHTML"></label><span class="step-total">' + $steps.Count + ' 手順</span></div>')
    [void]$sb.AppendLine('<div class="editor__actions">')
    if ($steps.Count -gt 0) {
        if ($steps.Count -gt 1) { [void]$sb.AppendLine('<span class="editor-shortcut"><kbd>Ctrl</kbd>+<kbd>Enter</kbd> 次の手順</span>') }
        [void]$sb.AppendLine('<button type="button" class="button button--secondary editor-add-step" data-add-step-end>＋ 手順を追加</button>')
        [void]$sb.AppendLine('<button type="button" class="button button--ghost editor-add-image" data-open-image-picker>＋ 画像から追加</button>')
    }
    [void]$sb.AppendLine('<details class="action-menu sheet-menu"><summary class="icon-button" title="シートメニュー" aria-label="シートのメニュー">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command" hx-post="/api/steps/add" hx-vals=''{"sheetId":"' + $sheetId + '"}'' hx-target="#workspace" hx-swap="outerHTML">文字だけ追加</button><button type="button" class="menu-command" data-sheet-duplicate data-sheet-id="' + $sheetId + '">シートを複製</button>')
    if (@($Project.sheets).Count -le 1) {
        [void]$sb.AppendLine('<button type="button" class="menu-command menu-command--danger" disabled title="最後のシートは削除できません">シートを削除</button>')
    } else {
        [void]$sb.AppendLine('<button type="button" class="menu-command menu-command--danger" data-sheet-delete data-sheet-id="' + $sheetId + '">シートを削除</button>')
    }
    [void]$sb.AppendLine('</div></details></div>')
    [void]$sb.AppendLine('</div>')

    if ($steps.Count -gt 0) {
        [void]$sb.AppendLine('<section class="step-review-toolbar" data-step-review-toolbar aria-labelledby="step-review-title"><div class="step-review-toolbar__heading"><strong id="step-review-title">分割結果を確認</strong><span>一覧ではスクロールするだけで前後の画面を見比べられます</span></div><div class="step-view-options" role="group" aria-label="手順の表示方法"><button type="button" class="step-view-option" data-step-view="review" aria-pressed="true">一覧で確認</button><button type="button" class="step-view-option" data-step-view="focus" aria-pressed="false">1件ずつ編集</button></div><div class="step-review-navigation" aria-label="手順の移動"><button type="button" data-step-previous aria-label="前の手順">← 前</button><span data-step-position aria-live="polite">1 / ' + $steps.Count + '</span><button type="button" data-step-next aria-label="次の手順">次 →</button></div><span class="step-review-shortcut"><kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>↑</kbd>/<kbd>↓</kbd></span></section>')
    }

    if ($projectStepCount -gt 0) {
        [void]$sb.AppendLine('<section class="finish-guide" data-finish-guide aria-labelledby="finish-guide-title"><div class="finish-guide__heading"><span class="finish-guide__mark" aria-hidden="true">✓</span><div><strong id="finish-guide-title">仕上げて出力</strong><span data-finish-summary>全シートの文章・画像・順番を確認します</span></div></div><div class="finish-guide__checks" aria-label="仕上げ状況"><button type="button" data-finish-check="text"><span>文章</span><strong data-finish-text>確認中</strong></button><button type="button" data-finish-check="image"><span>画像</span><strong data-finish-image>確認中</strong></button><button type="button" data-finish-check="annotation"><span>赤枠・番号</span><strong data-finish-annotation>確認中</strong></button><button type="button" data-finish-check="attention"><span>要確認</span><strong data-finish-attention>なし</strong></button><button type="button" data-step-organize-shortcut><span>順番・不要手順</span><strong>手順一覧で整理</strong></button></div><button type="button" class="button button--primary finish-guide__export" data-open-export-dialog>確認して出力</button></section>')
    }

    [void]$sb.AppendLine('<div class="editor-import-inputs"><input id="image-file-input" type="file" accept="image/png,image/jpeg,image/bmp" multiple hidden><input id="replacement-image-file-input" type="file" accept="image/png,image/jpeg,image/bmp" hidden><input id="result-image-file-input" type="file" accept="image/png,image/jpeg,image/bmp" hidden><input id="video-file-input" type="file" accept="video/mp4,video/webm" hidden></div>')
    [void]$sb.AppendLine('<section class="steps" aria-label="手順一覧">')
    if ($steps.Count -eq 0) {
        [void]$sb.AppendLine('<div class="empty-state drop-target"><p class="empty-state__eyebrow">最短の作り方</p><div class="empty-state__icon empty-state__icon--record" aria-hidden="true">●</div><h2>操作を記録して、手順書を作る</h2><p class="empty-state__lead">クリックした画面と操作箇所を保存し、編集できる手順をその場で作ります。</p><ol class="empty-state__flow" aria-label="作成の流れ"><li><span>1</span>操作を記録</li><li><span>2</span>自動で手順を作成</li><li><span>3</span>編集してExcelへ</li></ol><button type="button" class="button button--primary empty-state__main-button" data-record-operations>操作の記録を開始</button><p class="empty-state__privacy">記録と文章化はこのPC内で完了します。不要な手順や合わない赤枠は後から直せます。</p><div class="empty-state__alternatives"><span>すでに素材がある場合</span><div class="empty-state__actions"><button type="button" class="button button--secondary empty-state__button" data-open-video-picker>録画ファイルを取り込む <small>mp4・webm</small></button><button type="button" class="button button--ghost empty-state__button" data-open-image-picker>画像から作る</button><button type="button" class="button button--ghost empty-state__button" hx-post="/api/steps/add" hx-vals=''{"sheetId":"' + $sheetId + '"}'' hx-target="#workspace" hx-swap="outerHTML">空の手順を追加</button></div><p>録画ファイルや画像がなくても、空の手順から直接書き始められます。仕上げの文章調整にだけCopilotを任意で使えます。</p></div></div>')
    } else {
        for ($i = 0; $i -lt $steps.Count; $i++) {
            $canUndo = $UndoImageStepIds -contains [string]$steps[$i].id
            $stepVideo = Get-MbStepVideoEntry -Project $Project -Step $steps[$i]
            [void]$sb.AppendLine((ConvertTo-MbStepCardHtml -Step $steps[$i] -Number ($i + 1) -Total $steps.Count -Token $Token -CanUndoImageReplacement $canUndo -Video $stepVideo))
        }
    }
    [void]$sb.AppendLine('</section></main></div><div id="drop-overlay" class="drop-overlay" aria-hidden="true"><div>画像をドロップして手順を追加</div></div></div>')
    return $sb.ToString()
}

Export-ModuleMember -Function @(
    'ConvertTo-MbHtml',
    'ConvertTo-MbSaveStatusHtml',
    'ConvertTo-MbWatchStatusHtml',
    'ConvertTo-MbStepCardHtml',
    'ConvertTo-MbCaptureSnapshotHtml',
    'ConvertTo-MbProjectLibraryHtml',
    'ConvertTo-MbWorkspaceHtml'
)
