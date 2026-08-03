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
    [void]$sb.AppendLine('<div class="sheet-nav__guide" data-sheet-sort-guide aria-live="polite">⠿で並べ替え</div>')
    [void]$sb.AppendLine('<div class="sheet-nav__list">')
    foreach ($sheet in @($Project.sheets)) {
        $active = if ($sheet.id -eq $Project.selectedSheetId) { ' sheet-nav__item--active' } else { '' }
        $current = if ($sheet.id -eq $Project.selectedSheetId) { ' aria-current="page"' } else { '' }
        $name = ConvertTo-MbHtml $sheet.name
        $stepCount = @($sheet.steps).Count
        [void]$sb.AppendLine('<div class="sheet-nav__item' + $active + '" data-sheet-nav-item data-sheet-id="' + $sheet.id + '" data-sheet-drop-target>')
        [void]$sb.AppendLine('<button type="button" class="sheet-nav__drag" draggable="true" data-sheet-nav-drag-handle title="ドラッグしてシートを並べ替え" aria-label="' + $name + ' をドラッグして並べ替え">⠿</button>')
        [void]$sb.AppendLine('<button type="button" class="sheet-nav__main" hx-post="/api/sheets/select" hx-vals=''{"sheetId":"' + $sheet.id + '"}'' hx-target="#workspace" hx-swap="outerHTML"' + $current + '><span class="sheet-nav__name">' + $name + '</span><span class="sheet-nav__count">' + $stepCount + '</span></button>')
        [void]$sb.AppendLine('</div>')
    }
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<button type="button" class="button button--sidebar" hx-post="/api/sheets/add" hx-target="#workspace" hx-swap="outerHTML">＋ シートを追加</button>')
    [void]$sb.AppendLine('</nav>')
    return $sb.ToString()
}

function Render-MbStepNavigation {
    param([Parameter(Mandatory = $true)][object]$Sheet)

    $steps = @($Sheet.steps)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<nav class="step-nav" aria-label="選択中シートの手順">')
    [void]$sb.AppendLine('<div class="step-nav__heading"><span>手順</span><span class="count-badge">' + $steps.Count + '</span></div>')
    [void]$sb.AppendLine('<div class="step-nav__guide" data-step-sort-guide aria-live="polite">⠿で並べ替え・別シートへ移動</div>')
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
        $active = if ($i -eq 0) { ' step-nav__item--active' } else { '' }
        [void]$sb.AppendLine('<div class="step-nav__item step-nav__item--' + $status + $active + '" data-step-nav-item data-step-id="' + $stepId + '">')
        [void]$sb.AppendLine('<button type="button" class="step-nav__drag" draggable="true" data-step-nav-drag-handle title="ドラッグして並べ替え" aria-label="手順 ' + ($i + 1) + ' をドラッグして並べ替え">⠿</button>')
        [void]$sb.AppendLine('<button type="button" class="step-nav__main" data-step-jump="' + $stepId + '"><span class="step-nav__number">' + ($i + 1) + '</span><span class="' + $titleClass + '">' + (ConvertTo-MbHtml $title) + '</span><span class="step-nav__status" title="' + $statusLabel + '" aria-label="' + $statusLabel + '"></span></button>')
        [void]$sb.AppendLine('<div class="step-nav__actions"><button type="button" class="step-nav__delete" data-step-nav-delete title="この手順を削除" aria-label="手順 ' + ($i + 1) + ' を削除">×</button></div></div>')
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
    $sb = New-Object System.Text.StringBuilder

    # 動画はPowerPoint出力にだけ埋め込む。画面では添付されていることだけを示す。
    $videoRow = ''
    if ($Video) {
        $videoSizeText = if ([long]$Video.byteLength -ge (1024 * 1024)) {
            [string][Math]::Round([long]$Video.byteLength / 1MB, 1) + 'MB'
        } else {
            [string][Math]::Max(1, [Math]::Round([long]$Video.byteLength / 1KB)) + 'KB'
        }
        $videoLengthText = [string][Math]::Round([double]$Video.durationSec, 1) + '秒'
        $videoRow = '<div class="step-video" data-step-video><span class="step-video__mark" aria-hidden="true">▶</span><span class="step-video__text">動画つき ' + (ConvertTo-MbHtml $videoLengthText) + ' ・ ' + (ConvertTo-MbHtml $videoSizeText) + '<span class="step-video__hint">PowerPointで作成すると再生できます</span></span><button type="button" class="image-secondary-button" data-detach-video>動画を外す</button></div>'
    }

    $imageStateClass = if ($Step.imageId) { '' } else { ' step-card--no-image' }
    [void]$sb.AppendLine('<article class="step-card' + $imageStateClass + '" id="step-' + $stepId + '" data-step-id="' + $stepId + '">')
    [void]$sb.AppendLine('<input type="hidden" name="stepId" value="' + $stepId + '">')
    [void]$sb.AppendLine('<div class="step-card__bar">')
    [void]$sb.AppendLine('<span class="step-number">' + $Number + '</span>')
    [void]$sb.AppendLine('<span class="step-card__label">手順 ' + $Number + '</span>')
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<div class="step-card__body">')
    [void]$sb.AppendLine('<section class="image-panel" aria-label="スクリーンショット">')
    if ($Step.imageId) {
        $imageId = ConvertTo-MbHtml $Step.imageId
        $imageUrl = '/images/' + $imageId + '?token=' + (ConvertTo-MbHtml $Token)
        $cropBadge = if ($isCropped) { '<span class="image-edit-button__badge crop-badge">切り抜き済み</span>' } else { '' }
        $undoButton = if ($CanUndoImageReplacement) { '<button type="button" class="image-secondary-button image-secondary-button--undo" data-undo-image-replace>元の画像へ戻す</button>' } else { '' }
        $annotationBadge = if ($annotations.Count -gt 0) { '<span class="image-edit-button__badge">注釈 <span class="annotation-count">' + $annotations.Count + '</span></span>' } else { '' }
        [void]$sb.AppendLine('<div class="step-image-frame"><button type="button" class="step-image-button" data-image-preview="' + $imageUrl + '" aria-label="手順 ' + $Number + ' のスクリーンショットを拡大"><span class="step-image-viewport"><img class="step-image" src="' + $imageUrl + '" alt="手順 ' + $Number + ' のスクリーンショット" loading="lazy"><svg class="step-annotation-overlay" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true"></svg></span></button><div class="image-edit-actions" aria-label="画像の操作"><button type="button" class="image-edit-button" data-open-annotation title="切り抜き・赤枠・矢印・番号・黒塗り"><span class="image-edit-button__icon" aria-hidden="true">✎</span><span class="image-edit-button__copy"><strong>画像を編集</strong></span>' + $annotationBadge + $cropBadge + '</button><div class="image-secondary-actions"><button type="button" class="image-secondary-button" data-replace-image>差し替え</button>' + $undoButton + '</div>' + $videoRow + '</div></div>')
        [void]$sb.AppendLine('<textarea class="step-annotations-data" hidden>' + (ConvertTo-MbHtml $annotationsJson) + '</textarea>')
        [void]$sb.AppendLine('<textarea class="step-crop-data" hidden>' + (ConvertTo-MbHtml $cropJson) + '</textarea>')
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
    $title = if ($Directory) { 'スクリーンショット保存先: ' + $Directory } else { '保存先を検出できないため、貼り付け・ドロップ・画像選択を利用できます。' }
    return '<span id="watch-status" class="watch-status watch-status--' + $State + '" role="status" title="' + (ConvertTo-MbHtml $title) + '"><span aria-hidden="true">' + $icon + '</span> ' + (ConvertTo-MbHtml $label) + '</span>'
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
    [void]$sb.AppendLine('<div id="workspace" class="workspace project-library" data-app-version="0.27.0">')
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
            [void]$sb.AppendLine('<article class="project-card" data-project-card data-project-search-text="' + $searchText + '"><div class="project-card__header"><span class="project-card__mark" aria-hidden="true">▤</span>' + $lastBadge + '<details class="action-menu project-card__menu"><summary class="icon-button" aria-label="' + $title + ' の操作">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command" data-project-export data-project-key="' + $key + '"' + $disabled + '>ZIPに書き出す</button><button type="button" class="menu-command" hx-post="/api/projects/duplicate" hx-vals=''{"projectKey":"' + $key + '"}'' hx-target="#workspace" hx-swap="outerHTML"' + $disabled + '>複製</button><button type="button" class="menu-command menu-command--danger" hx-post="/api/projects/archive" hx-vals=''{"projectKey":"' + $key + '"}'' hx-confirm="このマニュアルをアーカイブしますか？ 後から復元できます。" hx-target="#workspace" hx-swap="outerHTML"' + $disabled + '>アーカイブ</button></div></details></div><div class="project-card__body"><h3>' + $title + '</h3><p>' + [int]$entry.sheetCount + ' シート ・ ' + [int]$entry.stepCount + ' 手順</p><span>更新 ' + $updated + '</span></div>')
            if ([bool]$entry.readable) {
                [void]$sb.AppendLine('<button type="button" class="project-card__open" hx-post="/api/projects/open" hx-vals=''{"projectKey":"' + $key + '"}'' hx-target="#workspace" hx-swap="outerHTML"><span>開く</span><span aria-hidden="true">→</span></button></article>')
            } else {
                [void]$sb.AppendLine('<div class="project-card__error" title="' + (ConvertTo-MbHtml $entry.error) + '">読み込みエラー。project.jsonを確認してください。</div></article>')
            }
        }
        [void]$sb.AppendLine('</div><p class="project-search-empty" data-project-search-empty hidden>一致するマニュアルがありません。</p>')
    }
    [void]$sb.AppendLine('</section>')

    if (@($ArchivedProjects).Count -gt 0) {
        [void]$sb.AppendLine('<details class="project-archive"><summary>アーカイブ <span>' + @($ArchivedProjects).Count + ' 件</span></summary><div class="project-archive__list">')
        foreach ($entry in @($ArchivedProjects)) {
            $key = ConvertTo-MbHtml $entry.key
            $title = ConvertTo-MbHtml $entry.title
            $disabled = if ([bool]$entry.readable) { '' } else { ' disabled' }
            [void]$sb.AppendLine('<div class="project-archive__item"><div><strong>' + $title + '</strong><span>' + [int]$entry.sheetCount + ' シート ・ ' + [int]$entry.stepCount + ' 手順</span></div><button type="button" class="button button--ghost" hx-post="/api/projects/restore" hx-vals=''{"projectKey":"' + $key + '"}'' hx-target="#workspace" hx-swap="outerHTML"' + $disabled + '>復元</button></div>')
        }
        [void]$sb.AppendLine('</div></details>')
    }
    [void]$sb.AppendLine('<footer class="project-library__footer"><span>v0.27.0</span></footer></main></div>')
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

    [void]$sb.AppendLine('<div id="workspace" class="workspace" data-app-version="0.27.0" data-revision="' + [int]$Project.revision + '" data-capture-version="' + $CaptureVersion + '">')
    [void]$sb.AppendLine('<header class="topbar">')
    [void]$sb.AppendLine('<button type="button" class="brand brand--home" data-project-home hx-post="/api/projects/home" hx-target="#workspace" hx-swap="outerHTML" title="マニュアル一覧へ戻る" aria-label="マニュアル一覧へ戻る"><span class="brand__mark" aria-hidden="true">M</span><span>ManualBuilder</span></button>')
    [void]$sb.AppendLine('<label class="project-title editable-name name-field" data-editable-name title="マニュアル名を編集"><span class="sr-only">マニュアル名</span><input type="text" name="title" maxlength="100" value="' + $title + '" aria-label="マニュアル名。入力して変更" hx-post="/api/project/title" hx-trigger="input changed delay:700ms, change" hx-target="#save-status" hx-swap="outerHTML"></label>')
    [void]$sb.AppendLine('<div class="topbar__state">' + (ConvertTo-MbSaveStatusHtml) + (ConvertTo-MbWatchStatusHtml -State $CaptureState -Role $CaptureRole -Directory $CaptureDirectory) + '</div>')
    [void]$sb.AppendLine('<div class="topbar__actions"><button type="button" class="button button--primary" data-export-excel>Excelで作成</button><button type="button" class="button button--secondary" data-export-html>HTMLで作成</button><details class="action-menu topbar-menu"><summary class="icon-button" title="その他" aria-label="その他の操作">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command" data-export-powerpoint>PowerPointで作成（動画つき）</button><button type="button" class="menu-command" data-export-word>Wordで作成</button><button type="button" class="menu-command menu-command--danger" hx-post="/api/shutdown" hx-target="body" hx-swap="none" hx-confirm="ManualBuilderを終了しますか？">ManualBuilderを終了</button></div></details></div>')
    [void]$sb.AppendLine('</header>')

    [void]$sb.AppendLine('<div class="app-layout">')
    [void]$sb.AppendLine('<aside class="sidebar">' + (Render-MbSheetNavigation -Project $Project) + (Render-MbStepNavigation -Sheet $sheet) + '<div class="sidebar__footer"><span class="sidebar__version">v0.27.0</span></div></aside>')
    [void]$sb.AppendLine('<main class="editor">')
    [void]$sb.AppendLine('<div class="editor__heading">')
    [void]$sb.AppendLine('<div class="sheet-heading"><input type="hidden" name="sheetId" value="' + $sheetId + '"><label class="editable-name editable-name--sheet name-field" data-editable-name title="シート名を編集"><span class="sr-only">シート名</span><input class="sheet-name-input" type="text" name="name" maxlength="50" value="' + $sheetName + '" aria-label="シート名。入力して変更" hx-post="/api/sheets/rename" hx-trigger="input changed delay:700ms, change" hx-include="closest .sheet-heading" hx-target="#save-status" hx-swap="outerHTML"></label><span class="step-total">' + $steps.Count + ' 手順</span></div>')
    [void]$sb.AppendLine('<div class="editor__actions">')
    if ($steps.Count -gt 0) {
        [void]$sb.AppendLine('<button type="button" class="button button--primary editor-add-image" data-open-image-picker>＋ 画像から追加</button>')
        [void]$sb.AppendLine('<button type="button" class="button button--ghost editor-add-video" data-open-video-picker title="録画から場面を選んで手順にします">動画から選ぶ</button>')
    }
    [void]$sb.AppendLine('<details class="action-menu sheet-menu"><summary class="icon-button" title="シートメニュー" aria-label="シートのメニュー">…</summary><div class="action-menu__panel action-menu__panel--right"><button type="button" class="menu-command" hx-post="/api/steps/add" hx-vals=''{"sheetId":"' + $sheetId + '"}'' hx-target="#workspace" hx-swap="outerHTML">文字だけ追加</button>')
    if (@($Project.sheets).Count -le 1) {
        [void]$sb.AppendLine('<button type="button" class="menu-command menu-command--danger" disabled title="最後のシートは削除できません">シートを削除</button>')
    } else {
        [void]$sb.AppendLine('<button type="button" class="menu-command menu-command--danger" hx-post="/api/sheets/delete" hx-vals=''{"sheetId":"' + $sheetId + '"}'' hx-confirm="このシートと中の手順を削除しますか？" hx-target="#workspace" hx-swap="outerHTML">シートを削除</button>')
    }
    [void]$sb.AppendLine('</div></details></div>')
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<div class="editor-import-inputs"><input id="image-file-input" type="file" accept="image/png,image/jpeg,image/bmp" multiple hidden><input id="replacement-image-file-input" type="file" accept="image/png,image/jpeg,image/bmp" hidden><input id="video-file-input" type="file" accept="video/mp4,video/webm" hidden></div>')
    [void]$sb.AppendLine('<section class="steps" aria-label="手順一覧">')
    if ($steps.Count -eq 0) {
        [void]$sb.AppendLine('<div class="empty-state drop-target"><div class="empty-state__icon" aria-hidden="true">▧</div><h2>ここにスクリーンショットを追加</h2><div class="empty-state__actions"><button type="button" class="button button--primary empty-state__button" data-open-image-picker>＋ 画像から追加</button><button type="button" class="button button--ghost empty-state__button" data-open-video-picker>動画から選ぶ</button></div><p>撮影・貼り付け・ドロップでも作成できます</p></div>')
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
