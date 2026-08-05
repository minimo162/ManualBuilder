# Phase 1 web rendering tests (no server, no browser).

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Web.psm1') -Force

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function New-MbTestStep {
    param([int]$AnnotationCount)
    $annotations = @()
    for ($i = 1; $i -le $AnnotationCount; $i++) {
        $annotations += [pscustomobject]@{
            id    = 'annotation-' + ([guid]::NewGuid().ToString('N'))
            type  = 'number'
            x1    = 0.5; y1 = 0.5; x2 = 0.5; y2 = 0.5
            label = $i
        }
    }
    return [pscustomobject]@{
        id          = 'step-' + ([guid]::NewGuid().ToString('N'))
        title       = 'テスト手順'
        description = 'テスト説明'
        note        = ''
        imageId     = 'image-' + ([guid]::NewGuid().ToString('N'))
        annotations = @($annotations)
        crop        = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
        review      = [pscustomobject]@{ required = $false; action = ''; reason = '' }
    }
}

function Get-MbAnnotationsJson {
    param([string]$Html)
    $match = [regex]::Match($Html, '<textarea class="step-annotations-data" hidden>(.*?)</textarea>')
    if (-not $match.Success) { throw '注釈データのtextareaが見つかりません。' }
    return [Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
}

Write-Host '--- 注釈データのJSON形式 ---' -ForegroundColor Cyan

# 1件のときにPowerShellのパイプが配列を解いてしまうと、画面側のArray.isArray判定で
# 注釈が空扱いになり、表示が消えたうえ画像編集の保存で失われる。
foreach ($count in @(1, 2, 3)) {
    $step = New-MbTestStep -AnnotationCount $count
    $html = ConvertTo-MbStepCardHtml -Step $step -Number 1 -Total 1 -Token 'testtoken'
    $json = Get-MbAnnotationsJson -Html $html
    Assert-Mb ($json.StartsWith('[')) "注釈が${count}件でもJSON配列で出力する"
    $parsed = $json | ConvertFrom-Json
    Assert-Mb (@($parsed).Count -eq $count) "注釈${count}件が欠落せずJSONへ入る"
}

$emptyStep = New-MbTestStep -AnnotationCount 0
$emptyHtml = ConvertTo-MbStepCardHtml -Step $emptyStep -Number 1 -Total 1 -Token 'testtoken'
Assert-Mb ((Get-MbAnnotationsJson -Html $emptyHtml) -eq '[]') '注釈が0件のときは空配列を出力する'

$resultStep = New-MbTestStep -AnnotationCount 0
$resultStep | Add-Member -NotePropertyName resultImageId -NotePropertyValue ('image-' + ([guid]::NewGuid().ToString('N')))
$resultStep | Add-Member -NotePropertyName imageLayout -NotePropertyValue 'side-by-side'
$resultStep | Add-Member -NotePropertyName imageOrder -NotePropertyValue 'after-before'
$resultHtml = ConvertTo-MbStepCardHtml -Step $resultStep -Number 1 -Total 1 -Token 'testtoken'
Assert-Mb ($resultHtml -match 'step-visual-item--before' -and $resultHtml -match 'step-visual-item--after') '操作前と操作後を同じ手順カードに表示する'
Assert-Mb ($resultHtml -match [regex]::Escape([string]$resultStep.resultImageId)) '結果画像を操作前画像と別に参照する'
Assert-Mb ($resultHtml -match 'step-visual-layout--side-by-side step-visual-layout--reverse') '選んだ左右配置と前後順を表示へ反映する'
Assert-Mb (([regex]::Matches($resultHtml, 'data-image-layout-option=')).Count -eq 4) '4種類の画像配置を選べる'
Assert-Mb ($resultHtml -match 'data-swap-image-order' -and $resultHtml -match 'data-remove-result-image') '比較画像の順序変更と取り外しができる'
Assert-Mb ($emptyHtml -match 'data-add-result-image') '1枚の手順から比較画像を追加できる'

Write-Host ''
Write-Host '--- 撮影監視の状態表示 ---' -ForegroundColor Cyan

# 狭い画面では文言を畳んでアイコンだけにする。文言は別の要素へ入れ、
# 状態名をツールチップの先頭へ置かないと、畳んだときに状態が分からなくなる（UX-04・UX-09）。
$watchHtml = ConvertTo-MbWatchStatusHtml -State 'active' -Role 'owner' -Directory 'C:\shots'
Assert-Mb ($watchHtml -match 'class="watch-status__label"') '監視状態の文言を畳める要素へ入れる'
Assert-Mb ($watchHtml -match 'title="監視中・このタブに追加 /') '状態名をツールチップの先頭へ入れる'
Assert-Mb ($watchHtml -match 'C:\\shots') 'ツールチップに保存先も残す'

$disabledHtml = ConvertTo-MbWatchStatusHtml -State 'disabled' -Role 'available'
Assert-Mb ($disabledHtml -match 'title="自動監視なし・貼り付け利用可 /') '保存先が無い場合も状態名をツールチップへ入れる'

Write-Host ''
Write-Host '--- 手順アウトラインの状態 ---' -ForegroundColor Cyan

function New-MbTestSheet {
    param([object[]]$Steps)
    return [pscustomobject]@{ id = 'sheet-1'; name = 'テストシート'; steps = @($Steps) }
}

$completeStep = New-MbTestStep -AnnotationCount 0
$incompleteStep = New-MbTestStep -AnnotationCount 0
$incompleteStep.description = ''
$imagelessStep = New-MbTestStep -AnnotationCount 0
$imagelessStep.imageId = ''

$testSheet = New-MbTestSheet -Steps @($completeStep, $incompleteStep, $imagelessStep)
$testProject = [pscustomobject]@{ selectedSheetId = $testSheet.id; sheets = @($testSheet) }
# Render-MbStepNavigation は公開していない補助関数のため、モジュールの内側で呼ぶ。
$navHtml = & (Get-Module ManualBuilder.Web) {
    param($Sheet, $Project)
    Render-MbStepNavigation -Sheet $Sheet -Project $Project
} $testSheet $testProject
# aria-label は role の無い span では支援技術へ届かない。目印には必ず role="img" を付ける（UX-10）。
Assert-Mb ($navHtml -notmatch '<span class="step-nav__status"(?![^>]*role=")') '未完了の目印にrole="img"を付ける'
Assert-Mb ($navHtml -match 'aria-label="説明未入力"') '説明未入力の手順に読み上げ可能な目印を出す'
Assert-Mb ($navHtml -match 'aria-label="画像なし"') '画像なしの手順に読み上げ可能な目印を出す'
Assert-Mb (([regex]::Matches($navHtml, 'step-nav__status')).Count -eq 2) '入力済みの手順には目印を出さない'
# ドラッグしか案内していないと、キーボードだけを使う人が並べ替えに気付けない（UX-08）。
Assert-Mb ($navHtml -match 'data-step-nav-drag-handle[^>]*↑↓ キー') '手順の取っ手がキーボード操作を案内する'
Assert-Mb ($navHtml -match 'data-step-select-mode[^>]*>複数選択') '複数選択を明示的に開始できる'
Assert-Mb ($navHtml -match 'data-step-selection-all[^>]*>すべて選択') '手順をすべて選択できる'
Assert-Mb (([regex]::Matches($navHtml, 'data-step-bulk-order=')).Count -eq 4) '選択した手順を一括で上下・先頭・末尾へ移動できる'

Write-Host ''
Write-Host '--- 録画から手順書を作る主導線 ---' -ForegroundColor Cyan

function New-MbTestProject {
    param([object[]]$Steps)
    $sheet = New-MbTestSheet -Steps $Steps
    return [pscustomobject]@{
        id              = 'project-1'
        title           = 'テストマニュアル'
        revision        = 1
        selectedSheetId = $sheet.id
        sheets          = @($sheet)
        videos          = @()
    }
}

$emptyWorkspaceHtml = ConvertTo-MbWorkspaceHtml -Project (New-MbTestProject -Steps @()) -Token 'testtoken'
Assert-Mb ($emptyWorkspaceHtml -match '操作を記録して、手順書を作る') '空の画面で主機能を成果が分かる見出しにする'
Assert-Mb ($emptyWorkspaceHtml -match 'empty-state__main-button[^>]*data-record-operations[^>]*>操作の記録を開始') '空の画面で操作記録を主ボタンにする'
Assert-Mb (([regex]::Matches($emptyWorkspaceHtml, 'button button--primary[^>]*data-record-operations')).Count -eq 1) '空の画面では操作記録の主ボタンを重複させない'
Assert-Mb ($emptyWorkspaceHtml -match 'data-open-video-picker>録画ファイルを取り込む') '既存録画の取り込みを同じ画面から選べる'
Assert-Mb (([regex]::Matches($emptyWorkspaceHtml, 'data-open-video-picker>録画ファイルを取り込む')).Count -eq 1) '空の画面では既存録画の入口を重複させない'
Assert-Mb ($emptyWorkspaceHtml -match '操作を記録.+使う操作を確認.+Copilotの提案を確認') '主機能の3段階を最初に示す'
Assert-Mb ($emptyWorkspaceHtml -match 'accept="video/mp4,video/webm"') '主ボタンから選べる動画形式を制限する'

$filledWorkspaceHtml = ConvertTo-MbWorkspaceHtml -Project (New-MbTestProject -Steps @((New-MbTestStep -AnnotationCount 0))) -Token 'testtoken'
Assert-Mb ($filledWorkspaceHtml -notmatch 'class="empty-state') '手順がある画面では開始案内を重複表示しない'
Assert-Mb ($filledWorkspaceHtml -match 'topbar__main-action[^>]*data-record-operations[^>]*>操作を記録') '編集中も主機能へスクロールせず戻れる'
Assert-Mb (([regex]::Matches($filledWorkspaceHtml, 'button button--primary[^>]*data-record-operations')).Count -eq 1) '編集中も操作記録の主ボタンを重複させない'
Assert-Mb ($filledWorkspaceHtml -match 'topbar__video-action[^>]*data-open-video-picker[^>]*>録画を取り込む') '編集中も既存録画をメニューを開かず取り込める'

Write-Host ''
Write-Host '--- Copilot後の仕上げ導線 ---' -ForegroundColor Cyan
Assert-Mb ($filledWorkspaceHtml -match 'data-finish-guide') '編集画面に仕上げ状況を常時表示する'
Assert-Mb ($filledWorkspaceHtml -match 'data-finish-check="text"') '説明なしの手順へ移動できる'
Assert-Mb ($filledWorkspaceHtml -match 'data-finish-check="annotation"') '赤枠・番号なしの手順を確認できる'
Assert-Mb ($filledWorkspaceHtml -match 'data-finish-check="attention"') 'Copilot後の要確認手順へ移動できる'
Assert-Mb ($filledWorkspaceHtml -match 'data-add-step-after[^>]*>＋ 手順をこの後に追加') '現在の手順の直後へ追加できる入口を表示する'
Assert-Mb (([regex]::Matches($filledWorkspaceHtml, 'data-open-export-dialog')).Count -eq 2) '上部と仕上げ欄から最終出力へ進める'
$cardHtml = ConvertTo-MbStepCardHtml -Step (New-MbTestStep -AnnotationCount 0) -Number 2 -Total 3 -Token 'testtoken'
Assert-Mb (([regex]::Matches($cardHtml, 'data-step-move=')).Count -eq 2) '手順カードから個別に上下移動できる'
Assert-Mb ($cardHtml -match 'data-step-card-delete[^>]*>削除') '手順カードから削除できる'
Assert-Mb ($cardHtml -match 'data-open-annotation[^>]*[\s\S]*赤枠・番号を追加') '赤枠・番号の入口を具体的な名前で表示する'
Assert-Mb ($cardHtml -match 'annotation-badge[^>]*>注釈 <span class="annotation-count">0') '注釈がない手順も状態を表示する'

$reviewStep = New-MbTestStep -AnnotationCount 0
$reviewStep.review = [pscustomobject]@{ required = $true; action = 'review'; reason = '赤枠の候補を特定できませんでした。' }
$reviewCardHtml = ConvertTo-MbStepCardHtml -Step $reviewStep -Number 1 -Total 1 -Token 'testtoken'
Assert-Mb ($reviewCardHtml -match 'data-step-review-notice') '要確認を手順カード上で見落とさない'
Assert-Mb ($reviewCardHtml -match 'data-step-review-resolve[^>]*>確認済みにする') '明示操作でだけ要確認を解決できる'

$secondSheetStep = New-MbTestStep -AnnotationCount 1
$secondSheetStep.description = ''
$secondSheetStep.annotations[0].type = 'arrow'
$secondSheetStep.review = [pscustomobject]@{ required = $true; action = 'review'; reason = '文章を確認してください。' }
$multiSheetProject = New-MbTestProject -Steps @((New-MbTestStep -AnnotationCount 1))
$secondSheet = New-MbTestSheet -Steps @($secondSheetStep)
$secondSheet.id = 'sheet-2'
$secondSheet.name = '未完成シート'
$multiSheetProject.sheets = @($multiSheetProject.sheets[0], $secondSheet)
$multiSheetHtml = ConvertTo-MbWorkspaceHtml -Project $multiSheetProject -Token 'testtoken'
Assert-Mb (([regex]::Matches($multiSheetHtml, 'data-sheet-delete')).Count -eq 1) '選択シートの削除を取り消し対応の操作として表示する'
$finishMatch = [regex]::Match($multiSheetHtml, '<textarea hidden data-project-finish-data>(.*?)</textarea>')
Assert-Mb $finishMatch.Success '全シートの仕上げ情報を画面へ埋め込む'
$finishItems = ([Net.WebUtility]::HtmlDecode($finishMatch.Groups[1].Value) | ConvertFrom-Json)
Assert-Mb (@($finishItems).Count -eq 2) '仕上げ情報が別シートの手順も含む'
$secondFinish = @($finishItems | Where-Object { $_.sheetId -eq 'sheet-2' })[0]
Assert-Mb ([bool]$secondFinish.missingText) '別シートの説明未入力を出力前確認へ渡す'
Assert-Mb (-not [bool]$secondFinish.hasFocusAnnotation) '矢印だけを赤枠・番号ありとして数えない'
Assert-Mb ([bool]$secondFinish.reviewRequired -and [string]$secondFinish.reviewAction -eq 'review') '別シートの要確認を再読込後も仕上げ確認へ渡す'

Write-Host ''
Write-Host 'Web rendering tests passed.' -ForegroundColor Green
exit 0
