# Phase 1 web rendering tests (no server, no browser).

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
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

Write-Host ''
Write-Host 'Web rendering tests passed.' -ForegroundColor Green
exit 0
