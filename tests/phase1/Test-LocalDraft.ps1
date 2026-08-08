# ローカル文章化は、操作記録の事実だけから安定した初稿を作る。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$modulePath = Join-Path $repoRoot 'src\ManualBuilder.LocalDraft.psm1'
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

Import-Module $modulePath -Force

$button = Get-MbLocalStepDraft -ActionKind 'click' -TargetName '保存' `
    -TargetType 'ControlType.Button' -WindowTitle 'Book1 - Excel' -TargetSource 'UIA' -TargetConfidence 'high'
Add-Result ($button.title -eq '保存') 'ボタン名を手順名にする'
Add-Result ($button.description -eq '［保存］をクリックします。') 'ボタン操作を短い日本語にする'
Add-Result (-not $button.reviewRequired) '根拠のある操作は要確認にしない'

$input = Get-MbLocalStepDraft -ActionKind 'input' -TargetName '検索（入力）' `
    -TargetType 'ControlType.Edit' -WindowTitle 'エクスプローラー' -TargetSource 'UIA' -TargetConfidence 'medium'
Add-Result ($input.title -eq '検索に入力') '入力欄の補足を重ねず手順名を作る'
Add-Result ($input.description -eq '［検索］に必要な内容を入力します。') '入力値を保存せず入力操作を説明する'
Add-Result ($input.reviewRequired -and $input.reasonCodes -contains 'TARGET_MEDIUM_CONFIDENCE') `
    '中信頼の対象候補は確認済みにせず要確認へ送る'

$cell = Get-MbLocalStepDraft -ActionKind 'click' -TargetName 'F8' `
    -TargetType 'ControlType.DataItem' -WindowTitle 'Book1 - Excel' -TargetSource 'UIA' -TargetConfidence 'high'
Add-Result ($cell.title -eq 'セルF8を選択') 'Excelセルをセル番地として説明する'
$cellInput = Get-MbLocalStepDraft -ActionKind 'input' -TargetName 'B2（入力）' `
    -TargetType 'ControlType.DataItem' -WindowTitle 'Book1 - Excel' -TargetSource 'UIA' -TargetConfidence 'high'
Add-Result ($cellInput.title -eq 'セルB2に入力' -and $cellInput.description -eq 'セルB2に必要な内容を入力します。') `
    'Excelセルへの入力をセル番地つきで説明する'
Add-Result ($cellInput.reviewRequired -and $cellInput.reasonCodes -contains 'CONTENT_NOT_RECORDED') `
    '入力内容を保存しない手順は文章を補えるよう要確認へ送る'

$fallback = Get-MbLocalStepDraft -ActionKind 'click' -TargetName '' `
    -TargetType 'ControlType.ClickPoint' -WindowTitle '申請画面 - Edge' -TargetSource 'click-point' -TargetConfidence 'low'
Add-Result ($fallback.title -eq '画面上の項目を選択') '対象不明でも空の手順を作らない'
Add-Result ($fallback.reviewRequired -and $fallback.reasonCodes -contains 'TARGET_UNKNOWN') '対象不明は削除せず要確認にする'

if ($errors.Count -gt 0) {
    Write-Host "`n$($errors.Count) 件の失敗" -ForegroundColor Red
    exit 1
}
Write-Host "`nローカル文章化テストに合格しました。" -ForegroundColor Green
