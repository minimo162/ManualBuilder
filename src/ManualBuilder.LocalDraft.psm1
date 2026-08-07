# 操作記録の事実から、外部AIを使わずに手順文を作る。
#
# このモジュールは文章の推測を広げず、操作種別・対象名・実クリック位置という
# ローカルで確認できる情報だけを使う。対象が曖昧でも手順を削除せず、要確認として残す。

Set-StrictMode -Version 2.0

function ConvertTo-MbLocalTargetLabel {
    param([AllowEmptyString()][string]$Value = '')

    $label = if ($null -eq $Value) { '' } else { [string]$Value }
    $label = $label.Replace("`r`n", ' ').Replace("`r", ' ').Replace("`n", ' ')
    $label = [regex]::Replace($label, '\s+', ' ').Trim()
    $label = [regex]::Replace($label, '\s*（(?:入力|右クリック)）\s*$', '')
    if ($label.Length -gt 80) { $label = $label.Substring(0, 79).TrimEnd() + '…' }
    return $label
}

function Get-MbLocalStepDraft {
    param(
        [AllowEmptyString()][string]$ActionKind = 'click',
        [AllowEmptyString()][string]$TargetName = '',
        [AllowEmptyString()][string]$TargetType = '',
        [AllowEmptyString()][string]$WindowTitle = '',
        [AllowEmptyString()][string]$TargetSource = '',
        [AllowEmptyString()][string]$TargetConfidence = ''
    )

    $label = ConvertTo-MbLocalTargetLabel -Value $TargetName
    $type = [string]$TargetType
    $kind = ([string]$ActionKind).ToLowerInvariant()
    $window = [string]$WindowTitle
    $hasLabel = -not [string]::IsNullOrWhiteSpace($label)
    $title = ''
    $description = ''

    if ($kind -in @('input', 'recorded-input')) {
        if ($type -eq 'ControlType.DataItem' -and $window -match '(?i)Excel' -and
            $hasLabel -and $label -match '^[A-Z]{1,3}\d+$') {
            $title = 'セル' + $label + 'に入力'
            $description = 'セル' + $label + 'に必要な内容を入力します。'
        } elseif ($hasLabel) {
            $title = $label + 'に入力'
            $description = '［' + $label + '］に必要な内容を入力します。'
        } else {
            $title = '内容を入力'
            $description = '赤枠で示した入力欄に必要な内容を入力します。'
        }
    } elseif ($kind -in @('right-click', 'recorded-right-click')) {
        if ($hasLabel) {
            $title = $label + 'のメニューを開く'
            $description = '［' + $label + '］を右クリックします。'
        } else {
            $title = 'メニューを開く'
            $description = '赤枠で示した位置を右クリックします。'
        }
    } elseif ($kind -in @('double-click', 'recorded-double-click')) {
        if ($hasLabel) {
            $title = $label + 'を開く'
            $description = '［' + $label + '］をダブルクリックします。'
        } else {
            $title = '項目を開く'
            $description = '赤枠で示した位置をダブルクリックします。'
        }
    } elseif ($type -eq 'ControlType.CheckBox') {
        $title = if ($hasLabel) { $label + 'を切り替える' } else { '設定を切り替える' }
        $description = if ($hasLabel) { '［' + $label + '］をオンまたはオフにします。' } else { '赤枠で示した設定をオンまたはオフにします。' }
    } elseif ($type -eq 'ControlType.RadioButton') {
        $title = if ($hasLabel) { $label + 'を選択' } else { '選択肢を選ぶ' }
        $description = if ($hasLabel) { '［' + $label + '］を選択します。' } else { '赤枠で示した選択肢を選択します。' }
    } elseif ($type -in @('ControlType.MenuItem', 'ControlType.ListItem', 'ControlType.ComboBox')) {
        $title = if ($hasLabel) { $label + 'を選択' } else { '項目を選択' }
        $description = if ($hasLabel) { '［' + $label + '］を選択します。' } else { '赤枠で示した項目を選択します。' }
    } elseif ($type -eq 'ControlType.TabItem') {
        $tabLabel = if ($hasLabel -and $label -notmatch 'タブ$') { $label + 'タブ' } else { $label }
        $title = if ($hasLabel) { $tabLabel + 'を開く' } else { 'タブを開く' }
        $description = if ($hasLabel) { '［' + $tabLabel + '］をクリックします。' } else { '赤枠で示したタブをクリックします。' }
    } elseif ($type -eq 'ControlType.DataItem' -and $window -match '(?i)Excel') {
        if ($hasLabel -and $label -match '^[A-Z]{1,3}\d+$') {
            $title = 'セル' + $label + 'を選択'
            $description = 'セル' + $label + 'をクリックします。'
        } elseif ($hasLabel) {
            $title = $label + 'を選択'
            $description = '［' + $label + '］を選択します。'
        } else {
            $title = 'セルを選択'
            $description = '赤枠で示したセルをクリックします。'
        }
    } elseif ($type -in @('ControlType.Button', 'ControlType.SplitButton', 'ControlType.Hyperlink')) {
        $title = if ($hasLabel) { $label } else { '項目を実行' }
        $description = if ($hasLabel) { '［' + $label + '］をクリックします。' } else { '赤枠で示した位置をクリックします。' }
    } else {
        $title = if ($hasLabel) { $label } else { '画面上の項目を選択' }
        $description = if ($hasLabel) { '［' + $label + '］をクリックします。' } else { '赤枠で示した位置をクリックします。' }
    }

    $reasonCodes = New-Object System.Collections.ArrayList
    if (-not $hasLabel) { [void]$reasonCodes.Add('TARGET_UNKNOWN') }
    $lowEvidence = [string]$TargetConfidence -eq 'low' -or [string]$TargetSource -eq 'click-point' -or
        [string]$TargetType -eq 'ControlType.ClickPoint'
    if ($lowEvidence -and -not $reasonCodes.Contains('TARGET_LOW_CONFIDENCE')) {
        [void]$reasonCodes.Add('TARGET_LOW_CONFIDENCE')
    }
    $reviewReason = ''
    if ($reasonCodes.Contains('TARGET_UNKNOWN')) {
        $reviewReason = '操作対象を特定できませんでした。赤枠と文章を確認してください。'
    } elseif ($reasonCodes.Contains('TARGET_LOW_CONFIDENCE')) {
        $reviewReason = 'クリック位置から作成した手順です。赤枠が操作箇所を示しているか確認してください。'
    }

    return [pscustomobject]@{
        title = $title
        description = $description
        note = ''
        reviewRequired = $reasonCodes.Count -gt 0
        reviewReason = $reviewReason
        reasonCodes = @($reasonCodes)
        source = 'local'
        version = '1'
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-MbLocalTargetLabel',
    'Get-MbLocalStepDraft'
)
