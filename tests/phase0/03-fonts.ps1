# =====================================================================
# 03-fonts.ps1  —  フォントの存在確認（V-13）
#
# BIZ UDPゴシックが「BIZ UDPゴシック」「BIZ UDPGothic」のどちらで
# 登録されているかは環境によって異なります。実測して確定させます。
#
# 使い方:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\03-fonts.ps1
# =====================================================================
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  V-13: フォントの存在確認' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''

$installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families |
             ForEach-Object { $_.Name }
Write-Host ("  インストール済みフォント数: {0}" -f $installed.Count)
Write-Host ''

$targets = @(
    'BIZ UDPゴシック', 'BIZ UDPGothic',
    'BIZ UDゴシック', 'BIZ UDGothic',
    'BIZ UDP明朝 Medium', 'BIZ UDPMincho',
    'Meiryo', 'Meiryo UI', 'メイリオ',
    'Yu Gothic UI', 'Yu Gothic', '游ゴシック',
    'MS Pゴシック', 'MS PGothic',
    'Segoe UI'
)

$rows = New-Object System.Collections.ArrayList
foreach ($t in $targets) {
    $has = $installed -contains $t
    [void]$rows.Add([pscustomobject]@{
        フォント名 = $t
        状態 = $(if ($has) { '存在する' } else { '無い' })
    })
}
$rows | Format-Table -AutoSize

# ---------------------------------------------------------------------
# UD 系を名前で総ざらい（表記ゆれの実測）
# ---------------------------------------------------------------------
Write-Host '  「BIZ」を含むフォント（実際の登録名）:' -ForegroundColor Cyan
$biz = @($installed | Where-Object { $_ -like '*BIZ*' } | Sort-Object)
if ($biz.Count -eq 0) {
    Write-Host '   （見つかりません）' -ForegroundColor Yellow
} else {
    $biz | ForEach-Object { Write-Host "   - $_" -ForegroundColor Green }
}
Write-Host ''

# ---------------------------------------------------------------------
# フォールバックの動作確認
# ---------------------------------------------------------------------
function Resolve-BodyFont {
    param([string]$Preferred = 'BIZ UDPゴシック')
    foreach ($f in @($Preferred, 'BIZ UDPゴシック', 'BIZ UDPGothic', 'BIZ UDゴシック',
                     'Meiryo', 'Yu Gothic UI', 'MS Pゴシック')) {
        if ($f -and ($installed -contains $f)) { return $f }
    }
    return 'MS Pゴシック'
}

Write-Host '  フォールバックの動作確認:' -ForegroundColor Cyan
foreach ($pref in @('BIZ UDPゴシック', 'BIZ UDPGothic', '存在しないフォント名XYZ', '')) {
    $r = Resolve-BodyFont -Preferred $pref
    Write-Host ("   希望「{0,-22}」 → 採用「{1}」" -f $(if ($pref) { $pref } else { '(空)' }), $r)
}
Write-Host ''

$decided = Resolve-BodyFont
Write-Host ("  この環境で使う本文フォント: {0}" -f $decided) -ForegroundColor Green
if ($decided -notlike 'BIZ*') {
    Write-Host '  → BIZ UD 系が無いため、代替フォントになります（動作に問題はありません）。' -ForegroundColor Yellow
}

$report = Join-Path $PSScriptRoot 'result-03-fonts.txt'
(($rows | Format-Table -AutoSize | Out-String -Width 120) +
 "`r`nBIZ を含む登録名:`r`n" + ($biz -join "`r`n") +
 "`r`n`r`n採用フォント: $decided`r`n") | Out-File -FilePath $report -Encoding UTF8
Write-Host ''
Write-Host "  レポート: $report" -ForegroundColor Cyan
Write-Host ''
