# =====================================================================
# 02-known-folders.ps1  —  スクリーンショット保存先の検出（V-9 / V-18）
#
# Snipping Tool の保存先は Pictures 既知フォルダに従い、OneDrive の画像
# 自動保存が有効だと OneDrive 配下に切り替わります。アプリ独自の保存先
# 設定は存在しないため、固定パス依存は危険です。
#
# このスクリプトは解決過程を全部表示し、実際に撮った結果と照合します。
#
# 使い方:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\02-known-folders.ps1
# =====================================================================
$ErrorActionPreference = 'Continue'

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  V-9 / V-18: スクリーンショット保存先の検出' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan

$candidates = New-Object System.Collections.ArrayList
function Add-Cand {
    param([string]$Source, [string]$Raw, [string]$Path)
    $exists = $false
    $count = 0
    if ($Path) {
        $exists = Test-Path -LiteralPath $Path
        if ($exists) {
            $count = @(Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue |
                       Where-Object { $_.Extension -match '^\.(png|jpg|jpeg|bmp)$' }).Count
        }
    }
    [void]$candidates.Add([pscustomobject]@{
        取得元 = $Source; 生の値 = $Raw; 解決後 = $Path
        存在 = $(if ($exists) { 'あり' } else { 'なし' }); 画像数 = $count
    })
}

# ---------------------------------------------------------------------
# 1. Screenshots 既知フォルダ（レジストリ）
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '[1] Screenshots 既知フォルダ（レジストリ）' -ForegroundColor Cyan
$guid = '{B7BEDE81-DF94-4682-A7D8-57A52620B86F}'
foreach ($key in @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')) {
    $raw = $null
    try {
        $raw = (Get-ItemProperty -Path $key -Name $guid -ErrorAction Stop).$guid
    } catch { }
    if ($raw) {
        $expanded = [Environment]::ExpandEnvironmentVariables($raw)
        Write-Host ("  {0}" -f (Split-Path $key -Leaf)) -ForegroundColor Gray
        Write-Host ("    生の値 : {0}" -f $raw)
        Write-Host ("    展開後 : {0}" -f $expanded)
        Add-Cand (Split-Path $key -Leaf) $raw $expanded
    } else {
        Write-Host ("  {0} → 値なし" -f (Split-Path $key -Leaf)) -ForegroundColor DarkGray
    }
}

# ---------------------------------------------------------------------
# 2. Pictures 既知フォルダ配下
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '[2] Pictures 既知フォルダ配下' -ForegroundColor Cyan
$pics = [Environment]::GetFolderPath('MyPictures')
Write-Host ("  MyPictures : {0}" -f $pics)
foreach ($name in @('Screenshots', 'スクリーンショット')) {
    Add-Cand "MyPictures\$name" $name (Join-Path $pics $name)
}

# ---------------------------------------------------------------------
# 3. OneDrive 配下
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '[3] OneDrive 環境変数' -ForegroundColor Cyan
$odVars = [ordered]@{
    OneDrive = $env:OneDrive
    OneDriveCommercial = $env:OneDriveCommercial
    OneDriveConsumer = $env:OneDriveConsumer
}
foreach ($k in $odVars.Keys) {
    Write-Host ("  {0,-20} {1}" -f $k, $(if ($odVars[$k]) { $odVars[$k] } else { '(未設定)' }))
}
foreach ($k in $odVars.Keys) {
    $base = $odVars[$k]
    if (-not $base) { continue }
    foreach ($name in @('Pictures\Screenshots', 'ピクチャ\スクリーンショット',
                        'Pictures\スクリーンショット', 'ドキュメント\..\Pictures\Screenshots')) {
        $p = Join-Path $base $name
        try { $p = [IO.Path]::GetFullPath($p) } catch { continue }
        Add-Cand "$k\$name" $name $p
    }
}

# ---------------------------------------------------------------------
# 4. 候補の一覧
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '[4] 候補の一覧' -ForegroundColor Cyan
$candidates | Format-Table -AutoSize -Wrap

$best = $candidates | Where-Object { $_.存在 -eq 'あり' } |
        Sort-Object -Property 画像数 -Descending | Select-Object -First 1
if ($best) {
    Write-Host ("  推定される保存先: {0}  （画像 {1} 件）" -f $best.解決後, $best.画像数) -ForegroundColor Green
} else {
    Write-Host '  存在する候補が見つかりませんでした。' -ForegroundColor Yellow
    Write-Host '  → Snipping Tool の設定で「スクリーンショットを自動保存する」がOFFの可能性があります。' -ForegroundColor Yellow
}

# ---------------------------------------------------------------------
# 5. 実測: 実際に撮ってどこに保存されるか（V-9d）
# ---------------------------------------------------------------------
Write-Host ''
Write-Host '[5] 実測 — 実際に撮って照合します' -ForegroundColor Cyan
Write-Host ''
Write-Host '  これから 15 秒待ちます。その間に Win+Shift+S でスクリーンショットを' -ForegroundColor Yellow
Write-Host '  1枚撮ってください（範囲はどこでもかまいません）。' -ForegroundColor Yellow
Write-Host ''

$since = Get-Date
$searchDirs = @($candidates | Where-Object { $_.存在 -eq 'あり' } | Select-Object -ExpandProperty 解決後 -Unique)
if ($searchDirs.Count -eq 0) { $searchDirs = @($pics) }

for ($s = 15; $s -ge 1; $s--) {
    Write-Host ("`r  残り {0,2} 秒 ... " -f $s) -NoNewline -ForegroundColor Yellow
    Start-Sleep -Seconds 1
}
Write-Host ''
Write-Host '  検索します（数秒かかります）...' -ForegroundColor Gray

# 候補フォルダ＋Pictures配下を再帰で探す
$found = New-Object System.Collections.ArrayList
$scanRoots = @($searchDirs + @($pics) + @($odVars.Values | Where-Object { $_ })) | Select-Object -Unique
foreach ($root in $scanRoots) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    try {
        Get-ChildItem -LiteralPath $root -File -Recurse -Depth 3 -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Extension -match '^\.(png|jpg|jpeg|bmp)$' -and $_.CreationTime -gt $since
            } | ForEach-Object {
                [void]$found.Add([pscustomobject]@{
                    ファイル = $_.Name
                    フォルダ = $_.DirectoryName
                    作成時刻 = $_.CreationTime.ToString('HH:mm:ss')
                    サイズKB = [int]($_.Length / 1KB)
                    属性 = "$($_.Attributes)"
                })
            }
    } catch { }
}

Write-Host ''
if ($found.Count -eq 0) {
    Write-Host '  新しい画像が見つかりませんでした。' -ForegroundColor Yellow
    Write-Host '  考えられる原因:' -ForegroundColor Yellow
    Write-Host '   ・Snipping Tool の「スクリーンショットを自動保存する」がOFF'
    Write-Host '     （Snipping Tool を開く → 右上の … → 設定 で確認できます）'
    Write-Host '   ・撮影しなかった / 撮影がキャンセルされた'
    Write-Host '   ・保存先が上の候補以外の場所'
    Write-Host ''
    Write-Host '  → 自動取り込み（V-7）は使えず、Ctrl+V 貼り付けが主経路になります。' -ForegroundColor Yellow
} else {
    Write-Host ("  新しい画像が {0} 件見つかりました:" -f $found.Count) -ForegroundColor Green
    $found | Format-Table -AutoSize -Wrap

    # V-18b: OneDrive プレースホルダかどうか
    $ph = @($found | Where-Object { $_.属性 -match 'Offline|RecallOnDataAccess|ReparsePoint' })
    if ($ph.Count -gt 0) {
        Write-Host '  ⚠ OneDrive のプレースホルダ属性が付いたファイルがあります。' -ForegroundColor Yellow
        Write-Host '    実体が来る前に読むと 0 バイトになるため、保留処理が必要です。' -ForegroundColor Yellow
    } else {
        Write-Host '  プレースホルダ属性はありません（実体がローカルにあります）。' -ForegroundColor Green
    }

    # 推定と一致したか
    $dirs = @($found | Select-Object -ExpandProperty フォルダ -Unique)
    if ($best -and ($dirs -contains $best.解決後)) {
        Write-Host ''
        Write-Host '  ✓ 推定した保存先と一致しました。自動検出は機能します。' -ForegroundColor Green
    } else {
        Write-Host ''
        Write-Host '  ✗ 推定した保存先と一致しませんでした。' -ForegroundColor Red
        Write-Host ("    推定: {0}" -f $(if ($best) { $best.解決後 } else { '(なし)' }))
        Write-Host ("    実際: {0}" -f ($dirs -join ' / '))
        Write-Host '    → 初回起動時にユーザーへ選ばせる方式にします。' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------
# レポート
# ---------------------------------------------------------------------
$report = Join-Path $PSScriptRoot 'result-02-folders.txt'
(($candidates | Format-Table -AutoSize -Wrap | Out-String -Width 200) + "`r`n実測:`r`n" +
 ($found | Format-Table -AutoSize -Wrap | Out-String -Width 200)) |
    Out-File -FilePath $report -Encoding UTF8
Write-Host ''
Write-Host "  レポート: $report" -ForegroundColor Cyan
Write-Host ''
