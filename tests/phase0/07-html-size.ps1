# =====================================================================
# 07-html-size.ps1  —  自己完結HTMLの容量と表示速度（V-19）
#
# base64 埋め込みは約33%増えるため、ステップが増えると単一HTMLが
# 実用的でなくなります。切り替え閾値（htmlEmbedLimitMB）を実測で決めます。
#
# 使い方:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\07-html-size.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\07-html-size.ps1 -Steps 10,50,100,200
# =====================================================================
[CmdletBinding()]
param(
    [int[]]$Steps = @(10, 50, 100),
    [int]$ImageW = 1600,
    [int]$ImageH = 900
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$outDir = Join-Path $PSScriptRoot 'out\html'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  V-19: 自己完結HTMLの容量と生成時間' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''

# ---------------------------------------------------------------------
# 代表画像を1枚作る（実際のスクショに近いサイズになるよう作る）
# ---------------------------------------------------------------------
function New-SampleImage {
    param([string]$Path, [int]$W, [int]$H)
    $bmp = New-Object System.Drawing.Bitmap $W, $H
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $pen = $null; $br = $null; $fnt = $null
    try {
        $g.Clear([System.Drawing.Color]::FromArgb(246, 247, 249))
        $g.FillRectangle([System.Drawing.Brushes]::White, 24, 24, $W - 48, $H - 48)
        $brT = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(60, 90, 150))
        $g.FillRectangle($brT, 24, 24, $W - 48, 52)
        $brT.Dispose()
        $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(150, 155, 160)), 2
        $br = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(40, 40, 40))
        $fnt = New-Object System.Drawing.Font 'Meiryo', 13
        # 文字と枠を大量に描いて、実際のスクショに近い圧縮率にする
        for ($y = 110; $y -lt ($H - 60); $y += 34) {
            $g.DrawRectangle($pen, 60, $y, [int](($W - 160) * 0.75), 26)
            $g.DrawString("項目 $y : 入力してください  ABCDEFG 0123456789 あいうえお", $fnt, $br, 70, [float]($y + 4))
        }
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        foreach ($d in @($pen, $br, $fnt)) { if ($d) { $d.Dispose() } }
        $g.Dispose(); $bmp.Dispose()
    }
}

$sample = Join-Path $outDir 'sample.png'
Write-Host '  代表画像を生成中...' -ForegroundColor Gray
New-SampleImage -Path $sample -W $ImageW -H $ImageH
$imgBytes = [IO.File]::ReadAllBytes($sample)
$imgKB = [int]($imgBytes.Length / 1KB)
Write-Host ("  代表画像: {0}x{1} / {2} KB" -f $ImageW, $ImageH, $imgKB)
Write-Host ''

$b64 = [Convert]::ToBase64String($imgBytes)
$b64KB = [int]($b64.Length / 1KB)
$overhead = [Math]::Round(100.0 * $b64.Length / $imgBytes.Length - 100, 1)
Write-Host ("  base64 化すると {0} KB（+{1}%）" -f $b64KB, $overhead) -ForegroundColor Yellow
Write-Host ''

# ---------------------------------------------------------------------
# HTML 生成（埋め込み方式 / フォルダ方式）
# ---------------------------------------------------------------------
Add-Type -AssemblyName System.Web

function New-ManualHtml {
    param([int]$StepCount, [string]$ImgSrcTemplate, [string]$Title)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">')
    [void]$sb.AppendLine("<title>$([System.Web.HttpUtility]::HtmlEncode($Title))</title>")
    [void]$sb.AppendLine(@'
<style>
body { font-family: "BIZ UDPGothic","Meiryo","Yu Gothic UI",sans-serif; font-size:15px;
       line-height:1.7; color:#222; max-width:900px; margin:0 auto; padding:32px; }
h1 { font-size:26px; border-bottom:3px solid #1f3864; padding-bottom:8px; }
h2 { font-size:19px; margin-top:36px; color:#1f3864; }
.step { break-inside: avoid; page-break-inside: avoid; margin-bottom:28px; }
.note { background:#fff9db; border:1px solid #d6aa00; border-radius:6px; padding:10px 14px; margin:10px 0; }
img { max-width:100%; border:1px solid #ccc; border-radius:4px; display:block; margin-top:10px; }
@media print { body { padding:0; } h2 { page-break-after: avoid; } }
</style></head><body>
'@)
    [void]$sb.AppendLine("<h1>$([System.Web.HttpUtility]::HtmlEncode($Title))</h1>")
    for ($i = 1; $i -le $StepCount; $i++) {
        $src = $ImgSrcTemplate -replace '\{N\}', $i
        [void]$sb.AppendLine('<div class="step">')
        [void]$sb.AppendLine("<h2>手順 $i  ログイン画面を開く</h2>")
        [void]$sb.AppendLine('<p>ブラウザで https://example.co.jp を開き、社員番号とパスワードを入力します。</p>')
        [void]$sb.AppendLine('<div class="note">補足: パスワードは初回のみ変更が必要です</div>')
        [void]$sb.AppendLine("<img src=""$src"" alt=""手順 $i の画面"">")
        [void]$sb.AppendLine('</div>')
    }
    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$rows = New-Object System.Collections.ArrayList

foreach ($n in ($Steps | Sort-Object)) {
    # --- 埋め込み方式 ---
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $html = New-ManualHtml -StepCount $n -ImgSrcTemplate "data:image/png;base64,$b64" `
                           -Title "埋め込み方式 $n ステップ"
    $p1 = Join-Path $outDir ("embed-{0:d3}steps.html" -f $n)
    [IO.File]::WriteAllText($p1, $html, $utf8NoBom)
    $sw.Stop()
    $mb1 = [Math]::Round((Get-Item -LiteralPath $p1).Length / 1MB, 1)
    $gen1 = [Math]::Round($sw.Elapsed.TotalSeconds, 2)

    # --- フォルダ方式 ---
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $folderDir = Join-Path $outDir ("folder-{0:d3}steps" -f $n)
    $imgSub = Join-Path $folderDir 'images'
    New-Item -ItemType Directory -Force -Path $imgSub | Out-Null
    # 実運用ではステップごとに別画像だが、容量比較のため同じ画像をコピーする
    for ($i = 1; $i -le $n; $i++) {
        Copy-Item -LiteralPath $sample -Destination (Join-Path $imgSub ("s-{0:d3}.png" -f $i)) -Force
    }
    $html2 = New-ManualHtml -StepCount $n -ImgSrcTemplate 'images/s-{N}.png' `
                            -Title "フォルダ方式 $n ステップ"
    # ファイル名のゼロ埋めに合わせる
    $html2 = [regex]::Replace($html2, 'images/s-(\d+)\.png', {
        param($m) 'images/s-{0:d3}.png' -f [int]$m.Groups[1].Value })
    $p2 = Join-Path $folderDir 'index.html'
    [IO.File]::WriteAllText($p2, $html2, $utf8NoBom)
    $sw2.Stop()
    $total2 = (Get-ChildItem -LiteralPath $folderDir -Recurse -File | Measure-Object -Property Length -Sum).Sum
    $mb2 = [Math]::Round($total2 / 1MB, 1)
    $gen2 = [Math]::Round($sw2.Elapsed.TotalSeconds, 2)

    $judge = if ($mb1 -le 25) { '実用圏' } elseif ($mb1 -le 60) { '要注意' } else { '実用外' }

    [void]$rows.Add([pscustomobject]@{
        ステップ数 = $n
        埋め込みMB = $mb1
        埋め込み生成秒 = $gen1
        フォルダ計MB = $mb2
        フォルダ生成秒 = $gen2
        判定 = $judge
    })
    Write-Host ("  {0,4} ステップ: 埋め込み {1,6} MB ({2}s) / フォルダ {3,6} MB ({4}s)  → {5}" -f `
        $n, $mb1, $gen1, $mb2, $gen2, $judge) -ForegroundColor $(
        switch ($judge) { '実用圏' { 'Green' } '要注意' { 'Yellow' } default { 'Red' } })
}

Write-Host ''
$rows | Format-Table -AutoSize

# ---------------------------------------------------------------------
# 推奨閾値
# ---------------------------------------------------------------------
$last = $rows | Where-Object { $_.判定 -eq '実用圏' } | Sort-Object ステップ数 -Descending | Select-Object -First 1
Write-Host ''
if ($last) {
    Write-Host ("  埋め込み方式が実用圏なのは {0} ステップ（{1} MB）まででした。" -f $last.ステップ数, $last.埋め込みMB) -ForegroundColor Green
} else {
    Write-Host '  埋め込み方式は最小構成でも大きすぎます。フォルダ方式を既定にしてください。' -ForegroundColor Yellow
}
Write-Host '  → htmlEmbedLimitMB は「画像の合計サイズ」で判定します（base64前の値）。' -ForegroundColor Cyan
Write-Host ("     代表画像 {0} KB なら、25MB は約 {1} ステップに相当します。" -f $imgKB, [int](25 * 1024 / [Math]::Max($imgKB,1))) -ForegroundColor Cyan

Write-Host ''
Write-Host '  生成した HTML をブラウザで開いて、体感の表示速度を確認してください:' -ForegroundColor Yellow
Get-ChildItem -LiteralPath $outDir -Filter 'embed-*.html' | ForEach-Object {
    Write-Host ("   {0}  ({1} MB)" -f $_.FullName, [Math]::Round($_.Length / 1MB, 1))
}
Write-Host ''
Write-Host '  確認すること:' -ForegroundColor Yellow
Write-Host '   ・開くまでに待たされないか'
Write-Host '   ・スクロールが滑らかか'
Write-Host '   ・Ctrl+P（印刷）でプレビューが出るか。手順が途中で切れていないか'

$report = Join-Path $PSScriptRoot 'result-07-html.txt'
(($rows | Format-Table -AutoSize | Out-String -Width 160) +
 "`r`n代表画像: ${ImageW}x${ImageH} / $imgKB KB / base64 +$overhead%`r`n") |
    Out-File -FilePath $report -Encoding UTF8
Write-Host ''
Write-Host "  レポート: $report" -ForegroundColor Cyan
Write-Host ''
Write-Host '  フォルダを開きますか？ [Y/n] ' -NoNewline -ForegroundColor Yellow
if ((Read-Host) -ne 'n') { Start-Process explorer.exe $outDir }
