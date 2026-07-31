# =====================================================================
# 04-word-layout.ps1  —  Word COM 検証（最優先）
#
# 検証項目:
#   V-4  【最重要】Word起動中はCOM生成前に中止し、ユーザーの Word を壊さないこと
#   V-3  表紙 / 目次 / 見出し / 本文 / 囲み枠 / 画像 / ページ番号 / 代替テキスト
#   V-5  日本語版Wordでスタイルの数値指定が効くか
#   V-11 ページ分断防止（KeepWithNext / KeepTogether）
#   要検証#2  囲み枠は「段落罫線」と「1x1の表」のどちらが良いか（両方出力）
#
# 使い方:
#   通常は run.cmd のメニュー4〜7（監督スクリプト経由）を使用する。
#   4-a  Word を閉じた状態で:
#          powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\04-word-layout-supervisor.ps1 -Runs 1
#          単発成功後に -Runs 10
#   4-b  Word で未保存の文書を開いた状態で（COM生成前に中止すること）:
#          powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\04-word-layout-supervisor.ps1 -Runs 1
#   4-f  キャンセル経路の確認:
#          powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\04-word-layout-supervisor.ps1 -Runs 1 -CancelAtStep 3
#   4-h  異常状態（既存インスタンス検出）の再現:
#          powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\04-word-layout-supervisor.ps1 -Runs 1 -SimulateExisting
# =====================================================================
[CmdletBinding()]
param(
    [int]$Repeat = 1,
    [int]$StepCount = 6,
    [int]$CancelAtStep = 0,        # 0 = キャンセルしない
    [switch]$SimulateExisting,     # 異常分岐を人為的に再現する
    [switch]$KeepVisible,
    [switch]$NoOpenPrompt,         # 監督スクリプトからの実行用
    [string]$StatusFile = ''       # 監督スクリプトへ結果コードを返す
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------
# ログ
# ---------------------------------------------------------------------
$log = New-Object System.Collections.ArrayList
function Add-Check {
    param([string]$Id, [string]$Name, [string]$Judge, [string]$Detail = '')
    [void]$log.Add([pscustomobject]@{ 項目 = $Id; 内容 = $Name; 判定 = $Judge; 詳細 = $Detail })
    $c = switch ($Judge) { 'OK' { 'Green' } 'NG' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    $d = if ($Detail) { " — $Detail" } else { '' }
    Write-Host ("  [{0,-4}] {1} {2}{3}" -f $Judge, $Id, $Name, $d) -ForegroundColor $c
}

# ---------------------------------------------------------------------
# Win32: ウィンドウハンドルからプロセスIDを得る
# ---------------------------------------------------------------------
$canPInvoke = $false
try {
    Add-Type -Namespace MBW -Name Api -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@ -ErrorAction Stop
    $canPInvoke = $true
} catch { }

function Get-ComProcessId {
    param($App)
    if (-not $canPInvoke) { return 0 }
    try {
        $hwnd = [IntPtr]([int64]$App.Hwnd)
        if ($hwnd -eq [IntPtr]::Zero) { return 0 }
        $procId = [uint32]0
        [void][MBW.Api]::GetWindowThreadProcessId($hwnd, [ref]$procId)
        return [int]$procId
    } catch { return 0 }
}

function Resolve-OwnedWordProcessId {
    param($App, [int[]]$PidsBefore, [int]$TimeoutMilliseconds = 4000)

    $deadline = (Get-Date).AddMilliseconds($TimeoutMilliseconds)
    do {
        # 第1候補: Application.Hwnd。文書未作成・非表示のWordでは0になる環境がある。
        $idFromHwnd = Get-ComProcessId $App
        if ($idFromHwnd -gt 0 -and $PidsBefore -notcontains $idFromHwnd) {
            return [pscustomobject]@{ Pid = $idFromHwnd; Mode = 'Hwnd' }
        }

        # 第2候補: 起動前にWordが0件であることを前提に、新しく現れた唯一のPIDを採用。
        $pidsNow = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                     Select-Object -ExpandProperty Id)
        $newPids = @($pidsNow | Where-Object { $PidsBefore -notcontains $_ })
        if ($newPids.Count -eq 1) {
            return [pscustomobject]@{ Pid = [int]$newPids[0]; Mode = '空ベースライン＋PID差分' }
        }
        if ($newPids.Count -gt 1) {
            return [pscustomobject]@{ Pid = 0; Mode = "新規PIDが複数: $($newPids -join ',')" }
        }
        Start-Sleep -Milliseconds 200
    } while ((Get-Date) -lt $deadline)

    return [pscustomobject]@{ Pid = 0; Mode = 'PIDを特定できず' }
}

# ---------------------------------------------------------------------
# Word 定数
# ---------------------------------------------------------------------
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleTitle = -63
$wdStyleSubtitle = -75
$wdAlignLeft = 0
$wdAlignCenter = 1
$wdPageBreak = 7
$wdFieldPage = 33
$wdHeaderFooterPrimary = 1
$wdLineStyleSingle = 1
$wdLineStyleNone = 0
$wdLineWidth075pt = 6
$wdFormatDocumentDefault = 16
$wdDoNotSaveChanges = 0
$wdStory = 6
$wdCollapseEnd = 0
$wdColorAutomatic = -16777216
$wdStatisticPages = 2
$msoTrue = -1

function ToBgr { param([int]$R, [int]$G, [int]$B) return ($B * 65536) + ($G * 256) + $R }
$colorNoteBg = ToBgr 255 249 219
$colorNoteLn = ToBgr 214 170 0
$colorImgLine = ToBgr 170 170 170

# ページ本文領域の概算高さ（A4縦・上下余白20mm）
$bodyAreaHeightPt = 700.0

# ---------------------------------------------------------------------
# スタイル適用（数値 → 失敗したら Styles.Item()）
# ---------------------------------------------------------------------
$script:StyleMode = $null
function Set-Style {
    param($Selection, $Doc, [int]$StyleId)
    if ($script:StyleMode -eq 'item') { $Selection.Style = $Doc.Styles.Item($StyleId); return }
    try {
        $Selection.Style = $StyleId
        if (-not $script:StyleMode) { $script:StyleMode = 'numeric' }
    } catch {
        $script:StyleMode = 'item'
        $Selection.Style = $Doc.Styles.Item($StyleId)
    }
}

function Move-SelectionToDocumentEnd {
    param($Selection, $Document)
    # Content.End-1 は文書末尾の段落記号の直前。表セル内のSelectionや
    # 現在のstoryに依存せず、必ず本文ストーリーの末尾へ戻す。
    $endPosition = [Math]::Max(0, ([int]$Document.Content.End - 1))
    $Selection.SetRange($endPosition, $endPosition)
}

function Get-FirstNonEmptyParagraphText {
    param($Document)
    for ($paragraphIndex = 1; $paragraphIndex -le $Document.Paragraphs.Count; $paragraphIndex++) {
        $text = [string]$Document.Paragraphs.Item($paragraphIndex).Range.Text
        $text = $text.Trim().Trim([char[]]@([char]7)).Trim()
        if ($text) { return $text }
    }
    return ''
}

# ---------------------------------------------------------------------
# フォント解決とスタイルへの適用（V-13 の一部・P1）
# ---------------------------------------------------------------------
function Resolve-BodyFont {
    param([string]$Preferred = 'BIZ UDPゴシック')
    Add-Type -AssemblyName System.Drawing
    $installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families |
                 ForEach-Object { $_.Name }
    foreach ($f in @($Preferred, 'BIZ UDPゴシック', 'BIZ UDPGothic', 'BIZ UDゴシック',
                     'Meiryo', 'Yu Gothic UI', 'MS Pゴシック')) {
        if ($f -and ($installed -contains $f)) { return $f }
    }
    return 'MS Pゴシック'
}

function Set-DocFonts {
    param($Doc, [string]$FontName)
    # 本文（Name だけでは日本語に効かないので NameFarEast も設定する）
    $f = $Doc.Content.Font
    $f.Name = $FontName; $f.NameFarEast = $FontName; $f.NameAscii = $FontName
    # スタイル定義も変える（これをしないと見出しが元のフォントのまま）
    $applied = @()
    foreach ($id in @($wdStyleNormal, $wdStyleHeading1, $wdStyleHeading2, -4, $wdStyleTitle, $wdStyleSubtitle)) {
        try {
            $sf = $Doc.Styles.Item($id).Font
            $sf.Name = $FontName; $sf.NameFarEast = $FontName; $sf.NameAscii = $FontName
            $applied += $id
        } catch { }
    }
    return $applied
}

# ---------------------------------------------------------------------
# ファイル名の安全化（P0-10）
# ---------------------------------------------------------------------
function Get-SafeFileName {
    param([string]$Name, [string]$Extension = '.docx', [string]$Dir)
    $s = $Name -replace '[\\/:*?"<>|]', '_'
    $s = ($s.ToCharArray() | ForEach-Object { if ([int]$_ -lt 32) { '_' } else { $_ } }) -join ''
    $s = $s.TrimEnd(' ', '.')
    $reserved = @('CON', 'PRN', 'AUX', 'NUL') +
                (1..9 | ForEach-Object { "COM$_" }) + (1..9 | ForEach-Object { "LPT$_" })
    if ($reserved -contains $s.ToUpperInvariant()) { $s = "_$s" }
    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'manual' }
    if ($s.Length -gt 100) { $s = $s.Substring(0, 100) }
    while ((Join-Path $Dir ($s + $Extension)).Length -gt 240 -and $s.Length -gt 8) {
        $s = $s.Substring(0, $s.Length - 8)
    }
    $c = $s; $i = 2
    while (Test-Path -LiteralPath (Join-Path $Dir ($c + $Extension))) { $c = "${s}_$i"; $i++ }
    return $c + $Extension
}

function Save-WordDocument {
    param($Document, [string]$Path, [int]$Format = 16)

    # 実機で SaveAs2(直接値) が戻らなくなったため、Windows PowerShell 5.1 で
    # 長く使われている SaveAs＋[ref] の2引数形式だけを使用する。
    $savePathRef = $Path
    $formatRef = $Format
    $Document.SaveAs([ref]$savePathRef, [ref]$formatRef)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Wordは保存完了を返しましたが、ファイルが見つかりません: $Path"
    }
}

# ---------------------------------------------------------------------
# テスト画像
# ---------------------------------------------------------------------
function New-TestImage {
    param([string]$Path, [int]$W, [int]$H, [string]$Caption)
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap $W, $H
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $penB = $null; $penF = $null; $brT = $null; $brk = $null; $fnt = $null; $fnt2 = $null
    try {
        $g.Clear([System.Drawing.Color]::FromArgb(245, 246, 248))
        $g.FillRectangle([System.Drawing.Brushes]::White, 20, 20, $W - 40, $H - 40)
        $penB = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(190, 195, 200)), 2
        $g.DrawRectangle($penB, 20, 20, $W - 40, $H - 40)
        $brT = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(60, 90, 150))
        $g.FillRectangle($brT, 20, 20, $W - 40, 48)
        $penF = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(150, 155, 160)), 2
        for ($i = 0; $i -lt 3; $i++) {
            $y = 120 + ($i * 70)
            if ($y -lt ($H - 90)) { $g.DrawRectangle($penF, 80, $y, [int](($W - 200) * 0.6), 40) }
        }
        $fnt = New-Object System.Drawing.Font 'Meiryo', 20, ([System.Drawing.FontStyle]::Bold)
        $brk = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 30, 30))
        $g.DrawString($Caption, $fnt, $brk, 40, [float]($H - 70))
        $fnt2 = New-Object System.Drawing.Font 'Meiryo', 14
        $g.DrawString("$W x $H px", $fnt2, $brk, 40, 80)
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        foreach ($d in @($penB, $penF, $brT, $brk, $fnt, $fnt2)) { if ($d) { $d.Dispose() } }
        $g.Dispose(); $bmp.Dispose()
    }
}

# =====================================================================
# 準備
# =====================================================================
$outDir = Join-Path $PSScriptRoot 'out'
$imgDir = Join-Path $outDir 'images'
$tmpDir = Join-Path $outDir '.tmp'
New-Item -ItemType Directory -Force -Path $imgDir | Out-Null
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  V-4 / V-3 / V-5 / V-11: Word COM 検証' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''

Add-Check 'V-4' 'Win32 API（GetWindowThreadProcessId）の利用' `
    $(if ($canPInvoke) { 'OK' } else { 'NG' }) `
    $(if (-not $canPInvoke) { 'PIDを特定できないため、Stop-Process は一切行いません' } else { '' })

$bodyFont = Resolve-BodyFont
Add-Check 'V-13' '本文フォントの解決' 'INFO' $bodyFont

Write-Host ''
Write-Host '[1] テスト画像を生成' -ForegroundColor Cyan
$specs = @(
    @{ W = 1600; H = 900;  C = 'ワイド画面 16:9'; Lines = 2 },
    @{ W = 800;  H = 1400; C = '縦長画面（高さ制限のテスト）'; Lines = 1 },
    @{ W = 1200; H = 260;  C = '横長・低い画面'; Lines = 8 },
    @{ W = 2560; H = 1440; C = '高解像度 2560x1440'; Lines = 1 },
    @{ W = 1400; H = 780;  C = 'ページ境界テスト用A'; Lines = 12 },
    @{ W = 900;  H = 1900; C = '極端な縦長（1ページ超・分断許容の確認）'; Lines = 1 }
)
$images = @()
for ($i = 0; $i -lt $StepCount; $i++) {
    $s = $specs[$i % $specs.Count]
    $p = Join-Path $imgDir ("test-{0:d3}.png" -f ($i + 1))
    New-TestImage -Path $p -W $s.W -H $s.H -Caption $s.C
    $images += [pscustomobject]@{ Path = $p; W = $s.W; H = $s.H; Caption = $s.C; Lines = $s.Lines }
}
Add-Check '-' 'テスト画像の生成' 'OK' "$($images.Count) 枚"

# =====================================================================
# 本体
# =====================================================================
$savedFiles = @()
$summaries = @()

for ($round = 1; $round -le $Repeat; $round++) {

    Write-Host ''
    Write-Host "[2] ラウンド $round / $Repeat" -ForegroundColor Cyan

    # ---- V-4: 起動前の WINWORD PID を記録 ----
    $pidsBefore = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                    Select-Object -ExpandProperty Id)
    Write-Host ("  起動前の WINWORD PID : {0}" -f $(if ($pidsBefore.Count) { $pidsBefore -join ', ' } else { 'なし' })) -ForegroundColor Gray

    $word = $null; $doc = $null; $sel = $null
    $ownPid = 0
    $ownershipMode = 'なし'
    $baselineClear = ($pidsBefore.Count -eq 0)
    $abnormal = $false
    $settingsApplied = $false
    $quitCalled = $false
    $killed = $false
    $savePath = ''
    $partialPath = ''
    $roundError = ''
    $stopRemainingRounds = $false

    try {
        # ---- 4-h: 異常分岐だけを安全に再現する ----
        # 旧実装は検証用に起動した Word を「既存扱い」にして Quit() を避けていたため、
        # 空の WINWORD が残る可能性があった。シミュレーションでは COM を起動しない。
        if ($SimulateExisting) {
            $abnormal = $true
            $roundError = '既存の Word インスタンスを検出した想定で出力を中止しました'
            Add-Check '4-e' '自インスタンスの PID 特定' 'INFO' 'シミュレーションのため COM は起動しません'
            Add-Check '4-b' '既存インスタンス検出（シミュレーション）' 'OK' '異常状態と判定 → 出力を中止します'
            Add-Check '4-c' 'DisplayAlerts / ScreenUpdating の変更' 'OK' '変更していません'
            Add-Check '4-h' '異常時に Quit() を呼ばないこと' 'OK' 'Word 自体を起動していません'
            throw 'ABORT_EXISTING_WORD'
        }

        # ---- 実機結果に基づく安全フォールバック ----
        # 空の非表示Wordでは Application.Hwnd が0になる環境がある。
        # 既存Wordがある状態では所有権を安全に証明できないため、COM生成前に中止する。
        if (-not $baselineClear) {
            $abnormal = $true
            $roundError = "Word が起動中です（PID: $($pidsBefore -join ',')）"
            Add-Check '4-b' '起動前の Word が0件' 'OK' '起動中を検出し、COMを生成する前に中止しました'
            Add-Check '4-c' '既存 Word への変更' 'OK' 'COMを生成していないため変更していません'
            Add-Check '4-h' '既存 Word に Quit() を呼ばないこと' 'OK' 'Wordを閉じてから再実行してください'
            throw 'ABORT_WORD_RUNNING'
        }
        Add-Check '4-b' '起動前の Word が0件' 'OK' '安全な空ベースラインを確認'

        # ---- COM 生成。ここでは何も設定しない（P0-2） ----
        $word = New-Object -ComObject Word.Application
        $resolved = Resolve-OwnedWordProcessId -App $word -PidsBefore $pidsBefore
        $ownPid = [int]$resolved.Pid
        $ownershipMode = [string]$resolved.Mode

        Write-Host ("  自インスタンスの PID : {0}" -f $(if ($ownPid) { $ownPid } else { '特定できず' })) -ForegroundColor Gray
        Add-Check '4-e' '自インスタンスの PID 特定' `
            $(if ($ownPid -gt 0) { 'OK' } else { 'WARN' }) `
            $(if ($ownPid -gt 0) { "PID $ownPid / $ownershipMode" } else { "$ownershipMode / Quit()はCOM参照にだけ実行し、Stop-Processは行いません" })

        # ---- ここから初めて設定してよい ----
        $word.Visible = [bool]$KeepVisible
        $word.DisplayAlerts = 0
        try { $word.ScreenUpdating = $false } catch { }
        $settingsApplied = $true
        Add-Check '4-c' 'インスタンス設定の適用' 'OK' 'PID判定の後に適用しました'
        Add-Check '-' 'Word 起動' 'OK' `
            "Ver $($word.Version) / PID $(if ($ownPid) { $ownPid } else { '不明' }) / $ownershipMode"

        $doc = $word.Documents.Add()
        $sel = $word.Selection

        # ---- フォント（P1） ----
        try {
            $applied = Set-DocFonts $doc $bodyFont
            Add-Check 'B-9' 'フォント（NameFarEast＋スタイル定義）' 'OK' `
                "$bodyFont / スタイル $($applied.Count) 個に適用"
        } catch {
            Add-Check 'B-9' 'フォント設定' 'NG' $_.Exception.Message
        }

        # ---- 余白（テンプレート差の吸収） ----
        try {
            $doc.PageSetup.TopMargin = 56.7
            $doc.PageSetup.BottomMargin = 56.7
            $doc.PageSetup.LeftMargin = 56.7
            $doc.PageSetup.RightMargin = 56.7
            $doc.Content.Font.Size = 10.5
            Add-Check 'B-10' '余白・本文サイズの明示設定' 'OK' '20mm / 10.5pt'
        } catch {
            Add-Check 'B-10' '余白の明示設定' 'WARN' $_.Exception.Message
        }

        # ---- 表紙 ----
        try {
            Set-Style $sel $doc $wdStyleTitle
            $sel.ParagraphFormat.Alignment = $wdAlignCenter
            $sel.TypeText([string]'経費精算システム 操作マニュアル')
            $sel.TypeParagraph()
            Set-Style $sel $doc $wdStyleSubtitle
            $sel.TypeText([string]'Ver. 1.0')
            $sel.TypeParagraph()
            Set-Style $sel $doc $wdStyleNormal
            $sel.ParagraphFormat.Alignment = $wdAlignCenter
            $sel.TypeText([string](Get-Date -Format 'yyyy年M月d日'))
            $sel.TypeParagraph()
            $sel.TypeText([string]'作成者: ManualBuilder')
            $sel.TypeParagraph()
            $sel.InsertBreak($wdPageBreak)
            Add-Check 'B-1' '表紙（タイトル/副題/中央揃え/改ページ）' 'OK' "スタイル指定方式: $script:StyleMode"
        } catch {
            Add-Check 'B-1' '表紙' 'NG' $_.Exception.Message
        }

        # ---- 目次 ----
        $tocOk = $false
        try {
            # 「目次」自体を目次項目へ含めないため、Heading 1にはしない。
            Set-Style $sel $doc $wdStyleNormal
            $sel.ParagraphFormat.Alignment = $wdAlignLeft
            $sel.Font.Bold = $true
            $sel.Font.Size = 18
            $sel.TypeText([string]'目次')
            $sel.TypeParagraph()
            Set-Style $sel $doc $wdStyleNormal
            $sel.Font.Bold = $false
            $sel.Font.Size = 10.5
            $null = $doc.TablesOfContents.Add($sel.Range, $true, 1, 3)
            Move-SelectionToDocumentEnd -Selection $sel -Document $doc
            $sel.TypeParagraph()
            $sel.InsertBreak($wdPageBreak)
            $tocOk = $true
            Add-Check 'B-2' '目次の挿入' 'OK' '見出し1〜3を対象'
        } catch {
            Add-Check 'B-2' '目次の挿入' 'NG' $_.Exception.Message
        }

        # ---- 概要 ----
        try {
            Set-Style $sel $doc $wdStyleHeading1
            $sel.TypeText([string]'概要')
            $sel.TypeParagraph()
            Set-Style $sel $doc $wdStyleNormal
            $sel.TypeText([string]("経費精算の申請から承認までの手順です。" + [char]11 + "対象: 経理部・各部門の申請者"))
            $sel.TypeParagraph()
        } catch { }

        # ---- 手順 ----
        $maxWpt = 400.0
        $maxHpt = 480.0
        $noteMethods = @()
        $keepTogetherOff = @()
        $explicitPageBreaks = @()
        $pageBreakWarnings = @()
        $altTextOk = 0
        # 概要の見出し＋本文が使う高さ。非表示Wordでは現在Y座標を取得できないため、
        # ページ位置をCOMから読まず、挿入した内容の概算高さを累積して判断する。
        $estimatedPageUsedPt = 75.0

        for ($i = 0; $i -lt $images.Count; $i++) {
            $n = $i + 1
            $img = $images[$i]
            Move-SelectionToDocumentEnd -Selection $sel -Document $doc

            # ---- 4-f: キャンセル経路の確認 ----
            if ($CancelAtStep -gt 0 -and $n -eq $CancelAtStep) {
                $roundError = "キャンセル要求（手順 $n の直前）"
                Add-Check '4-f' 'キャンセル要求で finally に入るか' 'OK' "手順 $n で中断します"
                throw 'CANCELLED'
            }

            # ---- 画像の想定高さから KeepTogether の可否を決める（V-11） ----
            # AddPicture 前に概算する（縦横比とmaxHptから逆算）
            $estImgH = [Math]::Min($maxHpt, $maxWpt * $img.H / $img.W)
            # 実機文書では長い手順を前ページ末尾へ配置し始め、先頭が印刷領域外へ
            # 欠けるケースがあった。折り返しと段落間隔を含め、やや保守的に見積もる。
            $estTotal = 32 + ($img.Lines * 16) + 52 + $estImgH      # 見出し+本文+補足+画像
            $keepTogether = ($estTotal -lt ($bodyAreaHeightPt * 0.85))
            if (-not $keepTogether) { $keepTogetherOff += $n }

            # ---- 手順全体が残り領域へ入らない場合は、見出し前で明示改ページ ----
            # v0.4.7はSelection.Informationで現在Y座標を読んだが、非表示Wordでは
            # 0または取得不能となり、改ページが一度も発動しなかった。ここでは
            # Wordの表示状態に依存しない累積高さで、同じ入力なら必ず同じ位置で改ページする。
            $pageBreakBeforeHeading = $false
            if ($estimatedPageUsedPt -gt 0 -and
                ($estimatedPageUsedPt + $estTotal + 18.0) -gt $bodyAreaHeightPt) {
                try {
                    # 空の改ページ段落を挿入すると、長いKeepWithNext連鎖をWordが
                    # ページ上端より上へ押し出す環境があった。見出し自体へ
                    # PageBreakBeforeを設定し、改ページ位置を段落境界へ固定する。
                    $pageBreakBeforeHeading = $true
                    $explicitPageBreaks += $n
                    $estimatedPageUsedPt = 0.0
                } catch {
                    $pageBreakWarnings += "手順${n}: $($_.Exception.Message)"
                }
            }
            $estimatedPageUsedPt += ($estTotal + 18.0)

            # ---- 見出し ----
            Set-Style $sel $doc $wdStyleHeading2
            $sel.ParagraphFormat.Alignment = $wdAlignLeft
            $sel.ParagraphFormat.LeftIndent = 0
            $sel.ParagraphFormat.PageBreakBefore = $pageBreakBeforeHeading
            $sel.TypeText([string]("手順 $n  $($img.Caption)"))
            $hp = $sel.Paragraphs.Item(1)
            $hp.KeepWithNext = $true                     # V-11
            $sel.TypeParagraph()

            # ---- 本文（行数を変えてページ境界のケースを作る） ----
            Set-Style $sel $doc $wdStyleNormal
            $sel.ParagraphFormat.PageBreakBefore = $false
            $sel.ParagraphFormat.LeftIndent = 0
            $sel.ParagraphFormat.RightIndent = 0
            $bodyLines = @()
            for ($k = 1; $k -le $img.Lines; $k++) {
                $bodyLines += "説明文の $k 行目です。ブラウザで https://example.co.jp を開いて操作します。"
            }
            $sel.TypeText([string]($bodyLines -join [char]11))
            $bp = $sel.Paragraphs.Item(1)
            # 見出し→本文は見出し側のKeepWithNextで維持する。本文まで補足・画像へ
            # 連鎖させると、大きな手順全体が1ブロック化しページ上端から欠ける
            # 環境があるため、ここで連鎖を切る。補足→画像は別ブロックで維持する。
            $bp.KeepWithNext = $false
            $bp.KeepTogether = $keepTogether
            $sel.TypeParagraph()

            # ---- 補足の囲み枠（奇数=段落罫線 / 偶数=1x1の表） ----
            $noteText = "補足: この文が囲み枠になっているか、書式が次の段落に漏れていないかを確認してください（手順 $n）"
            if ($n % 2 -eq 1) {
                try {
                    Set-Style $sel $doc $wdStyleNormal
                    $sel.TypeText([string]$noteText)
                    $par = $sel.Paragraphs.Item(1)
                    $par.LeftIndent = 14; $par.RightIndent = 14
                    $par.SpaceBefore = 6; $par.SpaceAfter = 6
                    $par.KeepWithNext = $true
                    foreach ($bi in @(-1, -2, -3, -4)) {
                        $bd = $par.Borders.Item($bi)
                        $bd.LineStyle = $wdLineStyleSingle
                        $bd.LineWidth = $wdLineWidth075pt
                        $bd.Color = $colorNoteLn
                    }
                    $par.Range.Shading.BackgroundPatternColor = $colorNoteBg
                    $sel.TypeParagraph()
                    # 次の段落へ書式を引き継がせない
                    $sel.ParagraphFormat.Reset()
                    Set-Style $sel $doc $wdStyleNormal
                    $sel.ParagraphFormat.LeftIndent = 0
                    $sel.ParagraphFormat.RightIndent = 0
                    foreach ($bi in @(-1, -2, -3, -4)) {
                        $sel.ParagraphFormat.Borders.Item($bi).LineStyle = $wdLineStyleNone
                    }
                    $sel.Range.Shading.BackgroundPatternColor = $wdColorAutomatic
                    $noteMethods += "手順${n}=段落罫線"
                } catch {
                    Add-Check 'B-4' "囲み枠(段落罫線) 手順$n" 'NG' $_.Exception.Message
                }
            } else {
                try {
                    # Selection.Rangeが表セルstoryへ残ることがあるため、文書末尾の
                    # 明示Rangeへ表を追加する。以降もEndKey(wdStory)は使わない。
                    $tableEnd = [Math]::Max(0, ([int]$doc.Content.End - 1))
                    $tableRange = $doc.Range($tableEnd, $tableEnd)
                    $tbl = $doc.Tables.Add($tableRange, 1, 1)
                    $tbl.Range.Shading.BackgroundPatternColor = $colorNoteBg
                    $tbl.Borders.InsideLineStyle = $wdLineStyleSingle
                    $tbl.Borders.OutsideLineStyle = $wdLineStyleSingle
                    $tbl.Borders.OutsideColor = $colorNoteLn
                    $tbl.Borders.InsideColor = $colorNoteLn
                    $tbl.Rows.AllowBreakAcrossPages = $false
                    $tbl.Cell(1, 1).Range.Text = [string]$noteText
                    $tbl.Cell(1, 1).TopPadding = 6
                    $tbl.Cell(1, 1).BottomPadding = 6
                    # 表の最後の段落と直後の画像を同じページへ送れるようにする。
                    $tbl.Cell(1, 1).Range.Paragraphs.Item(1).KeepWithNext = $true
                    $tbl.Range.InsertParagraphAfter()
                    Move-SelectionToDocumentEnd -Selection $sel -Document $doc
                    Set-Style $sel $doc $wdStyleNormal
                    $noteMethods += "手順${n}=1x1表"
                } catch {
                    Add-Check 'B-4' "囲み枠(1x1表) 手順$n" 'NG' $_.Exception.Message
                }
            }

            # ---- 画像 ----
            try {
                $shape = $sel.InlineShapes.AddPicture($img.Path, $false, $true)
                $natW = [double]$shape.Width
                $natH = [double]$shape.Height
                $shape.LockAspectRatio = $msoTrue
                $scale = [Math]::Min([Math]::Min($maxWpt / $natW, $maxHpt / $natH), 1.0)
                $shape.Width = [float]($natW * $scale)
                # 代替テキスト（B-8）
                try {
                    $shape.AlternativeText = [string]("手順 $n の画面: $($img.Caption)")
                    $altTextOk++
                } catch { }
                try {
                    $shape.Line.Visible = $true
                    $shape.Line.ForeColor.RGB = $colorImgLine
                    $shape.Line.Weight = 0.75
                } catch { }
                $ip = $sel.Paragraphs.Item(1)
                $ip.KeepTogether = $keepTogether          # V-11
                $ip.KeepWithNext = $false                 # 次の手順まで連鎖させない
                $sel.TypeParagraph()
                if ($n -eq 1) {
                    Add-Check 'B-6' '画像の挿入とサイズ調整' 'OK' `
                        ("元 {0:N0}x{1:N0}pt → 倍率 {2:N2} → 幅 {3:N0}pt" -f $natW, $natH, $scale, ($natW * $scale))
                }
            } catch {
                Add-Check 'B-6' "画像の挿入 手順$n" 'NG' $_.Exception.Message
            }
        }

        # ---- 文書順序の自己検査 ----
        # 実機で表セルstoryにSelectionが残り、手順6→2が表紙より前へ逆順挿入された。
        # 先頭の非空段落と、本文にある各手順見出しの最終出現位置を検査する。
        $firstParagraphText = Get-FirstNonEmptyParagraphText -Document $doc
        $documentText = [string]$doc.Content.Text
        $previousStepPosition = -1
        $stepOrderOk = $true
        for ($orderIndex = 0; $orderIndex -lt $images.Count; $orderIndex++) {
            $orderNumber = $orderIndex + 1
            $stepToken = "手順 $orderNumber  $($images[$orderIndex].Caption)"
            # 目次にも同じ文字列があるため、LastIndexOfで本文側の見出しを採る。
            $stepPosition = $documentText.LastIndexOf($stepToken)
            if ($stepPosition -le $previousStepPosition) {
                $stepOrderOk = $false
                break
            }
            $previousStepPosition = $stepPosition
        }
        if ($firstParagraphText -ne '経費精算システム 操作マニュアル' -or -not $stepOrderOk) {
            Add-Check 'B-11' '文書順序（表紙と手順1→6）' 'NG' `
                "先頭: $firstParagraphText / 手順順序: $(if ($stepOrderOk) { '正常' } else { '異常' })"
            throw 'DOC_ORDER_INVALID'
        }
        Add-Check 'B-11' '文書順序（表紙と手順1→6）' 'OK' '表紙→目次→概要→手順1→6'
        if ($images.Count -eq 6 -and (($explicitPageBreaks -join ',') -ne '2,3,5,6')) {
            $pageBreakWarnings += "既定6手順の改ページ位置が不正: $($explicitPageBreaks -join ',')（期待: 2,3,5,6）"
        }
        Add-Check 'B-12' '手順単位の事前改ページ' `
            $(if ($pageBreakWarnings.Count) { 'NG' } else { 'OK' }) `
            $(if ($pageBreakWarnings.Count) {
                $pageBreakWarnings -join ' / '
            } elseif ($explicitPageBreaks.Count) {
                "残り領域へ収まらない手順の見出しへPageBreakBeforeを設定: $($explicitPageBreaks -join ',')"
            } else {
                '追加改ページなし'
            })

        Add-Check 'B-4' '囲み枠（2方式を交互に出力）' 'OK' ($noteMethods -join ', ')
        Add-Check 'B-8' '画像の代替テキスト' `
            $(if ($altTextOk -eq $images.Count) { 'OK' } else { 'WARN' }) `
            "$altTextOk / $($images.Count) 件に設定"
        Add-Check 'V-11' 'KeepWithNext / KeepTogether' 'OK' `
            $(if ($keepTogetherOff.Count) { "手順 $($keepTogetherOff -join ',') は1ページを超えるため KeepTogether を外しました" } else { '全手順に適用' })

        # ---- ページ番号 ----
        try {
            $footer = $doc.Sections.Item(1).Footers.Item($wdHeaderFooterPrimary)
            $fr = $footer.Range
            $fr.ParagraphFormat.Alignment = $wdAlignCenter
            $fr.Text = '- '
            $fr.Collapse($wdCollapseEnd) | Out-Null
            $null = $footer.Range.Fields.Add($fr, $wdFieldPage)
            Add-Check 'B-7' 'ページ番号（フッター中央）' 'OK' ''
        } catch {
            Add-Check 'B-7' 'ページ番号' 'NG' $_.Exception.Message
        }

        # ---- 目次更新（P1: Repaginate を先に） ----
        try {
            $doc.Repaginate()
            if ($tocOk -and $doc.TablesOfContents.Count -gt 0) {
                $doc.TablesOfContents.Item(1).Update()
            }
            $doc.Fields.Update() | Out-Null
            # 目次とフィールド更新で行数が変わるため、更新後にも再改ページする。
            $doc.Repaginate()
            Add-Check 'B-2' '改ページ→目次／フィールド更新→再改ページ' 'OK' ''
        } catch {
            Add-Check 'B-2' '目次の更新' 'WARN' $_.Exception.Message
        }

        # ---- ページ数（V-11d の確認材料） ----
        try {
            $pages = $doc.ComputeStatistics($wdStatisticPages)
            Add-Check 'V-11' 'ページ数' 'INFO' "$pages ページ（$($images.Count) 手順）"
        } catch { }

        # ---- 保存（原子的: .tmp → リネーム） ----
        $tmpPath = Join-Path $tmpDir ("build-{0}-{1}.docx" -f $PID, $round)
        if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force }
        Write-Host "  Word保存開始: $tmpPath" -ForegroundColor Cyan
        $saveWatch = [Diagnostics.Stopwatch]::StartNew()
        Save-WordDocument -Document $doc -Path $tmpPath -Format $wdFormatDocumentDefault
        $saveWatch.Stop()
        $savedKb = (Get-Item -LiteralPath $tmpPath).Length / 1KB
        Write-Host ("  Word保存完了: {0:N1}秒 / {1:N0}KB" -f $saveWatch.Elapsed.TotalSeconds, $savedKb) -ForegroundColor Green
        Write-Host '  Word文書を閉じています...' -ForegroundColor Gray
        $doc.Close($wdDoNotSaveChanges)
        $doc = $null
        Write-Host '  Word文書を閉じました' -ForegroundColor Gray

        $safeName = Get-SafeFileName -Name ("手順書サンプル_r{0}_{1}" -f $round, (Get-Date -Format 'yyyyMMdd_HHmmss')) -Dir $outDir
        $savePath = Join-Path $outDir $safeName
        Move-Item -LiteralPath $tmpPath -Destination $savePath
        $savedFiles += $savePath
        Add-Check 'R-54' '原子的な保存（.tmp → リネーム）' 'OK' $safeName

    } catch {
        $msg = $_.Exception.Message
        if ($msg -eq 'ABORT_EXISTING_WORD') {
            Write-Host '  → 既存インスタンスを検出した想定で中止しました（シミュレーション）' -ForegroundColor Yellow
        } elseif ($msg -eq 'ABORT_WORD_RUNNING') {
            Write-Host '  → Word が起動中のため、COMを生成せず安全に中止しました' -ForegroundColor Yellow
            $stopRemainingRounds = $true
        } elseif ($msg -eq 'CANCELLED') {
            Write-Host '  → キャンセル要求により中断しました' -ForegroundColor Yellow
        } else {
            $roundError = $msg
            Add-Check '-' 'docx 生成' 'NG' $msg
        }
        # 途中まで作れていたら .partial.docx として残す（P0-8）
        if ($doc -and $settingsApplied) {
            try {
                $pName = Get-SafeFileName -Name ("手順書サンプル_r{0}_失敗" -f $round) -Extension '.partial.docx' -Dir $outDir
                $partialPath = Join-Path $outDir $pName
                Write-Host "  partial保存開始: $partialPath" -ForegroundColor Cyan
                Save-WordDocument -Document $doc -Path $partialPath -Format $wdFormatDocumentDefault
                Add-Check 'R-54' '失敗時は .partial.docx として残す' 'OK' $pName
            } catch { }
        }
    } finally {
        # ---- 後始末 ----
        try { if ($doc) { $doc.Close($wdDoNotSaveChanges) } } catch { }

        # 起動前にWordが0件だった場合だけ、今回得たCOM参照へQuit()を呼ぶ。
        # PIDが取れなくても、既存Wordへ接続した可能性は空ベースラインで排除できる。
        if ($word -and $baselineClear -and -not $SimulateExisting) {
            if ($settingsApplied) { try { $word.ScreenUpdating = $true } catch { } }
            try { $word.Quit(); $quitCalled = $true } catch { }
        }

        foreach ($o in @($sel, $doc, $word)) {
            if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        Remove-Variable sel, doc, word -ErrorAction SilentlyContinue
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()

        # ---- 残留の掃除: Hwndで直接特定できたPIDだけ（最大10秒待つ） ----
        # 空ベースライン＋PID差分は競合起動との理論上の取り違えが残るため、強制終了には使わない。
        if (-not $abnormal -and $ownPid -gt 0 -and $ownershipMode -eq 'Hwnd' -and `
            ($pidsBefore -notcontains $ownPid)) {
            $deadline = (Get-Date).AddSeconds(10)
            while ((Get-Date) -lt $deadline) {
                if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { break }
                Start-Sleep -Milliseconds 500
            }
            if (Get-Process -Id $ownPid -ErrorAction SilentlyContinue) {
                Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue
                $killed = $true
            }
        }

        $summaries += [pscustomobject]@{
            ラウンド = $round
            自PID = $(if ($ownPid) { $ownPid } else { '不明' })
            PID特定方式 = $ownershipMode
            空ベースライン = $baselineClear
            異常検出 = $abnormal
            設定適用 = $settingsApplied
            Quit呼出 = $quitCalled
            強制終了 = $killed
            結果 = $(if ($savePath) { 'ok' } elseif ($partialPath) { 'partial' } else { 'abort' })
            メモ = $roundError
        }
    }
    if ($stopRemainingRounds) {
        Write-Host '  → 残りの繰り返しは実行しません。Wordを閉じてから再実行してください。' -ForegroundColor Yellow
        break
    }
}

# =====================================================================
# V-4: ユーザーの Word が生きているか
# =====================================================================
Write-Host ''
Write-Host '[3] V-4: 残留と、ユーザーの Word の生存確認' -ForegroundColor Cyan

$ownPids = @($summaries | Where-Object { $_.自PID -ne '不明' } | Select-Object -ExpandProperty 自PID)
$executedRounds = @($summaries | Where-Object { $_.設定適用 })
$unknownExecuted = @($executedRounds | Where-Object { $_.自PID -eq '不明' })
$blockedRounds = @($summaries | Where-Object { $_.メモ -like 'Word が起動中です*' })

# WordはQuit()後も数秒だけプロセスが残る環境がある。即NGにせず、
# 最大10秒は終了を観測する（PID差分で特定したプロセスを強制終了はしない）。
$shutdownWatch = [Diagnostics.Stopwatch]::StartNew()
$shutdownDeadline = (Get-Date).AddSeconds(10)
do {
    $pidsAfterAll = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                      Select-Object -ExpandProperty Id)
    $leaked = @($pidsAfterAll | Where-Object { $ownPids -contains $_ })
    $unknownResidual = ($unknownExecuted.Count -gt 0 -and $pidsAfterAll.Count -gt 0)
    if ($executedRounds.Count -eq 0 -or ($leaked.Count -eq 0 -and -not $unknownResidual)) {
        break
    }
    Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $shutdownDeadline)
$shutdownWatch.Stop()

if ($executedRounds.Count -eq 0 -and $blockedRounds.Count -gt 0) {
    Add-Check '4-a' '既存 Word への非干渉' 'OK' `
        'COMを生成せず中止。起動中のWordはそのまま残します'
} elseif ($executedRounds.Count -eq 0 -and $SimulateExisting) {
    Add-Check '4-a' 'シミュレーション時の残留' 'OK' 'Wordを起動していません'
} else {
    $hasLeak = ($leaked.Count -gt 0 -or $unknownResidual)
    $leakDetail = if ($leaked.Count -gt 0) {
        "所有PIDの残留: $($leaked -join ',')"
    } elseif ($unknownResidual) {
        "PID不明の実行後に WINWORD が残留: $($pidsAfterAll -join ',')"
    } else {
        "$($executedRounds.Count) 回出力して残留 0（Quit後 $([Math]::Round($shutdownWatch.Elapsed.TotalSeconds, 1)) 秒確認）"
    }
    Add-Check '4-a' '自インスタンスの残留' $(if ($hasLeak) { 'NG' } else { 'OK' }) $leakDetail
}

$killedRounds = @($summaries | Where-Object { $_.強制終了 })
if ($killedRounds.Count -gt 0) {
    Add-Check '4-a' '強制終了が必要だった回数' 'WARN' `
        "$($killedRounds.Count) 回（Quit() だけでは終わらない環境）"
}

Write-Host ''
Write-Host '  V-4b の確認:' -ForegroundColor Yellow
Write-Host '   ・Word で未保存の文書を開いた状態では、出力を開始せず中止すること'
Write-Host '   ・その Word のウィンドウと未保存内容が、そのまま残っていること'
Write-Host ''

# =====================================================================
# 結果
# =====================================================================
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  ラウンド別サマリ' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
$summaries | Format-Table -AutoSize -Wrap

Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  検証項目' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
$log | Format-Table -AutoSize -Wrap

$ngCount = @($log | Where-Object { $_.判定 -eq 'NG' }).Count
Write-Host ''
if ($ngCount -eq 0) {
    Write-Host '  NG はありません。' -ForegroundColor Green
} else {
    Write-Host "  NG が $ngCount 件あります。" -ForegroundColor Red
}

$reportPath = Join-Path $outDir 'result-04-word.txt'
(($summaries | Format-Table -AutoSize -Wrap | Out-String -Width 200) + "`r`n" +
 ($log | Format-Table -AutoSize -Wrap | Out-String -Width 200)) |
    Out-File -FilePath $reportPath -Encoding UTF8

Write-Host ''
Write-Host "  レポート: $reportPath" -ForegroundColor Cyan
foreach ($f in $savedFiles) { Write-Host "  生成物  : $f" -ForegroundColor Cyan }

if ($savedFiles.Count -gt 0) {
    Write-Host ''
    Write-Host '  生成された docx を開いて、目視で確認してください:' -ForegroundColor Yellow
    Write-Host '   B-1  表紙のタイトルが大きく中央に出ているか'
    Write-Host '   B-2  目次にページ番号付きで手順が並んでいるか'
    Write-Host '   B-3  「手順 N ...」が見出しスタイルか（表示→ナビゲーションウィンドウ）'
    Write-Host '   B-4  補足が囲み枠か。★奇数=段落罫線 / 偶数=1x1の表。どちらが好みか'
    Write-Host '   B-4  ★囲み枠の書式が次の段落に漏れていないか（段落罫線側の弱点）'
    Write-Host '   B-5  手順の見出し・説明・画像が別ページに分断されていないか'
    Write-Host '   B-6  縦長画像がページに収まっているか（手順2）'
    Write-Host '   B-7  フッターにページ番号が出ているか'
    Write-Host '   B-8  画像を右クリック→代替テキストの編集で文が入っているか'
    Write-Host '   B-9  見出しも本文と同じフォントになっているか'
    Write-Host '   B-11 表紙が1ページ目で、手順1→6が順番どおりか'
    Write-Host '   B-12 手順2・5・6の見出しや本文先頭が欠けず、画像と不自然に分断されていないか'
    Write-Host ''
    if (-not $NoOpenPrompt) {
        Write-Host '  最後の docx を開きますか？ [Y/n] ' -NoNewline -ForegroundColor Yellow
        if ((Read-Host) -ne 'n') { Start-Process $savedFiles[-1] }
    }
} else {
    Write-Host ''
    Write-Host '  docx は生成されていません。上の中止理由を確認してください。' -ForegroundColor Yellow
}

$workerExitCode = $(if ($ngCount -gt 0) { 1 } else { 0 })
if ($StatusFile) {
    try {
        $statusDir = Split-Path -Parent $StatusFile
        if ($statusDir -and -not (Test-Path -LiteralPath $statusDir)) {
            New-Item -ItemType Directory -Force -Path $statusDir | Out-Null
        }
        [IO.File]::WriteAllText($StatusFile, [string]$workerExitCode, [Text.Encoding]::ASCII)
    } catch {
        Write-Host "  監督結果ファイルを書き込めません: $($_.Exception.Message)" -ForegroundColor Red
        $workerExitCode = 1
    }
}
exit $workerExitCode
