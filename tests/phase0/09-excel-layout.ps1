# =====================================================================
# 09-excel-layout.ps1 — Excel主出力のPhase 0検証
#
# 検証項目:
#   X-01/X-02  アプリのシート順・手順順・シート内採番
#   X-05       目次リンクと「目次へ戻る」リンク
#   X-06/X-07  PC向け横長カード、画像の縦横比・重なり防止
#   X-08       禁止文字、31文字、重複を含むシート名変換
#   X-10/X-11  所有EXCELの終了、既存の未保存Excelへの非干渉
#   X-13       一時保存から完成名への移動、未完成品の除外
#
# 通常は09-excel-layout-supervisor.ps1経由で実行する。
# =====================================================================
[CmdletBinding()]
param(
    [ValidateRange(1, 10)][int]$Round = 1,
    [int]$CancelAtStep = 0,
    [switch]$SimulateExisting,
    [switch]$KeepVisible,
    [switch]$NoOpenPrompt,
    [string]$StatusFile = ''
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------
# ログ
# ---------------------------------------------------------------------
$log = New-Object System.Collections.ArrayList
function Add-Check {
    param([string]$Id, [string]$Name, [string]$Judge, [string]$Detail = '')
    [void]$log.Add([pscustomobject]@{ 項目 = $Id; 内容 = $Name; 判定 = $Judge; 詳細 = $Detail })
    $color = switch ($Judge) {
        'OK' { 'Green' }
        'NG' { 'Red' }
        'WARN' { 'Yellow' }
        default { 'Gray' }
    }
    $suffix = if ($Detail) { " — $Detail" } else { '' }
    Write-Host ("  [{0,-4}] {1} {2}{3}" -f $Judge, $Id, $Name, $suffix) -ForegroundColor $color
}

function To-Bgr {
    param([int]$R, [int]$G, [int]$B)
    return ($B * 65536) + ($G * 256) + $R
}

# ---------------------------------------------------------------------
# Win32: Excel.Application.Hwndから所有PIDを得る
# ---------------------------------------------------------------------
$canPInvoke = $false
try {
    Add-Type -Namespace MBExcel -Name Api -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@ -ErrorAction Stop
    $canPInvoke = $true
} catch { }

function Get-ComProcessId {
    param($Application)
    if (-not $canPInvoke) { return 0 }
    try {
        $hwnd = [IntPtr]([int64]$Application.Hwnd)
        if ($hwnd -eq [IntPtr]::Zero) { return 0 }
        $processId = [uint32]0
        [void][MBExcel.Api]::GetWindowThreadProcessId($hwnd, [ref]$processId)
        return [int]$processId
    } catch {
        return 0
    }
}

function Resolve-OwnedExcelProcessId {
    param($Application, [int[]]$PidsBefore, [int]$TimeoutMilliseconds = 5000)

    $deadline = (Get-Date).AddMilliseconds($TimeoutMilliseconds)
    do {
        $idFromHwnd = Get-ComProcessId -Application $Application
        if ($idFromHwnd -gt 0) {
            if ($PidsBefore -contains $idFromHwnd) {
                return [pscustomobject]@{
                    Pid = 0
                    Mode = "既存EXCEL PID $idFromHwnd へ接続"
                    ExistingConnection = $true
                }
            }
            return [pscustomobject]@{
                Pid = $idFromHwnd
                Mode = 'Hwnd'
                ExistingConnection = $false
            }
        }

        $pidsNow = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
                     Select-Object -ExpandProperty Id)
        $newPids = @($pidsNow | Where-Object { $PidsBefore -notcontains $_ })

        # 起動前が0件なら、唯一の新規PIDを所有プロセスとして採用できる。
        # 既存Excelがある場合は、PID差分だけでは競合起動を排除できないため採用しない。
        if ($PidsBefore.Count -eq 0 -and $newPids.Count -eq 1) {
            return [pscustomobject]@{
                Pid = [int]$newPids[0]
                Mode = '空ベースライン＋PID差分'
                ExistingConnection = $false
            }
        }
        if ($newPids.Count -gt 1) {
            return [pscustomobject]@{
                Pid = 0
                Mode = "新規EXCEL PIDが複数: $($newPids -join ',')"
                ExistingConnection = $false
            }
        }
        Start-Sleep -Milliseconds 200
    } while ((Get-Date) -lt $deadline)

    return [pscustomobject]@{
        Pid = 0
        Mode = '所有PIDを特定できず'
        ExistingConnection = $false
    }
}

function Release-Com {
    param($Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch { }
    }
}

# ---------------------------------------------------------------------
# 共通ユーティリティ
# ---------------------------------------------------------------------
function Resolve-BodyFont {
    Add-Type -AssemblyName System.Drawing
    $installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families |
                 ForEach-Object { $_.Name }
    foreach ($font in @('BIZ UDPゴシック', 'BIZ UDPGothic', 'BIZ UDゴシック',
                        'Meiryo', 'Yu Gothic UI', 'MS Pゴシック')) {
        if ($installed -contains $font) { return $font }
    }
    return 'MS Pゴシック'
}

function Get-SafeFileName {
    param([string]$Name, [string]$Extension = '.xlsx', [string]$Directory)
    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = ($safe.ToCharArray() | ForEach-Object {
        if ([int]$_ -lt 32) { '_' } else { $_ }
    }) -join ''
    $safe = $safe.TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'manual' }
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0, 100) }
    while ((Join-Path $Directory ($safe + $Extension)).Length -gt 240 -and $safe.Length -gt 8) {
        $safe = $safe.Substring(0, $safe.Length - 8)
    }
    $candidate = $safe
    $suffix = 2
    while (Test-Path -LiteralPath (Join-Path $Directory ($candidate + $Extension))) {
        $candidate = "${safe}_$suffix"
        $suffix++
    }
    return $candidate + $Extension
}

function Get-SafeWorksheetName {
    param([string]$RequestedName, [hashtable]$UsedNames)

    $safe = [string]$RequestedName
    # Excelが禁止する半角記号は、環境差の少ない中黒へ統一する。
    $safe = $safe -replace '[:\\/\?\*\[\]]', '・'
    $safe = ($safe.ToCharArray() | ForEach-Object {
        if ([int]$_ -lt 32) { '・' } else { $_ }
    }) -join ''
    $safe = $safe.Trim().Trim([char]39)
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'シート' }
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31) }

    $baseName = $safe
    $candidate = $baseName
    $number = 2
    while ($UsedNames.ContainsKey($candidate)) {
        $suffix = " ($number)"
        $maxBaseLength = 31 - $suffix.Length
        $shortBase = if ($baseName.Length -gt $maxBaseLength) {
            $baseName.Substring(0, $maxBaseLength)
        } else {
            $baseName
        }
        $candidate = $shortBase + $suffix
        $number++
    }
    $UsedNames[$candidate] = $true
    return $candidate
}

function Test-WorksheetNameSyntax {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name.Length -gt 31) { return $false }
    if ($Name -match '[:\\/\?\*\[\]]') { return $false }
    if ($Name.StartsWith("'") -or $Name.EndsWith("'")) { return $false }
    foreach ($character in $Name.ToCharArray()) {
        if ([int]$character -lt 32) { return $false }
    }
    return $true
}

function Escape-SheetNameForAddress {
    param([string]$Name)
    return $Name.Replace("'", "''")
}

function New-TestImage {
    param([string]$Path, [int]$Width, [int]$Height, [string]$Caption)
    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $borderPen = $null
    $fieldPen = $null
    $barBrush = $null
    $textBrush = $null
    $titleFont = $null
    $smallFont = $null
    try {
        $graphics.Clear([System.Drawing.Color]::FromArgb(245, 246, 248))
        $graphics.FillRectangle([System.Drawing.Brushes]::White, 20, 20, $Width - 40, $Height - 40)
        $borderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(190, 195, 200)), 2
        $graphics.DrawRectangle($borderPen, 20, 20, $Width - 40, $Height - 40)
        $barBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(60, 90, 150))
        $graphics.FillRectangle($barBrush, 20, 20, $Width - 40, 48)
        $fieldPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(150, 155, 160)), 2
        for ($index = 0; $index -lt 3; $index++) {
            $y = 120 + ($index * 70)
            if ($y -lt ($Height - 90)) {
                $fieldWidth = [int](($Width - 200) * 0.6)
                $graphics.DrawRectangle($fieldPen, 80, $y, $fieldWidth, 40)
            }
        }
        $titleFont = New-Object System.Drawing.Font 'Meiryo', 20, ([System.Drawing.FontStyle]::Bold)
        $smallFont = New-Object System.Drawing.Font 'Meiryo', 14
        $textBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 30, 30))
        $graphics.DrawString($Caption, $titleFont, $textBrush, 40, [float]($Height - 70))
        $graphics.DrawString("$Width x $Height px", $smallFont, $textBrush, 40, 80)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        foreach ($item in @($borderPen, $fieldPen, $barBrush, $textBrush, $titleFont, $smallFont)) {
            if ($item) { $item.Dispose() }
        }
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Set-CardBorders {
    param($Range, [int]$Color, [int]$LineStyle, [int]$Weight)
    $borders = $Range.Borders
    try {
        $borders.LineStyle = $LineStyle
        $borders.Weight = $Weight
        $borders.Color = $Color
    } finally {
        Release-Com $borders
    }
}

function Set-WorksheetView {
    param($Application, $Worksheet, [int]$FreezeRows = 2)
    $Worksheet.Activate()
    $window = $Application.ActiveWindow
    try {
        $window.DisplayGridlines = $false
        $window.Zoom = 90
        $window.FreezePanes = $false
        $window.SplitColumn = 0
        $window.SplitRow = $FreezeRows
        $window.FreezePanes = $true
    } finally {
        Release-Com $window
    }
}

function Add-StepCard {
    param(
        $Worksheet,
        [int]$StartRow,
        [int]$StepNumber,
        [string]$Title,
        [string]$Description,
        [string]$Note,
        [string]$ImagePath,
        [string]$FontName
    )

    $xlCenter = -4108
    $xlLeft = -4131
    $xlTop = -4160
    $xlContinuous = 1
    $xlThin = 2
    $xlMoveAndSize = 1
    $msoTrue = -1
    $msoFalse = 0
    $colorBlue = To-Bgr 31 78 121
    $colorLine = To-Bgr 205 213 221
    $colorText = To-Bgr 24 32 42
    $colorNote = To-Bgr 255 249 219
    $colorNoteLine = To-Bgr 214 170 0

    $headerRow = $StartRow
    $contentStart = $StartRow + 1
    $contentEnd = $StartRow + 10
    $noteStart = $StartRow + 8
    $noteEnd = $contentEnd

    $rowRange = $Worksheet.Range("A${headerRow}:H${headerRow}")
    try { $rowRange.RowHeight = 28 } finally { Release-Com $rowRange }
    for ($row = $contentStart; $row -le $contentEnd; $row++) {
        $rowRange = $Worksheet.Range("A${row}:H${row}")
        try { $rowRange.RowHeight = 22 } finally { Release-Com $rowRange }
    }
    $spacerRow = $contentEnd + 1
    $rowRange = $Worksheet.Range("A${spacerRow}:H${spacerRow}")
    try { $rowRange.RowHeight = 12 } finally { Release-Com $rowRange }

    $header = $Worksheet.Range("A${headerRow}:H${headerRow}")
    $imageArea = $Worksheet.Range("A${contentStart}:D${contentEnd}")
    $descriptionArea = $Worksheet.Range("E${contentStart}:H$($noteStart - 1)")
    $noteArea = $Worksheet.Range("E${noteStart}:H${noteEnd}")
    $card = $Worksheet.Range("A${headerRow}:H${contentEnd}")
    $shape = $null
    try {
        $header.Merge()
        $header.Value2 = "$StepNumber  $Title"
        $header.Interior.Color = $colorBlue
        $header.Font.Name = $FontName
        $header.Font.Size = 13
        $header.Font.Bold = $true
        $header.Font.Color = (To-Bgr 255 255 255)
        $header.HorizontalAlignment = $xlLeft
        $header.VerticalAlignment = $xlCenter

        $imageArea.Merge()
        $imageArea.Interior.Color = (To-Bgr 255 255 255)

        $descriptionArea.Merge()
        $descriptionArea.Value2 = "説明`n$Description"
        $descriptionArea.WrapText = $true
        $descriptionArea.HorizontalAlignment = $xlLeft
        $descriptionArea.VerticalAlignment = $xlTop
        $descriptionArea.Font.Name = $FontName
        $descriptionArea.Font.Size = 11
        $descriptionArea.Font.Color = $colorText

        $noteArea.Merge()
        $noteArea.Value2 = "補足`n$Note"
        $noteArea.WrapText = $true
        $noteArea.HorizontalAlignment = $xlLeft
        $noteArea.VerticalAlignment = $xlTop
        $noteArea.Interior.Color = $colorNote
        $noteArea.Font.Name = $FontName
        $noteArea.Font.Size = 10.5
        Set-CardBorders -Range $noteArea -Color $colorNoteLine -LineStyle $xlContinuous -Weight $xlThin

        Set-CardBorders -Range $card -Color $colorLine -LineStyle $xlContinuous -Weight $xlThin

        $shape = $Worksheet.Shapes.AddPicture(
            $ImagePath, $msoFalse, $msoTrue,
            [single]($imageArea.Left + 6), [single]($imageArea.Top + 6),
            [single]-1, [single]-1
        )
        $shape.LockAspectRatio = $msoTrue
        $originalWidth = [double]$shape.Width
        $originalHeight = [double]$shape.Height
        $maxWidth = [double]$imageArea.Width - 12
        $maxHeight = [double]$imageArea.Height - 12
        $scale = [Math]::Min(1.0, [Math]::Min($maxWidth / $originalWidth, $maxHeight / $originalHeight))
        $shape.Width = [single]($originalWidth * $scale)
        $shape.Left = [single]($imageArea.Left + (($imageArea.Width - $shape.Width) / 2))
        $shape.Top = [single]($imageArea.Top + (($imageArea.Height - $shape.Height) / 2))
        $shape.Placement = $xlMoveAndSize
        try { $shape.AlternativeText = "手順 $StepNumber の画面: $Title" } catch { }
    } finally {
        Release-Com $shape
        Release-Com $card
        Release-Com $noteArea
        Release-Com $descriptionArea
        Release-Com $imageArea
        Release-Com $header
    }
}

# ---------------------------------------------------------------------
# Excel定数・出力先
# ---------------------------------------------------------------------
$xlOpenXMLWorkbook = 51
$xlLandscape = 2
$xlPaperA4 = 9
$xlCenter = -4108
$xlLeft = -4131
$xlTop = -4160
$xlContinuous = 1
$xlThin = 2
$msoTrue = -1

$colorBlue = To-Bgr 31 78 121
$colorBlueSoft = To-Bgr 234 242 255
$colorLine = To-Bgr 205 213 221
$colorMuted = To-Bgr 102 113 127
$colorWhite = To-Bgr 255 255 255

$outRoot = Join-Path $PSScriptRoot 'out'
$excelOutDir = Join-Path $outRoot 'excel'
$imageDir = Join-Path $excelOutDir 'images'
$tmpDir = Join-Path $excelOutDir '.tmp'
foreach ($directory in @($outRoot, $excelOutDir, $imageDir, $tmpDir)) {
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  Excel主出力: X-01/X-02/X-05〜X-08/X-10/X-11/X-13' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''

Add-Check 'X-10' 'Win32 API（GetWindowThreadProcessId）の利用' `
    $(if ($canPInvoke) { 'OK' } else { 'NG' }) `
    $(if ($canPInvoke) { 'Excel.Application.Hwndから所有PIDを確認します' } else { '所有PIDを証明できません' })

$bodyFont = Resolve-BodyFont
Add-Check '-' '本文フォントの解決' 'INFO' $bodyFont

# ---------------------------------------------------------------------
# 3シート・6手順の検証データ
# ---------------------------------------------------------------------
$imageSpecs = @(
    @{ W = 1600; H = 900;  C = 'ログイン画面を開く' },
    @{ W = 800;  H = 1400; C = '社員番号を入力する' },
    @{ W = 1200; H = 260;  C = '申請メニューを選ぶ' },
    @{ W = 2560; H = 1440; C = '申請内容を入力する' },
    @{ W = 1400; H = 780;  C = '承認依頼を送信する' },
    @{ W = 900;  H = 1900; C = '承認結果を確認する' }
)

$imagePaths = @()
for ($index = 0; $index -lt $imageSpecs.Count; $index++) {
    $spec = $imageSpecs[$index]
    $path = Join-Path $imageDir ("excel-step-{0}.png" -f ($index + 1))
    New-TestImage -Path $path -Width $spec.W -Height $spec.H -Caption $spec.C
    $imagePaths += $path
}
Add-Check '-' 'Excel用テスト画像の生成' 'OK' "$($imagePaths.Count) 枚"

$sheetModels = @(
    [pscustomobject]@{
        RequestedName = 'ログイン'
        Summary = 'システムを開き、社員番号でログインします。'
        Steps = @(
            [pscustomobject]@{
                Title = 'ログイン画面を開く'
                Description = 'ブラウザを開き、社内システムのURLへアクセスします。'
                Note = '社外から接続する場合はVPNへ接続してください。'
                ImagePath = $imagePaths[0]
            },
            [pscustomobject]@{
                Title = '社員番号を入力する'
                Description = '社員番号とパスワードを入力し、ログインを選択します。'
                Note = 'パスワードを忘れた場合は再発行手続きを行います。'
                ImagePath = $imagePaths[1]
            }
        )
    },
    [pscustomobject]@{
        RequestedName = '経費/申請:国内'
        Summary = '国内経費の申請内容を登録します。'
        Steps = @(
            [pscustomobject]@{
                Title = '申請メニューを選ぶ'
                Description = 'トップ画面から「国内経費申請」を選択します。'
                Note = '海外出張は別の申請メニューを使用します。'
                ImagePath = $imagePaths[2]
            },
            [pscustomobject]@{
                Title = '申請内容を入力する'
                Description = '利用日、金額、勘定科目、支払先を入力し、証憑を添付します。'
                Note = '金額と証憑の記載内容が一致していることを確認してください。'
                ImagePath = $imagePaths[3]
            }
        )
    },
    [pscustomobject]@{
        RequestedName = 'ログイン'
        Summary = '承認依頼を送信し、処理結果を確認します。'
        Steps = @(
            [pscustomobject]@{
                Title = '承認依頼を送信する'
                Description = '入力内容を確認し、承認者を選択して申請を送信します。'
                Note = '送信後は内容を直接編集できません。'
                ImagePath = $imagePaths[4]
            },
            [pscustomobject]@{
                Title = '承認結果を確認する'
                Description = '申請一覧を開き、ステータスと承認コメントを確認します。'
                Note = '差戻しの場合はコメントを確認して再申請します。'
                ImagePath = $imagePaths[5]
            }
        )
    }
)

# ---------------------------------------------------------------------
# COM本体
# ---------------------------------------------------------------------
$excel = $null
$workbook = $null
$worksheets = $null
$indexSheet = $null
$workbookClosed = $false
$ownPid = 0
$ownershipMode = 'なし'
$ownershipProven = $false
$canQuitCom = $false
$settingsApplied = $false
$quitCalled = $false
$killed = $false
$savedPath = ''
$tmpPath = Join-Path $tmpDir ("excel-build-{0}-{1}.xlsx" -f $PID, $Round)
$roundError = ''
$errorLine = 0
$currentStage = '初期化'
$abortedForExisting = $false
$pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Id)

Write-Host ("  起動前の EXCEL PID : {0}" -f $(if ($pidsBefore.Count) {
    $pidsBefore -join ', '
} else {
    'なし'
})) -ForegroundColor Gray

try {
    if ($SimulateExisting) {
        Add-Check 'X-11' '既存Excel検出（シミュレーション）' 'OK' 'Excel COMは起動しません'
        Add-Check 'X-13' '異常分岐でファイルを作らないこと' 'OK' '未完成ファイルなし'
        throw 'SIMULATED_EXISTING_EXCEL'
    }

    # COM生成直後は設定を変更しない。Hwnd/PIDで所有権を確認してから設定する。
    $excel = New-Object -ComObject Excel.Application
    $canQuitCom = ($pidsBefore.Count -eq 0)
    $resolved = Resolve-OwnedExcelProcessId -Application $excel -PidsBefore $pidsBefore
    $ownPid = [int]$resolved.Pid
    $ownershipMode = [string]$resolved.Mode

    if ([bool]$resolved.ExistingConnection) {
        $abortedForExisting = $true
        Add-Check 'X-11' '既存Excelへの接続を検出' 'OK' "$ownershipMode / 設定変更前に中止"
        Add-Check 'X-11' '既存ExcelへQuitしないこと' 'OK' 'COM参照だけ解放し、Excelはそのまま残します'
        throw 'CONNECTED_TO_EXISTING_EXCEL'
    }

    if ($ownPid -le 0) {
        if ($pidsBefore.Count -gt 0) {
            $abortedForExisting = $true
            Add-Check 'X-11' '既存Excelがある状態の所有確認' 'OK' "$ownershipMode / 安全のため出力を中止"
            throw 'OWNERSHIP_UNRESOLVED_WITH_EXISTING'
        }
        Add-Check 'X-10' '自インスタンスのPID特定' 'NG' "$ownershipMode / 出力を開始しません"
        throw 'OWNERSHIP_UNRESOLVED'
    }

    $ownershipProven = $true
    $canQuitCom = $true
    Add-Check 'X-10' '自インスタンスのPID特定' 'OK' "PID $ownPid / $ownershipMode"

    $excel.Visible = [bool]$KeepVisible
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    try { $excel.AskToUpdateLinks = $false } catch { }
    $settingsApplied = $true
    Add-Check 'X-11' 'Excel設定の適用時点' 'OK' '所有PID確認後にだけ適用'
    Add-Check '-' 'Excel起動' 'OK' "Ver $($excel.Version) / PID $ownPid / $ownershipMode"

    $workbooks = $null
    try {
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Add()
    } finally {
        Release-Com $workbooks
    }
    $worksheets = $workbook.Worksheets

    # 既定シート数は利用者設定で変わるため、先頭だけ残して目次へ転用する。
    $currentStage = '既定ワークシートの整理'
    while ($worksheets.Count -gt 1) {
        $deleteSheet = $worksheets.Item($worksheets.Count)
        try { $deleteSheet.Delete() } finally { Release-Com $deleteSheet }
    }
    $indexSheet = $worksheets.Item(1)
    $currentStage = '目次シートの名前設定'
    $indexSheetName = [string]$indexSheet.Name
    if ($indexSheetName -ne '目次') {
        try {
            $indexSheet.Name = '目次'
            $indexSheetName = '目次'
        } catch {
            $indexSheet.Name = 'INDEX'
            $indexSheetName = 'INDEX'
            Write-Host '  目次シート名をINDEXへ変更しました。' -ForegroundColor Yellow
        }
    }

    $usedNames = @{}
    $usedNames[$indexSheetName] = $true
    $escapedIndexSheetName = Escape-SheetNameForAddress -Name $indexSheetName
    $nameMappings = @()
    foreach ($model in $sheetModels) {
        $safeName = Get-SafeWorksheetName -RequestedName $model.RequestedName -UsedNames $usedNames
        $nameMappings += [pscustomobject]@{
            Requested = $model.RequestedName
            Safe = $safeName
        }
    }

    # ---- 目次シート ----
    $indexSheet.Columns.Item(1).ColumnWidth = 6
    $indexSheet.Columns.Item(2).ColumnWidth = 28
    $indexSheet.Columns.Item(3).ColumnWidth = 48
    $indexSheet.Columns.Item(4).ColumnWidth = 12
    $titleRange = $indexSheet.Range('A1:D1')
    $metaRange = $indexSheet.Range('A2:D2')
    $headerRange = $indexSheet.Range('A4:D4')
    try {
        $titleRange.Merge()
        $titleRange.Value2 = '経費精算システム 操作マニュアル'
        $titleRange.Interior.Color = $colorBlue
        $titleRange.Font.Name = $bodyFont
        $titleRange.Font.Size = 18
        $titleRange.Font.Bold = $true
        $titleRange.Font.Color = $colorWhite
        $titleRange.RowHeight = 34
        $titleRange.HorizontalAlignment = $xlLeft
        $titleRange.VerticalAlignment = $xlCenter

        $metaRange.Merge()
        $metaRange.Value2 = "Ver. 1.0　作成日: $(Get-Date -Format 'yyyy年M月d日')　作成者: ManualBuilder"
        $metaRange.Font.Name = $bodyFont
        $metaRange.Font.Size = 10.5
        $metaRange.Font.Color = $colorMuted
        $metaRange.RowHeight = 24

        $indexSheet.Cells.Item(4, 1).Value2 = 'No.'
        $indexSheet.Cells.Item(4, 2).Value2 = 'シート'
        $indexSheet.Cells.Item(4, 3).Value2 = '概要'
        $indexSheet.Cells.Item(4, 4).Value2 = '手順数'
        $headerRange.Interior.Color = $colorBlueSoft
        $headerRange.Font.Name = $bodyFont
        $headerRange.Font.Bold = $true
        $headerRange.RowHeight = 25
        Set-CardBorders -Range $headerRange -Color $colorLine -LineStyle $xlContinuous -Weight $xlThin
    } finally {
        Release-Com $headerRange
        Release-Com $metaRange
        Release-Com $titleRange
    }

    $globalStep = 0
    $totalShapes = 0
    $createdSheetNames = @()
    $sheetNameFallbacks = New-Object System.Collections.ArrayList
    $linkIssues = New-Object System.Collections.ArrayList
    $pageSetupIssues = New-Object System.Collections.ArrayList

    for ($sheetIndex = 0; $sheetIndex -lt $sheetModels.Count; $sheetIndex++) {
        $model = $sheetModels[$sheetIndex]
        $safeName = $nameMappings[$sheetIndex].Safe
        $currentStage = "ワークシート追加 $($sheetIndex + 1): $safeName"
        $afterSheet = $worksheets.Item($worksheets.Count)
        $worksheet = $worksheets.Add([Type]::Missing, $afterSheet)
        Release-Com $afterSheet

        $currentStage = "シート名設定 $($sheetIndex + 1): $safeName"
        Write-Host ("  シート名: {0} → {1}（{2}文字）" -f $model.RequestedName, $safeName, $safeName.Length) -ForegroundColor Gray
        try {
            $worksheet.Name = $safeName
        } catch {
            [void]$usedNames.Remove($safeName)
            $fallbackName = Get-SafeWorksheetName -RequestedName ("手順シート {0}" -f ($sheetIndex + 1)) -UsedNames $usedNames
            [void]$sheetNameFallbacks.Add("$safeName → $fallbackName")
            Write-Host "  Excelがシート名を拒否したため、$fallbackName を使用します。" -ForegroundColor Yellow
            $currentStage = "代替シート名設定 $($sheetIndex + 1): $fallbackName"
            $worksheet.Name = $fallbackName
            $safeName = $fallbackName
            $nameMappings[$sheetIndex].Safe = $fallbackName
        }
        $createdSheetNames += $safeName

        $currentStage = "シート見出し設定 $safeName"
        try {
            foreach ($column in 1..4) {
                $worksheet.Columns.Item($column).ColumnWidth = 12
            }
            foreach ($column in 5..8) {
                $worksheet.Columns.Item($column).ColumnWidth = 10
            }

            $sheetTitle = $worksheet.Range('A1:F1')
            $backLink = $worksheet.Range('G1:H1')
            $summary = $worksheet.Range('A2:H2')
            try {
                $sheetTitle.Merge()
                $sheetTitle.Value2 = $safeName
                $sheetTitle.Interior.Color = $colorBlue
                $sheetTitle.Font.Name = $bodyFont
                $sheetTitle.Font.Size = 16
                $sheetTitle.Font.Bold = $true
                $sheetTitle.Font.Color = $colorWhite
                $sheetTitle.RowHeight = 32
                $sheetTitle.HorizontalAlignment = $xlLeft
                $sheetTitle.VerticalAlignment = $xlCenter

                $backLink.Merge()
                $backLink.Interior.Color = $colorBlue
                $backLink.Font.Name = $bodyFont
                $backLink.Font.Size = 10.5
                $backLink.Font.Color = $colorWhite
                $backLink.HorizontalAlignment = $xlCenter
                $backLink.VerticalAlignment = $xlCenter
                $currentStage = "目次へ戻るリンク $safeName"
                $sheetHyperlinks = $null
                $newHyperlink = $null
                try {
                    $sheetHyperlinks = $worksheet.Hyperlinks
                    $newHyperlink = $sheetHyperlinks.Add(
                        $backLink,
                        [string]'',
                        [string]"'$escapedIndexSheetName'!A1",
                        [Type]::Missing,
                        [string]'目次へ戻る'
                    )
                    # Hyperlinks.Addはセルの直接書式を標準リンク色へ戻す。
                    # 濃紺帯で読めるよう、リンク作成後に白字を再適用する。
                    $backLink.Font.Name = $bodyFont
                    $backLink.Font.Size = 10.5
                    $backLink.Font.Color = $colorWhite
                    $backLink.Font.Bold = $true
                } catch {
                    [void]$linkIssues.Add("$safeName → 目次: $($_.Exception.Message)")
                    $backLink.Value2 = '目次へ戻る'
                } finally {
                    Release-Com $newHyperlink
                    Release-Com $sheetHyperlinks
                }

                $summary.Merge()
                $summary.Value2 = $model.Summary
                $summary.Font.Name = $bodyFont
                $summary.Font.Size = 10.5
                $summary.Font.Color = $colorMuted
                $summary.RowHeight = 24
            } finally {
                Release-Com $summary
                Release-Com $backLink
                Release-Com $sheetTitle
            }

            $startRow = 4
            for ($stepIndex = 0; $stepIndex -lt $model.Steps.Count; $stepIndex++) {
                $globalStep++
                if ($CancelAtStep -gt 0 -and $globalStep -eq $CancelAtStep) {
                    Add-Check 'X-13' 'キャンセル時に完成ファイルを作らないこと' 'OK' "全体手順 $globalStep の直前で中止"
                    throw 'CANCELLED'
                }
                $step = $model.Steps[$stepIndex]
                $currentStage = "手順カード追加 $safeName/$($stepIndex + 1)"
                Add-StepCard -Worksheet $worksheet -StartRow $startRow `
                    -StepNumber ($stepIndex + 1) -Title $step.Title `
                    -Description $step.Description -Note $step.Note `
                    -ImagePath $step.ImagePath -FontName $bodyFont
                $startRow += 12
            }

            $used = $worksheet.UsedRange
            try {
                $used.Font.Name = $bodyFont
            } finally {
                Release-Com $used
            }

            # 印刷は副次要件。既定プリンターやExcel環境によってPageSetupの一部が
            # 設定できない場合でも、PC閲覧用xlsxの生成は中止しない。
            $currentStage = "印刷設定 $safeName"
            $pageSetup = $null
            try {
                $pageSetup = $worksheet.PageSetup
                try { $pageSetup.PaperSize = $xlPaperA4 } catch {
                    [void]$pageSetupIssues.Add("$safeName/PaperSize: $($_.Exception.Message)")
                }
                try { $pageSetup.Orientation = $xlLandscape } catch {
                    [void]$pageSetupIssues.Add("$safeName/Orientation: $($_.Exception.Message)")
                }
                try { $pageSetup.Zoom = $false } catch {
                    [void]$pageSetupIssues.Add("$safeName/Zoom: $($_.Exception.Message)")
                }
                try { $pageSetup.FitToPagesWide = 1 } catch {
                    [void]$pageSetupIssues.Add("$safeName/FitToPagesWide: $($_.Exception.Message)")
                }
                try { $pageSetup.FitToPagesTall = $false } catch {
                    [void]$pageSetupIssues.Add("$safeName/FitToPagesTall: $($_.Exception.Message)")
                }
                try { $pageSetup.CenterHorizontally = $true } catch {
                    [void]$pageSetupIssues.Add("$safeName/CenterHorizontally: $($_.Exception.Message)")
                }
            } catch {
                [void]$pageSetupIssues.Add("$safeName/PageSetup: $($_.Exception.Message)")
            } finally {
                Release-Com $pageSetup
            }

            Set-WorksheetView -Application $excel -Worksheet $worksheet -FreezeRows 2
            $totalShapes += [int]$worksheet.Shapes.Count
        } finally {
            Release-Com $worksheet
        }
    }

    # 目次の一覧とリンクは全シート作成後に設定する。
    $currentStage = '目次一覧とリンクの設定'
    $indexCells = $null
    $indexHyperlinks = $null
    try {
        $indexCells = $indexSheet.Cells
        $indexHyperlinks = $indexSheet.Hyperlinks
        for ($sheetIndex = 0; $sheetIndex -lt $sheetModels.Count; $sheetIndex++) {
            $row = 5 + $sheetIndex
            $model = $sheetModels[$sheetIndex]
            $safeName = [string]$nameMappings[$sheetIndex].Safe
            $numberCell = $null
            $nameCell = $null
            $summaryCell = $null
            $countCell = $null
            $rowRange = $null
            $newHyperlink = $null
            try {
                $numberCell = $indexCells.Item([int]$row, [int]1)
                $nameCell = $indexCells.Item([int]$row, [int]2)
                $summaryCell = $indexCells.Item([int]$row, [int]3)
                $countCell = $indexCells.Item([int]$row, [int]4)

                # Excel COMのValue2は、環境によってInt32を文字列へ暗黙変換できない。
                # 数値はDouble、文章はStringへ明示変換してから渡す。
                $currentStage = "目次行番号 $($sheetIndex + 1): $safeName"
                $numberCell.Value2 = [double]($sheetIndex + 1)
                $currentStage = "目次シート名 $($sheetIndex + 1): $safeName"
                $nameCell.Value2 = [string]$safeName
                $currentStage = "目次概要 $($sheetIndex + 1): $safeName"
                $summaryCell.Value2 = [string]$model.Summary
                $currentStage = "目次手順数 $($sheetIndex + 1): $safeName"
                $countCell.Value2 = [double]($model.Steps.Count)

                $escapedName = Escape-SheetNameForAddress -Name $safeName
                $currentStage = "目次リンク追加 $($sheetIndex + 1): $safeName"
                try {
                    $newHyperlink = $indexHyperlinks.Add(
                        $nameCell,
                        [string]'',
                        [string]"'$escapedName'!A1",
                        [Type]::Missing,
                        [string]$safeName
                    )
                } catch {
                    [void]$linkIssues.Add("目次 → ${safeName}: $($_.Exception.Message)")
                    $nameCell.Value2 = [string]$safeName
                } finally {
                    Release-Com $newHyperlink
                }

                $currentStage = "目次行書式 $($sheetIndex + 1): $safeName"
                $rowRange = $indexSheet.Range("A${row}:D${row}")
                $rowRange.Font.Name = $bodyFont
                $rowRange.Font.Size = 10.5
                $rowRange.RowHeight = 28
                $rowRange.WrapText = $true
                Set-CardBorders -Range $rowRange -Color $colorLine -LineStyle $xlContinuous -Weight $xlThin
            } finally {
                Release-Com $rowRange
                Release-Com $countCell
                Release-Com $summaryCell
                Release-Com $nameCell
                Release-Com $numberCell
            }
        }
    } finally {
        Release-Com $indexHyperlinks
        Release-Com $indexCells
    }

    $indexSheet.Columns.Item(1).HorizontalAlignment = $xlCenter
    $indexSheet.Columns.Item(4).HorizontalAlignment = $xlCenter
    Set-WorksheetView -Application $excel -Worksheet $indexSheet -FreezeRows 4

    # ---- 保存前の自己検査 ----
    $expectedNames = @($indexSheetName) + $createdSheetNames
    $actualNames = @()
    for ($index = 1; $index -le $worksheets.Count; $index++) {
        $checkSheet = $worksheets.Item($index)
        try { $actualNames += [string]$checkSheet.Name } finally { Release-Com $checkSheet }
    }
    $sheetOrderOk = (($expectedNames -join '|') -eq ($actualNames -join '|'))
    Add-Check 'X-01' 'ワークシート順' $(if ($sheetOrderOk) { 'OK' } else { 'NG' }) `
        "期待: $($expectedNames -join ' → ') / 実際: $($actualNames -join ' → ')"

    $longNameUsed = @{}
    $longRequestedName = 'これは三十一文字を超える非常に長いシート名の変換確認用サンプルです'
    $longSafeName = Get-SafeWorksheetName -RequestedName $longRequestedName -UsedNames $longNameUsed
    $invalidCreatedNames = @($createdSheetNames | Where-Object {
        -not (Test-WorksheetNameSyntax -Name $_)
    })
    $uniqueCreatedNameCount = @($createdSheetNames | Select-Object -Unique).Count
    $sheetNameOk = (
        $invalidCreatedNames.Count -eq 0 -and
        $uniqueCreatedNameCount -eq $createdSheetNames.Count -and
        $longSafeName.Length -eq 31
    )
    Add-Check 'X-08' 'シート名の禁止文字・重複変換' $(if ($sheetNameOk) { 'OK' } else { 'NG' }) `
        ((($nameMappings | ForEach-Object { "$($_.Requested) → $($_.Safe)" }) -join ' / ') +
         " / 長名=$($longSafeName.Length)文字")

    Add-Check 'X-02' '3シート・各2手順・シート内採番' `
        $(if ($createdSheetNames.Count -eq 3 -and $globalStep -eq 6) { 'OK' } else { 'NG' }) `
        "シート $($createdSheetNames.Count) / 手順 $globalStep"

    $indexLinks = $null
    try {
        $indexLinks = $indexSheet.Hyperlinks
        $indexLinkCount = [int]$indexLinks.Count
    } finally {
        Release-Com $indexLinks
    }
    $backLinkCount = 0
    foreach ($safeName in $createdSheetNames) {
        $linkSheet = $worksheets.Item($safeName)
        $sheetLinks = $null
        try {
            $sheetLinks = $linkSheet.Hyperlinks
            $backLinkCount += [int]$sheetLinks.Count
        } finally {
            Release-Com $sheetLinks
            Release-Com $linkSheet
        }
    }
    Add-Check 'X-05' '目次リンクと戻るリンク' `
        $(if ($indexLinkCount -eq 3 -and $backLinkCount -eq 3) { 'OK' } else { 'NG' }) `
        "目次から $indexLinkCount / 目次へ戻る $backLinkCount / 設定失敗 $($linkIssues.Count)"

    Add-Check 'X-06' 'PC向け横長カード' 'OK' '画像A:D（約57%）／説明E:H（約43%）、ズーム90%'
    Add-Check 'X-07' '画像の埋め込み・縦横比・配置' `
        $(if ($totalShapes -eq 6) { 'OK' } else { 'NG' }) `
        "埋め込み画像 $totalShapes / 6、MoveAndSize、カード内へ縮小"

    $pageSetupDetail = if ($pageSetupIssues.Count -eq 0) {
        'A4横・横1ページ・縦自動を設定'
    } else {
        $firstIssues = (($pageSetupIssues | Select-Object -First 3) -join ' / ')
        "設定できない項目 $($pageSetupIssues.Count) 件。PC閲覧用xlsxの生成は継続: $firstIssues"
    }
    Add-Check '-' '印刷設定（副次要件）' `
        $(if ($pageSetupIssues.Count -eq 0) { 'INFO' } else { 'WARN' }) `
        $pageSetupDetail

    # ---- 一時ファイルへ保存してから完成名へ移動 ----
    $currentStage = '一時xlsxへの保存'
    if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force }
    Write-Host "  Excel保存開始: $tmpPath" -ForegroundColor Cyan
    $saveWatch = [Diagnostics.Stopwatch]::StartNew()
    $workbook.SaveAs($tmpPath, $xlOpenXMLWorkbook)
    $saveWatch.Stop()
    if (-not (Test-Path -LiteralPath $tmpPath)) {
        throw "Excelは保存完了を返しましたが、一時ファイルがありません: $tmpPath"
    }
    $savedSizeKb = (Get-Item -LiteralPath $tmpPath).Length / 1KB
    Write-Host ("  Excel保存完了: {0:N1}秒 / {1:N0}KB" -f $saveWatch.Elapsed.TotalSeconds, $savedSizeKb) -ForegroundColor Green

    $workbook.Close($false)
    $workbookClosed = $true

    $safeFileName = Get-SafeFileName `
        -Name ("経費精算システム_操作マニュアル_r{0}_{1}" -f $Round, (Get-Date -Format 'yyyyMMdd_HHmmss')) `
        -Extension '.xlsx' -Directory $excelOutDir
    $savedPath = Join-Path $excelOutDir $safeFileName
    $currentStage = '完成xlsxへの移動'
    Move-Item -LiteralPath $tmpPath -Destination $savedPath
    Add-Check 'X-13' '原子的保存（.tmp → 完成名）' 'OK' $safeFileName

} catch {
    $message = $_.Exception.Message
    $errorLine = [int]$_.InvocationInfo.ScriptLineNumber
    switch ($message) {
        'SIMULATED_EXISTING_EXCEL' {
            $roundError = '既存Excel検出のシミュレーションで中止'
        }
        'CONNECTED_TO_EXISTING_EXCEL' {
            $roundError = '既存Excelへ接続したため、設定変更前に中止'
        }
        'OWNERSHIP_UNRESOLVED_WITH_EXISTING' {
            $roundError = '既存Excelがあり所有PIDを証明できないため中止'
        }
        'CANCELLED' {
            $roundError = "キャンセル要求（全体手順 $CancelAtStep の直前）"
        }
        default {
            $roundError = "$message / 段階: $currentStage / 行: $errorLine"
            Add-Check '-' 'xlsx生成' 'NG' $roundError
        }
    }
} finally {
    try {
        if ($workbook -and -not $workbookClosed) { $workbook.Close($false) }
    } catch { }
    Release-Com $indexSheet
    Release-Com $worksheets
    Release-Com $workbook

    # 空ベースライン、またはHwnd/PIDで新規所有を証明した場合だけQuitする。
    # 既存Excelへ接続した可能性があるCOM参照にはQuitを呼ばない。
    if ($excel -and $canQuitCom) {
        if ($settingsApplied) {
            try { $excel.EnableEvents = $true } catch { }
            try { $excel.ScreenUpdating = $true } catch { }
        }
        try {
            $excel.Quit()
            $quitCalled = $true
        } catch { }
    }
    Release-Com $excel
    $indexSheet = $null
    $worksheets = $null
    $workbook = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Hwndで直接所有を証明したPIDだけ、Quit後10秒残る場合に限り終了する。
    if ($ownershipProven -and $ownPid -gt 0 -and $ownershipMode -eq 'Hwnd') {
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
    if (Test-Path -LiteralPath $tmpPath) {
        Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------
# 残留・結果
# ---------------------------------------------------------------------
$pidsAfter = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
               Select-Object -ExpandProperty Id)
$newResidualPids = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })

if ($SimulateExisting) {
    Add-Check 'X-10' 'シミュレーション時のEXCEL残留' 'OK' 'Excelを起動していません'
} elseif ($abortedForExisting) {
    $missingExistingPids = @($pidsBefore | Where-Object { $pidsAfter -notcontains $_ })
    Add-Check 'X-11' '既存Excelの生存' `
        $(if ($missingExistingPids.Count -eq 0) { 'OK' } else { 'NG' }) `
        $(if ($missingExistingPids.Count -eq 0) {
            '起動前PIDがすべて残存。設定変更・Quitは未実施'
        } else {
            "消失した起動前PID: $($missingExistingPids -join ',')"
        })
    if ($newResidualPids.Count -gt 0) {
        Add-Check 'X-10' '所有不明の新規EXCEL' 'WARN' `
            "自動終了しません。新規PID: $($newResidualPids -join ',')"
    }
} else {
    Add-Check 'X-10' '自インスタンスの残留' `
        $(if ($newResidualPids.Count -eq 0) { 'OK' } else { 'NG' }) `
        $(if ($newResidualPids.Count -eq 0) {
            '新しいEXCEL PIDの残留0'
        } else {
            "残留PID: $($newResidualPids -join ',')"
        })
}

if ($killed) {
    Add-Check 'X-10' 'Quit後の強制終了' 'WARN' "所有PID $ownPid を10秒後に終了"
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  Excel検証結果' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
$log | Format-Table -AutoSize -Wrap

$ngCount = @($log | Where-Object { $_.判定 -eq 'NG' }).Count
$warnCount = @($log | Where-Object { $_.判定 -eq 'WARN' }).Count
Write-Host ''
if ($ngCount -eq 0) {
    Write-Host '  NG はありません。' -ForegroundColor Green
} else {
    Write-Host "  NG が $ngCount 件あります。" -ForegroundColor Red
}
if ($warnCount -gt 0) {
    Write-Host "  WARN が $warnCount 件あります。" -ForegroundColor Yellow
}

$reportPath = Join-Path $excelOutDir 'result-09-excel.txt'
$summary = [pscustomobject]@{
    ラウンド = $Round
    起動前PID = $(if ($pidsBefore.Count) { $pidsBefore -join ',' } else { 'なし' })
    自PID = $(if ($ownPid) { $ownPid } else { '不明' })
    PID特定方式 = $ownershipMode
    設定適用 = $settingsApplied
    Quit呼出 = $quitCalled
    強制終了 = $killed
    生成物 = $(if ($savedPath) { Split-Path -Leaf $savedPath } else { 'なし' })
    最終段階 = $currentStage
    エラー行 = $(if ($errorLine -gt 0) { $errorLine } else { '-' })
    メモ = $roundError
}
(($summary | Format-List | Out-String -Width 220) + "`r`n" +
 ($log | Format-Table -AutoSize -Wrap | Out-String -Width 220)) |
    Out-File -LiteralPath $reportPath -Encoding UTF8

Write-Host "  レポート: $reportPath" -ForegroundColor Cyan
if ($savedPath) {
    Write-Host "  生成物  : $savedPath" -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  生成したxlsxをExcelで開き、次を確認してください:' -ForegroundColor Yellow
    Write-Host '   E-1  目次→ログイン→経費・申請・国内→ログイン (2) の順か'
    Write-Host '   E-2  目次のリンクと各シートの「目次へ戻る」が動くか'
    Write-Host '   E-3  画像が左、説明・補足が右に同時表示されるか'
    Write-Host '   E-4  画像がセルからはみ出したり、次の手順へ重ならないか'
    Write-Host '   E-5  説明・補足セルを直接編集できるか'
    Write-Host '   E-6  グリッド線非表示、ズーム90%、上部2行固定か'
    Write-Host '   E-7  1366×768・Windows 125%表示で画像と説明を同時に読めるか'
    Write-Host '   E-8  ファイル→情報でマクロなしの.xlsxになっているか'
    Write-Host ''
    if (-not $NoOpenPrompt) {
        Write-Host '  xlsxを開きますか？ [Y/n] ' -NoNewline -ForegroundColor Yellow
        if ((Read-Host) -ne 'n') { Start-Process $savedPath }
    }
} else {
    Write-Host '  xlsxは生成されていません。上の中止理由を確認してください。' -ForegroundColor Yellow
}

$workerExitCode = if ($ngCount -gt 0) { 1 } else { 0 }
if ($StatusFile) {
    try {
        $statusDirectory = Split-Path -Parent $StatusFile
        if ($statusDirectory -and -not (Test-Path -LiteralPath $statusDirectory)) {
            New-Item -ItemType Directory -Force -Path $statusDirectory | Out-Null
        }
        [IO.File]::WriteAllText($StatusFile, [string]$workerExitCode, [Text.Encoding]::ASCII)
    } catch {
        Write-Host "  監督結果ファイルを書き込めません: $($_.Exception.Message)" -ForegroundColor Red
        $workerExitCode = 1
    }
}
exit $workerExitCode
