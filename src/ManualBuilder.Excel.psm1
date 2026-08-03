# ManualBuilder Excel exporter.

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Drawing

# 動画つきの手順があるときだけ、ブックと動画をこのフォルダー名でまとめて出力する。
$script:MbExcelVideoFolderName = '動画'

$script:MbExcelPInvokeAvailable = $false
try {
    Add-Type -Namespace ManualBuilderExcel -Name NativeMethods -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@ -ErrorAction Stop
    $script:MbExcelPInvokeAvailable = $true
} catch {
    try {
        [void][ManualBuilderExcel.NativeMethods]
        $script:MbExcelPInvokeAvailable = $true
    } catch { }
}

function ConvertTo-MbExcelBgr {
    param([int]$R, [int]$G, [int]$B)
    return ($B * 65536) + ($G * 256) + $R
}

function Release-MbExcelComObject {
    param([AllowNull()][object]$InputObject)
    if ($null -ne $InputObject) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($InputObject) } catch { }
    }
}

function Get-MbExcelProcessId {
    param([Parameter(Mandatory = $true)][object]$Application)
    if (-not $script:MbExcelPInvokeAvailable) { return 0 }
    try {
        $hwnd = [IntPtr]([int64]$Application.Hwnd)
        if ($hwnd -eq [IntPtr]::Zero) { return 0 }
        $processId = [uint32]0
        [void][ManualBuilderExcel.NativeMethods]::GetWindowThreadProcessId($hwnd, [ref]$processId)
        return [int]$processId
    } catch {
        return 0
    }
}

function Resolve-MbOwnedExcelProcess {
    param(
        [Parameter(Mandatory = $true)][object]$Application,
        [int[]]$PidsBefore,
        [int]$TimeoutMilliseconds = 5000
    )

    $deadline = (Get-Date).AddMilliseconds($TimeoutMilliseconds)
    do {
        $idFromHwnd = Get-MbExcelProcessId -Application $Application
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

        $pidsNow = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $newPids = @($pidsNow | Where-Object { $PidsBefore -notcontains $_ })
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

function Resolve-MbExcelBodyFont {
    $collection = New-Object Drawing.Text.InstalledFontCollection
    try {
        $installed = @($collection.Families | ForEach-Object { $_.Name })
        foreach ($font in @('BIZ UDPゴシック', 'BIZ UDPGothic', 'BIZ UDゴシック', 'Meiryo', 'Yu Gothic UI', 'MS Pゴシック')) {
            if ($installed -contains $font) { return $font }
        }
        return 'MS Pゴシック'
    } finally {
        $collection.Dispose()
    }
}

function Get-MbSafeExcelFileName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Directory,
        [string]$Extension = '.xlsx'
    )

    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = ($safe.ToCharArray() | ForEach-Object { if ([int]$_ -lt 32) { '_' } else { $_ } }) -join ''
    $safe = $safe.TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'manual' }
    if ($safe -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $safe = '_' + $safe }
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

function Get-MbSafeExcelFolderName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Directory
    )
    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = ($safe.ToCharArray() | ForEach-Object { if ([int]$_ -lt 32) { '_' } else { $_ } }) -join ''
    $safe = $safe.TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'manual' }
    if ($safe -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $safe = '_' + $safe }
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0, 100) }
    while ((Join-Path $Directory $safe).Length -gt 200 -and $safe.Length -gt 8) {
        $safe = $safe.Substring(0, $safe.Length - 8)
    }
    $candidate = $safe
    $suffix = 2
    while (Test-Path -LiteralPath (Join-Path $Directory $candidate)) {
        $candidate = "${safe}_$suffix"
        $suffix++
    }
    return $candidate
}

function Get-MbExcelVideoPlan {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath
    )

    # Excelは動画を埋め込めないが、同じフォルダーへ動画を置けば =HYPERLINK() の相対パスから再生できる。
    # 相対パスはクリック時にブックの場所を基準に解決されるため、フォルダーごと移動しても効く。
    # （Hyperlinks.Addは保存時に絶対パスへ変換されるため使わない。）
    $stepLinks = @{}
    $files = New-Object System.Collections.ArrayList
    $namesByVideoId = @{}
    if ($Project.PSObject.Properties.Name -notcontains 'videos') {
        return [pscustomobject]@{ StepLinks = $stepLinks; Files = @(); Count = 0; FolderName = $script:MbExcelVideoFolderName }
    }
    $videoRoot = Join-Path (Split-Path -Parent $ProjectPath) 'videos'
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            if ($step.PSObject.Properties.Name -notcontains 'videoId') { continue }
            $videoId = [string]$step.videoId
            if ([string]::IsNullOrWhiteSpace($videoId)) { continue }
            if (-not $namesByVideoId.ContainsKey($videoId)) {
                $video = @($Project.videos | Where-Object { $_.id -eq $videoId }) | Select-Object -First 1
                if (-not $video -or [string]$video.fileName -notmatch '^video-[a-f0-9]{32}\.(mp4|webm)$') {
                    throw "手順が参照する動画が見つかりません: $($step.id)"
                }
                $sourcePath = Join-Path $videoRoot ([string]$video.fileName)
                if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "動画ファイルが見つかりません: $($video.fileName)" }
                # 同じ動画を複数の手順へ付けても、ファイルは1本だけ置く。
                $fileName = ('動画{0:d3}' -f ($namesByVideoId.Count + 1)) + [IO.Path]::GetExtension([string]$video.fileName).ToLowerInvariant()
                $namesByVideoId[$videoId] = $fileName
                [void]$files.Add([pscustomobject]@{ SourcePath = $sourcePath; FileName = $fileName })
            }
            $stepLinks[[string]$step.id] = $script:MbExcelVideoFolderName + '\' + $namesByVideoId[$videoId]
        }
    }
    return [pscustomobject]@{
        StepLinks = $stepLinks
        Files = @($files)
        Count = @($files).Count
        FolderName = $script:MbExcelVideoFolderName
    }
}

function Get-MbSafeExcelWorksheetName {
    param(
        [AllowEmptyString()][string]$RequestedName,
        [Parameter(Mandatory = $true)][hashtable]$UsedNames
    )

    $safe = [string]$RequestedName
    $safe = $safe -replace '[:\\/\?\*\[\]]', '・'
    $safe = ($safe.ToCharArray() | ForEach-Object { if ([int]$_ -lt 32) { '・' } else { $_ } }) -join ''
    $safe = $safe.Trim().Trim([char]39)
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'シート' }
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31) }

    $baseName = $safe
    $candidate = $baseName
    $number = 2
    while ($UsedNames.ContainsKey($candidate)) {
        $suffix = " ($number)"
        $maxBaseLength = 31 - $suffix.Length
        $shortBase = if ($baseName.Length -gt $maxBaseLength) { $baseName.Substring(0, $maxBaseLength) } else { $baseName }
        $candidate = $shortBase + $suffix
        $number++
    }
    $UsedNames[$candidate] = $true
    return $candidate
}

function Test-MbExcelWorksheetName {
    param([AllowEmptyString()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name.Length -gt 31) { return $false }
    if ($Name -match '[:\\/\?\*\[\]]') { return $false }
    if ($Name.StartsWith("'") -or $Name.EndsWith("'")) { return $false }
    foreach ($character in $Name.ToCharArray()) {
        if ([int]$character -lt 32) { return $false }
    }
    return $true
}

function ConvertTo-MbExcelSheetAddress {
    param([Parameter(Mandatory = $true)][string]$Name)
    return $Name.Replace("'", "''")
}

function Write-MbExcelStatus {
    param(
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][object]$Status
    )

    $Status.updatedAt = [DateTime]::UtcNow.ToString('o')
    $directory = Split-Path -Parent $StatusPath
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    $temporary = Join-Path $directory ('.status-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backup = Join-Path $directory ('.status-backup-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $json = $Status | ConvertTo-Json -Depth 8
    try {
        [IO.File]::WriteAllText($temporary, $json, (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $StatusPath) {
            [IO.File]::Replace($temporary, $StatusPath, $backup, $true)
        } else {
            [IO.File]::Move($temporary, $StatusPath)
        }
    } finally {
        foreach ($path in @($temporary, $backup)) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Set-MbExcelStatusProgress {
    param(
        [Parameter(Mandatory = $true)][object]$Status,
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$Phase,
        [Parameter(Mandatory = $true)][string]$Message,
        [int]$CurrentStep,
        [int]$TotalSteps,
        [int]$Percent
    )

    $Status.phase = $Phase
    $Status.message = $Message
    $Status.currentStep = $CurrentStep
    $Status.totalSteps = $TotalSteps
    $Status.percent = [Math]::Max(0, [Math]::Min(100, $Percent))
    Write-MbExcelStatus -StatusPath $StatusPath -Status $Status
}

function Test-MbExcelCancellation {
    param([Parameter(Mandatory = $true)][string]$CancelPath)
    if (Test-Path -LiteralPath $CancelPath -PathType Leaf) { throw 'MB_EXPORT_CANCELLED' }
}

function New-MbRoundedRectanglePath {
    param(
        [single]$Left,
        [single]$Top,
        [single]$Width,
        [single]$Height,
        [single]$Radius
    )

    $path = New-Object Drawing.Drawing2D.GraphicsPath
    $safeRadius = [single][Math]::Max(0, [Math]::Min($Radius, [Math]::Min($Width, $Height) / 2.0))
    if ($safeRadius -le 0.01) {
        [void]$path.AddRectangle((New-Object Drawing.RectangleF -ArgumentList @($Left, $Top, $Width, $Height)))
        return $path
    }

    $diameter = [single]($safeRadius * 2.0)
    [void]$path.AddArc($Left, $Top, $diameter, $diameter, 180, 90)
    [void]$path.AddArc($Left + $Width - $diameter, $Top, $diameter, $diameter, 270, 90)
    [void]$path.AddArc($Left + $Width - $diameter, $Top + $Height - $diameter, $diameter, $diameter, 0, 90)
    [void]$path.AddArc($Left, $Top + $Height - $diameter, $diameter, $diameter, 90, 90)
    [void]$path.CloseFigure()
    return $path
}

function New-MbAnnotatedImage {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Annotations,
        [AllowNull()][object]$Crop,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [ValidateRange(100, 4000)][int]$TargetDisplayWidth = 760,
        [ValidateRange(100, 4000)][int]$TargetDisplayHeight = 880,
        [ValidateRange(1.0, 2.0)][double]$MaximumDisplayScale = 2.0,
        # 番号注釈は編集画面のSVGと同じ基準フォントで描く。呼び出し元の解決済みフォントを受け取る。
        [AllowEmptyString()][string]$NumberFontName = ''
    )
    if ([string]::IsNullOrWhiteSpace($NumberFontName)) { $NumberFontName = Resolve-MbExcelBodyFont }

    $cropX = if ($null -ne $Crop -and $Crop.PSObject.Properties.Name -contains 'x') { [double]$Crop.x } else { 0.0 }
    $cropY = if ($null -ne $Crop -and $Crop.PSObject.Properties.Name -contains 'y') { [double]$Crop.y } else { 0.0 }
    $cropWidth = if ($null -ne $Crop -and $Crop.PSObject.Properties.Name -contains 'width') { [double]$Crop.width } else { 1.0 }
    $cropHeight = if ($null -ne $Crop -and $Crop.PSObject.Properties.Name -contains 'height') { [double]$Crop.height } else { 1.0 }
    $requiresCrop = $cropX -gt 0.000001 -or $cropY -gt 0.000001 -or $cropWidth -lt 0.999999 -or $cropHeight -lt 0.999999
    if (@($Annotations).Count -eq 0 -and -not $requiresCrop) { return $SourcePath }
    $source = $null
    $bitmap = $null
    $graphics = $null
    $rectPen = $null
    $arrowPen = $null
    $redBrush = $null
    $blackBrush = $null
    $whiteBrush = $null
    $numberFont = $null
    $numberFormat = $null
    try {
        $source = [Drawing.Image]::FromFile($SourcePath)

        # 先に切り抜き領域を確定し、注釈は切り抜き後の画像へ描画する。
        # 最終表示倍率から線幅と番号径を逆算することで、全画面・縦長・細長い切り抜きでも
        # Office上の注釈サイズをほぼ一定に保つ。
        $cropLeft = if ($requiresCrop) { [int][Math]::Floor($cropX * $source.Width) } else { 0 }
        $cropTop = if ($requiresCrop) { [int][Math]::Floor($cropY * $source.Height) } else { 0 }
        $cropLeft = [Math]::Max(0, [Math]::Min($cropLeft, $source.Width - 1))
        $cropTop = [Math]::Max(0, [Math]::Min($cropTop, $source.Height - 1))
        $cropPixelWidth = if ($requiresCrop) { [int][Math]::Max(1, [Math]::Round($cropWidth * $source.Width)) } else { [int]$source.Width }
        $cropPixelHeight = if ($requiresCrop) { [int][Math]::Max(1, [Math]::Round($cropHeight * $source.Height)) } else { [int]$source.Height }
        $cropPixelWidth = [Math]::Min($cropPixelWidth, $source.Width - $cropLeft)
        $cropPixelHeight = [Math]::Min($cropPixelHeight, $source.Height - $cropTop)

        $bitmap = New-Object Drawing.Bitmap -ArgumentList @($cropPixelWidth, $cropPixelHeight, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        try { $bitmap.SetResolution(96, 96) } catch { }
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $destinationRectangle = New-Object Drawing.Rectangle -ArgumentList @(0, 0, $cropPixelWidth, $cropPixelHeight)
        $sourceRectangle = New-Object Drawing.Rectangle -ArgumentList @($cropLeft, $cropTop, $cropPixelWidth, $cropPixelHeight)
        $graphics.DrawImage($source, $destinationRectangle, $sourceRectangle, [Drawing.GraphicsUnit]::Pixel)

        $displayScale = [Math]::Min($MaximumDisplayScale, [Math]::Min($TargetDisplayWidth / [double]$cropPixelWidth, $TargetDisplayHeight / [double]$cropPixelHeight))
        $displayScale = [Math]::Max(0.01, $displayScale)
        # Web編集画面と同じ共通トークンを使う。画像の縦横比に左右されず、
        # 通常画面・横長・縦長の間で赤枠・矢印・番号の視覚ウェイトを揃える。
        $annotationUnit = 0.62
        $rectLineWidth = [single][Math]::Max(1.5, (7.0 * $annotationUnit) / $displayScale)
        $arrowLineWidth = [single][Math]::Max(1.5, (8.0 * $annotationUnit) / $displayScale)
        $cornerRadius = [single][Math]::Max(2.0, (4.0 * $annotationUnit) / $displayScale)
        $badgeRadius = [single][Math]::Max(7.0, (24.0 * $annotationUnit) / $displayScale)
        $red = [Drawing.Color]::FromArgb(217, 45, 32)
        $rectPen = New-Object Drawing.Pen $red, $rectLineWidth
        $rectPen.LineJoin = [Drawing.Drawing2D.LineJoin]::Round
        $arrowPen = New-Object Drawing.Pen $red, $arrowLineWidth
        $arrowPen.LineJoin = [Drawing.Drawing2D.LineJoin]::Round
        $arrowPen.StartCap = [Drawing.Drawing2D.LineCap]::Round
        $arrowPen.EndCap = [Drawing.Drawing2D.LineCap]::Round
        $redBrush = New-Object Drawing.SolidBrush $red
        $blackBrush = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(17, 24, 39))
        $whiteBrush = New-Object Drawing.SolidBrush ([Drawing.Color]::White)
        # 基準フォント(BIZ UDPゴシック)はArialより数字が広いため、2桁でも円内に余白が残る24を使う。
        # 編集画面のSVG(app.js)と同じ値にすること。
        $fontSize = [single][Math]::Max(9, (24.0 * $annotationUnit) / $displayScale)
        $numberOffsetY = [single]((1.5 * $annotationUnit) / $displayScale)
        $numberFont = New-Object Drawing.Font $NumberFontName, $fontSize, ([Drawing.FontStyle]::Bold), ([Drawing.GraphicsUnit]::Pixel)
        $numberFormat = New-Object Drawing.StringFormat
        $numberFormat.Alignment = [Drawing.StringAlignment]::Center
        $numberFormat.LineAlignment = [Drawing.StringAlignment]::Center

        foreach ($annotation in @($Annotations)) {
            $x1 = [single](([double]$annotation.x1 * $source.Width) - $cropLeft)
            $y1 = [single](([double]$annotation.y1 * $source.Height) - $cropTop)
            $x2 = [single](([double]$annotation.x2 * $source.Width) - $cropLeft)
            $y2 = [single](([double]$annotation.y2 * $source.Height) - $cropTop)
            $left = [single][Math]::Min($x1, $x2)
            $top = [single][Math]::Min($y1, $y2)
            $width = [single][Math]::Max(1, [Math]::Abs($x2 - $x1))
            $height = [single][Math]::Max(1, [Math]::Abs($y2 - $y1))
            switch ([string]$annotation.type) {
                'rect' {
                    $rectPath = $null
                    try {
                        $rectPath = New-MbRoundedRectanglePath -Left $left -Top $top -Width $width -Height $height -Radius $cornerRadius
                        $graphics.DrawPath($rectPen, $rectPath)
                    } finally {
                        if ($rectPath) { $rectPath.Dispose() }
                    }
                }
                'blackout' {
                    $graphics.FillRectangle($blackBrush, $left, $top, $width, $height)
                }
                'number' {
                    $graphics.FillEllipse($redBrush, $x1 - $badgeRadius, $y1 - $badgeRadius, $badgeRadius * 2, $badgeRadius * 2)
                    $textBounds = New-Object Drawing.RectangleF -ArgumentList @(
                        [single]($x1 - $badgeRadius), [single]($y1 - $badgeRadius + $numberOffsetY),
                        [single]($badgeRadius * 2), [single]($badgeRadius * 2)
                    )
                    $graphics.DrawString([string][int]$annotation.label, $numberFont, $whiteBrush, $textBounds, $numberFormat)
                }
                'arrow' {
                    $graphics.DrawLine($arrowPen, $x1, $y1, $x2, $y2)
                    $angle = [Math]::Atan2($y2 - $y1, $x2 - $x1)
                    $headLength = (24.0 * $annotationUnit) / $displayScale
                    $headWidth = (12.0 * $annotationUnit) / $displayScale
                    $baseX = $x2 - ([Math]::Cos($angle) * $headLength)
                    $baseY = $y2 - ([Math]::Sin($angle) * $headLength)
                    $sideX = [Math]::Sin($angle) * $headWidth
                    $sideY = -[Math]::Cos($angle) * $headWidth
                    $points = New-Object 'Drawing.PointF[]' 3
                    $points[0] = New-Object Drawing.PointF -ArgumentList @([single]$x2, [single]$y2)
                    $points[1] = New-Object Drawing.PointF -ArgumentList @([single]($baseX + $sideX), [single]($baseY + $sideY))
                    $points[2] = New-Object Drawing.PointF -ArgumentList @([single]($baseX - $sideX), [single]($baseY - $sideY))
                    $graphics.FillPolygon($redBrush, $points)
                }
            }
        }

        $directory = Split-Path -Parent $DestinationPath
        if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
        $bitmap.Save($DestinationPath, [Drawing.Imaging.ImageFormat]::Png)
        return $DestinationPath
    } finally {
        foreach ($item in @($numberFormat, $numberFont, $whiteBrush, $blackBrush, $redBrush, $arrowPen, $rectPen, $graphics, $bitmap, $source)) {
            if ($item) { try { $item.Dispose() } catch { } }
        }
    }
}

function Set-MbExcelBorders {
    param([Parameter(Mandatory = $true)][object]$Range, [int]$Color, [int]$LineStyle = 1, [int]$Weight = 2)
    $borders = $null
    try {
        $borders = $Range.Borders
        $borders.LineStyle = $LineStyle
        $borders.Weight = $Weight
        $borders.Color = $Color
    } finally {
        Release-MbExcelComObject $borders
    }
}

function Set-MbExcelEdgeBorder {
    param(
        [Parameter(Mandatory = $true)][object]$Range,
        [Parameter(Mandatory = $true)][int[]]$Edges,
        [int]$Color,
        [int]$LineStyle = 1,
        [int]$Weight = 2
    )
    foreach ($edge in $Edges) {
        $border = $null
        try {
            $border = $Range.Borders.Item([int]$edge)
            $border.LineStyle = $LineStyle
            $border.Weight = $Weight
            $border.Color = $Color
        } finally { Release-MbExcelComObject $border }
    }
}

function Set-MbExcelWorksheetView {
    param([Parameter(Mandatory = $true)][object]$Application, [Parameter(Mandatory = $true)][object]$Worksheet, [int]$FreezeRows = 2)
    $window = $null
    try {
        $Worksheet.Activate()
        $window = $Application.ActiveWindow
        $window.DisplayGridlines = $false
        $window.Zoom = 100
        $window.FreezePanes = $false
        $window.SplitColumn = 0
        $window.SplitRow = $FreezeRows
        $window.FreezePanes = $true
    } finally {
        Release-MbExcelComObject $window
    }
}

function Get-MbExcelTextLineEstimate {
    param(
        [AllowEmptyString()][string]$Text,
        [ValidateRange(10, 200)][int]$CharactersPerLine = 68
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return 1 }
    $normalized = ([string]$Text).Replace("`r`n", "`n").Replace("`r", "`n")
    $count = 0
    foreach ($line in @($normalized -split "`n")) {
        # Excel上では英数字・半角記号は日本語全角文字より狭い。
        # UTF-16の文字数をそのまま使うと英数字を含む長文ほど行数を過大評価するため、
        # 半角文字を0.55文字相当として表示幅を近似する。
        $visualLength = 0.0
        foreach ($character in ([string]$line).ToCharArray()) {
            $visualLength += if ([int][char]$character -le 255) { 0.55 } else { 1.0 }
        }
        $count += [Math]::Max(1, [int][Math]::Ceiling($visualLength / [double]$CharactersPerLine))
    }
    return [Math]::Max(1, $count)
}

function Get-MbExcelStepCardLayout {
    param(
        [AllowEmptyString()][string]$Description = '',
        [AllowEmptyString()][string]$Note = '',
        [ValidateRange(0, 100000)][int]$ImageWidth = 0,
        [ValidateRange(0, 100000)][int]$ImageHeight = 0,
        # 画像のない手順はカード全幅を文章に使うため、画像領域の最低行数を確保しない。
        [bool]$HasImage = $true
    )

    $descriptionLines = Get-MbExcelTextLineEstimate -Text $Description -CharactersPerLine 68
    # H:Lの実幅と12pt本文に合わせ、表示行高16pt＋上下余裕4ptを26pt行へ換算する。
    # Excel実画面で残っていた1〜2行相当の余白を詰めつつ、折返し分の安全余裕は残す。
    $descriptionBodyRows = [int][Math]::Ceiling((($descriptionLines * 16.0) + 4.0) / 26.0)
    $descriptionBodyRows = [Math]::Max(3, [Math]::Min(60, $descriptionBodyRows))
    $hasNote = -not [string]::IsNullOrWhiteSpace($Note)
    $noteBodyRows = 0
    if ($hasNote) {
        $noteLines = Get-MbExcelTextLineEstimate -Text $Note -CharactersPerLine 72
        $noteBodyRows = [int][Math]::Ceiling((($noteLines * 15.0) + 4.0) / 26.0)
        $noteBodyRows = [Math]::Max(2, [Math]::Min(30, $noteBodyRows))
    }

    # 画像列の実幅（約570pt）から必要な高さを逆算する。
    # 先頭の本文行は20pt、以降は26pt、画像の上下余白は合計18ptとすることで、
    # 16:9は14行、16:10は15行に収まり、画像の下へ不要な空白を残さない。
    $imageRows = if ($HasImage) { 6 } else { 0 }
    if ($HasImage -and $ImageWidth -gt 0 -and $ImageHeight -gt 0) {
        $imageAspectRatio = $ImageWidth / [double]$ImageHeight
        $requiredImageHeightPoints = (570.0 * ($ImageHeight / [double]$ImageWidth)) + 18.0
        $imageRows = 1 + [int][Math]::Ceiling(([Math]::Max(0.0, $requiredImageHeightPoints - 20.0) / 26.0) - 0.000001)
        # 高さが幅の3倍以上の画像は、カード全体が縦に伸びすぎないよう22行で止める。
        # 一般的な縦長（16:9を回転した程度）は従来どおり26行まで使用する。
        $maximumImageRows = if ($imageAspectRatio -le (1.0 / 3.0)) { 22 } else { 26 }
        $imageRows = [Math]::Max(6, [Math]::Min($maximumImageRows, $imageRows))
    }

    $textRows = 1 + $descriptionBodyRows
    if ($hasNote) { $textRows += 1 + $noteBodyRows }
    $contentRows = [Math]::Max($imageRows, $textRows)
    return [pscustomobject]@{
        ContentRows = [int]$contentRows
        ImageRows = [int]$imageRows
        DescriptionBodyRows = [int]$descriptionBodyRows
        NoteBodyRows = [int]$noteBodyRows
        HasNote = [bool]$hasNote
        NextRowOffset = [int]($contentRows + 2)
    }
}

function Add-MbExcelStepCard {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [int]$StartRow,
        [int]$StepNumber,
        [AllowEmptyString()][string]$Title,
        [AllowEmptyString()][string]$Description,
        [AllowEmptyString()][string]$Note,
        [AllowEmptyString()][string]$ImagePath,
        [Parameter(Mandatory = $true)][string]$FontName,
        # 動画つきの手順だけ、見出しの右へ「▶ 動画を見る」の相対リンクを置く。
        [AllowEmptyString()][string]$VideoLinkPath = ''
    )

    $xlCenter = -4108
    $xlLeft = -4131
    $xlRight = -4152
    $xlTop = -4160
    $xlMoveAndSize = 1
    $msoTrue = -1
    $msoFalse = 0
    $colorAccent = ConvertTo-MbExcelBgr 58 91 160
    $colorAccentDark = ConvertTo-MbExcelBgr 38 57 104
    $colorLine = ConvertTo-MbExcelBgr 216 222 232
    $colorText = ConvertTo-MbExcelBgr 24 32 51
    $colorMuted = ConvertTo-MbExcelBgr 102 112 133
    $colorImage = ConvertTo-MbExcelBgr 246 248 251
    $colorNote = ConvertTo-MbExcelBgr 255 249 226
    $colorNoteLine = ConvertTo-MbExcelBgr 234 217 168
    $colorWhite = ConvertTo-MbExcelBgr 255 255 255

    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = '手順名未入力' }
    $hasNote = -not [string]::IsNullOrWhiteSpace($Note)
    $imageWidth = 0
    $imageHeight = 0
    $hasImage = [bool]($ImagePath -and (Test-Path -LiteralPath $ImagePath -PathType Leaf))
    if ($hasImage) {
        $layoutImage = $null
        try {
            $layoutImage = [Drawing.Image]::FromFile($ImagePath)
            $imageWidth = [int]$layoutImage.Width
            $imageHeight = [int]$layoutImage.Height
        } finally {
            if ($layoutImage) { $layoutImage.Dispose() }
        }
    }
    # 画像のない手順は左半分を空けず、説明と補足をカード全幅で読ませる。
    # 編集画面が「画像なし手順の空白を縮小する」のと同じ考え方に揃える。
    $textColumn = if ($hasImage) { 'H' } else { 'A' }
    $layout = Get-MbExcelStepCardLayout -Description $Description -Note $Note -ImageWidth $imageWidth -ImageHeight $imageHeight -HasImage $hasImage
    $headerRow = $StartRow
    $contentStart = $StartRow + 1
    $contentEnd = $contentStart + [int]$layout.ContentRows - 1
    $descriptionLabelRow = $contentStart
    $descriptionStart = $contentStart + 1
    $descriptionEnd = [Math]::Min($contentEnd, $descriptionStart + [int]$layout.DescriptionBodyRows - 1)
    $noteLabelRow = if ($hasNote) { $descriptionEnd + 1 } else { 0 }
    $noteStart = if ($hasNote) { $noteLabelRow + 1 } else { 0 }
    $noteEnd = if ($hasNote) { [Math]::Min($contentEnd, $noteStart + [int]$layout.NoteBodyRows - 1) } else { 0 }
    $spacerRow = $contentEnd + 1

    for ($row = $headerRow; $row -le $spacerRow; $row++) {
        $rowRange = $null
        try {
            $rowRange = $Worksheet.Range("A${row}:L${row}")
            $rowRange.RowHeight = if ($row -eq $headerRow) {
                30
            } elseif ($row -eq $spacerRow) {
                16
            } elseif ($row -eq $descriptionLabelRow -or ($hasNote -and $row -eq $noteLabelRow)) {
                20
            } else {
                26
            }
        } finally { Release-MbExcelComObject $rowRange }
    }

    $hasVideoLink = -not [string]::IsNullOrWhiteSpace($VideoLinkPath)
    $headerBand = $null
    $numberCell = $null
    $titleArea = $null
    $videoCell = $null
    $imageArea = $null
    $descriptionLabel = $null
    $descriptionArea = $null
    $noteLabel = $null
    $noteArea = $null
    $noteBlock = $null
    $textPanel = $null
    $card = $null
    $shape = $null
    try {
        $headerBand = $Worksheet.Range("A${headerRow}:L${headerRow}")
        $numberCell = $Worksheet.Range("A${headerRow}:A${headerRow}")
        $titleArea = if ($hasVideoLink) { $Worksheet.Range("B${headerRow}:I${headerRow}") } else { $Worksheet.Range("B${headerRow}:L${headerRow}") }
        if ($hasVideoLink) { $videoCell = $Worksheet.Range("J${headerRow}:L${headerRow}") }
        if ($hasImage) { $imageArea = $Worksheet.Range("A${contentStart}:G${contentEnd}") }
        $descriptionLabel = $Worksheet.Range("${textColumn}${descriptionLabelRow}:L${descriptionLabelRow}")
        $descriptionArea = $Worksheet.Range("${textColumn}${descriptionStart}:L${descriptionEnd}")
        $textPanel = $Worksheet.Range("${textColumn}${contentStart}:L${contentEnd}")
        $card = $Worksheet.Range("A${headerRow}:L${contentEnd}")

        $headerBand.Interior.Color = $colorWhite

        $numberCell.Value2 = [double]$StepNumber
        $numberCell.NumberFormat = '"STEP "00'
        $numberCell.Interior.Color = $colorAccent
        $numberCell.Font.Name = $FontName
        $numberCell.Font.Size = 10
        $numberCell.Font.Bold = $true
        $numberCell.Font.Color = $colorWhite
        $numberCell.HorizontalAlignment = $xlCenter
        $numberCell.VerticalAlignment = $xlCenter

        $titleArea.Merge()
        $titleArea.NumberFormat = '@'
        $titleArea.Value2 = [string]$Title
        $titleArea.Font.Name = $FontName
        $titleArea.Font.Size = 13
        $titleArea.Font.Bold = $true
        $titleArea.Font.Color = $colorText
        $titleArea.HorizontalAlignment = $xlLeft
        $titleArea.VerticalAlignment = $xlCenter

        if ($hasVideoLink) {
            $videoCell.Merge()
            # ここは数式のまま保存する。表示形式を文字列にすると数式が文字として残り、リンクにならない。
            # =HYPERLINK() の相対パスはクリック時にブックの場所を基準に解決されるため、
            # フォルダーごと共有フォルダーへコピーしても、そのまま再生できる。
            $videoCell.Formula = '=HYPERLINK("' + ($VideoLinkPath -replace '"', '""') + '","▶ 動画を見る")'
            $videoCell.Font.Name = $FontName
            $videoCell.Font.Size = 11
            $videoCell.Font.Bold = $true
            $videoCell.Font.Color = $colorAccent
            $videoCell.Font.Underline = 2
            $videoCell.Interior.Color = $colorWhite
            $videoCell.HorizontalAlignment = $xlRight
            $videoCell.VerticalAlignment = $xlCenter
            $videoCell.IndentLevel = 1
        }

        if ($hasImage) {
            $imageArea.Merge()
            $imageArea.Interior.Color = $colorImage
        }

        # 文章側は下端まで白の一枚面にする。必要な行だけ白くすると、短い説明のときに
        # 白い帯がグレーの中へ浮いて見え、カードが未完成の印象になるため。
        $textPanel.Interior.Color = $colorWhite

        $descriptionLabel.Merge()
        $descriptionLabel.NumberFormat = '@'
        $descriptionLabel.Value2 = '説明'
        $descriptionLabel.Font.Name = $FontName
        $descriptionLabel.Font.Size = 10
        $descriptionLabel.Font.Bold = $true
        $descriptionLabel.Font.Color = $colorAccentDark
        $descriptionLabel.Interior.Color = $colorWhite
        $descriptionLabel.HorizontalAlignment = $xlLeft
        $descriptionLabel.VerticalAlignment = $xlCenter
        # 文字が区切り線へ張り付かないよう、文章側は一段字下げする。
        $descriptionLabel.IndentLevel = 1

        $descriptionArea.Merge()
        $descriptionArea.NumberFormat = '@'
        $descriptionArea.Value2 = if ([string]::IsNullOrWhiteSpace($Description)) { '（説明未入力）' } else { [string]$Description }
        $descriptionArea.WrapText = $true
        $descriptionArea.HorizontalAlignment = $xlLeft
        $descriptionArea.VerticalAlignment = $xlTop
        $descriptionArea.Font.Name = $FontName
        $descriptionArea.Font.Size = 12
        $descriptionArea.Font.Color = if ([string]::IsNullOrWhiteSpace($Description)) { $colorMuted } else { $colorText }
        $descriptionArea.Interior.Color = $colorWhite
        $descriptionArea.IndentLevel = 1

        if ($hasNote) {
            $noteLabel = $Worksheet.Range("${textColumn}${noteLabelRow}:L${noteLabelRow}")
            $noteArea = $Worksheet.Range("${textColumn}${noteStart}:L${noteEnd}")
            $noteBlock = $Worksheet.Range("${textColumn}${noteLabelRow}:L${noteEnd}")
            $noteLabel.Merge()
            $noteLabel.NumberFormat = '@'
            $noteLabel.Value2 = '補足'
            $noteLabel.Interior.Color = $colorNote
            $noteLabel.Font.Name = $FontName
            $noteLabel.Font.Size = 10
            $noteLabel.Font.Bold = $true
            $noteLabel.Font.Color = ConvertTo-MbExcelBgr 122 83 0
            $noteLabel.HorizontalAlignment = $xlLeft
            $noteLabel.VerticalAlignment = $xlCenter
            $noteLabel.IndentLevel = 1

            $noteArea.Merge()
            $noteArea.NumberFormat = '@'
            $noteArea.Value2 = [string]$Note
            $noteArea.WrapText = $true
            $noteArea.HorizontalAlignment = $xlLeft
            $noteArea.VerticalAlignment = $xlTop
            $noteArea.Interior.Color = $colorNote
            $noteArea.Font.Name = $FontName
            $noteArea.Font.Size = 11
            $noteArea.Font.Color = $colorText
            $noteArea.IndentLevel = 1
            Set-MbExcelEdgeBorder -Range $noteBlock -Edges @(7) -Color $colorNoteLine -Weight 2
        }

        Set-MbExcelEdgeBorder -Range $card -Edges @(7, 8, 9, 10) -Color $colorLine
        Set-MbExcelEdgeBorder -Range $headerBand -Edges @(9) -Color $colorLine -Weight 2
        # 区切り線は画像と文章が並ぶときだけ引く。全幅の文章カードには不要。
        if ($hasImage) { Set-MbExcelEdgeBorder -Range $textPanel -Edges @(7) -Color $colorLine -Weight -4138 }

        if ($hasImage) {
            $shape = $Worksheet.Shapes.AddPicture(
                $ImagePath, $msoFalse, $msoTrue,
                [single]($imageArea.Left + 9), [single]($imageArea.Top + 9), [single]-1, [single]-1
            )
            $shape.LockAspectRatio = $msoTrue
            $originalWidth = [double]$shape.Width
            $originalHeight = [double]$shape.Height
            $maxWidth = [double]$imageArea.Width - 18
            $maxHeight = [double]$imageArea.Height - 18
            # 小さな元画像は原寸感と鮮明さを保つため、Excelでは最大1.5倍に抑える。
            # 3:1以上の細長い画像は、長辺方向いっぱいに広がらないよう通常寸法の85%へ制限する。
            $maximumImageScale = 1.5
            $compactImageWidthRatio = if (($originalWidth / $originalHeight) -ge 3.0) { 0.85 } else { 1.0 }
            $compactImageHeightRatio = if (($originalHeight / $originalWidth) -ge 3.0) { 0.85 } else { 1.0 }
            $effectiveMaxWidth = $maxWidth * $compactImageWidthRatio
            $effectiveMaxHeight = $maxHeight * $compactImageHeightRatio
            $scale = [Math]::Min($maximumImageScale, [Math]::Min($effectiveMaxWidth / $originalWidth, $effectiveMaxHeight / $originalHeight))
            $shape.Width = [single]($originalWidth * $scale)
            $shape.Left = [single]($imageArea.Left + (($imageArea.Width - $shape.Width) / 2))
            $shape.Top = [single]($imageArea.Top + (($imageArea.Height - $shape.Height) / 2))
            $shape.Placement = $xlMoveAndSize
            try { $shape.AlternativeText = "手順 $StepNumber の画面: $Title" } catch { }
        }
    } finally {
        Release-MbExcelComObject $shape
        Release-MbExcelComObject $card
        Release-MbExcelComObject $textPanel
        Release-MbExcelComObject $noteBlock
        Release-MbExcelComObject $noteArea
        Release-MbExcelComObject $noteLabel
        Release-MbExcelComObject $descriptionArea
        Release-MbExcelComObject $descriptionLabel
        Release-MbExcelComObject $imageArea
        Release-MbExcelComObject $videoCell
        Release-MbExcelComObject $titleArea
        Release-MbExcelComObject $numberCell
        Release-MbExcelComObject $headerBand
    }
    return [int]($spacerRow + 1)
}

function Invoke-MbExcelExport {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$CancelPath,
        [Parameter(Mandatory = $true)][string]$JobId
    )

    $startedAt = [DateTime]::UtcNow.ToString('o')
    $totalSteps = 0
    foreach ($sheet in @($Project.sheets)) { $totalSteps += @($sheet.steps).Count }
    $status = [pscustomobject]@{
        jobId = $JobId
        state = 'running'
        phase = 'starting'
        message = 'Excel出力を準備しています'
        percent = 0
        currentStep = 0
        totalSteps = $totalSteps
        outputPath = ''
        outputName = ''
        outputFolder = ''
        outputFolderName = ''
        videoCount = 0
        outputDirectory = $OutputDirectory
        sheetNameMappings = @()
        ownedExcelPid = 0
        ownedExcelStartTimeUtc = ''
        ownershipMode = ''
        ownershipProven = $false
        startedAt = $startedAt
        updatedAt = $startedAt
        completedAt = ''
        errorCode = ''
    }
    Write-MbExcelStatus -StatusPath $StatusPath -Status $status

    $excel = $null
    $workbook = $null
    $worksheets = $null
    $indexSheet = $null
    $workbookClosed = $false
    $canQuitCom = $false
    $settingsApplied = $false
    $ownershipProven = $false
    $ownershipMode = ''
    $ownPid = 0
    $temporaryPath = ''
    $outputPath = ''
    $stagingDirectory = ''
    $videoPlan = Get-MbExcelVideoPlan -Project $Project -ProjectPath $ProjectPath
    # 動画つきの手順があるときだけフォルダーで出す。無い場合は今までどおり単体のxlsxのまま。
    $usesFolderOutput = [int]$videoPlan.Count -gt 0
    $renderDirectory = Join-Path (Split-Path -Parent $StatusPath) 'rendered-images'
    # 注釈を合成した派生画像は出力専用のため、成功・失敗・中止のいずれでもfinallyで削除する。
    $generatedImages = New-Object System.Collections.ArrayList
    $pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $xlOpenXMLWorkbook = 51
    $xlCenter = -4108
    $xlLeft = -4131
    $xlContinuous = 1
    $xlThin = 2
    $colorAccent = ConvertTo-MbExcelBgr 58 91 160
    $colorAccentDark = ConvertTo-MbExcelBgr 38 57 104
    $colorAccentSoft = ConvertTo-MbExcelBgr 238 243 252
    $colorLine = ConvertTo-MbExcelBgr 216 222 232
    $colorText = ConvertTo-MbExcelBgr 24 32 51
    $colorMuted = ConvertTo-MbExcelBgr 102 112 133
    $colorWhite = ConvertTo-MbExcelBgr 255 255 255

    try {
        Test-MbExcelCancellation -CancelPath $CancelPath
        if (-not $script:MbExcelPInvokeAvailable) { throw 'MB_EXCEL_OWNERSHIP_API_UNAVAILABLE' }
        if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }
        if (-not (Test-Path -LiteralPath $renderDirectory)) { [void](New-Item -ItemType Directory -Path $renderDirectory -Force) }
        $bodyFont = Resolve-MbExcelBodyFont

        Set-MbExcelStatusProgress -Status $status -StatusPath $StatusPath -Phase 'starting-excel' `
            -Message '安全確認のためExcelを起動しています' -CurrentStep 0 -TotalSteps $totalSteps -Percent 3
        $excel = New-Object -ComObject Excel.Application
        $canQuitCom = ($pidsBefore.Count -eq 0)
        $resolved = Resolve-MbOwnedExcelProcess -Application $excel -PidsBefore $pidsBefore
        $ownPid = [int]$resolved.Pid
        $ownershipMode = [string]$resolved.Mode
        if ([bool]$resolved.ExistingConnection) { throw 'MB_CONNECTED_TO_EXISTING_EXCEL' }
        if ($ownPid -le 0) {
            if ($pidsBefore.Count -gt 0) { throw 'MB_EXCEL_OWNERSHIP_UNRESOLVED_WITH_EXISTING' }
            throw 'MB_EXCEL_OWNERSHIP_UNRESOLVED'
        }

        $ownershipProven = $true
        $canQuitCom = $true
        $status.ownedExcelPid = $ownPid
        $status.ownershipMode = $ownershipMode
        $status.ownershipProven = $true
        try { $status.ownedExcelStartTimeUtc = (Get-Process -Id $ownPid -ErrorAction Stop).StartTime.ToUniversalTime().ToString('o') } catch { }
        Write-MbExcelStatus -StatusPath $StatusPath -Status $status

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        try { $excel.AskToUpdateLinks = $false } catch { }
        $settingsApplied = $true

        $workbooks = $null
        try {
            $workbooks = $excel.Workbooks
            $workbook = $workbooks.Add()
        } finally { Release-MbExcelComObject $workbooks }
        $worksheets = $workbook.Worksheets
        while ($worksheets.Count -gt 1) {
            $deleteSheet = $worksheets.Item($worksheets.Count)
            try { $deleteSheet.Delete() } finally { Release-MbExcelComObject $deleteSheet }
        }
        $indexSheet = $worksheets.Item(1)
        $indexSheetName = [string]$indexSheet.Name
        try { $indexSheet.Name = '目次'; $indexSheetName = '目次' } catch { $indexSheet.Name = 'INDEX'; $indexSheetName = 'INDEX' }

        $usedNames = @{}
        $usedNames[$indexSheetName] = $true
        $nameMappings = @()
        foreach ($sheet in @($Project.sheets)) {
            $safeName = Get-MbSafeExcelWorksheetName -RequestedName ([string]$sheet.name) -UsedNames $usedNames
            $nameMappings += [pscustomobject]@{ Requested = [string]$sheet.name; Safe = $safeName }
        }
        $status.sheetNameMappings = @($nameMappings | ForEach-Object {
            [pscustomobject]@{ requested = [string]$_.Requested; output = [string]$_.Safe }
        })
        Write-MbExcelStatus -StatusPath $StatusPath -Status $status

        Set-MbExcelStatusProgress -Status $status -StatusPath $StatusPath -Phase 'building-index' `
            -Message '目次を作成しています' -CurrentStep 0 -TotalSteps $totalSteps -Percent 8
        $indexSheet.Columns.Item(1).ColumnWidth = 9
        $indexSheet.Columns.Item(2).ColumnWidth = 32
        $indexSheet.Columns.Item(3).ColumnWidth = 58
        $indexSheet.Columns.Item(4).ColumnWidth = 14
        $titleRange = $null
        $metaRange = $null
        $headerRange = $null
        try {
            $titleRange = $indexSheet.Range('A1:D1')
            $metaRange = $indexSheet.Range('A2:D2')
            $headerRange = $indexSheet.Range('A4:D4')
            $titleRange.Merge()
            $titleRange.NumberFormat = '@'
            $titleRange.Value2 = [string]$Project.title
            $titleRange.Font.Name = $bodyFont
            $titleRange.Font.Size = 21
            $titleRange.Font.Bold = $true
            $titleRange.Font.Color = $colorWhite
            $titleRange.Interior.Color = $colorAccentDark
            $titleRange.RowHeight = 46
            $titleRange.HorizontalAlignment = $xlLeft
            $titleRange.VerticalAlignment = $xlCenter

            $metaRange.Merge()
            $metaRange.NumberFormat = '@'
            $metaRange.Value2 = [string]("更新 " + (Get-Date -Format 'yyyy年M月d日 HH:mm') + "　｜　$(@($Project.sheets).Count) シート　｜　$totalSteps 手順")
            $metaRange.Font.Name = $bodyFont
            $metaRange.Font.Size = 10
            $metaRange.Font.Color = $colorMuted
            $metaRange.Interior.Color = $colorWhite
            $metaRange.RowHeight = 24
            Set-MbExcelEdgeBorder -Range $metaRange -Edges @(9) -Color $colorLine -Weight 2

            $indexSheet.Cells.Item(4, 1).Value2 = 'No.'
            $indexSheet.Cells.Item(4, 2).Value2 = 'シート'
            $indexSheet.Cells.Item(4, 3).Value2 = '最初の手順'
            $indexSheet.Cells.Item(4, 4).Value2 = '手順数'
            $headerRange.Interior.Color = $colorAccentSoft
            $headerRange.Font.Name = $bodyFont
            $headerRange.Font.Bold = $true
            $headerRange.Font.Color = $colorAccentDark
            $headerRange.RowHeight = 28
            Set-MbExcelEdgeBorder -Range $headerRange -Edges @(9) -Color $colorAccent -LineStyle $xlContinuous -Weight 2
            try { $indexSheet.Tab.Color = $colorAccentDark } catch { }
        } finally {
            Release-MbExcelComObject $headerRange
            Release-MbExcelComObject $metaRange
            Release-MbExcelComObject $titleRange
        }

        $globalStep = 0
        $expectedShapes = 0
        $createdSheetNames = @()
        $escapedIndexName = ConvertTo-MbExcelSheetAddress -Name $indexSheetName
        for ($sheetIndex = 0; $sheetIndex -lt @($Project.sheets).Count; $sheetIndex++) {
            Test-MbExcelCancellation -CancelPath $CancelPath
            $sheetModel = @($Project.sheets)[$sheetIndex]
            $safeName = [string]$nameMappings[$sheetIndex].Safe
            $afterSheet = $worksheets.Item($worksheets.Count)
            $worksheet = $null
            try {
                $worksheet = $worksheets.Add([Type]::Missing, $afterSheet)
            } finally { Release-MbExcelComObject $afterSheet }
            try {
                try { $worksheet.Name = $safeName } catch {
                    [void]$usedNames.Remove($safeName)
                    $safeName = Get-MbSafeExcelWorksheetName -RequestedName ("手順シート {0}" -f ($sheetIndex + 1)) -UsedNames $usedNames
                    $worksheet.Name = $safeName
                    $nameMappings[$sheetIndex].Safe = $safeName
                }
                $createdSheetNames += $safeName
                try { $worksheet.Tab.Color = $colorAccent } catch { }
                # 画像はA:G、説明はH:Lで約半幅ずつ。1920×1080・100%表示で画像と説明を同時に読む前提。
                $worksheet.Columns.Item(1).ColumnWidth = 9
                foreach ($column in 2..7) { $worksheet.Columns.Item($column).ColumnWidth = 15 }
                foreach ($column in 8..12) { $worksheet.Columns.Item($column).ColumnWidth = 20 }

                $sheetHeader = $null
                $sheetTitle = $null
                $backLink = $null
                $summaryRange = $null
                $sheetHyperlinks = $null
                $newHyperlink = $null
                try {
                    $sheetHeader = $worksheet.Range('A1:L1')
                    $sheetTitle = $worksheet.Range('A1:J1')
                    $backLink = $worksheet.Range('K1:L1')
                    $sheetTitle.Merge()
                    $sheetTitle.NumberFormat = '@'
                    $sheetTitle.Value2 = [string]$safeName
                    $sheetTitle.Font.Name = $bodyFont
                    $sheetTitle.Font.Size = 16
                    $sheetTitle.Font.Bold = $true
                    $sheetTitle.Font.Color = $colorText
                    $sheetTitle.RowHeight = 40
                    $sheetTitle.IndentLevel = 1
                    $sheetTitle.HorizontalAlignment = $xlLeft
                    $sheetTitle.VerticalAlignment = $xlCenter

                    $sheetHeader.Interior.Color = $colorWhite

                    $backLink.Merge()
                    $backLink.Font.Name = $bodyFont
                    $backLink.Font.Size = 10.5
                    $backLink.Font.Color = $colorAccent
                    $backLink.Font.Bold = $false
                    $backLink.HorizontalAlignment = -4152
                    $backLink.VerticalAlignment = $xlCenter
                    try {
                        $sheetHyperlinks = $worksheet.Hyperlinks
                        $newHyperlink = $sheetHyperlinks.Add($backLink, [string]'', [string]("'$escapedIndexName'!A1"), [Type]::Missing, [string]'目次へ戻る')
                    } catch { $backLink.Value2 = '目次へ戻る' }
                    $backLink.Font.Color = $colorAccent
                    try { $backLink.Font.Underline = -4142 } catch { }

                    $summaryText = if ($sheetModel.PSObject.Properties.Name -contains 'summary') { [string]$sheetModel.summary } else { '' }
                    if (-not [string]::IsNullOrWhiteSpace($summaryText)) {
                        $summaryRange = $worksheet.Range('A2:L2')
                        $summaryRange.Merge()
                        $summaryRange.NumberFormat = '@'
                        $summaryRange.Value2 = $summaryText
                        $summaryRange.Font.Name = $bodyFont
                        $summaryRange.Font.Size = 10.5
                        # 淡い青地の上ではcolorMutedがAA(4.5:1)を下回るため、本文色で表示する。
                        $summaryRange.Font.Color = $colorText
                        $summaryRange.Interior.Color = $colorAccentSoft
                        $summaryRange.RowHeight = 28
                        $summaryRange.WrapText = $true
                        $summaryRange.VerticalAlignment = -4108
                        Set-MbExcelEdgeBorder -Range $summaryRange -Edges @(7) -Color $colorAccent -Weight 4
                    }
                    Set-MbExcelEdgeBorder -Range $sheetHeader -Edges @(9) -Color $colorAccentDark -Weight -4138
                    Set-MbExcelEdgeBorder -Range $sheetHeader -Edges @(7) -Color $colorAccent -Weight 4
                } finally {
                    Release-MbExcelComObject $newHyperlink
                    Release-MbExcelComObject $sheetHyperlinks
                    Release-MbExcelComObject $summaryRange
                    Release-MbExcelComObject $backLink
                    Release-MbExcelComObject $sheetTitle
                    Release-MbExcelComObject $sheetHeader
                }

                $startRow = if ([string]::IsNullOrWhiteSpace($summaryText)) { 2 } else { 3 }
                $steps = @($sheetModel.steps)
                for ($stepIndex = 0; $stepIndex -lt $steps.Count; $stepIndex++) {
                    Test-MbExcelCancellation -CancelPath $CancelPath
                    $step = $steps[$stepIndex]
                    $globalStep++
                    $percent = 10
                    if ($totalSteps -gt 0) { $percent = 10 + [int][Math]::Floor(($globalStep / [double]$totalSteps) * 76) }
                    Set-MbExcelStatusProgress -Status $status -StatusPath $StatusPath -Phase 'building-steps' `
                        -Message ("{0}：手順 {1} / {2}" -f $safeName, ($stepIndex + 1), $steps.Count) `
                        -CurrentStep $globalStep -TotalSteps $totalSteps -Percent $percent

                    $imagePath = ''
                    if ($step.imageId) {
                        $image = @($Project.images | Where-Object { $_.id -eq $step.imageId }) | Select-Object -First 1
                        if (-not $image -or [string]$image.fileName -notmatch '^image-[a-f0-9]{32}\.(png|jpg|bmp)$') {
                            throw "手順が参照する画像が見つかりません: $($step.id)"
                        }
                        $sourcePath = Join-Path (Join-Path (Split-Path -Parent $ProjectPath) 'images') ([string]$image.fileName)
                        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "画像ファイルが見つかりません: $($image.fileName)" }
                        $imagePath = $sourcePath
                        $annotations = @($step.annotations)
                        $crop = if ($step.PSObject.Properties.Name -contains 'crop') { $step.crop } else { $null }
                        $renderedPath = Join-Path $renderDirectory ("$($step.id).png")
                        $cropWidthRatio = if ($null -ne $crop -and $crop.PSObject.Properties.Name -contains 'width') { [double]$crop.width } else { 1.0 }
                        $cropHeightRatio = if ($null -ne $crop -and $crop.PSObject.Properties.Name -contains 'height') { [double]$crop.height } else { 1.0 }
                        $effectivePixelWidth = [double]$image.width * $cropWidthRatio
                        $effectivePixelHeight = [double]$image.height * $cropHeightRatio
                        $renderTargetWidth = if ($effectivePixelHeight -gt 0 -and ($effectivePixelWidth / $effectivePixelHeight) -ge 3.0) { 646 } else { 760 }
                        # 22行の画像領域（余白と85%制限を反映）は約620px相当。
                        # 注釈の線幅・番号径も最終配置倍率と一致させる。
                        $renderTargetHeight = if ($effectivePixelWidth -gt 0 -and ($effectivePixelHeight / $effectivePixelWidth) -ge 3.0) { 620 } else { 880 }
                        $imagePath = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations $annotations -Crop $crop `
                            -DestinationPath $renderedPath -TargetDisplayWidth $renderTargetWidth -TargetDisplayHeight $renderTargetHeight -MaximumDisplayScale 1.5 -NumberFontName $bodyFont
                        if ($imagePath -eq $renderedPath) { [void]$generatedImages.Add($renderedPath) }
                        $expectedShapes++
                    }

                    $videoLinkPath = if ($videoPlan.StepLinks.ContainsKey([string]$step.id)) { [string]$videoPlan.StepLinks[[string]$step.id] } else { '' }
                    $startRow = Add-MbExcelStepCard -Worksheet $worksheet -StartRow $startRow -StepNumber ($stepIndex + 1) `
                        -Title ([string]$step.title) -Description ([string]$step.description) -Note ([string]$step.note) `
                        -ImagePath $imagePath -FontName $bodyFont -VideoLinkPath $videoLinkPath
                }

                if ($steps.Count -eq 0) {
                    $emptyArea = $null
                    try {
                        $emptyArea = $worksheet.Range("A${startRow}:L$($startRow + 2)")
                        $emptyArea.Merge()
                        $emptyArea.NumberFormat = '@'
                        $emptyArea.Value2 = 'このシートには手順がありません'
                        $emptyArea.Font.Name = $bodyFont
                        $emptyArea.Font.Size = 11
                        $emptyArea.Font.Color = $colorMuted
                        $emptyArea.Interior.Color = ConvertTo-MbExcelBgr 248 249 251
                        $emptyArea.HorizontalAlignment = $xlCenter
                        $emptyArea.VerticalAlignment = -4108
                        Set-MbExcelEdgeBorder -Range $emptyArea -Edges @(7, 8, 9, 10) -Color $colorLine
                    } finally { Release-MbExcelComObject $emptyArea }
                }

                $used = $null
                try { $used = $worksheet.UsedRange; $used.Font.Name = $bodyFont } finally { Release-MbExcelComObject $used }
                Set-MbExcelWorksheetView -Application $excel -Worksheet $worksheet -FreezeRows 1
            } finally { Release-MbExcelComObject $worksheet }
        }

        Set-MbExcelStatusProgress -Status $status -StatusPath $StatusPath -Phase 'finishing-index' `
            -Message '目次リンクと出力順を確認しています' -CurrentStep $globalStep -TotalSteps $totalSteps -Percent 89
        $indexCells = $null
        $indexHyperlinks = $null
        try {
            $indexCells = $indexSheet.Cells
            $indexHyperlinks = $indexSheet.Hyperlinks
            for ($sheetIndex = 0; $sheetIndex -lt @($Project.sheets).Count; $sheetIndex++) {
                $row = 5 + $sheetIndex
                $sheetModel = @($Project.sheets)[$sheetIndex]
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
                    $numberCell.Value2 = [double]($sheetIndex + 1)
                    $numberCell.NumberFormat = '00'
                    $nameCell.NumberFormat = '@'
                    $nameCell.Value2 = [string]$safeName
                    $summaryCell.NumberFormat = '@'
                    $summaryText = if ($sheetModel.PSObject.Properties.Name -contains 'summary') { [string]$sheetModel.summary } else { '' }
                    if ([string]::IsNullOrWhiteSpace($summaryText)) {
                        $firstStep = @($sheetModel.steps) | Select-Object -First 1
                        if ($firstStep -and -not [string]::IsNullOrWhiteSpace([string]$firstStep.title)) {
                            $summaryText = [string]$firstStep.title
                        }
                    }
                    $summaryCell.Value2 = $summaryText
                    $countCell.Value2 = [double](@($sheetModel.steps).Count)
                    $countCell.NumberFormat = '0 "手順"'
                    try {
                        $escapedName = ConvertTo-MbExcelSheetAddress -Name $safeName
                        $newHyperlink = $indexHyperlinks.Add($nameCell, [string]'', [string]("'$escapedName'!A1"), [Type]::Missing, [string]$safeName)
                    } catch { $nameCell.Value2 = [string]$safeName }
                    $rowRange = $indexSheet.Range("A${row}:D${row}")
                    $rowRange.Font.Name = $bodyFont
                    $rowRange.Font.Size = 10.5
                    $rowRange.RowHeight = 34
                    $rowRange.WrapText = $true
                    $numberCell.Interior.Color = $colorAccentSoft
                    $numberCell.Font.Bold = $true
                    $numberCell.Font.Color = $colorAccent
                    $nameCell.Font.Bold = $true
                    $nameCell.Font.Color = $colorAccentDark
                    $summaryCell.Font.Color = $colorMuted
                    $countCell.Font.Color = $colorMuted
                    Set-MbExcelEdgeBorder -Range $rowRange -Edges @(9) -Color $colorLine
                } finally {
                    Release-MbExcelComObject $newHyperlink
                    Release-MbExcelComObject $rowRange
                    Release-MbExcelComObject $countCell
                    Release-MbExcelComObject $summaryCell
                    Release-MbExcelComObject $nameCell
                    Release-MbExcelComObject $numberCell
                }
            }
        } finally {
            Release-MbExcelComObject $indexHyperlinks
            Release-MbExcelComObject $indexCells
        }
        $indexSheet.Columns.Item(1).HorizontalAlignment = $xlCenter
        $indexSheet.Columns.Item(4).HorizontalAlignment = $xlCenter
        Set-MbExcelWorksheetView -Application $excel -Worksheet $indexSheet -FreezeRows 4

        $expectedNames = @($indexSheetName) + $createdSheetNames
        $actualNames = @()
        $actualShapes = 0
        for ($index = 1; $index -le $worksheets.Count; $index++) {
            $checkSheet = $worksheets.Item($index)
            $checkShapes = $null
            try {
                $actualNames += [string]$checkSheet.Name
                if ($index -gt 1) {
                    $checkShapes = $checkSheet.Shapes
                    $actualShapes += [int]$checkShapes.Count
                }
            } finally {
                Release-MbExcelComObject $checkShapes
                Release-MbExcelComObject $checkSheet
            }
        }
        if (($expectedNames -join '|') -ne ($actualNames -join '|')) { throw '出力シート順の自己検査に失敗しました。' }
        if ($globalStep -ne $totalSteps) { throw '出力手順数の自己検査に失敗しました。' }
        if ($actualShapes -ne $expectedShapes) { throw '出力画像数の自己検査に失敗しました。' }
        foreach ($createdName in $createdSheetNames) {
            if (-not (Test-MbExcelWorksheetName -Name $createdName)) { throw "不正な出力シート名です: $createdName" }
        }

        Test-MbExcelCancellation -CancelPath $CancelPath
        Set-MbExcelStatusProgress -Status $status -StatusPath $StatusPath -Phase 'saving' `
            -Message 'Excelファイルを保存しています' -CurrentStep $globalStep -TotalSteps $totalSteps -Percent 94
        $temporaryPath = Join-Path $OutputDirectory ('.ManualBuilder-' + $JobId + '.tmp.xlsx')
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        $workbook.SaveAs($temporaryPath, $xlOpenXMLWorkbook)
        if (-not (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) { throw 'Excelは保存完了を返しましたが、一時ファイルがありません。' }
        $workbook.Close($false)
        $workbookClosed = $true
        Test-MbExcelCancellation -CancelPath $CancelPath

        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        if ($usesFolderOutput) {
            # 動画つき。ブックと動画を1つのフォルダーへまとめ、相対リンクで再生できるようにする。
            # 途中の状態を見せないよう、別名で組み立ててから最後にフォルダー名を差し替える。
            $outputFolderName = Get-MbSafeExcelFolderName -Name (([string]$Project.title) + '_' + $stamp) -Directory $OutputDirectory
            $outputFolderPath = Join-Path $OutputDirectory $outputFolderName
            $stagingDirectory = Join-Path $OutputDirectory ('.mb-excel-' + $JobId)
            if (Test-Path -LiteralPath $stagingDirectory) { Remove-Item -LiteralPath $stagingDirectory -Recurse -Force }
            [void](New-Item -ItemType Directory -Path $stagingDirectory -Force)
            $videoDirectory = Join-Path $stagingDirectory ([string]$videoPlan.FolderName)
            [void](New-Item -ItemType Directory -Path $videoDirectory -Force)
            foreach ($videoFile in @($videoPlan.Files)) {
                [IO.File]::Copy([string]$videoFile.SourcePath, (Join-Path $videoDirectory ([string]$videoFile.FileName)), $true)
            }
            $copiedVideoCount = @(Get-ChildItem -LiteralPath $videoDirectory -File).Count
            if ($copiedVideoCount -ne [int]$videoPlan.Count) { throw '出力した動画数の自己検査に失敗しました。' }

            $outputName = Get-MbSafeExcelFileName -Name ([string]$Project.title) -Directory $stagingDirectory
            [IO.File]::Move($temporaryPath, (Join-Path $stagingDirectory $outputName))
            $temporaryPath = ''
            [IO.Directory]::Move($stagingDirectory, $outputFolderPath)
            $stagingDirectory = ''
            $outputPath = Join-Path $outputFolderPath $outputName
            $status.outputFolder = $outputFolderPath
            $status.outputFolderName = $outputFolderName
            $status.videoCount = [int]$videoPlan.Count
        } else {
            $outputName = Get-MbSafeExcelFileName -Name (([string]$Project.title) + '_' + $stamp) -Directory $OutputDirectory
            $outputPath = Join-Path $OutputDirectory $outputName
            [IO.File]::Move($temporaryPath, $outputPath)
            $temporaryPath = ''
        }
        $status.state = 'finalizing'
        $status.phase = 'finalizing'
        $status.message = 'Excelを安全に終了しています'
        $status.percent = 99
        $status.outputPath = $outputPath
        $status.outputName = $outputName
        Write-MbExcelStatus -StatusPath $StatusPath -Status $status
    } catch {
        $message = [string]$_.Exception.Message
        if ($message -eq 'MB_EXPORT_CANCELLED') {
            $status.state = 'cancelled'
            $status.phase = 'cancelled'
            $status.message = 'Excel作成を中止しました'
            $status.errorCode = 'CANCELLED'
        } else {
            $status.state = 'failed'
            $status.phase = 'failed'
            $status.errorCode = $message
            $status.message = switch ($message) {
                'MB_CONNECTED_TO_EXISTING_EXCEL' { '既存のExcelへ接続したため、安全のため作成を中止しました。Excelを閉じて再実行してください。' }
                'MB_EXCEL_OWNERSHIP_UNRESOLVED_WITH_EXISTING' { 'Excelが起動中で、作成用Excelを安全に識別できません。Excelを閉じて再実行してください。' }
                'MB_EXCEL_OWNERSHIP_UNRESOLVED' { '作成用Excelの安全確認ができませんでした。' }
                'MB_EXCEL_OWNERSHIP_API_UNAVAILABLE' { 'Excelの所有確認に必要なWindows機能を利用できません。' }
                default { 'Excelファイルを作成できませんでした: ' + $message }
            }
        }
    } finally {
        try { if ($workbook -and -not $workbookClosed) { $workbook.Close($false) } } catch { }
        Release-MbExcelComObject $indexSheet
        Release-MbExcelComObject $worksheets
        Release-MbExcelComObject $workbook
        if ($excel -and $canQuitCom) {
            if ($settingsApplied) {
                try { $excel.EnableEvents = $true } catch { }
                try { $excel.ScreenUpdating = $true } catch { }
            }
            try { $excel.Quit() } catch { }
        }
        Release-MbExcelComObject $excel
        $indexSheet = $null
        $worksheets = $null
        $workbook = $null
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        # Stop-ProcessはHwndで所有を証明したPIDにのみ許可する。
        # PID差分は「起動直後に利用者がExcelを開いた」場合に他人のプロセスを指し得るため、
        # 未保存ブックを失わせないよう強制終了の対象にしない（Quitと参照解放に任せる）。
        if ($ownershipProven -and $ownPid -gt 0 -and $ownershipMode -eq 'Hwnd') {
            $deadline = (Get-Date).AddSeconds(10)
            do {
                if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { break }
                Start-Sleep -Milliseconds 500
            } while ((Get-Date) -lt $deadline)
            $ownedProcess = Get-Process -Id $ownPid -ErrorAction SilentlyContinue
            if ($ownedProcess -and $ownedProcess.ProcessName -eq 'EXCEL') {
                Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue
            }
        }
        # 失敗・中止のときに、組み立て途中のフォルダーを保存先へ残さない。
        if ($stagingDirectory -and (Test-Path -LiteralPath $stagingDirectory)) {
            Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
        if ($temporaryPath -and (Test-Path -LiteralPath $temporaryPath)) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
        foreach ($generatedImage in @($generatedImages)) {
            if (Test-Path -LiteralPath $generatedImage) {
                Remove-Item -LiteralPath $generatedImage -Force -ErrorAction SilentlyContinue
            }
        }
    }

    if ($status.state -eq 'finalizing') {
        $status.state = 'completed'
        $status.phase = 'completed'
        $status.message = 'Excelファイルを作成しました'
        $status.percent = 100
    }
    $status.completedAt = [DateTime]::UtcNow.ToString('o')
    Write-MbExcelStatus -StatusPath $StatusPath -Status $status
    return $status
}

Export-ModuleMember -Function @(
    'Get-MbSafeExcelFileName',
    'Get-MbSafeExcelFolderName',
    'Get-MbExcelVideoPlan',
    'Get-MbSafeExcelWorksheetName',
    'Test-MbExcelWorksheetName',
    'Get-MbExcelStepCardLayout',
    'New-MbAnnotatedImage',
    'Invoke-MbExcelExport'
)
