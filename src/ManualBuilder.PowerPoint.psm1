# ManualBuilder PowerPoint exporter.
# PowerPointだけは動画をファイルの中へ取り込める（ExcelとWordはリンクかアイコンにしかならない）。
# 手順ごとに1枚のスライドを作り、動画のある手順は動画を、無い手順は編集済み画像を貼る。
# WordのExport実装（ManualBuilder.Word.psm1）と同じ安全境界を踏襲する。

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Drawing

$script:MbPowerPointPInvokeAvailable = $false
try {
    Add-Type -Namespace ManualBuilderPowerPoint -Name NativeMethods -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@ -ErrorAction Stop
    $script:MbPowerPointPInvokeAvailable = $true
} catch {
    try { [void][ManualBuilderPowerPoint.NativeMethods]; $script:MbPowerPointPInvokeAvailable = $true } catch { }
}

# スライドは16:9（960x540pt）。PowerPoint 2013以降の既定と同じ。
$script:MbSlideWidth = 960.0
$script:MbSlideHeight = 540.0
$script:MbMediaLeft = 40.0
$script:MbMediaTop = 96.0
$script:MbMediaWidth = 570.0
$script:MbMediaHeight = 400.0
$script:MbTextLeft = 630.0
$script:MbTextWidth = 290.0

function ConvertTo-MbPowerPointRgb {
    param([int]$R, [int]$G, [int]$B)
    return ($B * 65536) + ($G * 256) + $R
}

function Release-MbPowerPointComObject {
    param([AllowNull()][object]$InputObject)
    if ($null -ne $InputObject) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($InputObject) } catch { }
    }
}

function Get-MbPowerPointProcessId {
    param([Parameter(Mandatory = $true)][object]$Application)
    if (-not $script:MbPowerPointPInvokeAvailable) { return 0 }
    $handle = [int64]0
    try { $handle = [int64]$Application.HWND } catch { $handle = [int64]0 }
    if ($handle -eq 0) { return 0 }
    try {
        $processId = [uint32]0
        [void][ManualBuilderPowerPoint.NativeMethods]::GetWindowThreadProcessId([IntPtr]$handle, [ref]$processId)
        return [int]$processId
    } catch { return 0 }
}

function Resolve-MbOwnedPowerPointProcess {
    param(
        [Parameter(Mandatory = $true)][object]$Application,
        [int[]]$PidsBefore,
        [int]$TimeoutMilliseconds = 5000
    )
    $deadline = (Get-Date).AddMilliseconds($TimeoutMilliseconds)
    do {
        $idFromHwnd = Get-MbPowerPointProcessId -Application $Application
        if ($idFromHwnd -gt 0) {
            if ($PidsBefore -contains $idFromHwnd) {
                return [pscustomobject]@{ Pid = 0; Mode = "既存POWERPNT PID $idFromHwnd へ接続"; ExistingConnection = $true }
            }
            return [pscustomobject]@{ Pid = $idFromHwnd; Mode = 'Hwnd'; ExistingConnection = $false }
        }
        $current = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $added = @($current | Where-Object { $PidsBefore -notcontains $_ })
        if ($added.Count -eq 1) {
            return [pscustomobject]@{ Pid = [int]$added[0]; Mode = 'PID差分'; ExistingConnection = $false }
        }
        Start-Sleep -Milliseconds 200
    } while ((Get-Date) -lt $deadline)
    return [pscustomobject]@{ Pid = 0; Mode = ''; ExistingConnection = $false }
}

function Resolve-MbPowerPointBodyFont {
    $candidates = @('BIZ UDPGothic', 'BIZ UDPゴシック', 'Meiryo', 'Yu Gothic UI', 'MS Pゴシック')
    $installed = $null
    try { $installed = New-Object Drawing.Text.InstalledFontCollection } catch { return 'Meiryo' }
    try {
        $names = @($installed.Families | ForEach-Object { $_.Name })
        foreach ($candidate in $candidates) {
            if ($names -contains $candidate) { return $candidate }
        }
    } finally {
        if ($installed) { $installed.Dispose() }
    }
    return 'Meiryo'
}

function Write-MbPowerPointStatus {
    param([string]$StatusPath, [object]$Status)
    $directory = Split-Path -Parent $StatusPath
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    $temporaryPath = Join-Path $directory ('.ppt-status-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = Join-Path $directory ('.ppt-status-backup-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $Status.updatedAt = [DateTime]::UtcNow.ToString('o')
        [IO.File]::WriteAllText($temporaryPath, ($Status | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $StatusPath) { [IO.File]::Replace($temporaryPath, $StatusPath, $backupPath, $true) }
        else { [IO.File]::Move($temporaryPath, $StatusPath) }
    } finally {
        foreach ($path in @($temporaryPath, $backupPath)) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Set-MbPowerPointStatusProgress {
    param($Status, [string]$StatusPath, [string]$Phase, [string]$Message, [int]$CurrentStep, [int]$TotalSteps, [int]$Percent)
    $Status.phase = $Phase; $Status.message = $Message; $Status.currentStep = $CurrentStep
    $Status.totalSteps = $TotalSteps; $Status.percent = $Percent
    Write-MbPowerPointStatus -StatusPath $StatusPath -Status $Status
}

function Test-MbPowerPointCancellation {
    param([string]$CancelPath)
    if (Test-Path -LiteralPath $CancelPath -PathType Leaf) { throw 'MB_EXPORT_CANCELLED' }
}

function Get-MbPowerPointContentSheets {
    param([Parameter(Mandatory = $true)][object]$Project)
    return @($Project.sheets | Where-Object { @($_.steps).Count -gt 0 })
}

function Add-MbPowerPointTextBox {
    param(
        [Parameter(Mandatory = $true)][object]$Slide,
        [Parameter(Mandatory = $true)][string]$Text,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$FontSize = 14,
        [bool]$Bold = $false,
        [int]$Color = 0,
        [string]$FontName = 'Meiryo'
    )
    $shapes = $null; $shape = $null; $textFrame = $null; $textRange = $null; $font = $null; $paragraphFormat = $null
    try {
        $shapes = $Slide.Shapes
        # 1 = msoTextOrientationHorizontal
        $shape = $shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
        $textFrame = $shape.TextFrame
        $textFrame.WordWrap = -1
        $textFrame.AutoSize = 0
        $textFrame.MarginLeft = 0; $textFrame.MarginRight = 0
        $textFrame.MarginTop = 0; $textFrame.MarginBottom = 0
        $textRange = $textFrame.TextRange
        $textRange.Text = $Text
        $font = $textRange.Font
        $font.Name = $FontName
        try { $font.NameFarEast = $FontName } catch { }
        $font.Size = $FontSize
        $font.Bold = if ($Bold) { -1 } else { 0 }
        $font.Color.RGB = $Color
        $paragraphFormat = $textRange.ParagraphFormat
        try { $paragraphFormat.SpaceWithin = 1.1 } catch { }
        return $shape
    } finally {
        Release-MbPowerPointComObject $paragraphFormat
        Release-MbPowerPointComObject $font
        Release-MbPowerPointComObject $textRange
        Release-MbPowerPointComObject $textFrame
        Release-MbPowerPointComObject $shapes
    }
}

function Set-MbPowerPointMediaLayout {
    param([Parameter(Mandatory = $true)][object]$Shape)
    # 縦横比を保ったまま、決めた枠へ収める。
    $naturalWidth = [double]$Shape.Width
    $naturalHeight = [double]$Shape.Height
    if ($naturalWidth -le 0 -or $naturalHeight -le 0) { return }
    $scale = [Math]::Min($script:MbMediaWidth / $naturalWidth, $script:MbMediaHeight / $naturalHeight)
    # 小さすぎる素材でも間延びしないよう、拡大は1.5倍までに抑える（Excel出力と同じ考え方）。
    $scale = [Math]::Min(1.5, $scale)
    $Shape.LockAspectRatio = -1
    $Shape.Width = [single]($naturalWidth * $scale)
    $Shape.Left = [single]($script:MbMediaLeft + (($script:MbMediaWidth - [double]$Shape.Width) / 2))
    $Shape.Top = [single]($script:MbMediaTop + (($script:MbMediaHeight - [double]$Shape.Height) / 2))
}

function Get-MbSafePowerPointFileName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Directory,
        [string]$Extension = '.pptx'
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

function Invoke-MbPowerPointExport {
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
        jobId = $JobId; state = 'running'; phase = 'starting'; message = 'PowerPoint出力を準備しています'; percent = 0
        currentStep = 0; totalSteps = $totalSteps; outputPath = ''; outputName = ''; outputDirectory = $OutputDirectory
        ownedPowerPointPid = 0; ownedPowerPointStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false
        slideCount = 0; videoCount = 0
        startedAt = $startedAt; updatedAt = $startedAt; completedAt = ''; errorCode = ''
    }
    Write-MbPowerPointStatus $StatusPath $status

    $powerPoint = $null; $presentation = $null
    $presentationClosed = $false; $canQuitCom = $false; $ownershipProven = $false
    $ownPid = 0; $ownershipMode = ''; $temporaryPath = ''; $outputPath = ''
    $renderDirectory = Join-Path (Split-Path -Parent $StatusPath) 'rendered-ppt-images'
    $generatedImages = New-Object System.Collections.ArrayList
    $pidsBefore = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

    try {
        Test-MbPowerPointCancellation $CancelPath
        if ($pidsBefore.Count -gt 0) { throw 'MB_POWERPOINT_RUNNING' }
        if (-not $script:MbPowerPointPInvokeAvailable) { throw 'MB_POWERPOINT_OWNERSHIP_API_UNAVAILABLE' }
        if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }
        if (-not (Test-Path -LiteralPath $renderDirectory)) { [void](New-Item -ItemType Directory -Path $renderDirectory -Force) }
        $fontName = Resolve-MbPowerPointBodyFont
        $colorText = ConvertTo-MbPowerPointRgb 24 32 51
        $colorMuted = ConvertTo-MbPowerPointRgb 90 99 115
        $colorAccent = ConvertTo-MbPowerPointRgb 58 91 160

        Set-MbPowerPointStatusProgress $status $StatusPath 'starting-powerpoint' '安全確認のためPowerPointを起動しています' 0 $totalSteps 3
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $canQuitCom = $true
        $resolved = Resolve-MbOwnedPowerPointProcess -Application $powerPoint -PidsBefore $pidsBefore
        $ownPid = [int]$resolved.Pid; $ownershipMode = [string]$resolved.Mode
        if ([bool]$resolved.ExistingConnection) { throw 'MB_CONNECTED_TO_EXISTING_POWERPOINT' }
        if ($ownPid -le 0) { throw 'MB_POWERPOINT_OWNERSHIP_UNRESOLVED' }
        $ownershipProven = $true
        $status.ownedPowerPointPid = $ownPid; $status.ownershipMode = $ownershipMode; $status.ownershipProven = $true
        try { $status.ownedPowerPointStartTimeUtc = (Get-Process -Id $ownPid -ErrorAction Stop).StartTime.ToUniversalTime().ToString('o') } catch { }
        Write-MbPowerPointStatus $StatusPath $status

        try { $powerPoint.DisplayAlerts = 1 } catch { }
        # PowerPointは Visible = $false を受け付けない。ウィンドウ無しでプレゼンテーションを開けば
        # 画面には出ないため、Presentations.Add に msoFalse を渡す。
        $presentations = $null
        try { $presentations = $powerPoint.Presentations; $presentation = $presentations.Add(0) }
        finally { Release-MbPowerPointComObject $presentations }

        if ($ownershipMode -ne 'Hwnd') {
            $hwndOwnedPid = Get-MbPowerPointProcessId -Application $powerPoint
            if ($hwndOwnedPid -gt 0 -and $hwndOwnedPid -eq $ownPid) {
                $ownershipMode = 'Hwnd'
                $status.ownershipMode = $ownershipMode
                Write-MbPowerPointStatus $StatusPath $status
            }
        }

        $pageSetup = $null
        try {
            $pageSetup = $presentation.PageSetup
            $pageSetup.SlideWidth = [single]$script:MbSlideWidth
            $pageSetup.SlideHeight = [single]$script:MbSlideHeight
        } finally { Release-MbPowerPointComObject $pageSetup }

        Set-MbPowerPointStatusProgress $status $StatusPath 'building-cover' '表紙を作成しています' 0 $totalSteps 8
        $slideIndex = 0
        $slides = $null
        try {
            $slides = $presentation.Slides
            # 12 = ppLayoutBlank
            $coverSlide = $null
            try {
                $slideIndex++
                $coverSlide = $slides.Add($slideIndex, 12)
                $shape = Add-MbPowerPointTextBox -Slide $coverSlide -Text ([string]$Project.title) `
                    -Left 60 -Top 200 -Width 840 -Height 80 -FontSize 36 -Bold $true -Color $colorText -FontName $fontName
                Release-MbPowerPointComObject $shape
                $shape = Add-MbPowerPointTextBox -Slide $coverSlide -Text ('操作マニュアル　' + (Get-Date -Format 'yyyy年M月d日')) `
                    -Left 60 -Top 292 -Width 840 -Height 40 -FontSize 16 -Color $colorMuted -FontName $fontName
                Release-MbPowerPointComObject $shape
            } finally { Release-MbPowerPointComObject $coverSlide }

            $globalStep = 0
            $videoCount = 0
            $includedSheets = @(Get-MbPowerPointContentSheets -Project $Project)
            if ($includedSheets.Count -eq 0) { throw 'PowerPointへ出力する手順がありません。' }
            foreach ($sheet in $includedSheets) {
                Test-MbPowerPointCancellation $CancelPath
                $sheetSlide = $null
                try {
                    $slideIndex++
                    $sheetSlide = $slides.Add($slideIndex, 12)
                    $shape = Add-MbPowerPointTextBox -Slide $sheetSlide -Text ([string]$sheet.name) `
                        -Left 60 -Top 220 -Width 840 -Height 60 -FontSize 28 -Bold $true -Color $colorAccent -FontName $fontName
                    Release-MbPowerPointComObject $shape
                    if ($sheet.PSObject.Properties.Name -contains 'summary' -and -not [string]::IsNullOrWhiteSpace([string]$sheet.summary)) {
                        $shape = Add-MbPowerPointTextBox -Slide $sheetSlide -Text ([string]$sheet.summary) `
                            -Left 60 -Top 288 -Width 840 -Height 60 -FontSize 14 -Color $colorMuted -FontName $fontName
                        Release-MbPowerPointComObject $shape
                    }
                } finally { Release-MbPowerPointComObject $sheetSlide }

                $sheetSteps = @($sheet.steps)
                for ($stepIndex = 0; $stepIndex -lt $sheetSteps.Count; $stepIndex++) {
                    Test-MbPowerPointCancellation $CancelPath
                    $step = $sheetSteps[$stepIndex]; $globalStep++
                    $title = if ([string]::IsNullOrWhiteSpace([string]$step.title)) { '手順名未入力' } else { [string]$step.title }
                    Set-MbPowerPointStatusProgress $status $StatusPath 'building-steps' ("手順 $globalStep / $totalSteps を作成しています") `
                        $globalStep $totalSteps ([Math]::Min(90, 10 + [int](80 * $globalStep / [Math]::Max(1, $totalSteps))))

                    $stepSlide = $null
                    try {
                        $slideIndex++
                        $stepSlide = $slides.Add($slideIndex, 12)
                        $headingText = [string]("手順 {0}　{1}" -f ($stepIndex + 1), $title)
                        $shape = Add-MbPowerPointTextBox -Slide $stepSlide -Text $headingText `
                            -Left 40 -Top 34 -Width 880 -Height 44 -FontSize 22 -Bold $true -Color $colorText -FontName $fontName
                        Release-MbPowerPointComObject $shape

                        $bodyText = [string]$step.description
                        if (-not [string]::IsNullOrWhiteSpace([string]$step.note)) {
                            if (-not [string]::IsNullOrWhiteSpace($bodyText)) { $bodyText += "`r`r" }
                            $bodyText += ('補足: ' + [string]$step.note)
                        }
                        if (-not [string]::IsNullOrWhiteSpace($bodyText)) {
                            $shape = Add-MbPowerPointTextBox -Slide $stepSlide -Text ($bodyText.Replace("`r`n", "`r").Replace("`n", "`r")) `
                                -Left $script:MbTextLeft -Top $script:MbMediaTop -Width $script:MbTextWidth -Height $script:MbMediaHeight `
                                -FontSize 14 -Color $colorText -FontName $fontName
                            Release-MbPowerPointComObject $shape
                        }

                        $renderedPath = ''
                        if (-not [string]::IsNullOrWhiteSpace([string]$step.imageId)) {
                            $sourcePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$step.imageId)
                            if (-not $sourcePath -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "画像が見つかりません: $($step.imageId)" }
                            $destination = Join-Path $renderDirectory ("ppt-image-{0:d4}.png" -f $globalStep)
                            $renderedPath = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations @($step.annotations) -Crop $step.crop `
                                -DestinationPath $destination -TargetDisplayWidth 760 -TargetDisplayHeight 540 -MaximumDisplayScale 1.5 -NumberFontName $fontName
                            if ($renderedPath -eq $destination) { [void]$generatedImages.Add($destination) }
                        }

                        $videoPath = ''
                        if ($step.PSObject.Properties.Name -contains 'videoId' -and -not [string]::IsNullOrWhiteSpace([string]$step.videoId)) {
                            $videoPath = Get-MbVideoFilePath -Project $Project -ProjectPath $ProjectPath -VideoId ([string]$step.videoId)
                            if (-not $videoPath -or -not (Test-Path -LiteralPath $videoPath -PathType Leaf)) { throw "動画が見つかりません: $($step.videoId)" }
                        }

                        if ($videoPath) {
                            $mediaShape = $null; $mediaFormat = $null; $stepShapes = $null
                            try {
                                $stepShapes = $stepSlide.Shapes
                                # LinkToFile=msoFalse, SaveWithDocument=msoTrue でpptxの中へ取り込む。
                                $mediaShape = $stepShapes.AddMediaObject2($videoPath, 0, -1,
                                    [single]$script:MbMediaLeft, [single]$script:MbMediaTop,
                                    [single]$script:MbMediaWidth, [single]$script:MbMediaHeight)
                                Set-MbPowerPointMediaLayout -Shape $mediaShape
                                if ($renderedPath) {
                                    # 再生前は編集済みの画像（赤枠・番号つき）を表紙として見せる。
                                    try {
                                        $mediaFormat = $mediaShape.MediaFormat
                                        $mediaFormat.SetDisplayPicture($renderedPath)
                                    } catch { }
                                }
                                try { $mediaShape.AlternativeText = "手順 $($stepIndex + 1) の操作動画: $title" } catch { }
                                $videoCount++
                            } finally {
                                Release-MbPowerPointComObject $mediaFormat
                                Release-MbPowerPointComObject $mediaShape
                                Release-MbPowerPointComObject $stepShapes
                            }
                            $shape = Add-MbPowerPointTextBox -Slide $stepSlide -Text '▶ 画像をクリックすると操作の動画を再生できます' `
                                -Left $script:MbMediaLeft -Top ($script:MbMediaTop + $script:MbMediaHeight + 8) -Width $script:MbMediaWidth -Height 24 `
                                -FontSize 12 -Color $colorAccent -FontName $fontName
                            Release-MbPowerPointComObject $shape
                        } elseif ($renderedPath) {
                            $pictureShape = $null; $stepShapes = $null
                            try {
                                $stepShapes = $stepSlide.Shapes
                                $pictureShape = $stepShapes.AddPicture($renderedPath, 0, -1,
                                    [single]$script:MbMediaLeft, [single]$script:MbMediaTop,
                                    [single]$script:MbMediaWidth, [single]$script:MbMediaHeight)
                                Set-MbPowerPointMediaLayout -Shape $pictureShape
                                try { $pictureShape.AlternativeText = "手順 $($stepIndex + 1) の画面: $title" } catch { }
                            } finally {
                                Release-MbPowerPointComObject $pictureShape
                                Release-MbPowerPointComObject $stepShapes
                            }
                        }
                    } finally { Release-MbPowerPointComObject $stepSlide }
                }
            }

            Test-MbPowerPointCancellation $CancelPath
            $status.slideCount = [int]$slides.Count
            $status.videoCount = $videoCount
            if ([int]$slides.Count -ne $slideIndex) { throw '出力スライド数の自己検査に失敗しました。' }
            if ($globalStep -ne $totalSteps) { throw '出力手順数の自己検査に失敗しました。' }
        } finally { Release-MbPowerPointComObject $slides }

        Set-MbPowerPointStatusProgress $status $StatusPath 'saving' 'PowerPointファイルを保存しています' $totalSteps $totalSteps 94
        $temporaryPath = Join-Path $OutputDirectory ('.ManualBuilder-' + $JobId + '.tmp.pptx')
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        # 24 = ppSaveAsOpenXMLPresentation
        $presentation.SaveAs($temporaryPath, 24)
        $presentation.Close(); $presentationClosed = $true
        Test-MbPowerPointCancellation $CancelPath
        $outputName = Get-MbSafePowerPointFileName -Name (([string]$Project.title) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss')) -Directory $OutputDirectory
        $outputPath = Join-Path $OutputDirectory $outputName
        [IO.File]::Move($temporaryPath, $outputPath); $temporaryPath = ''
        $status.state = 'finalizing'; $status.phase = 'finalizing'; $status.message = 'PowerPointを安全に終了しています'; $status.percent = 99
        $status.outputPath = $outputPath; $status.outputName = $outputName
        Write-MbPowerPointStatus $StatusPath $status
    } catch {
        $message = [string]$_.Exception.Message
        if ($message -eq 'MB_EXPORT_CANCELLED') {
            $status.state = 'cancelled'; $status.phase = 'cancelled'; $status.message = 'PowerPoint作成を中止しました'; $status.errorCode = 'CANCELLED'
        } else {
            $status.state = 'failed'; $status.phase = 'failed'; $status.errorCode = $message
            $status.message = switch ($message) {
                'MB_POWERPOINT_RUNNING' { 'PowerPointが開いているため、安全のため作成を開始しませんでした。PowerPointを閉じて再実行するか、Excelで作成してください。' }
                'MB_CONNECTED_TO_EXISTING_POWERPOINT' { '既存のPowerPointへ接続したため、安全のため作成を中止しました。PowerPointを閉じて再実行してください。' }
                'MB_POWERPOINT_OWNERSHIP_UNRESOLVED' { '作成用PowerPointの安全確認ができませんでした。' }
                'MB_POWERPOINT_OWNERSHIP_API_UNAVAILABLE' { 'PowerPointの所有確認に必要なWindows機能を利用できません。' }
                default { 'PowerPointファイルを作成できませんでした: ' + $message }
            }
        }
    } finally {
        try { if ($presentation -and -not $presentationClosed) { $presentation.Close() } } catch { }
        Release-MbPowerPointComObject $presentation
        if ($powerPoint -and $canQuitCom) {
            try { $powerPoint.Quit() } catch { }
        }
        Release-MbPowerPointComObject $powerPoint
        $presentation = $null; $powerPoint = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # Stop-ProcessはHwndで所有を証明したPIDにのみ許可する（Word出力と同じ安全境界）。
        if ($ownershipProven -and $ownPid -gt 0 -and $ownershipMode -eq 'Hwnd') {
            $deadline = (Get-Date).AddSeconds(10)
            do {
                if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { break }
                Start-Sleep -Milliseconds 500
            } while ((Get-Date) -lt $deadline)
            $ownedProcess = Get-Process -Id $ownPid -ErrorAction SilentlyContinue
            if ($ownedProcess -and $ownedProcess.ProcessName -eq 'POWERPNT') { Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue }
        }
        if ($temporaryPath -and (Test-Path -LiteralPath $temporaryPath)) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
        foreach ($imagePath in @($generatedImages)) {
            if (Test-Path -LiteralPath $imagePath) { Remove-Item -LiteralPath $imagePath -Force -ErrorAction SilentlyContinue }
        }
    }
    if ($status.state -eq 'finalizing') {
        $status.state = 'completed'; $status.phase = 'completed'; $status.message = 'PowerPointファイルを作成しました'; $status.percent = 100
    }
    $status.completedAt = [DateTime]::UtcNow.ToString('o')
    Write-MbPowerPointStatus $StatusPath $status
    return $status
}

Export-ModuleMember -Function @(
    'Get-MbSafePowerPointFileName',
    'Invoke-MbPowerPointExport'
)
