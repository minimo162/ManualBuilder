# ManualBuilder Word exporter.

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Drawing

$script:MbWordPInvokeAvailable = $false
try {
    Add-Type -Namespace ManualBuilderWord -Name NativeMethods -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@ -ErrorAction Stop
    $script:MbWordPInvokeAvailable = $true
} catch {
    try { [void][ManualBuilderWord.NativeMethods]; $script:MbWordPInvokeAvailable = $true } catch { }
}

function ConvertTo-MbWordBgr {
    param([int]$R, [int]$G, [int]$B)
    return ($B * 65536) + ($G * 256) + $R
}

function Release-MbWordComObject {
    param([AllowNull()][object]$InputObject)
    if ($null -ne $InputObject) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($InputObject) } catch { }
    }
}

function Get-MbWordProcessId {
    param([Parameter(Mandatory = $true)][object]$Application)
    if (-not $script:MbWordPInvokeAvailable) { return 0 }
    # WordのApplicationはExcelと違い Hwnd を公開せず、常に0になる。
    # 文書を1つ開いた後の ActiveWindow.Hwnd なら所有プロセスを特定できるため、そちらを優先する。
    $handle = [int64]0
    try { $handle = [int64]$Application.ActiveWindow.Hwnd } catch { $handle = [int64]0 }
    if ($handle -eq 0) {
        try { $handle = [int64]$Application.Hwnd } catch { $handle = [int64]0 }
    }
    if ($handle -eq 0) { return 0 }
    try {
        $processId = [uint32]0
        [void][ManualBuilderWord.NativeMethods]::GetWindowThreadProcessId([IntPtr]$handle, [ref]$processId)
        return [int]$processId
    } catch { return 0 }
}

function Resolve-MbOwnedWordProcess {
    param(
        [Parameter(Mandatory = $true)][object]$Application,
        [int[]]$PidsBefore,
        [int]$TimeoutMilliseconds = 5000
    )
    $deadline = (Get-Date).AddMilliseconds($TimeoutMilliseconds)
    do {
        $idFromHwnd = Get-MbWordProcessId -Application $Application
        if ($idFromHwnd -gt 0) {
            if ($PidsBefore -contains $idFromHwnd) {
                return [pscustomobject]@{ Pid = 0; Mode = "既存WINWORD PID $idFromHwnd へ接続"; ExistingConnection = $true }
            }
            return [pscustomobject]@{ Pid = $idFromHwnd; Mode = 'Hwnd'; ExistingConnection = $false }
        }
        $pidsNow = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $newPids = @($pidsNow | Where-Object { $PidsBefore -notcontains $_ })
        if ($PidsBefore.Count -eq 0 -and $newPids.Count -eq 1) {
            return [pscustomobject]@{ Pid = [int]$newPids[0]; Mode = '空ベースライン＋PID差分'; ExistingConnection = $false }
        }
        if ($newPids.Count -gt 1) {
            return [pscustomobject]@{ Pid = 0; Mode = "新規WINWORD PIDが複数: $($newPids -join ',')"; ExistingConnection = $false }
        }
        Start-Sleep -Milliseconds 200
    } while ((Get-Date) -lt $deadline)
    return [pscustomobject]@{ Pid = 0; Mode = '所有PIDを特定できず'; ExistingConnection = $false }
}

function Resolve-MbWordBodyFont {
    $collection = New-Object Drawing.Text.InstalledFontCollection
    try {
        $installed = @($collection.Families | ForEach-Object { $_.Name })
        foreach ($font in @('BIZ UDPゴシック', 'BIZ UDPGothic', 'BIZ UDゴシック', 'Meiryo', 'Yu Gothic UI', 'MS Pゴシック')) {
            if ($installed -contains $font) { return $font }
        }
        return 'MS Pゴシック'
    } finally { $collection.Dispose() }
}

function Get-MbSafeWordFileName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Directory,
        [string]$Extension = '.docx'
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

function Write-MbWordStatus {
    param([string]$StatusPath, [object]$Status)
    $directory = Split-Path -Parent $StatusPath
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    $temporaryPath = Join-Path $directory ('.word-status-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = Join-Path $directory ('.word-status-backup-' + [guid]::NewGuid().ToString('N') + '.tmp')
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

function Set-MbWordStatusProgress {
    param($Status, [string]$StatusPath, [string]$Phase, [string]$Message, [int]$CurrentStep, [int]$TotalSteps, [int]$Percent)
    $Status.phase = $Phase; $Status.message = $Message; $Status.currentStep = $CurrentStep
    $Status.totalSteps = $TotalSteps; $Status.percent = $Percent
    Write-MbWordStatus -StatusPath $StatusPath -Status $Status
}

function Test-MbWordCancellation {
    param([string]$CancelPath)
    if (Test-Path -LiteralPath $CancelPath -PathType Leaf) { throw 'MB_EXPORT_CANCELLED' }
}

function Set-MbWordStyle {
    param($Selection, $Document, [int]$StyleId)
    try { $Selection.Style = $StyleId }
    catch { $Selection.Style = $Document.Styles.Item($StyleId) }
}

function Move-MbWordSelectionToEnd {
    param($Selection, $Document)
    $endPosition = [Math]::Max(0, ([int]$Document.Content.End - 1))
    $Selection.SetRange($endPosition, $endPosition)
}

function Get-MbWordFirstNonEmptyParagraphText {
    param($Document)
    for ($index = 1; $index -le $Document.Paragraphs.Count; $index++) {
        $paragraph = $null; $range = $null
        try {
            $paragraph = $Document.Paragraphs.Item($index)
            $range = $paragraph.Range
            $text = ([string]$range.Text).Trim().Trim([char[]]@([char]7)).Trim()
            if ($text) { return $text }
        } finally {
            Release-MbWordComObject $range
            Release-MbWordComObject $paragraph
        }
    }
    return ''
}

function Get-MbWordContentSheets {
    param([Parameter(Mandatory = $true)][object]$Project)
    return @($Project.sheets | Where-Object { @($_.steps).Count -gt 0 })
}

function Set-MbWordStyleDefinition {
    param($Document, [int]$StyleId, [string]$FontName, [double]$Size, [bool]$Bold, [int]$Color, [double]$Before, [double]$After)
    $style = $null
    $font = $null
    $paragraph = $null
    try {
        $style = $Document.Styles.Item($StyleId)
        $font = $style.Font
        $font.Name = $FontName; $font.NameFarEast = $FontName; $font.NameAscii = $FontName
        $font.Size = $Size; $font.Bold = $Bold; $font.Color = $Color
        $paragraph = $style.ParagraphFormat
        $paragraph.SpaceBefore = $Before; $paragraph.SpaceAfter = $After
        try { $paragraph.LineSpacingRule = 0 } catch { }
    } finally {
        Release-MbWordComObject $paragraph
        Release-MbWordComObject $font
        Release-MbWordComObject $style
    }
}

function Set-MbWordDocumentDesign {
    param($Document, [string]$FontName)
    # ExcelとWordは同じproject.jsonから作るため、テーマ色をExcel側（ManualBuilder.Excel.psm1）と揃える。
    $colorText = ConvertTo-MbWordBgr 24 32 51
    $colorAccent = ConvertTo-MbWordBgr 58 91 160
    Set-MbWordStyleDefinition $Document -1 $FontName 10.5 $false $colorText 0 6
    Set-MbWordStyleDefinition $Document -2 $FontName 16 $true $colorAccent 18 10
    Set-MbWordStyleDefinition $Document -3 $FontName 13 $true $colorAccent 14 7
    Set-MbWordStyleDefinition $Document -63 $FontName 28 $true $colorText 0 12
    Set-MbWordStyleDefinition $Document -75 $FontName 12 $false (ConvertTo-MbWordBgr 91 98 110) 0 8
    $contentFont = $null
    try {
        $contentFont = $Document.Content.Font
        $contentFont.Name = $FontName; $contentFont.NameFarEast = $FontName; $contentFont.NameAscii = $FontName
        $contentFont.Size = 10.5
    } finally { Release-MbWordComObject $contentFont }
}

function Add-MbWordTextParagraph {
    param($Selection, $Document, [AllowEmptyString()][string]$Text, [bool]$KeepTogether = $false)
    Set-MbWordStyle $Selection $Document -1
    $Selection.ParagraphFormat.Alignment = 0
    $Selection.ParagraphFormat.LeftIndent = 0
    $Selection.ParagraphFormat.RightIndent = 0
    $lines = @(([string]$Text) -split "`r?`n")
    if ($lines.Count -eq 0) { $lines = @('') }
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if ($index -gt 0) { $Selection.TypeText([string][char]11) }
        $Selection.TypeText([string]$lines[$index])
    }
    $paragraph = $null
    try {
        $paragraph = $Selection.Paragraphs.Item(1)
        $paragraph.KeepTogether = $KeepTogether
        $paragraph.KeepWithNext = $false
    } finally { Release-MbWordComObject $paragraph }
    $Selection.TypeParagraph()
}

function Add-MbWordNote {
    param($Selection, $Document, [string]$Text, [string]$FontName)
    $tableRange = $null; $table = $null; $cell = $null; $cellRange = $null; $font = $null; $paragraph = $null
    try {
        $endPosition = [Math]::Max(0, ([int]$Document.Content.End - 1))
        $tableRange = $Document.Range($endPosition, $endPosition)
        $table = $Document.Tables.Add($tableRange, 1, 1)
        $table.Range.Shading.BackgroundPatternColor = ConvertTo-MbWordBgr 255 249 226
        $table.Borders.OutsideLineStyle = 1
        $table.Borders.InsideLineStyle = 1
        $table.Borders.OutsideColor = ConvertTo-MbWordBgr 234 217 168
        $table.Borders.InsideColor = ConvertTo-MbWordBgr 234 217 168
        $table.Rows.AllowBreakAcrossPages = $false
        $cell = $table.Cell(1, 1)
        $cell.TopPadding = 6; $cell.BottomPadding = 6; $cell.LeftPadding = 8; $cell.RightPadding = 8
        $cellRange = $cell.Range
        $cellRange.Text = '補足　' + $Text
        $font = $cellRange.Font
        $font.Name = $FontName; $font.NameFarEast = $FontName; $font.NameAscii = $FontName
        $font.Size = 10.5; $font.Color = ConvertTo-MbWordBgr 92 72 20
        $paragraph = $cellRange.Paragraphs.Item(1)
        $paragraph.KeepWithNext = $true
        $table.Range.InsertParagraphAfter()
    } finally {
        Release-MbWordComObject $paragraph
        Release-MbWordComObject $font
        Release-MbWordComObject $cellRange
        Release-MbWordComObject $cell
        Release-MbWordComObject $table
        Release-MbWordComObject $tableRange
    }
    Move-MbWordSelectionToEnd $Selection $Document
    Set-MbWordStyle $Selection $Document -1
}

function Save-MbWordDocument {
    param($Document, [string]$Path)
    $savePathRef = $Path
    $formatRef = 16
    $Document.SaveAs([ref]$savePathRef, [ref]$formatRef)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Wordは保存完了を返しましたが、一時ファイルがありません。' }
}

function Invoke-MbWordExport {
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
        jobId = $JobId; state = 'running'; phase = 'starting'; message = 'Word出力を準備しています'; percent = 0
        currentStep = 0; totalSteps = $totalSteps; outputPath = ''; outputName = ''; outputDirectory = $OutputDirectory
        ownedWordPid = 0; ownedWordStartTimeUtc = ''; ownershipMode = ''; ownershipProven = $false; pageCount = 0
        startedAt = $startedAt; updatedAt = $startedAt; completedAt = ''; errorCode = ''
    }
    Write-MbWordStatus $StatusPath $status

    $word = $null; $document = $null; $selection = $null
    $documentClosed = $false; $canQuitCom = $false; $settingsApplied = $false; $ownershipProven = $false
    $ownPid = 0; $ownershipMode = ''; $temporaryPath = ''; $outputPath = ''
    $renderDirectory = Join-Path (Split-Path -Parent $StatusPath) 'rendered-word-images'
    $generatedImages = New-Object System.Collections.ArrayList
    $pidsBefore = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

    try {
        Test-MbWordCancellation $CancelPath
        if ($pidsBefore.Count -gt 0) { throw 'MB_WORD_RUNNING' }
        if (-not $script:MbWordPInvokeAvailable) { throw 'MB_WORD_OWNERSHIP_API_UNAVAILABLE' }
        if (-not (Test-Path -LiteralPath $OutputDirectory)) { [void](New-Item -ItemType Directory -Path $OutputDirectory -Force) }
        if (-not (Test-Path -LiteralPath $renderDirectory)) { [void](New-Item -ItemType Directory -Path $renderDirectory -Force) }
        $fontName = Resolve-MbWordBodyFont

        Set-MbWordStatusProgress $status $StatusPath 'starting-word' '安全確認のためWordを起動しています' 0 $totalSteps 3
        $word = New-Object -ComObject Word.Application
        $canQuitCom = $true
        $resolved = Resolve-MbOwnedWordProcess -Application $word -PidsBefore $pidsBefore
        $ownPid = [int]$resolved.Pid; $ownershipMode = [string]$resolved.Mode
        if ([bool]$resolved.ExistingConnection) { throw 'MB_CONNECTED_TO_EXISTING_WORD' }
        if ($ownPid -le 0) { throw 'MB_WORD_OWNERSHIP_UNRESOLVED' }
        $ownershipProven = $true
        $status.ownedWordPid = $ownPid; $status.ownershipMode = $ownershipMode; $status.ownershipProven = $true
        try { $status.ownedWordStartTimeUtc = (Get-Process -Id $ownPid -ErrorAction Stop).StartTime.ToUniversalTime().ToString('o') } catch { }
        Write-MbWordStatus $StatusPath $status

        $word.Visible = $false; $word.DisplayAlerts = 0
        try { $word.ScreenUpdating = $false } catch { }
        $settingsApplied = $true
        $documents = $null
        try { $documents = $word.Documents; $document = $documents.Add() } finally { Release-MbWordComObject $documents }

        # 文書ができるとActiveWindow経由でHwndを取得できる。PID差分で特定した所有PIDと
        # 一致した場合だけ「Hwndで所有を証明した」状態へ引き上げ、終了時の安全網を有効にする。
        # 一致しない場合は引き上げず、強制終了の対象にしない。
        if ($ownershipMode -ne 'Hwnd') {
            $hwndOwnedPid = Get-MbWordProcessId -Application $word
            if ($hwndOwnedPid -gt 0 -and $hwndOwnedPid -eq $ownPid) {
                $ownershipMode = 'Hwnd'
                $status.ownershipMode = $ownershipMode
                Write-MbWordStatus $StatusPath $status
            }
        }

        $selection = $word.Selection
        Set-MbWordDocumentDesign $document $fontName
        $pageSetup = $null
        try {
            $pageSetup = $document.PageSetup
            $pageSetup.TopMargin = 56.7; $pageSetup.BottomMargin = 56.7
            $pageSetup.LeftMargin = 56.7; $pageSetup.RightMargin = 56.7
        } finally { Release-MbWordComObject $pageSetup }

        Set-MbWordStatusProgress $status $StatusPath 'building-cover' '表紙と目次を作成しています' 0 $totalSteps 8
        Set-MbWordStyle $selection $document -63
        $selection.ParagraphFormat.Alignment = 1
        $selection.TypeText([string]$Project.title); $selection.TypeParagraph()
        Set-MbWordStyle $selection $document -75
        $selection.TypeText('操作マニュアル'); $selection.TypeParagraph()
        Set-MbWordStyle $selection $document -1
        $selection.ParagraphFormat.Alignment = 1
        $selection.TypeText((Get-Date -Format 'yyyy年M月d日')); $selection.TypeParagraph()
        $selection.TypeText('ManualBuilder'); $selection.TypeParagraph()
        $selection.InsertBreak(7)

        Set-MbWordStyle $selection $document -1
        $selection.ParagraphFormat.Alignment = 0; $selection.Font.Bold = $true; $selection.Font.Size = 18
        $selection.TypeText('目次'); $selection.TypeParagraph()
        Set-MbWordStyle $selection $document -1
        $selection.Font.Bold = $false; $selection.Font.Size = 10.5
        $toc = $null
        try { $toc = $document.TablesOfContents.Add($selection.Range, $true, 1, 2) } finally { Release-MbWordComObject $toc }
        Move-MbWordSelectionToEnd $selection $document
        $selection.TypeParagraph(); $selection.InsertBreak(7)

        $globalStep = 0; $expectedImages = 0
        $expectedOrderTokens = New-Object System.Collections.ArrayList
        $includedSheets = @(Get-MbWordContentSheets -Project $Project)
        if ($includedSheets.Count -eq 0) { throw 'Wordへ出力する手順がありません。' }
        for ($sheetIndex = 0; $sheetIndex -lt $includedSheets.Count; $sheetIndex++) {
            Test-MbWordCancellation $CancelPath
            $sheet = $includedSheets[$sheetIndex]
            [void]$expectedOrderTokens.Add([string]$sheet.name)
            Move-MbWordSelectionToEnd $selection $document
            Set-MbWordStyle $selection $document -2
            $selection.ParagraphFormat.Alignment = 0
            $selection.ParagraphFormat.PageBreakBefore = ($sheetIndex -gt 0)
            $selection.TypeText([string]$sheet.name)
            $heading = $null
            try { $heading = $selection.Paragraphs.Item(1); $heading.KeepWithNext = $true } finally { Release-MbWordComObject $heading }
            $selection.TypeParagraph()
            if ($sheet.PSObject.Properties.Name -contains 'summary' -and -not [string]::IsNullOrWhiteSpace([string]$sheet.summary)) {
                Add-MbWordTextParagraph $selection $document ([string]$sheet.summary) $true
            }

            $sheetSteps = @($sheet.steps)
            for ($stepIndex = 0; $stepIndex -lt $sheetSteps.Count; $stepIndex++) {
                Test-MbWordCancellation $CancelPath
                $step = $sheetSteps[$stepIndex]; $globalStep++
                $title = if ([string]::IsNullOrWhiteSpace([string]$step.title)) { '手順名未入力' } else { [string]$step.title }
                $stepHeadingText = [string]("手順 {0}　{1}" -f ($stepIndex + 1), $title)
                [void]$expectedOrderTokens.Add($stepHeadingText)
                Set-MbWordStatusProgress $status $StatusPath 'building-steps' ("手順 $globalStep / $totalSteps を作成しています") $globalStep $totalSteps ([Math]::Min(90, 10 + [int](80 * $globalStep / [Math]::Max(1, $totalSteps))))
                Move-MbWordSelectionToEnd $selection $document
                Set-MbWordStyle $selection $document -3
                $selection.ParagraphFormat.Alignment = 0
                $selection.ParagraphFormat.PageBreakBefore = $false
                $selection.TypeText($stepHeadingText)
                $stepHeading = $null
                try { $stepHeading = $selection.Paragraphs.Item(1); $stepHeading.KeepWithNext = $true } finally { Release-MbWordComObject $stepHeading }
                $selection.TypeParagraph()

                if (-not [string]::IsNullOrWhiteSpace([string]$step.description)) {
                    Add-MbWordTextParagraph $selection $document ([string]$step.description) $true
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$step.note)) {
                    Add-MbWordNote $selection $document ([string]$step.note) $fontName
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$step.imageId)) {
                    $sourcePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$step.imageId)
                    if (-not $sourcePath -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "画像が見つかりません: $($step.imageId)" }
                    $renderedPath = Join-Path $renderDirectory ("word-image-{0:d4}.png" -f $globalStep)
                    $imagePath = New-MbAnnotatedImage -SourcePath $sourcePath -Annotations @($step.annotations) -Crop $step.crop `
                        -DestinationPath $renderedPath -TargetDisplayWidth 600 -TargetDisplayHeight 680 -NumberFontName $fontName
                    if ($imagePath -eq $renderedPath) { [void]$generatedImages.Add($renderedPath) }
                    Move-MbWordSelectionToEnd $selection $document
                    $shape = $null; $imageParagraph = $null
                    try {
                        $selection.ParagraphFormat.Alignment = 1
                        $shape = $selection.InlineShapes.AddPicture($imagePath, $false, $true)
                        $naturalWidth = [double]$shape.Width; $naturalHeight = [double]$shape.Height
                        $shape.LockAspectRatio = -1
                        $scale = [Math]::Min(2.0, [Math]::Min(450.0 / $naturalWidth, 510.0 / $naturalHeight))
                        $shape.Width = [single]($naturalWidth * $scale)
                        try { $shape.AlternativeText = "手順 $($stepIndex + 1) の画面: $title" } catch { }
                        try {
                            $shape.Line.Visible = $true
                            $shape.Line.ForeColor.RGB = ConvertTo-MbWordBgr 210 214 220
                            $shape.Line.Weight = 0.75
                        } catch { }
                        $imageParagraph = $selection.Paragraphs.Item(1)
                        $imageParagraph.Alignment = 1
                        $imageParagraph.KeepTogether = $true; $imageParagraph.KeepWithNext = $false
                        $expectedImages++
                    } finally {
                        Release-MbWordComObject $imageParagraph
                        Release-MbWordComObject $shape
                    }
                    $selection.TypeParagraph()
                    $selection.ParagraphFormat.Alignment = 0
                }
            }
        }

        Test-MbWordCancellation $CancelPath
        $primaryFooter = $null; $footerRange = $null; $footerParagraph = $null; $footerParagraphRange = $null
        $fields = $null; $field = $null
        try {
            $primaryFooter = $document.Sections.Item(1).Footers.Item(1)
            $footerRange = $primaryFooter.Range
            $footerRange.ParagraphFormat.Alignment = 1
            $fields = $footerRange.Fields
            $field = $fields.Add($footerRange, 33)
            $footerParagraph = $primaryFooter.Range.Paragraphs.Item(1)
            $footerParagraph.Alignment = 1
            $footerParagraphRange = $footerParagraph.Range
            $footerParagraphRange.ParagraphFormat.Alignment = 1
        } finally {
            Release-MbWordComObject $field; Release-MbWordComObject $fields
            Release-MbWordComObject $footerParagraphRange; Release-MbWordComObject $footerParagraph
            Release-MbWordComObject $footerRange; Release-MbWordComObject $primaryFooter
        }
        [void]$document.Repaginate()
        if ($document.TablesOfContents.Count -gt 0) {
            $tocToUpdate = $null
            try { $tocToUpdate = $document.TablesOfContents.Item(1); [void]$tocToUpdate.Update() }
            finally { Release-MbWordComObject $tocToUpdate }
        }
        [void]$document.Fields.Update()
        [void]$document.Repaginate()
        $status.pageCount = [int]$document.ComputeStatistics(2)
        if ([int]$document.InlineShapes.Count -ne $expectedImages) { throw '出力画像数の自己検査に失敗しました。' }
        if ($globalStep -ne $totalSteps) { throw '出力手順数の自己検査に失敗しました。' }
        if ((Get-MbWordFirstNonEmptyParagraphText $document) -ne [string]$Project.title) { throw '表紙タイトルの自己検査に失敗しました。' }
        $documentText = [string]$document.Content.Text
        $orderCursor = 0
        foreach ($orderToken in @($expectedOrderTokens)) {
            $foundAt = $documentText.IndexOf([string]$orderToken, $orderCursor, [StringComparison]::Ordinal)
            if ($foundAt -lt 0) { throw "出力順の自己検査に失敗しました: $orderToken" }
            $orderCursor = $foundAt + ([string]$orderToken).Length
        }

        Set-MbWordStatusProgress $status $StatusPath 'saving' 'Wordファイルを保存しています' $globalStep $totalSteps 94
        $temporaryPath = Join-Path $OutputDirectory ('.ManualBuilder-' + $JobId + '.tmp.docx')
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        Save-MbWordDocument $document $temporaryPath
        $document.Close(0); $documentClosed = $true
        Test-MbWordCancellation $CancelPath
        $outputName = Get-MbSafeWordFileName -Name (([string]$Project.title) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss')) -Directory $OutputDirectory
        $outputPath = Join-Path $OutputDirectory $outputName
        [IO.File]::Move($temporaryPath, $outputPath); $temporaryPath = ''
        $status.state = 'finalizing'; $status.phase = 'finalizing'; $status.message = 'Wordを安全に終了しています'; $status.percent = 99
        $status.outputPath = $outputPath; $status.outputName = $outputName
        Write-MbWordStatus $StatusPath $status
    } catch {
        $message = [string]$_.Exception.Message
        if ($message -eq 'MB_EXPORT_CANCELLED') {
            $status.state = 'cancelled'; $status.phase = 'cancelled'; $status.message = 'Word作成を中止しました'; $status.errorCode = 'CANCELLED'
        } else {
            $status.state = 'failed'; $status.phase = 'failed'; $status.errorCode = $message
            $status.message = switch ($message) {
                'MB_WORD_RUNNING' { 'Wordが開いているため、安全のため作成を開始しませんでした。Wordを閉じて再実行するか、Excelで作成してください。' }
                'MB_CONNECTED_TO_EXISTING_WORD' { '既存のWordへ接続したため、安全のため作成を中止しました。Wordを閉じて再実行してください。' }
                'MB_WORD_OWNERSHIP_UNRESOLVED' { '作成用Wordの安全確認ができませんでした。' }
                'MB_WORD_OWNERSHIP_API_UNAVAILABLE' { 'Wordの所有確認に必要なWindows機能を利用できません。' }
                default { 'Wordファイルを作成できませんでした: ' + $message }
            }
        }
    } finally {
        try { if ($document -and -not $documentClosed) { $document.Close(0) } } catch { }
        Release-MbWordComObject $selection
        Release-MbWordComObject $document
        if ($word -and $canQuitCom) {
            if ($settingsApplied) { try { $word.ScreenUpdating = $true } catch { } }
            try { $word.Quit(0) } catch { }
        }
        Release-MbWordComObject $word
        $selection = $null; $document = $null; $word = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # Stop-ProcessはHwndで所有を証明したPIDにのみ許可する（WORD-OUTPUT-DESIGN-v0.12 の安全境界）。
        # PID差分モードは他人のWINWORDを指し得るため、未保存文書を守る目的で強制終了しない。
        if ($ownershipProven -and $ownPid -gt 0 -and $ownershipMode -eq 'Hwnd') {
            $deadline = (Get-Date).AddSeconds(10)
            do {
                if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { break }
                Start-Sleep -Milliseconds 500
            } while ((Get-Date) -lt $deadline)
            $ownedProcess = Get-Process -Id $ownPid -ErrorAction SilentlyContinue
            if ($ownedProcess -and $ownedProcess.ProcessName -eq 'WINWORD') { Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue }
        }
        if ($temporaryPath -and (Test-Path -LiteralPath $temporaryPath)) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
        foreach ($imagePath in @($generatedImages)) {
            if (Test-Path -LiteralPath $imagePath) { Remove-Item -LiteralPath $imagePath -Force -ErrorAction SilentlyContinue }
        }
    }
    if ($status.state -eq 'finalizing') {
        $status.state = 'completed'; $status.phase = 'completed'; $status.message = 'Wordファイルを作成しました'; $status.percent = 100
    }
    $status.completedAt = [DateTime]::UtcNow.ToString('o')
    Write-MbWordStatus $StatusPath $status
    return $status
}

Export-ModuleMember -Function @(
    'Get-MbSafeWordFileName',
    'Invoke-MbWordExport'
)
