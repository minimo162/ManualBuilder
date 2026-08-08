# ManualBuilder explicit Excel COM export test. Close Excel before running.

[CmdletBinding()]
param([ValidateRange(30, 300)][int]$TimeoutSeconds = 120)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$workerPath = Join-Path $repoRoot 'src\Export-ManualBuilderExcel.ps1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-ExcelExportTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'
$outputDirectory = Join-Path $testRoot 'output'
$statusPath = Join-Path $testRoot 'status.json'
$cancelPath = Join-Path $testRoot 'cancel.requested'
$jobId = 'test-' + [guid]::NewGuid().ToString('N')
$worker = $null
Add-Type -AssemblyName System.Drawing

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function New-MbTestPngBytes {
    param(
        [Drawing.Color]$Color,
        [ValidateRange(100, 4000)][int]$Width = 1920,
        [ValidateRange(100, 4000)][int]$Height = 1080,
        [ValidateSet('Png', 'Jpeg')][string]$Format = 'Png'
    )
    $bitmap = New-Object Drawing.Bitmap $Width, $Height
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $stream = New-Object IO.MemoryStream
    $font = $null
    try {
        $graphics.Clear($Color)
        $graphics.FillRectangle(
            [Drawing.Brushes]::White,
            [int]($Width * 0.06), [int]($Height * 0.11),
            [int]($Width * 0.88), [int]($Height * 0.78)
        )
        $font = New-Object Drawing.Font 'Arial', 48
        $graphics.DrawString('ManualBuilder Excel Test', $font, [Drawing.Brushes]::Navy, [single]($Width * 0.12), [single]($Height * 0.43))
        $imageFormat = if ($Format -eq 'Jpeg') { [Drawing.Imaging.ImageFormat]::Jpeg } else { [Drawing.Imaging.ImageFormat]::Png }
        $bitmap.Save($stream, $imageFormat)
        return $stream.ToArray()
    } finally {
        if ($font) { $font.Dispose() }
        $graphics.Dispose()
        $bitmap.Dispose()
        $stream.Dispose()
    }
}

try {
    if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
        Write-Host 'Excelをすべて閉じてから、もう一度このテストを実行してください。' -ForegroundColor Yellow
        exit 2
    }

    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
    $project = Get-MbProject -Path $projectPath
    Set-MbProjectTitle -Project $project -Title '経費精算システム 操作マニュアル'
    Rename-MbSheet -Project $project -SheetId $project.sheets[0].id -Name '経費/申請:国内'
    $first = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id `
        -Bytes (New-MbTestPngBytes -Color ([Drawing.Color]::LightSteelBlue)) -Source file
    Update-MbStep -Project $project -StepId $first.Step.id -Title '=不正な式ではなくタイトル' `
        -Description '申請メニューを選択し、必要事項を入力します。' -Note '赤枠で入力欄を示しています。'
    Set-MbStepAnnotations -Project $project -StepId $first.Step.id -AnnotationsJson '[{"id":"annotation-00000000000000000000000000000031","type":"rect","x1":0.1,"y1":0.1,"x2":0.55,"y2":0.55,"label":0},{"id":"annotation-00000000000000000000000000000032","type":"number","x1":0.2,"y1":0.25,"x2":0.2,"y2":0.25,"label":1}]'
    [void](Set-MbStepResultImage -Project $project -ProjectPath $projectPath -StepId $first.Step.id `
        -Bytes (New-MbTestPngBytes -Color ([Drawing.Color]::LightCyan)) -Source file)
    Assert-Mb ([string]$first.Step.imageLayout -eq 'side-by-side') '操作後画像つきの手順を左右比較で準備する'

    # 現行形式の操作証拠から作った手順を、そのままExcelへ出せることも同じ試験で確認する。
    # 旧形式との互換分岐は持たず、証拠形式v2だけを正とする。
    $sourceSessionId = 'record-' + [guid]::NewGuid().ToString('N')
    $sourceEvidenceId = 'evidence-' + [guid]::NewGuid().ToString('N')
    $sessionRoot = Join-Path (Join-Path $testRoot 'evidence') $sourceSessionId
    $sessionImages = Join-Path $sessionRoot 'images'
    [void](New-Item -ItemType Directory -Path $sessionImages -Force)
    [IO.File]::WriteAllBytes((Join-Path $sessionImages ($sourceEvidenceId + '.jpg')), (New-MbTestPngBytes -Color ([Drawing.Color]::LightSteelBlue) -Format Jpeg))
    [IO.File]::WriteAllLines((Join-Path $sessionRoot 'evidence-ledger.jsonl'), @(
        ([ordered]@{ recordType='capture-start'; formatVersion=2; sessionId=$sourceSessionId; completeness='no-known-gaps' } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='operation'; id=$sourceEvidenceId; sessionId=$sourceSessionId; kind='click'; timeMs=1000; image=($sourceEvidenceId + '.jpg') } | ConvertTo-Json -Compress),
        ([ordered]@{ recordType='capture-end'; formatVersion=2; sessionId=$sourceSessionId; operationCount=1; reason='stopped'; completeness='no-known-gaps'; warning='' } | ConvertTo-Json -Compress)
    ), [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $sessionRoot 'transformations.jsonl'),
        (([ordered]@{ recordType='transformation'; sessionId=$sourceSessionId; proposalId='local-test'; evidenceIds=@($sourceEvidenceId); accepted=$true; reviewed=$true; reason='1件のクリックを1手順として採用'; decisionSource='user-review' } | ConvertTo-Json -Compress) + [Environment]::NewLine),
        [Text.UTF8Encoding]::new($false))
    $project.evidenceSessions = @([pscustomobject]@{
        id=$sourceSessionId; operationCount=1; undoneCount=0; formatVersion=2
        captureCompleteness='no-known-gaps'; captureWarning=''; retention='project-lifetime'
        ledgerFile=('evidence/' + $sourceSessionId + '/evidence-ledger.jsonl')
        imageDirectory=('evidence/' + $sourceSessionId + '/images')
        decisionsFile=('evidence/' + $sourceSessionId + '/transformations.jsonl')
        importedAt=[DateTime]::UtcNow.ToString('o')
    })
    [void](Set-MbStepCapture -Project $project -StepId $first.Step.id -Kind 'recorded-local' `
        -SourceSessionId $sourceSessionId -EvidenceIdsJson (ConvertTo-Json @($sourceEvidenceId) -Compress) `
        -SourceOperationCount 1 -TransformationReason '1件のクリックを1手順として採用しました。')
    Assert-Mb ([string]$first.Step.capture.sourceSessionId -eq $sourceSessionId -and
        @($first.Step.capture.evidenceIds).Count -eq 1) 'Excelへ出す手順から現行形式の操作証拠をたどれる'

    $wide = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id `
        -Bytes (New-MbTestPngBytes -Color ([Drawing.Color]::PaleGreen) -Width 1920 -Height 500) -Source file
    Update-MbStep -Project $project -StepId $wide.Step.id -Title '横長の画面を確認する' `
        -Description '横長画像ではカード下の余白を抑えます。' -Note ''

    $portrait = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id `
        -Bytes (New-MbTestPngBytes -Color ([Drawing.Color]::LightGoldenrodYellow) -Width 1080 -Height 1920) -Source file
    Update-MbStep -Project $project -StepId $portrait.Step.id -Title '縦長の画面を確認する' `
        -Description '縦長画像ではカードを広げ、画像を読みやすくします。' -Note ''

    $longDescription = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id `
        -Bytes (New-MbTestPngBytes -Color ([Drawing.Color]::Lavender) -Width 1920 -Height 1080) -Source file
    Update-MbStep -Project $project -StepId $longDescription.Step.id -Title '長い説明を確認する' `
        -Description ('この説明はExcelの結合セル内で末尾が欠けないことを確認するための文章です。' * 30) -Note ''

    $secondSheet = Add-MbSheet -Project $project
    Rename-MbSheet -Project $project -SheetId $secondSheet.id -Name '経費/申請:国内'
    $second = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $secondSheet.id `
        -Bytes (New-MbTestPngBytes -Color ([Drawing.Color]::MistyRose)) -Source file
    Update-MbStep -Project $project -StepId $second.Step.id -Title '内容を確認する' `
        -Description '入力内容を確認して申請ボタンを押します。' -Note ''
    [void](Save-MbProject -Project $project -Path $projectPath)

    $arguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA',
        '-File', ('"' + $workerPath + '"'), '-ProjectPath', ('"' + $projectPath + '"'),
        '-OutputDirectory', ('"' + $outputDirectory + '"'), '-StatusPath', ('"' + $statusPath + '"'),
        '-CancelPath', ('"' + $cancelPath + '"'), '-JobId', $jobId
    )
    $worker = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -PassThru -WindowStyle Hidden
    if (-not $worker.WaitForExit($TimeoutSeconds * 1000)) {
        [IO.File]::WriteAllText($cancelPath, 'timeout', (New-Object Text.UTF8Encoding($false)))
        [void]$worker.WaitForExit(15000)
        throw "Excel出力が${TimeoutSeconds}秒以内に完了しませんでした。"
    }
    Assert-Mb (Test-Path -LiteralPath $statusPath -PathType Leaf) 'Excel出力が結果ファイルを返す'
    $status = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$status.state -eq 'completed') ("Excel出力が完了する: " + [string]$status.message)
    Assert-Mb ([bool]$status.ownershipProven -and [int]$status.ownedExcelPid -gt 0) '出力用Excelの所有PIDを確認する'
    Assert-Mb (@($status.sheetNameMappings).Count -eq 2) '出力結果へシート名変換一覧を記録する'
    Assert-Mb (Test-Path -LiteralPath ([string]$status.outputPath) -PathType Leaf) 'xlsxファイルを生成する'

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead([string]$status.outputPath)
    try {
        $workbookEntry = $archive.GetEntry('xl/workbook.xml')
        Assert-Mb ($null -ne $workbookEntry) 'xlsx内部のworkbook.xmlを確認できる'
        $reader = New-Object IO.StreamReader($workbookEntry.Open(), [Text.Encoding]::UTF8)
        try { $workbookXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        Assert-Mb ($workbookXml -match '目次') '目次シートを出力する'
        Assert-Mb ($workbookXml -match '経費・申請・国内' -and $workbookXml -match '経費・申請・国内 \(2\)') '禁止文字と重複を処理したシート名を出力する'
        Assert-Mb ($workbookXml -match '_xlnm\.Print_Titles') '印刷時に目次・セクション見出しを繰り返す'
        $sharedStringsEntry = $archive.GetEntry('xl/sharedStrings.xml')
        Assert-Mb ($null -ne $sharedStringsEntry) 'Excel内の表示文言を確認できる'
        $sharedStringsReader = New-Object IO.StreamReader($sharedStringsEntry.Open(), [Text.Encoding]::UTF8)
        try { $sharedStringsXml = $sharedStringsReader.ReadToEnd() } finally { $sharedStringsReader.Dispose() }
        Assert-Mb ($sharedStringsXml -match '使い方' -and $sharedStringsXml -match 'シート名をクリックして開き') '目次にマニュアルの読み方を表示する'
        Assert-Mb ($sharedStringsXml -match '開く　経費・申請・国内' -and $sharedStringsXml -match '← 目次へ戻る') '目次と各セクションの往復操作を明示する'
        Assert-Mb ($sharedStringsXml -match '>操作<' -and $sharedStringsXml -match 'ポイント・注意') '操作本文と注意情報を明確に分ける'
        Assert-Mb ($sharedStringsXml -match '全 4 手順' -and $sharedStringsXml -match '全 1 手順') '各セクションに総手順数を表示する'
        $drawingEntries = @($archive.Entries | Where-Object { $_.FullName -like 'xl/drawings/drawing*.xml' })
        $mediaEntries = @($archive.Entries | Where-Object { $_.FullName -like 'xl/media/*' })
        Assert-Mb ($drawingEntries.Count -ge 2 -and $mediaEntries.Count -ge 5) '各シートへスクリーンショットを埋め込む'
        $hasCompactCard = $false
        $hasExpandedCard = $false
        $hasLandscapeFitToWidth = $false
        $hasPageNumberFooter = $false
        $worksheetEntries = @($archive.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' })
        foreach ($worksheetEntry in $worksheetEntries) {
            $worksheetReader = New-Object IO.StreamReader($worksheetEntry.Open(), [Text.Encoding]::UTF8)
            try { $worksheetXml = $worksheetReader.ReadToEnd() } finally { $worksheetReader.Dispose() }
            if ($worksheetXml -match '<pageSetup[^>]*orientation="landscape"' -and $worksheetXml -match 'fitToPage="1"') { $hasLandscapeFitToWidth = $true }
            if ($worksheetXml -match '&amp;P / &amp;N') { $hasPageNumberFooter = $true }
            foreach ($merge in [regex]::Matches($worksheetXml, '<mergeCell ref="A(\d+):G(\d+)"')) {
                $span = ([int]$merge.Groups[2].Value - [int]$merge.Groups[1].Value) + 1
                if ($span -eq 7) { $hasCompactCard = $true }
                if ($span -ge 26) { $hasExpandedCard = $true }
            }
        }
        Assert-Mb $hasCompactCard '横長画像のExcelカードを7行へ縮める'
        Assert-Mb $hasExpandedCard '縦長画像または長文のExcelカードを26行以上へ広げる'
        Assert-Mb $hasLandscapeFitToWidth '手順シートを横向き1ページ幅で印刷できる'
        Assert-Mb $hasPageNumberFooter '印刷時のフッターにページ番号を表示する'
        $hasReadableImageExtent = $false
        foreach ($drawingEntry in $drawingEntries) {
            $drawingReader = New-Object IO.StreamReader($drawingEntry.Open(), [Text.Encoding]::UTF8)
            try { $drawingXml = $drawingReader.ReadToEnd() } finally { $drawingReader.Dispose() }
            foreach ($extent in [regex]::Matches($drawingXml, '<(?:xdr|a):ext[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"')) {
                if ([int64]$extent.Groups[1].Value -ge 6800000 -and [int64]$extent.Groups[2].Value -ge 3800000) {
                    $hasReadableImageExtent = $true
                    break
                }
            }
            if ($hasReadableImageExtent) { break }
        }
        Assert-Mb $hasReadableImageExtent '1920x1080画像を約760x428px相当で配置する'
    } finally {
        $archive.Dispose()
    }

    Start-Sleep -Seconds 2
    Assert-Mb (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -eq 0) '出力後にExcelプロセスを残さない'
    Write-Host ''
    Write-Host 'Excel COM export test passed.' -ForegroundColor Cyan
} finally {
    if ($worker -and -not $worker.HasExited) { Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue }
    if ($worker) { $worker.Dispose() }
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
