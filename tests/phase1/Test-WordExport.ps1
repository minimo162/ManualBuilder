# ManualBuilder explicit Word COM export test. Close Word before running.

[CmdletBinding()]
param([ValidateRange(30, 300)][int]$TimeoutSeconds = 120)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$workerPath = Join-Path $repoRoot 'src\Export-ManualBuilderWord.ps1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-WordExportTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'
$outputDirectory = Join-Path $testRoot 'output'
$statusPath = Join-Path $testRoot 'status.json'
$cancelPath = Join-Path $testRoot 'cancel.requested'
$jobId = 'word-test-' + [guid]::NewGuid().ToString('N')
$worker = $null
Add-Type -AssemblyName System.Drawing

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function New-MbWordTestPngBytes {
    param([Drawing.Color]$Color, [string]$Caption)
    $bitmap = New-Object Drawing.Bitmap 1920, 1080
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $stream = New-Object IO.MemoryStream
    $font = $null
    try {
        $graphics.Clear($Color)
        $graphics.FillRectangle([Drawing.Brushes]::White, 120, 120, 1680, 840)
        $font = New-Object Drawing.Font 'Arial', 48
        $graphics.DrawString($Caption, $font, [Drawing.Brushes]::Navy, 240, 470)
        $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    } finally {
        if ($font) { $font.Dispose() }
        $graphics.Dispose(); $bitmap.Dispose(); $stream.Dispose()
    }
}

try {
    if (@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Count -gt 0) {
        Write-Host 'Wordをすべて閉じてから、もう一度このテストを実行してください。' -ForegroundColor Yellow
        exit 2
    }
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
    $project = Get-MbProject -Path $projectPath
    Set-MbProjectTitle $project '経費精算システム 操作マニュアル'
    Rename-MbSheet $project $project.sheets[0].id 'ログイン'
    $first = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $project.sheets[0].id `
        -Bytes (New-MbWordTestPngBytes ([Drawing.Color]::LightSteelBlue) 'ManualBuilder Word Test 1') -Source file
    Update-MbStep $project $first.Step.id 'ログイン画面を開く' 'ブラウザーから対象システムを開きます。' '社外からはVPNが必要です。'
    Set-MbStepImageEdits -Project $project -StepId $first.Step.id `
        -AnnotationsJson '[{"id":"annotation-00000000000000000000000000000041","type":"rect","x1":0.1,"y1":0.1,"x2":0.55,"y2":0.55,"label":0},{"id":"annotation-00000000000000000000000000000042","type":"number","x1":0.2,"y1":0.25,"x2":0.2,"y2":0.25,"label":1}]' `
        -CropJson '{"x":0.05,"y":0.05,"width":0.9,"height":0.9}'
    $secondSheet = Add-MbSheet $project
    Rename-MbSheet $project $secondSheet.id '申請内容の確認'
    $second = Add-MbImageStep -Project $project -ProjectPath $projectPath -SheetId $secondSheet.id `
        -Bytes (New-MbWordTestPngBytes ([Drawing.Color]::MistyRose) 'ManualBuilder Word Test 2') -Source file
    Update-MbStep $project $second.Step.id '申請内容を確認する' '入力内容を確認して申請ボタンを押します。' ''
    $emptySheet = Add-MbSheet $project
    Rename-MbSheet $project $emptySheet.id '出力しない空シート'
    [void](Save-MbProject $project $projectPath)

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
        throw "Word出力が${TimeoutSeconds}秒以内に完了しませんでした。"
    }
    Assert-Mb (Test-Path $statusPath -PathType Leaf) 'Word出力が結果ファイルを返す'
    $status = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$status.state -eq 'completed') ('Word出力が完了する: ' + [string]$status.message)
    Assert-Mb ([bool]$status.ownershipProven -and [int]$status.ownedWordPid -gt 0) '出力用Wordの所有PIDを確認する'
    Assert-Mb ([int]$status.pageCount -ge 3) '表紙・目次・本文のページ数を記録する'
    Assert-Mb (Test-Path ([string]$status.outputPath) -PathType Leaf) 'docxファイルを生成する'

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead([string]$status.outputPath)
    try {
        $documentEntry = $archive.GetEntry('word/document.xml')
        Assert-Mb ($null -ne $documentEntry) 'docx内部のdocument.xmlを確認できる'
        $reader = New-Object IO.StreamReader($documentEntry.Open(), [Text.Encoding]::UTF8)
        try { $documentXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        Assert-Mb ($documentXml -match '経費精算システム 操作マニュアル') '表紙タイトルを出力する'
        Assert-Mb ($documentXml -match 'TOC' -and $documentXml -match 'ログイン' -and $documentXml -match '申請内容の確認') '目次とシート見出しを出力する'
        Assert-Mb ($documentXml -notmatch '出力しない空シート') '空シートを本文と目次から除外する'
        Assert-Mb ($documentXml -match '手順 1' -and $documentXml -match 'ログイン画面を開く') '手順見出しと説明を出力する'
        Assert-Mb ($documentXml -match '補足' -and $documentXml -match 'VPN') '補足欄を出力する'
        $mediaEntries = @($archive.Entries | Where-Object { $_.FullName -like 'word/media/*' })
        Assert-Mb ($mediaEntries.Count -eq 2) '編集済みスクリーンショットを各手順へ埋め込む'
        Assert-Mb ($documentXml -match '手順 1 の画面') '画像へ代替テキストを設定する'
        Assert-Mb ($documentXml -match 'cx="5715000"') '1920x1080画像をWord本文幅450ptで配置する'
        $footerXml = ''
        foreach ($footerEntry in @($archive.Entries | Where-Object { $_.FullName -like 'word/footer*.xml' })) {
            $footerReader = New-Object IO.StreamReader($footerEntry.Open(), [Text.Encoding]::UTF8)
            try { $footerXml += $footerReader.ReadToEnd() } finally { $footerReader.Dispose() }
        }
        Assert-Mb ($footerXml -match 'PAGE' -and $footerXml -match '<w:jc w:val="center"') 'フッターのページ番号を中央揃えにする'
    } finally { $archive.Dispose() }
    Start-Sleep -Seconds 2
    Assert-Mb (@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Count -eq 0) '出力後にWordプロセスを残さない'
    Write-Host ''
    Write-Host 'Word COM export test passed.' -ForegroundColor Cyan
} finally {
    if ($worker -and -not $worker.HasExited) { Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue }
    if ($worker) { $worker.Dispose() }
    if (Test-Path $testRoot) { Remove-Item $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
