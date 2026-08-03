# Phase 1 video attachment storage test.
# 動画はPowerPoint出力にだけ埋め込む。ここでは保存・重複排除・付け外し・後片付けを確かめる。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-VideoTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function New-MbTestMp4 {
    param([int]$FillerLength = 512, [byte]$Filler = 0x21)
    # 先頭12バイトだけ本物のmp4と同じ形（ftypボックス）にする。中身は判定に使わない。
    $header = [byte[]]@(0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0x69, 0x73, 0x6F, 0x6D)
    $body = New-Object byte[] $FillerLength
    for ($i = 0; $i -lt $body.Length; $i++) { $body[$i] = $Filler }
    return @($header + $body)
}

function New-MbTestWebm {
    $header = [byte[]]@(0x1A, 0x45, 0xDF, 0xA3)
    $body = New-Object byte[] 256
    return @($header + $body)
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    $project = New-MbProject
    $project = Save-MbProject -Project $project -Path $projectPath
    $sheetId = [string]$project.sheets[0].id
    $step = Add-MbStep -Project $project -SheetId $sheetId
    $stepId = [string]$step.id
    $project = Save-MbProject -Project $project -Path $projectPath

    Assert-Mb (@($project.videos).Count -eq 0) '新しいマニュアルに動画は無い'
    Assert-Mb ($project.sheets[0].steps[0].PSObject.Properties.Name -contains 'videoId') '手順にvideoIdを持つ'

    # --- 形式の判定 ---
    Assert-Mb ((Get-MbVideoKind -Bytes (New-MbTestMp4)).MimeType -eq 'video/mp4') 'mp4を判定できる'
    Assert-Mb ((Get-MbVideoKind -Bytes (New-MbTestWebm)).MimeType -eq 'video/webm') 'webmを判定できる'
    Assert-Mb ($null -eq (Get-MbVideoKind -Bytes ([byte[]]@(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)))) '対応しない形式は受け付けない'

    $rejected = $false
    try { [void](Add-MbVideoAsset -Project $project -ProjectPath $projectPath -Bytes ([byte[]]@(1, 2, 3, 4)) -DurationSec 1) }
    catch { $rejected = $true }
    Assert-Mb $rejected 'mp4・webm以外は保存しない'

    # --- 添付 ---
    $result = Set-MbStepVideo -Project $project -ProjectPath $projectPath -StepId $stepId -Bytes (New-MbTestMp4) -DurationSec 12.34
    $project = Save-MbProject -Project $project -Path $projectPath
    $videoId = [string]$result.Video.id
    Assert-Mb ($videoId -match '^video-[a-f0-9]{32}$') '動画IDを採番する'
    Assert-Mb ([string]$project.sheets[0].steps[0].videoId -eq $videoId) '手順へ動画を紐づける'
    Assert-Mb ([double]$result.Video.durationSec -eq 12.3) '動画の長さを0.1秒単位で持つ'

    $videoPath = Get-MbVideoFilePath -Project $project -ProjectPath $projectPath -VideoId $videoId
    Assert-Mb (Test-Path -LiteralPath $videoPath -PathType Leaf) '動画をvideosフォルダーへ保存する'
    Assert-Mb ((Split-Path -Leaf (Split-Path -Parent $videoPath)) -eq 'videos') '画像とは別のフォルダーへ保存する'
    Assert-Mb ((Get-MbVideoTotalBytes -Project $project) -eq [long]$result.Video.byteLength) '合計サイズを集計できる'

    # --- 重複排除 ---
    $step2 = Add-MbStep -Project $project -SheetId $sheetId
    $result2 = Set-MbStepVideo -Project $project -ProjectPath $projectPath -StepId ([string]$step2.id) -Bytes (New-MbTestMp4) -DurationSec 12.34
    $project = Save-MbProject -Project $project -Path $projectPath
    Assert-Mb ([string]$result2.Video.id -eq $videoId) '同じ内容の動画は使い回す'
    Assert-Mb (@($project.videos).Count -eq 1) '同じ動画のファイルを二重に持たない'

    # --- 差し替え ---
    $result3 = Set-MbStepVideo -Project $project -ProjectPath $projectPath -StepId $stepId -Bytes (New-MbTestMp4 -Filler 0x55) -DurationSec 3
    $project = Save-MbProject -Project $project -Path $projectPath
    Assert-Mb ([string]$result3.Video.id -ne $videoId) '別の動画へ差し替えられる'
    Assert-Mb ($null -eq $result3.RemovedPath) '他の手順が使っている動画は消さない'
    Assert-Mb (@($project.videos).Count -eq 2) '差し替え後も元の動画が残る'

    # --- 外す ---
    $removed = Remove-MbStepVideo -Project $project -ProjectPath $projectPath -StepId $stepId
    $project = Save-MbProject -Project $project -Path $projectPath
    Assert-Mb ($null -eq [string]$project.sheets[0].steps[0].videoId -or '' -eq [string]$project.sheets[0].steps[0].videoId) '手順から動画を外せる'
    Assert-Mb ($null -ne $removed.RemovedPath) '参照が無くなった動画は削除対象になる'
    Assert-Mb (@($project.videos).Count -eq 1) '外した動画をプロジェクトから取り除く'

    # --- 未参照の後片付け ---
    $orphan = Add-MbVideoAsset -Project $project -ProjectPath $projectPath -Bytes (New-MbTestWebm) -DurationSec 1
    Assert-Mb (@($project.videos).Count -eq 2) '手順に紐づかない動画も一旦は登録される'
    $orphanPaths = @(Remove-MbUnreferencedVideos -Project $project -ProjectPath $projectPath)
    Assert-Mb ($orphanPaths.Count -eq 1) '未参照の動画を洗い出せる'
    Assert-Mb (@($project.videos).Count -eq 1) '未参照の動画をプロジェクトから外す'
    Assert-Mb ([string]$orphanPaths[0] -eq (Join-Path (Join-Path $testRoot 'videos') ([string]$orphan.Video.fileName))) '削除対象のパスを返す'

    # --- 保存できる形式かどうか ---
    $project = Save-MbProject -Project $project -Path $projectPath
    $reloaded = Get-MbProject -Path $projectPath
    Assert-Mb (@($reloaded.videos).Count -eq 1) '動画の情報がproject.jsonへ保存される'
    Assert-Mb ([string]$reloaded.sheets[0].steps[1].videoId -eq [string]$reloaded.videos[0].id) '保存後も手順と動画の紐づけが残る'

    # --- 上限 ---
    $tooLarge = $false
    try { [void](Add-MbVideoAsset -Project $reloaded -ProjectPath $projectPath -Bytes (New-MbTestMp4 -FillerLength (31 * 1024 * 1024)) -DurationSec 1) }
    catch { $tooLarge = $true }
    Assert-Mb $tooLarge '30MBを超える動画は受け付けない'

    $tooLong = $false
    try { [void](Add-MbVideoAsset -Project $reloaded -ProjectPath $projectPath -Bytes (New-MbTestMp4 -Filler 0x77) -DurationSec 7200) }
    catch { $tooLong = $true }
    Assert-Mb $tooLong '1時間を超える長さは受け付けない'

    # --- 旧形式のproject.jsonとの互換 ---
    $legacy = [pscustomobject]@{
        schemaVersion   = 1
        revision        = 3
        id              = 'project-' + [guid]::NewGuid().ToString('N')
        title           = '旧形式'
        selectedSheetId = $null
        sheets          = @([pscustomobject]@{
            id    = 'sheet-' + [guid]::NewGuid().ToString('N')
            name  = '手順1'
            steps = @([pscustomobject]@{ id = 'step-' + [guid]::NewGuid().ToString('N'); title = '旧手順' })
        })
        images          = @()
    }
    $legacyPath = Join-Path $testRoot 'legacy.json'
    $legacy = Save-MbProject -Project $legacy -Path $legacyPath
    Assert-Mb ($legacy.PSObject.Properties.Name -contains 'videos') '動画欄が無い旧データにも動画欄を補える'
    Assert-Mb (@($legacy.videos).Count -eq 0) '旧データの読み込みで動画は空になる'
    Assert-Mb ($legacy.sheets[0].steps[0].PSObject.Properties.Name -contains 'videoId') '旧データの手順にもvideoIdを補える'

    Write-Host ''
    Write-Host 'Video attachment tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
