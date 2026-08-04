# 録画から切り出した場面が、Copilotの下書きを経て保存済み手順になるまでを検査する。
# 実際の動画・HttpListener・Microsoft 365 Copilotは使わず、JPEGと固定JSONで主経路を通す。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$srcRoot = Join-Path $repoRoot 'src'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-VideoToManual-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'
$errors = New-Object 'System.Collections.Generic.List[string]'

Import-Module (Join-Path $srcRoot 'ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.CopilotServer.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.Copilot.psm1') -Force
Import-Module (Join-Path $srcRoot 'ManualBuilder.CopilotJob.psm1') -Force

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

function New-MbTestJpeg {
    param(
        [Parameter(Mandatory = $true)][Drawing.Color]$Background,
        [Parameter(Mandatory = $true)][Drawing.Color]$Button
    )

    $bitmap = New-Object Drawing.Bitmap 96, 64
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $stream = New-Object IO.MemoryStream
    try {
        $graphics.Clear($Background)
        $brush = New-Object Drawing.SolidBrush($Button)
        try { $graphics.FillRectangle($brush, 28, 22, 40, 20) } finally { $brush.Dispose() }
        $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Jpeg)
        return $stream.ToArray()
    } finally {
        $stream.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    Add-Type -AssemblyName System.Drawing

    $project = New-MbProject
    $project.title = '録画から作る申請手順'
    $project = Save-MbProject -Project $project -Path $projectPath
    $sheetId = [string]$project.selectedSheetId

    # A -> B -> A。元の画面へ戻る場面も操作結果として必要なので、3手順を残す。
    # 1件目と3件目は同じJPEGを使い、画像実体だけを共有する。
    $sceneA = New-MbTestJpeg -Background ([Drawing.Color]::White) -Button ([Drawing.Color]::RoyalBlue)
    $sceneB = New-MbTestJpeg -Background ([Drawing.Color]::WhiteSmoke) -Button ([Drawing.Color]::SeaGreen)
    $first = Import-MbVideoScene -Project $project -ProjectPath $projectPath -SheetId $sheetId `
        -Bytes $sceneA -TimeMs 1000 -RectJson '{"x1":0.28,"y1":0.30,"x2":0.72,"y2":0.70}' -SkipOcr
    $second = Import-MbVideoScene -Project $project -ProjectPath $projectPath -SheetId $sheetId `
        -Bytes $sceneB -TimeMs 2500 -RectJson '{"x1":0.20,"y1":0.20,"x2":0.55,"y2":0.55}' -SkipOcr
    $third = Import-MbVideoScene -Project $project -ProjectPath $projectPath -SheetId $sheetId `
        -Bytes $sceneA -TimeMs 4000 -RectJson '{"x1":0.40,"y1":0.35,"x2":0.80,"y2":0.75}' -SkipOcr

    Add-Result ([string]$first.status -eq 'added' -and [string]$second.status -eq 'added') '異なる場面を手順として取り込める'
    Add-Result ([string]$third.status -eq 'added') '元の画面へ戻った場面も別の手順として残す'
    Add-Result (@($project.sheets[0].steps).Count -eq 3) 'A→B→Aを3件の手順として並べる'
    Add-Result (@($project.images).Count -eq 2) '同じ場面Aの画像実体を重複保存しない'

    if (@($project.sheets[0].steps).Count -eq 3) {
        $stepsBeforeSave = @($project.sheets[0].steps)
        Add-Result ([string]$stepsBeforeSave[0].imageId -eq [string]$stepsBeforeSave[2].imageId) '1件目と3件目で同じ画像を共有する'
        Add-Result ([string]$stepsBeforeSave[0].imageId -ne [string]$stepsBeforeSave[1].imageId) '異なる場面Bは別の画像を参照する'
    }

    $project = Save-MbProject -Project $project -Path $projectPath
    $project = Get-MbProject -Path $projectPath
    $steps = @($project.sheets[0].steps)
    Add-Result ($steps.Count -eq 3) '保存後も3件の手順が残る'

    if ($steps.Count -eq 3) {
        Add-Result ([int]$steps[0].capture.videoTimeMs -eq 1000 -and
            [int]$steps[1].capture.videoTimeMs -eq 2500 -and
            [int]$steps[2].capture.videoTimeMs -eq 4000) '録画内の時刻を手順ごとに保存する'
        Add-Result (@($steps | Where-Object { @($_.annotations).Count -eq 1 }).Count -eq 3) '各手順に操作位置の赤枠を保存する'
        Add-Result (@($steps | Where-Object { [string]$_.capture.kind -eq 'video-scene' }).Count -eq 3) '録画由来の手順として識別できる'
    }

    $imageDirectory = Join-Path $testRoot 'images'
    Add-Result (@(Get-ChildItem -LiteralPath $imageDirectory -File).Count -eq 2) '共有した画像を2ファイルだけ保存する'

    # 保存したプロジェクトを、Copilotへ渡す平坦な一覧とパケットへ変換する。
    $copilotSteps = @(Get-MbCopilotStepList -Project $project)
    Add-Result ($copilotSteps.Count -eq 3) '録画由来の3手順がCopilotの対象になる'
    $packets = Get-MbCopilotPackets -Steps $copilotSteps -StepsPerPacket 2
    Add-Result (@($packets).Count -eq 2) 'Copilotの対象を指定件数で分割する'
    if (@($packets).Count -eq 2) {
        Add-Result (@($packets[0]).Count -eq 2 -and @($packets[1]).Count -eq 1) 'Copilotパケットの端数を残す'
    }

    # 実Copilotの代わりに、依頼したidをそのまま返す固定JSONを使う。
    $answerSteps = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $copilotSteps.Count; $i++) {
        [void]$answerSteps.Add([pscustomobject]@{
            id          = [string]$copilotSteps[$i].id
            keep        = $true
            title       = ('録画手順' + ($i + 1))
            description = ('画面の操作' + ($i + 1) + 'を行います。')
            note        = ''
            confident   = $true
            reason      = ''
        })
    }
    $answerJson = ([pscustomobject]@{ steps = @($answerSteps) } | ConvertTo-Json -Depth 6 -Compress)
    $answer = Get-MbStepAnswerJson -Text ("回答です。`n" + $answerJson + "`nMB_END")
    $drafts = @(ConvertFrom-MbCopilotStepAnswer -Answer $answer -PacketSteps $copilotSteps)
    Add-Result ($drafts.Count -eq 3) '固定Copilot回答から3件の下書きを受け取る'

    $selection = [pscustomobject]@{ accept = @($drafts | ForEach-Object {
        [pscustomobject]@{ id = $_.id; title = $_.title; description = $_.description; note = $_.note }
    }) } | ConvertTo-Json -Depth 6 -Compress
    $applied = Set-MbCopilotDraftSelection -Project $project -SelectionJson $selection
    Add-Result ($applied -eq 3) '3件の下書きを手順へ反映する'

    $project = Save-MbProject -Project $project -Path $projectPath
    $reloaded = Get-MbProject -Path $projectPath
    $finalSteps = @($reloaded.sheets[0].steps)
    Add-Result ($finalSteps.Count -eq 3) '下書き反映後も手順数を維持する'
    if ($finalSteps.Count -eq 3) {
        Add-Result ([string]$finalSteps[0].title -eq '録画手順1' -and
            [string]$finalSteps[1].title -eq '録画手順2' -and
            [string]$finalSteps[2].title -eq '録画手順3') '採用した手順名を再読込できる'
        Add-Result ([string]$finalSteps[0].description -eq '画面の操作1を行います。' -and
            [string]$finalSteps[2].description -eq '画面の操作3を行います。') '採用した説明を再読込できる'
        Add-Result ([string]$finalSteps[0].imageId -eq [string]$finalSteps[2].imageId) '下書き反映後も共有画像の参照を維持する'
        Add-Result (@($finalSteps | Where-Object { @($_.annotations).Count -eq 1 }).Count -eq 3) '下書き反映で赤枠を失わない'
    }
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ''
if ($errors.Count -eq 0) {
    Write-Host '録画から手順書までの統合検査はすべて成功しました。' -ForegroundColor Green
    exit 0
}
Write-Host ("失敗: " + $errors.Count + " 件") -ForegroundColor Red
exit 1
