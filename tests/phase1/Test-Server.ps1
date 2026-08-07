# Phase 1 localhost server integration test.

[CmdletBinding()]
param([ValidateRange(1024, 65500)][int]$Port = 18765)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$serverScript = Join-Path $repoRoot 'src\Start-ManualBuilder.ps1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-ServerTest-' + [guid]::NewGuid().ToString('N'))
$projectPath = Join-Path $testRoot 'project.json'
$stdoutPath = Join-Path $testRoot 'stdout.txt'
$stderrPath = Join-Path $testRoot 'stderr.txt'
$baseUrl = "http://localhost:$Port"
$child = $null

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    $arguments = @(
        '-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA',
        '-File', ('"' + $serverScript + '"'),
        '-Port', $Port,
        '-ProjectPath', ('"' + $projectPath + '"'),
        '-DisableScreenshotWatcher',
        '-NoBrowser', '-SkipCopilotWarmup', '-AllowParallelTestInstance'
    )
    $child = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath

    $ready = $false
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 250
        try {
            $health = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/health" -TimeoutSec 2
            if ($health.StatusCode -eq 200) { $ready = $true; break }
        } catch { }
        if ($child.HasExited) { break }
    }
    Assert-Mb $ready 'localhostサーバーが起動する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $testRoot 'runtime.json') -PathType Leaf) '明示プロジェクトの実行時情報をテスト領域へ分離する'

    $shell = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/" -TimeoutSec 5
    Assert-Mb ($shell.Content -match 'ManualBuilder') 'アプリシェルを取得できる'
    $tokenMatch = [regex]::Match($shell.Content, 'X-Manual-Token":"(?<token>[a-f0-9]{32})')
    Assert-Mb $tokenMatch.Success '画面へセッショントークンが埋め込まれる'
    $headers = @{
        'X-Manual-Token' = $tokenMatch.Groups['token'].Value
        'X-Tab-Id' = 'phase1-server-test-tab'
    }

    $workspace = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/ui/workspace" -Headers $headers -TimeoutSec 5
    Assert-Mb ($workspace.Content -match 'id="workspace"') '編集画面を取得できる'

    $exportStatus = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/export/excel/status" -Headers $headers -TimeoutSec 5
    $exportStatusJson = $exportStatus.Content | ConvertFrom-Json
    Assert-Mb ([string]$exportStatusJson.state -eq 'idle') 'Excel出力の待機状態を取得できる'

    $wordExportStatus = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/export/word/status" -Headers $headers -TimeoutSec 5
    $wordExportStatusJson = $wordExportStatus.Content | ConvertFrom-Json
    Assert-Mb ([string]$wordExportStatusJson.state -eq 'idle') 'Word出力の待機状態を取得できる'

    $project = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $sheetId = [string]$project.selectedSheetId
    $added = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $sheetId } -TimeoutSec 5
    Assert-Mb ($added.Content -match 'class="step-card(?:\s|\")') '手順カードを追加できる'

    $project = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $stepId = [string]@($project.sheets[0].steps)[0].id
    $updated = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/update" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        stepId = $stepId
        title = 'ログイン画面を開く'
        description = 'ブラウザーから対象システムを開きます。'
        note = '社外からはVPNが必要です。'
    } -TimeoutSec 5
    Assert-Mb ($updated.Content -match '保存済み') '手順の文章を保存できる'

    $reloaded = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$reloaded.sheets[0].steps[0].title -eq 'ログイン画面を開く') '保存した文章がproject.jsonへ反映される'

    $attentionBody = '{"accept":[],"attention":[{"id":"' + $stepId + '","action":"review","reason":"赤枠を確認してください。"}]}'
    $attentionResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/copilot/draft/apply" -Method Post -Headers $headers -ContentType 'application/json; charset=UTF-8' -Body $attentionBody -TimeoutSec 5
    Assert-Mb (($attentionResponse.Content | ConvertFrom-Json).applied -eq 0) '採用0件でもCopilot要確認を受け付ける'
    $afterAttention = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([bool]$afterAttention.sheets[0].steps[0].review.required -and [string]$afterAttention.sheets[0].steps[0].review.action -eq 'review') 'Copilot要確認をproject.jsonへ保存する'
    $attentionWorkspace = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/ui/workspace" -Headers $headers -TimeoutSec 5
    Assert-Mb ($attentionWorkspace.Content -match 'data-step-review-notice') '再読込後も要確認を手順カードへ表示する'
    $resolvedReview = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/review/resolve" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepId = $stepId } -TimeoutSec 5
    Assert-Mb ($resolvedReview.Content -notmatch 'data-step-review-notice') '確認済み操作で要確認表示を外す'
    $afterResolve = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (-not [bool]$afterResolve.sheets[0].steps[0].review.required) '確認済み状態をproject.jsonへ保存する'

    $heartbeat = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/capture/heartbeat" -Method Post -Headers $headers -Body '' -TimeoutSec 5
    Assert-Mb ($heartbeat.Content -match 'watch-status') '撮影対象タブのハートビートを受け付ける'

    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object Drawing.Bitmap 12, 8
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $imageStream = New-Object IO.MemoryStream
    try {
        $graphics.Clear([Drawing.Color]::SteelBlue)
        $bitmap.Save($imageStream, [Drawing.Imaging.ImageFormat]::Png)
        $pngBytes = $imageStream.ToArray()
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
        $imageStream.Dispose()
    }

    $imageHeaders = @{
        'X-Manual-Token' = $headers['X-Manual-Token']
        'X-Tab-Id' = $headers['X-Tab-Id']
        'X-Sheet-Id' = $sheetId
        'X-Image-Source' = 'paste'
    }
    $imported = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/import" -Method Post -Headers $imageHeaders -ContentType 'image/png' -Body $pngBytes -TimeoutSec 5
    Assert-Mb ($imported.Content -match 'data-import-status="added"') '生バイトPOSTで画像を追加できる'
    Assert-Mb ($imported.Content -match 'class="step-image"') '追加画像を手順カードに表示できる'

    $withImage = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($withImage.images).Count -eq 1) '画像メタデータがproject.jsonへ反映される'
    Assert-Mb (@($withImage.sheets[0].steps).Count -eq 2) '画像付き手順が選択シート末尾へ追加される'
    $imageStepId = [string]$withImage.sheets[0].steps[1].id
    $insertedResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        sheetId = $sheetId
        afterStepId = $stepId
    } -TimeoutSec 5
    Assert-Mb ($insertedResponse.Content -match 'class="step-card(?:\s|\")') '指定位置への手順追加APIを実行できる'
    $afterInsert = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $insertedStepId = [string]$afterInsert.sheets[0].steps[1].id
    Assert-Mb (@($afterInsert.sheets[0].steps).Count -eq 3 -and $insertedStepId -ne $imageStepId) '現在手順の直後へ空の手順を追加する'
    $invalidAfterRejected = $false
    try {
        [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $sheetId; afterStepId = 'step-does-not-exist' } -TimeoutSec 5)
    } catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 400) { $invalidAfterRejected = $true }
    }
    $afterInvalidInsert = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ($invalidAfterRejected -and @($afterInvalidInsert.sheets[0].steps).Count -eq 3) '無効な直後指定を400で拒否し手順を増やさない'
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/delete" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepId = $insertedStepId } -TimeoutSec 5)
    $reordered = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/reorder" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        sheetId = $sheetId
        orderedIds = "$imageStepId,$stepId"
    } -TimeoutSec 5
    Assert-Mb ($reordered.Content -match '保存済み') '手順の並べ替えAPIを実行できる'
    $afterReorder = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$afterReorder.sheets[0].steps[0].id -eq $imageStepId) '並べ替え順がproject.jsonへ反映される'

    $annotationJson = '[{"id":"annotation-00000000000000000000000000000011","type":"rect","x1":0.1,"y1":0.1,"x2":0.6,"y2":0.5,"label":0},{"id":"annotation-00000000000000000000000000000012","type":"arrow","x1":0.8,"y1":0.2,"x2":0.5,"y2":0.4,"label":0},{"id":"annotation-00000000000000000000000000000013","type":"number","x1":0.25,"y1":0.3,"x2":0.25,"y2":0.3,"label":1},{"id":"annotation-00000000000000000000000000000014","type":"blackout","x1":0.7,"y1":0.7,"x2":0.9,"y2":0.9,"label":0}]'
    $annotated = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/annotations" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        stepId = $imageStepId
        annotations = $annotationJson
        crop = '{"x":0.1,"y":0.1,"width":0.8,"height":0.8}'
    } -TimeoutSec 5
    Assert-Mb ($annotated.Content -match '保存済み') '注釈保存APIを実行できる'
    $afterAnnotation = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($afterAnnotation.sheets[0].steps[0].annotations).Count -eq 4) '4種類の注釈がproject.jsonへ反映される'
    Assert-Mb ([double]$afterAnnotation.sheets[0].steps[0].crop.width -eq 0.8) '切り抜き範囲がproject.jsonへ反映される'

    $invalidAnnotationRejected = $false
    try {
        [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/annotations" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
            stepId = $imageStepId
            annotations = '[{"id":"annotation-00000000000000000000000000000015","type":"rect","x1":2,"y1":0,"x2":1,"y2":1,"label":0}]'
        } -TimeoutSec 5)
    } catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 400) { $invalidAnnotationRejected = $true }
    }
    Assert-Mb $invalidAnnotationRejected '注釈保存APIが範囲外座標を拒否する'

    $invalidCropRejected = $false
    try {
        [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/annotations" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
            stepId = $imageStepId
            annotations = $annotationJson
            crop = '{"x":0.98,"y":0,"width":0.1,"height":1}'
        } -TimeoutSec 5)
    } catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 400) { $invalidCropRejected = $true }
    }
    Assert-Mb $invalidCropRejected '注釈保存APIが画像外の切り抜きを拒否する'

    $imageId = [string]$withImage.images[0].id
    $imagePath = Join-Path (Join-Path $testRoot 'images') ([string]$withImage.images[0].fileName)
    $replacementBitmap = New-Object Drawing.Bitmap 12, 8
    $replacementGraphics = [Drawing.Graphics]::FromImage($replacementBitmap)
    $replacementStream = New-Object IO.MemoryStream
    try {
        $replacementGraphics.Clear([Drawing.Color]::OrangeRed)
        $replacementBitmap.Save($replacementStream, [Drawing.Imaging.ImageFormat]::Png)
        $replacementBytes = $replacementStream.ToArray()
    } finally {
        $replacementGraphics.Dispose()
        $replacementBitmap.Dispose()
        $replacementStream.Dispose()
    }
    $replaceHeaders = $imageHeaders.Clone()
    $replaceHeaders['X-Step-Id'] = $imageStepId
    $replacedResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/replace" -Method Post -Headers $replaceHeaders -ContentType 'image/png' -Body $replacementBytes -TimeoutSec 5
    $replacedJson = $replacedResponse.Content | ConvertFrom-Json
    Assert-Mb ([string]$replacedJson.state -eq 'replaced' -and [bool]$replacedJson.canUndo) '既存手順の画像差し替えAPIを実行できる'
    $afterReplace = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($afterReplace.sheets[0].steps).Count -eq 2 -and [string]$afterReplace.sheets[0].steps[0].id -eq $imageStepId) '画像差し替えで手順数と手順IDを維持する'
    Assert-Mb (@($afterReplace.sheets[0].steps[0].annotations).Count -eq 0 -and [double]$afterReplace.sheets[0].steps[0].crop.width -eq 1.0) '画像差し替えで新画像の編集状態を初期化する'
    $replacementImageId = [string]$afterReplace.sheets[0].steps[0].imageId
    Assert-Mb ($replacementImageId -ne $imageId) '差し替え画像をプロジェクトへ保存する'

    $undoResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/replace/undo" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepId = $imageStepId } -TimeoutSec 5
    $undoJson = $undoResponse.Content | ConvertFrom-Json
    Assert-Mb ([string]$undoJson.state -eq 'restored' -and -not [bool]$undoJson.canUndo) '画像差し替えを1段階元に戻せる'
    $afterUndo = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$afterUndo.sheets[0].steps[0].imageId -eq $imageId) '元の画像参照を復元する'
    Assert-Mb (@($afterUndo.sheets[0].steps[0].annotations).Count -eq 4 -and [double]$afterUndo.sheets[0].steps[0].crop.width -eq 0.8) '元画像の注釈と切り抜きを復元する'
    Assert-Mb (@($afterUndo.images | Where-Object { $_.id -eq $replacementImageId }).Count -eq 0) '復元後の差し替え画像を整理する'

    $resultResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/result" -Method Post -Headers $replaceHeaders -ContentType 'image/png' -Body $replacementBytes -TimeoutSec 5
    $resultJson = $resultResponse.Content | ConvertFrom-Json
    Assert-Mb ([string]$resultJson.state -eq 'set' -and [string]$resultJson.imageLayout -eq 'side-by-side') '操作後画像の追加APIを実行できる'
    $layoutResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/image-layout" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        stepId = $imageStepId; layout = 'stacked'; order = 'after-before'
    } -TimeoutSec 5
    Assert-Mb (($layoutResponse.Content | ConvertFrom-Json).state -eq 'saved') '比較画像の配置保存APIを実行できる'
    $afterLayout = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$afterLayout.sheets[0].steps[0].imageLayout -eq 'stacked' -and [string]$afterLayout.sheets[0].steps[0].imageOrder -eq 'after-before') '比較画像の配置と順序をproject.jsonへ保存する'
    $removeResultResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/result/remove" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepId = $imageStepId } -TimeoutSec 5
    Assert-Mb (($removeResultResponse.Content | ConvertFrom-Json).state -eq 'removed') '操作後画像の取り外しAPIを実行できる'
    $afterResultRemoval = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]::IsNullOrWhiteSpace([string]$afterResultRemoval.sheets[0].steps[0].resultImageId) -and [string]$afterResultRemoval.sheets[0].steps[0].imageLayout -eq 'before') '操作後画像の取り外しをproject.jsonへ保存する'

    $servedImage = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/images/${imageId}?token=$($headers['X-Manual-Token'])" -TimeoutSec 5
    Assert-Mb ($servedImage.StatusCode -eq 200 -and $servedImage.Headers['Content-Type'] -match 'image/png') '保存画像をセッショントークン付きで取得できる'

    $polled = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/capture/poll?sheetId=$sheetId&version=0" -Headers $headers -TimeoutSec 5
    Assert-Mb ($polled.Content -match 'data-capture-version="5"') '監視追加を部分更新用スナップショットで取得できる'

    $duplicate = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/import" -Method Post -Headers $imageHeaders -ContentType 'image/png' -Body $pngBytes -TimeoutSec 5
    Assert-Mb ($duplicate.Content -match 'data-import-status="duplicate"') '同じ画像の再取込みをスキップする'

    $deleted = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/delete" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepId = $imageStepId } -TimeoutSec 5
    Assert-Mb ($deleted.Content -match 'id="workspace"') '画像付き手順を削除できる'
    $afterDelete = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($afterDelete.sheets[0].steps).Count -eq 1) '削除結果がproject.jsonへ反映される'
    Assert-Mb (@($afterDelete.images).Count -eq 0) '削除した手順だけが使う画像メタデータを整理する'
    Assert-Mb (Test-Path -LiteralPath $imagePath) '取り消しに備えて削除した手順の画像ファイルを保持する'
    $deleteStatus = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/deletions/status" -Headers $headers -TimeoutSec 5
    $deleteStatusJson = $deleteStatus.Content | ConvertFrom-Json
    Assert-Mb ([bool]$deleteStatusJson.available -and [string]$deleteStatusJson.label -match '手順') '直前の削除を元に戻せると通知する'
    $deleteUndo = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/deletions/undo" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5
    Assert-Mb ($deleteUndo.Content -match 'id="workspace"') '削除した手順を元に戻すAPIを実行できる'
    $afterDeleteUndo = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($afterDeleteUndo.sheets[0].steps).Count -eq 2 -and [string]$afterDeleteUndo.sheets[0].steps[0].id -eq $imageStepId) '手順を削除前の位置へ復元する'
    Assert-Mb (@($afterDeleteUndo.images | Where-Object { $_.id -eq $imageId }).Count -eq 1 -and (Test-Path -LiteralPath $imagePath)) '削除した手順の画像参照とファイルを復元する'
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/delete" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepId = $imageStepId } -TimeoutSec 5)

    $emptyStepHeaders = $imageHeaders.Clone()
    $emptyStepHeaders['X-Step-Id'] = $stepId
    $emptyStepImageResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/replace" -Method Post -Headers $emptyStepHeaders -ContentType 'image/png' -Body $replacementBytes -TimeoutSec 5
    $emptyStepImageJson = $emptyStepImageResponse.Content | ConvertFrom-Json
    Assert-Mb ([string]$emptyStepImageJson.state -eq 'replaced' -and -not [bool]$emptyStepImageJson.canUndo) '空の手順へ画像を追加できる'
    $afterEmptyStepImage = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($afterEmptyStepImage.sheets[0].steps).Count -eq 1 -and [string]$afterEmptyStepImage.sheets[0].steps[0].imageId) '空の手順への画像追加で手順数を維持する'

    $reimported = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/images/import" -Method Post -Headers $imageHeaders -ContentType 'image/png' -Body $pngBytes -TimeoutSec 5
    Assert-Mb ($reimported.Content -match 'data-import-status="added"') '削除後は同じ画面を撮り直せる'
    $beforeBulkSetup = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $sourceBeforeSetup = @($beforeBulkSetup.sheets | Where-Object { $_.id -eq $sheetId })[0]
    $reimportedStepId = [string]@($sourceBeforeSetup.steps | Where-Object { $_.id -ne $stepId })[0].id
    $reimportedImageId = [string]@($sourceBeforeSetup.steps | Where-Object { $_.id -eq $reimportedStepId })[0].imageId
    $reimportedImage = @($beforeBulkSetup.images | Where-Object { $_.id -eq $reimportedImageId })[0]
    $reimportedImagePath = Join-Path (Join-Path $testRoot 'images') ([string]$reimportedImage.fileName)
    Assert-Mb (Test-Path -LiteralPath $reimportedImagePath) '一括削除テスト用の画像ファイルが存在する'
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $sheetId } -TimeoutSec 5)
    $afterFirstBulkAdd = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $sourceAfterFirstBulkAdd = @($afterFirstBulkAdd.sheets | Where-Object { $_.id -eq $sheetId })[0]
    $bulkSpacerStepId = [string]$sourceAfterFirstBulkAdd.steps[-1].id
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $sheetId } -TimeoutSec 5)
    $afterSecondBulkAdd = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $sourceAfterSecondBulkAdd = @($afterSecondBulkAdd.sheets | Where-Object { $_.id -eq $sheetId })[0]
    $bulkSecondStepId = [string]$sourceAfterSecondBulkAdd.steps[-1].id

    $addedSheetResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/sheets/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5
    Assert-Mb ($addedSheetResponse.Content -match 'data-sheet-nav-item') '移動先シートを追加できる'
    $afterSheetAdd = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $targetSheetId = [string]$afterSheetAdd.selectedSheetId
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/add" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $targetSheetId } -TimeoutSec 5)
    $afterTargetSeed = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $targetBeforeBulk = @($afterTargetSeed.sheets | Where-Object { $_.id -eq $targetSheetId })[0]
    $targetExistingStepId = [string]$targetBeforeBulk.steps[0].id
    $sheetDuplicated = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/sheets/duplicate" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $targetSheetId } -TimeoutSec 5
    Assert-Mb ($sheetDuplicated.Content -match 'id="workspace"' -and $sheetDuplicated.Content -match 'のコピー') 'シート複製APIを実行できる'
    $afterSheetDuplicate = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $duplicatedSheetId = [string]$afterSheetDuplicate.selectedSheetId
    $duplicatedSheet = @($afterSheetDuplicate.sheets | Where-Object { $_.id -eq $duplicatedSheetId })[0]
    Assert-Mb ($duplicatedSheetId -ne $targetSheetId -and @($duplicatedSheet.steps).Count -eq 1) '複製シートを選択し手順を維持する'
    Assert-Mb ([string]$duplicatedSheet.steps[0].id -ne $targetExistingStepId) '複製した手順へ新しいIDを割り当てる'
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/sheets/delete" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $duplicatedSheetId } -TimeoutSec 5)

    $bulkStepIds = @($reimportedStepId, $bulkSecondStepId)
    $bulkMovedResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/move-many" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        stepIds = ($bulkStepIds -join ',')
        targetSheetId = $targetSheetId
    } -TimeoutSec 5
    Assert-Mb ($bulkMovedResponse.Content -match 'class="step-card(?:\s|\")') '複数手順の移動APIを実行できる'
    $afterBulkMove = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $bulkTarget = @($afterBulkMove.sheets | Where-Object { $_.id -eq $targetSheetId }) | Select-Object -First 1
    $bulkSource = @($afterBulkMove.sheets | Where-Object { $_.id -eq $sheetId }) | Select-Object -First 1
    Assert-Mb ((@($bulkTarget.steps | ForEach-Object { [string]$_.id }) -join ',') -eq (@($targetExistingStepId) + $bulkStepIds -join ',')) '一括移動で移動先の既存順と要求順を維持する'
    Assert-Mb ((@($bulkSource.steps | ForEach-Object { [string]$_.id }) -join ',') -eq (@($stepId, $bulkSpacerStepId) -join ',')) '一括移動で移動元の未選択順を維持する'
    $bulkDeletedResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/delete-many" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepIds = ($bulkStepIds -join ',') } -TimeoutSec 5
    Assert-Mb ($bulkDeletedResponse.Content -match 'id="workspace"') '複数手順の削除APIを実行できる'
    $afterBulkDelete = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $remainingBulkSteps = @($afterBulkDelete.sheets | ForEach-Object { @($_.steps) } | Where-Object { $_.id -in $bulkStepIds })
    Assert-Mb ($remainingBulkSteps.Count -eq 0) '一括削除した2手順が再読込後も残らない'
    Assert-Mb (@($afterBulkDelete.images | Where-Object { $_.id -eq $reimportedImageId }).Count -eq 0) '一括削除で孤立した画像メタデータを整理する'
    Assert-Mb (Test-Path -LiteralPath $reimportedImagePath) '一括削除した画像ファイルを取り消し用に保持する'
    $bulkUndoResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/deletions/undo" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5
    Assert-Mb ($bulkUndoResponse.Content -match 'id="workspace"') '複数手順の削除を元に戻せる'
    $afterBulkUndo = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $bulkRestored = @($afterBulkUndo.sheets | Where-Object { $_.id -eq $targetSheetId })[0]
    Assert-Mb ((@($bulkRestored.steps | ForEach-Object { [string]$_.id }) -join ',') -eq (@($targetExistingStepId) + $bulkStepIds -join ',')) '一括削除した手順を元の順序へ復元する'
    Assert-Mb (@($afterBulkUndo.images | Where-Object { $_.id -eq $reimportedImageId }).Count -eq 1 -and (Test-Path -LiteralPath $reimportedImagePath)) '一括削除した画像を復元する'
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/delete-many" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepIds = ($bulkStepIds -join ',') } -TimeoutSec 5)
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/delete-many" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ stepIds = "$targetExistingStepId,$bulkSpacerStepId" } -TimeoutSec 5)
    $sheetOrder = @($targetSheetId, $sheetId) -join ','
    $sheetReordered = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/sheets/reorder" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ orderedIds = $sheetOrder } -TimeoutSec 5
    Assert-Mb ($sheetReordered.Content -match '保存済み') 'シートの並べ替えAPIを実行できる'
    $afterSheetReorder = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$afterSheetReorder.sheets[0].id -eq $targetSheetId) 'シート順がproject.jsonへ反映される'

    $movedStepResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/steps/move" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{
        stepId = $stepId
        targetSheetId = $targetSheetId
    } -TimeoutSec 5
    Assert-Mb ($movedStepResponse.Content -match 'class="step-card(?:\s|\")') '手順を別シートへ移動するAPIを実行できる'
    $afterStepMove = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $movedTarget = @($afterStepMove.sheets | Where-Object { $_.id -eq $targetSheetId }) | Select-Object -First 1
    $movedSource = @($afterStepMove.sheets | Where-Object { $_.id -eq $sheetId }) | Select-Object -First 1
    Assert-Mb (@($movedTarget.steps).Count -eq 1 -and [string]$movedTarget.steps[0].id -eq $stepId) '移動先シートへ同じ手順IDを保持する'
    Assert-Mb (@($movedSource.steps | Where-Object { $_.id -eq $stepId }).Count -eq 0) '移動元シートから手順を取り除く'
    Assert-Mb ([string]$movedTarget.steps[0].title -eq 'ログイン画面を開く') 'シート移動後も文章を維持する'

    $sheetDeleted = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/sheets/delete" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ sheetId = $targetSheetId } -TimeoutSec 5
    Assert-Mb ($sheetDeleted.Content -match 'id="workspace"') '手順を含むシートを削除できる'
    $afterSheetDelete = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb (@($afterSheetDelete.sheets | Where-Object { $_.id -eq $targetSheetId }).Count -eq 0) 'シート削除をproject.jsonへ反映する'
    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/deletions/undo" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5)
    $afterSheetUndo = [IO.File]::ReadAllText($projectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $restoredSheet = @($afterSheetUndo.sheets | Where-Object { $_.id -eq $targetSheetId }) | Select-Object -First 1
    Assert-Mb ($restoredSheet -and @($restoredSheet.steps).Count -eq 1 -and [string]$restoredSheet.steps[0].id -eq $stepId) '削除したシートと中の手順をまとめて復元する'

    $badTokenRejected = $false
    try {
        [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/ui/workspace" -Headers @{ 'X-Manual-Token' = 'invalid' } -TimeoutSec 5)
    } catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 403) { $badTokenRejected = $true }
    }
    Assert-Mb $badTokenRejected '不正なセッショントークンを拒否する'

    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/shutdown" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5)
    [void]$child.WaitForExit(5000)
    Assert-Mb $child.HasExited '終了APIでサーバーが停止する'
    Write-Host ''
    Write-Host 'Server integration tests passed.' -ForegroundColor Cyan
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    if (Test-Path -LiteralPath $stdoutPath) {
        Write-Host '--- server stdout ---' -ForegroundColor Yellow
        Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $stderrPath) {
        Write-Host '--- server stderr ---' -ForegroundColor Yellow
        Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue
    }
    throw
} finally {
    if ($child -and -not $child.HasExited) {
        Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
