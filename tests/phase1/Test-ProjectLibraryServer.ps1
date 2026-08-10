# ManualBuilder project-library server integration test.

[CmdletBinding()]
param([ValidateRange(1024, 65500)][int]$Port = 18766)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$serverScript = Join-Path $repoRoot 'src\Start-ManualBuilder.ps1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-ProjectLibraryServer-' + [guid]::NewGuid().ToString('N'))
$dataRoot = Join-Path $testRoot 'data'
$legacyRoot = Join-Path $testRoot 'legacy-empty'
$stdoutPath = Join-Path $testRoot 'stdout.txt'
$stderrPath = Join-Path $testRoot 'stderr.txt'
$transferZipPath = Join-Path $testRoot 'project-transfer.zip'
$baseUrl = "http://localhost:$Port"
$child = $null

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    [void](New-Item -ItemType Directory -Path $legacyRoot -Force)
    $arguments = @(
        '-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA',
        '-File', ('"' + $serverScript + '"'), '-Port', $Port,
        '-DataRoot', ('"' + $dataRoot + '"'), '-LegacyAppRoot', ('"' + $legacyRoot + '"'),
        '-DisableScreenshotWatcher', '-NoBrowser', '-AllowParallelTestInstance'
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
    Assert-Mb $ready 'プロジェクト一覧モードでサーバーが起動する'

    # 画面本体もトークンで守るため、無認証の `/` からは取れない。
    $runtimeInfo = [IO.File]::ReadAllText((Join-Path $dataRoot 'runtime.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    $entryUrl = [string]$runtimeInfo.entryUrl
    Assert-Mb ($entryUrl -match '\?token=(?<token>[a-f0-9]{32})$') '一覧モードでもセッショントークンを発行する'
    $sessionToken = $Matches['token']
    $shell = Invoke-WebRequest -UseBasicParsing -Uri $entryUrl -TimeoutSec 5
    $headers = @{ 'X-Manual-Token' = $sessionToken; 'X-Tab-Id' = 'project-library-test-tab' }

    $library = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/ui/workspace" -Headers $headers -TimeoutSec 5
    Assert-Mb ($library.Content -match 'class="workspace project-library"') '通常起動でマニュアル一覧を表示する'
    Assert-Mb ($library.Content -match '>新しいマニュアルを作る</button>') '一覧から新しいマニュアルを作成できる'
    Assert-Mb ($library.Content -match 'data-project-home') '一覧画面のアプリロゴをホーム操作として表示する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $dataRoot 'projects\default\project.json') -PathType Leaf) '従来のdefaultマニュアルを残す'

    $createdResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/create" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ title = '部内経費精算' } -TimeoutSec 5
    Assert-Mb ($createdResponse.Content -match 'class="[^"]*\bworkspace\b[^"]*"') '新規作成後に編集画面を開く'
    Assert-Mb ($createdResponse.Content -match '部内経費精算') '新規作成した名前を編集画面に表示する'
    Assert-Mb (($createdResponse.Content -match 'data-project-home') -and ($createdResponse.Content -notmatch 'project-home-button')) '編集画面のアプリロゴから一覧へ戻れる'
    $projectDirectories = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'projects') -Directory | Where-Object { $_.Name -ne 'default' })
    Assert-Mb ($projectDirectories.Count -eq 1) '新規マニュアルを個別フォルダーへ保存する'
    $createdKey = [string]$projectDirectories[0].Name

    $libraryHomeResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/home" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5
    Assert-Mb ($libraryHomeResponse.Content -match 'class="workspace project-library"') '編集画面から一覧へ戻る'

    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/export" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -OutFile $transferZipPath -TimeoutSec 10)
    Assert-Mb ((Test-Path -LiteralPath $transferZipPath -PathType Leaf) -and (Get-Item -LiteralPath $transferZipPath).Length -gt 0) '一覧APIからマニュアルZIPを書き出す'

    # ZIP取込みは展開・全ファイル検証・原子的保存までを同期して完了させる。
    # Defenderが新規ZIPと展開ファイルを走査する環境でも実処理の成否を判定できる猶予を持たせる。
    $importedResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/import" -Method Post -Headers $headers -ContentType 'application/zip' -InFile $transferZipPath -TimeoutSec 60
    Assert-Mb ($importedResponse.Content -match '部内経費精算') 'マニュアルZIPを新しい項目として取り込む'
    $afterImport = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'projects') -Directory)
    Assert-Mb ($afterImport.Count -eq 3) 'ZIP取込みで既存マニュアルを上書きしない'

    $duplicated = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/duplicate" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($duplicated.Content -match '部内経費精算 - コピー') '一覧からマニュアルを複製する'
    $afterCopy = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'projects') -Directory)
    Assert-Mb ($afterCopy.Count -eq 4) '複製先を独立フォルダーへ保存する'

    $opened = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/open" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($opened.Content -match '部内経費精算') '作成したマニュアルを開く'
    $settings = [IO.File]::ReadAllText((Join-Path $dataRoot 'settings.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$settings.lastOpenedProjectKey -eq $createdKey) '開いたマニュアルを次回の目印として記録する'

    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/home" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5)
    $deleteResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/delete" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($deleteResponse.Content -match 'class="workspace project-library"') '削除後もマニュアル一覧を表示する'
    Assert-Mb (-not (Test-Path -LiteralPath (Join-Path $dataRoot "projects\$createdKey") -PathType Container)) '削除したマニュアルの全ファイルを消す'
    Assert-Mb (-not (Test-Path -LiteralPath (Join-Path $dataRoot "projects-archive\$createdKey") -PathType Container)) '削除したマニュアルをごみ箱へ移さない'
    Assert-Mb ($deleteResponse.Content -notmatch 'ごみ箱|アーカイブ') '一覧にごみ箱とアーカイブを表示しない'

    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/shutdown" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5)
    [void]$child.WaitForExit(5000)
    Assert-Mb $child.HasExited '一覧モードのサーバーを終了できる'
    Write-Host ''
    Write-Host 'Project library server tests passed.' -ForegroundColor Cyan
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue }
    throw
} finally {
    if ($child -and -not $child.HasExited) { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
