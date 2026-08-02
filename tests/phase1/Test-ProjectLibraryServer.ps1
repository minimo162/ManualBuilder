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
        '-DisableScreenshotWatcher', '-NoBrowser'
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

    $shell = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/" -TimeoutSec 5
    $tokenMatch = [regex]::Match($shell.Content, 'X-Manual-Token":"(?<token>[a-f0-9]{32})')
    Assert-Mb $tokenMatch.Success '一覧モードでもセッショントークンを発行する'
    $headers = @{ 'X-Manual-Token' = $tokenMatch.Groups['token'].Value; 'X-Tab-Id' = 'project-library-test-tab' }

    $library = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/ui/workspace" -Headers $headers -TimeoutSec 5
    Assert-Mb ($library.Content -match 'class="workspace project-library"') '通常起動でマニュアル一覧を表示する'
    Assert-Mb ($library.Content -match '新規作成') '一覧から新規作成できる'
    Assert-Mb ($library.Content -match 'data-project-home') '一覧画面のアプリロゴをホーム操作として表示する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $dataRoot 'projects\default\project.json') -PathType Leaf) '従来のdefaultマニュアルを残す'

    $createdResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/create" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ title = '部内経費精算' } -TimeoutSec 5
    Assert-Mb ($createdResponse.Content -match 'class="workspace"') '新規作成後に編集画面を開く'
    Assert-Mb ($createdResponse.Content -match '部内経費精算') '新規作成した名前を編集画面に表示する'
    Assert-Mb (($createdResponse.Content -match 'data-project-home') -and ($createdResponse.Content -notmatch 'project-home-button')) '編集画面のアプリロゴから一覧へ戻れる'
    $projectDirectories = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'projects') -Directory | Where-Object { $_.Name -ne 'default' })
    Assert-Mb ($projectDirectories.Count -eq 1) '新規マニュアルを個別フォルダーへ保存する'
    $createdKey = [string]$projectDirectories[0].Name

    $libraryHomeResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/home" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body '' -TimeoutSec 5
    Assert-Mb ($libraryHomeResponse.Content -match 'class="workspace project-library"') '編集画面から一覧へ戻る'

    [void](Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/export" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -OutFile $transferZipPath -TimeoutSec 10)
    Assert-Mb ((Test-Path -LiteralPath $transferZipPath -PathType Leaf) -and (Get-Item -LiteralPath $transferZipPath).Length -gt 0) '一覧APIからマニュアルZIPを書き出す'

    $importedResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/import" -Method Post -Headers $headers -ContentType 'application/zip' -InFile $transferZipPath -TimeoutSec 10
    Assert-Mb ($importedResponse.Content -match '部内経費精算') 'マニュアルZIPを新しい項目として取り込む'
    $afterImport = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'projects') -Directory)
    Assert-Mb ($afterImport.Count -eq 3) 'ZIP取込みで既存マニュアルを上書きしない'

    $duplicated = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/duplicate" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($duplicated.Content -match '部内経費精算 - コピー') '一覧からマニュアルを複製する'
    $afterCopy = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'projects') -Directory)
    Assert-Mb ($afterCopy.Count -eq 4) '複製先を独立フォルダーへ保存する'

    $archiveResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/archive" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($archiveResponse.Content -match 'アーカイブ') 'アーカイブ後も一覧を表示する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $dataRoot "projects-archive\$createdKey\project.json") -PathType Leaf) 'アーカイブは削除せず別フォルダーへ移す'

    $restoreResponse = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/restore" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($restoreResponse.Content -match '部内経費精算') 'アーカイブからマニュアルを復元する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $dataRoot "projects\$createdKey\project.json") -PathType Leaf) '復元後は現役フォルダーへ戻す'

    $opened = Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/api/projects/open" -Method Post -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body @{ projectKey = $createdKey } -TimeoutSec 5
    Assert-Mb ($opened.Content -match '部内経費精算') '復元したマニュアルを開く'
    $settings = [IO.File]::ReadAllText((Join-Path $dataRoot 'settings.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$settings.lastOpenedProjectKey -eq $createdKey) '開いたマニュアルを次回の目印として記録する'

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
