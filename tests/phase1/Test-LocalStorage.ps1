# Phase 1 per-user local storage tests.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$storageModule = Join-Path $repoRoot 'src\ManualBuilder.Storage.psm1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-StorageTest-' + [guid]::NewGuid().ToString('N'))

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    Import-Module $storageModule -Force

    $localApplicationData = Join-Path $testRoot 'local-profile'
    $defaultDataRoot = Get-MbDefaultDataRoot -LocalApplicationData $localApplicationData
    $expectedDefaultRoot = [IO.Path]::GetFullPath((Join-Path $localApplicationData 'ManualBuilder\data'))
    Assert-Mb ($defaultDataRoot -eq $expectedDefaultRoot) '既定保存先をLocalApplicationData配下にする'

    $appRoot = Join-Path $testRoot 'shared-app'
    $legacyDirectory = Join-Path $appRoot 'data\projects\default'
    $legacyImages = Join-Path $legacyDirectory 'images'
    [void](New-Item -ItemType Directory -Path $legacyImages -Force)
    $legacyProjectPath = Join-Path $legacyDirectory 'project.json'
    [IO.File]::WriteAllText($legacyProjectPath, '{"title":"移行対象","revision":7}', (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllBytes((Join-Path $legacyImages 'sample.png'), [byte[]](1, 2, 3, 4))

    $dataRoot = Join-Path $localApplicationData 'ManualBuilder\data'
    $layout = Get-MbStorageLayout -AppRoot $appRoot -DataRoot $dataRoot
    Assert-Mb ([string]$layout.RuntimePath -eq (Join-Path $dataRoot 'runtime.json')) 'runtime情報をユーザーデータ配下にする'
    Assert-Mb ([string]$layout.ExportJobsRoot -eq (Join-Path $dataRoot 'export-jobs')) 'Office一時ジョブをユーザーデータ配下にする'

    $migration = Initialize-MbUserStorage -Layout $layout -MigrateLegacy
    Assert-Mb $migration.Migrated '既存のdefaultプロジェクトを初回だけ移行する'
    Assert-Mb (Test-Path -LiteralPath ([string]$layout.ProjectPath) -PathType Leaf) 'ローカルへproject.jsonをコピーする'
    Assert-Mb (Test-Path -LiteralPath (Join-Path (Split-Path -Parent ([string]$layout.ProjectPath)) 'images\sample.png') -PathType Leaf) 'ローカルへ画像をコピーする'
    Assert-Mb (Test-Path -LiteralPath $legacyProjectPath -PathType Leaf) '移行後も元のproject.jsonを削除しない'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $legacyImages 'sample.png') -PathType Leaf) '移行後も元画像を削除しない'
    $migrationMarker = Join-Path $dataRoot 'legacy-default-migration.completed'
    Assert-Mb (Test-Path -LiteralPath $migrationMarker -PathType Leaf) '初回移行済みの印をローカルへ残す'

    [IO.File]::WriteAllText([string]$layout.ProjectPath, '{"title":"ローカル編集後"}', (New-Object Text.UTF8Encoding($false)))
    $secondMigration = Initialize-MbUserStorage -Layout $layout -MigrateLegacy
    $localProject = ([IO.File]::ReadAllText([string]$layout.ProjectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json)
    Assert-Mb (-not $secondMigration.Migrated) '2回目の起動では移行を繰り返さない'
    Assert-Mb ([string]$localProject.title -eq 'ローカル編集後') '既存のローカルデータを旧データで上書きしない'

    Remove-Item -LiteralPath ([string]$layout.ProjectPath) -Force
    $afterArchiveMigration = Initialize-MbUserStorage -Layout $layout -MigrateLegacy
    Assert-Mb (-not $afterArchiveMigration.Migrated) 'defaultをアーカイブした後も旧データを再移行しない'
    Assert-Mb (-not (Test-Path -LiteralPath ([string]$layout.ProjectPath) -PathType Leaf)) 'アーカイブ後のdefaultを勝手に復活させない'

    $invalidAppRoot = Join-Path $testRoot 'shared-app-invalid'
    $invalidLegacyDirectory = Join-Path $invalidAppRoot 'data\projects\default'
    [void](New-Item -ItemType Directory -Path $invalidLegacyDirectory -Force)
    $invalidLegacyProject = Join-Path $invalidLegacyDirectory 'project.json'
    [IO.File]::WriteAllText($invalidLegacyProject, '{invalid-json', (New-Object Text.UTF8Encoding($false)))
    $invalidDataRoot = Join-Path $testRoot 'invalid-local\ManualBuilder\data'
    $invalidLayout = Get-MbStorageLayout -AppRoot $invalidAppRoot -DataRoot $invalidDataRoot
    $migrationFailed = $false
    try {
        [void](Initialize-MbUserStorage -Layout $invalidLayout -MigrateLegacy)
    } catch {
        $migrationFailed = $true
    }
    Assert-Mb $migrationFailed '検証できない旧プロジェクトの移行を中止する'
    Assert-Mb (-not (Test-Path -LiteralPath ([string]$invalidLayout.ProjectPath) -PathType Leaf)) '移行失敗時に空のローカル正本を作らない'
    Assert-Mb (Test-Path -LiteralPath $invalidLegacyProject -PathType Leaf) '移行失敗時も旧プロジェクトを残す'

    $explicitProjectPath = Join-Path $testRoot 'explicit\project.json'
    $explicitLayout = Get-MbStorageLayout -AppRoot $appRoot -ProjectPath $explicitProjectPath
    $expectedExplicitDataRoot = Split-Path -Parent ([IO.Path]::GetFullPath($explicitProjectPath))
    Assert-Mb ([string]$explicitLayout.DataRoot -eq $expectedExplicitDataRoot) 'テスト用ProjectPath指定時は同じフォルダーへ実行時データを分離する'
    Assert-Mb (-not $explicitLayout.UsesDefaultUserData) '明示パス指定時は自動移行を行わない'

    $cachedAppRoot = Join-Path $testRoot 'local-profile\ManualBuilder\app'
    $sharedLegacyRoot = Join-Path $testRoot 'shared-distribution'
    $cachedLayout = Get-MbStorageLayout -AppRoot $cachedAppRoot -DataRoot $dataRoot -LegacyAppRoot $sharedLegacyRoot
    $expectedSharedLegacyPath = [IO.Path]::GetFullPath((Join-Path $sharedLegacyRoot 'data\projects\default\project.json'))
    Assert-Mb ([string]$cachedLayout.LegacyProjectPath -eq $expectedSharedLegacyPath) 'ローカル実行時も共有元の旧データ位置を引き継ぐ'

    Write-Host ''
    Write-Host 'Local storage tests passed.' -ForegroundColor Green
} finally {
    Remove-Module ManualBuilder.Storage -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
