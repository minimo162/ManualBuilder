# Phase 1 verified local application cache tests.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$launcherModule = Join-Path $repoRoot 'src\ManualBuilder.Launcher.psm1'
$testRoot = Join-Path $env:TEMP ('ManualBuilder-AppCacheTest-' + [guid]::NewGuid().ToString('N'))

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Set-MbFakeApplication {
    param([string]$Root, [string]$Version, [string]$Marker)

    [void](New-Item -ItemType Directory -Path (Join-Path $Root 'src') -Force)
    [void](New-Item -ItemType Directory -Path (Join-Path $Root 'web') -Force)
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText((Join-Path $Root 'app-version.json'), ('{"schemaVersion":1,"appVersion":"' + $Version + '"}'), $utf8)
    [IO.File]::WriteAllText((Join-Path $Root 'run.cmd'), '@echo off', [Text.Encoding]::ASCII)
    [IO.File]::WriteAllText((Join-Path $Root 'src\Start-ManualBuilder.ps1'), ('# ' + $Marker), $utf8)
    [IO.File]::WriteAllText((Join-Path $Root 'src\Start-ManualBuilderLauncher.ps1'), ('# launcher ' + $Marker), $utf8)
    [IO.File]::WriteAllText((Join-Path $Root 'web\index.html'), ('<p>' + $Marker + '</p>'), $utf8)
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    Import-Module $launcherModule -Force

    $sourceRoot = Join-Path $testRoot 'shared-app'
    $cacheRoot = Join-Path $testRoot 'local-profile\ManualBuilder\app'
    Set-MbFakeApplication -Root $sourceRoot -Version '1.0.0' -Marker 'first'
    [void](New-Item -ItemType Directory -Path (Join-Path $sourceRoot 'data\projects\default') -Force)
    [IO.File]::WriteAllText((Join-Path $sourceRoot 'data\projects\default\project.json'), '{"secret":"do-not-copy"}', (New-Object Text.UTF8Encoding($false)))

    $first = Install-MbLocalApplication -SourceRoot $sourceRoot -CacheRoot $cacheRoot
    Assert-Mb $first.Updated '初回起動で共有版をローカルキャッシュへコピーする'
    Assert-Mb ([string]$first.Version -eq '1.0.0') 'ローカルキャッシュへ配布元バージョンを記録する'
    Assert-Mb (Test-MbCachedApplication -CacheRoot $cacheRoot -ExpectedVersion '1.0.0') '初回キャッシュの全ファイルを検証できる'
    Assert-Mb (-not (Test-Path -LiteralPath (Join-Path $cacheRoot 'data'))) 'プロジェクトデータをアプリキャッシュへコピーしない'
    Assert-Mb (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $cacheRoot) 'ManualBuilder.cmd') -PathType Leaf) '共有フォルダーなしでも使えるローカル起動ファイルを作る'

    $second = Install-MbLocalApplication -SourceRoot $sourceRoot -CacheRoot $cacheRoot
    Assert-Mb (-not $second.Updated) '同じバージョンでは共有ファイルを再コピーしない'

    Set-MbFakeApplication -Root $sourceRoot -Version '1.0.1' -Marker 'second'
    $updated = Install-MbLocalApplication -SourceRoot $sourceRoot -CacheRoot $cacheRoot
    Assert-Mb $updated.Updated '配布元のバージョン更新時だけローカル版を更新する'
    Assert-Mb (Test-MbCachedApplication -CacheRoot $cacheRoot -ExpectedVersion '1.0.1') '更新後のローカル版を検証できる'
    Assert-Mb (Test-MbCachedApplication -CacheRoot ($cacheRoot + '.previous') -ExpectedVersion '1.0.0') '更新後も直前の検証済みローカル版を残す'

    [IO.File]::WriteAllText((Join-Path $cacheRoot 'web\index.html'), '<p>corrupted</p>', (New-Object Text.UTF8Encoding($false)))
    Assert-Mb (-not (Test-MbCachedApplication -CacheRoot $cacheRoot -ExpectedVersion '1.0.1')) 'ローカル版の破損をSHA-256で検出する'
    $repaired = Install-MbLocalApplication -SourceRoot $sourceRoot -CacheRoot $cacheRoot
    Assert-Mb $repaired.Updated '同じバージョンでも破損したローカル版を再作成する'
    Assert-Mb (Test-MbCachedApplication -CacheRoot $cacheRoot -ExpectedVersion '1.0.1') '再作成後のローカル版を検証できる'

    Set-MbFakeApplication -Root $sourceRoot -Version '1.0.2' -Marker 'broken-source'
    Remove-Item -LiteralPath (Join-Path $sourceRoot 'web') -Recurse -Force
    $brokenUpdateFailed = $false
    try {
        [void](Install-MbLocalApplication -SourceRoot $sourceRoot -CacheRoot $cacheRoot)
    } catch {
        $brokenUpdateFailed = $true
    }
    Assert-Mb $brokenUpdateFailed '不完全な共有版への更新を中止する'
    Assert-Mb (Test-MbCachedApplication -CacheRoot $cacheRoot -ExpectedVersion '1.0.1') '更新失敗後も現在の検証済みローカル版を維持する'
    Assert-Mb (Test-Path -LiteralPath (Join-Path $sourceRoot 'data\projects\default\project.json') -PathType Leaf) 'キャッシュ更新で共有元のデータを変更しない'

    $recoverySource = Join-Path $testRoot 'shared-app-recovery'
    $recoveryCache = Join-Path $testRoot 'recovery-profile\ManualBuilder\app'
    Set-MbFakeApplication -Root $recoverySource -Version '2.0.0' -Marker 'recovery-old'
    [void](Install-MbLocalApplication -SourceRoot $recoverySource -CacheRoot $recoveryCache)
    [IO.Directory]::Move($recoveryCache, ($recoveryCache + '.previous'))
    Set-MbFakeApplication -Root $recoverySource -Version '2.0.1' -Marker 'recovery-new'
    [void](Install-MbLocalApplication -SourceRoot $recoverySource -CacheRoot $recoveryCache)
    Assert-Mb (Test-MbCachedApplication -CacheRoot $recoveryCache -ExpectedVersion '2.0.1') '中断状態から新しいローカル版を復旧する'
    Assert-Mb (Test-MbCachedApplication -CacheRoot ($recoveryCache + '.previous') -ExpectedVersion '2.0.0') '中断状態の直前版を消さずに更新する'

    Write-Host ''
    Write-Host 'Local application cache tests passed.' -ForegroundColor Green
} finally {
    Remove-Module ManualBuilder.Launcher -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
