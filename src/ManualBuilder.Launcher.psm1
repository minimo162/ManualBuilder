# ManualBuilder verified per-user application cache.

Set-StrictMode -Version 2.0

$script:MbAppCacheEntries = @('src', 'web', 'run.cmd', 'app-version.json')

function Get-MbApplicationVersion {
    param([Parameter(Mandatory = $true)][string]$AppRoot)

    $manifestPath = Join-Path ([IO.Path]::GetFullPath($AppRoot)) 'app-version.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "アプリのバージョン情報が見つかりません: $manifestPath"
    }
    $manifest = [IO.File]::ReadAllText($manifestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $version = [string]$manifest.appVersion
    if ([int]$manifest.schemaVersion -ne 1 -or $version -notmatch '^\d+\.\d+\.\d+$') {
        throw "アプリのバージョン情報が不正です: $manifestPath"
    }
    return $version
}

function Get-MbDefaultAppCacheRoot {
    param([AllowEmptyString()][string]$LocalApplicationData = '')

    if ([string]::IsNullOrWhiteSpace($LocalApplicationData)) {
        $LocalApplicationData = [Environment]::GetFolderPath('LocalApplicationData')
    }
    if ([string]::IsNullOrWhiteSpace($LocalApplicationData)) {
        throw 'ユーザーのLocalApplicationDataフォルダーを取得できません。'
    }
    return [IO.Path]::GetFullPath((Join-Path $LocalApplicationData 'ManualBuilder\app'))
}

function Get-MbAppPackageFiles {
    param([Parameter(Mandatory = $true)][string]$AppRoot)

    $resolvedRoot = [IO.Path]::GetFullPath($AppRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $trimCharacters = [char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $files = New-Object System.Collections.ArrayList
    foreach ($entryName in $script:MbAppCacheEntries) {
        $entryPath = Join-Path $resolvedRoot $entryName
        if (-not (Test-Path -LiteralPath $entryPath)) {
            throw "アプリの必須項目が見つかりません: $entryName"
        }
        $entry = Get-Item -LiteralPath $entryPath -Force -ErrorAction Stop
        if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "アプリ内のリンクはローカルキャッシュへコピーできません: $entryName"
        }
        $candidates = if ($entry.PSIsContainer) {
            @(Get-ChildItem -LiteralPath $entry.FullName -Recurse -File -Force -ErrorAction Stop)
        } else {
            @($entry)
        }
        foreach ($candidate in $candidates) {
            if (($candidate.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "アプリ内のリンクはローカルキャッシュへコピーできません: $($candidate.FullName)"
            }
            $relativePath = $candidate.FullName.Substring($resolvedRoot.Length).TrimStart($trimCharacters)
            [void]$files.Add([pscustomobject]@{
                RelativePath = $relativePath
                FullName = $candidate.FullName
                Length = [long]$candidate.Length
            })
        }
    }
    return @($files | Sort-Object RelativePath)
}

function Test-MbCachedApplication {
    param(
        [Parameter(Mandatory = $true)][string]$CacheRoot,
        [AllowEmptyString()][string]$ExpectedVersion = ''
    )

    try {
        $resolvedCacheRoot = [IO.Path]::GetFullPath($CacheRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
        $installPath = Join-Path $resolvedCacheRoot '.install.json'
        if (-not (Test-Path -LiteralPath $installPath -PathType Leaf)) { return $false }
        $version = Get-MbApplicationVersion -AppRoot $resolvedCacheRoot
        if ($ExpectedVersion -and $version -ne $ExpectedVersion) { return $false }

        $install = [IO.File]::ReadAllText($installPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        if ([int]$install.schemaVersion -ne 1 -or [string]$install.appVersion -ne $version) { return $false }
        $installedFiles = @($install.files | Sort-Object path)
        if ($installedFiles.Count -lt 1) { return $false }
        $actualFiles = @(Get-MbAppPackageFiles -AppRoot $resolvedCacheRoot)
        if ($actualFiles.Count -ne $installedFiles.Count) { return $false }

        $cachePrefix = $resolvedCacheRoot + [IO.Path]::DirectorySeparatorChar
        for ($index = 0; $index -lt $installedFiles.Count; $index++) {
            $file = $installedFiles[$index]
            $actualFile = $actualFiles[$index]
            $relativePath = [string]$file.path
            if ([string]::IsNullOrWhiteSpace($relativePath) -or [IO.Path]::IsPathRooted($relativePath) -or
                $relativePath -match '(^|[\\/])\.\.([\\/]|$)') { return $false }
            if (-not $relativePath.Equals([string]$actualFile.RelativePath, [StringComparison]::OrdinalIgnoreCase)) { return $false }
            $fullPath = [IO.Path]::GetFullPath((Join-Path $resolvedCacheRoot $relativePath))
            if (-not $fullPath.StartsWith($cachePrefix, [StringComparison]::OrdinalIgnoreCase)) { return $false }
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { return $false }
            $item = Get-Item -LiteralPath $fullPath -Force -ErrorAction Stop
            if ([long]$item.Length -ne [long]$file.length) { return $false }
            $hash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not $hash.Equals([string]$file.sha256, [StringComparison]::OrdinalIgnoreCase)) { return $false }
        }
        return $true
    } catch {
        return $false
    }
}

function Write-MbLocalLauncher {
    param([Parameter(Mandatory = $true)][string]$ProductRoot)

    $launcherPath = Join-Path $ProductRoot 'ManualBuilder.cmd'
    $temporaryPath = $launcherPath + '.tmp-' + [guid]::NewGuid().ToString('N')
    $backupPath = $launcherPath + '.bak'
    $content = @'
@echo off
setlocal
set "MB_SCRIPT=%~dp0app\src\Start-ManualBuilderLauncher.ps1"
if not exist "%MB_SCRIPT%" set "MB_SCRIPT=%~dp0app.previous\src\Start-ManualBuilderLauncher.ps1"
if not exist "%MB_SCRIPT%" (
  echo ManualBuilder local cache is not available.
  pause
  exit /b 1
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%MB_SCRIPT%"
set "MB_EXIT=%ERRORLEVEL%"
if not "%MB_EXIT%"=="0" pause
exit /b %MB_EXIT%
'@
    $content = ($content -replace "`r?`n", "`r`n") + "`r`n"
    try {
        [IO.File]::WriteAllText($temporaryPath, $content, [Text.Encoding]::ASCII)
        if (Test-Path -LiteralPath $launcherPath -PathType Leaf) {
            if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                Remove-Item -LiteralPath $backupPath -Force -ErrorAction Stop
            }
            [IO.File]::Replace($temporaryPath, $launcherPath, $backupPath)
            if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
            }
        } else {
            [IO.File]::Move($temporaryPath, $launcherPath)
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
    return $launcherPath
}

function Install-MbLocalApplication {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$CacheRoot
    )

    $resolvedSourceRoot = [IO.Path]::GetFullPath($SourceRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $resolvedCacheRoot = [IO.Path]::GetFullPath($CacheRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if ($resolvedSourceRoot.Equals($resolvedCacheRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{ Updated = $false; Version = (Get-MbApplicationVersion -AppRoot $resolvedSourceRoot); CacheRoot = $resolvedCacheRoot; SourceIsCache = $true }
    }

    $sourceVersion = Get-MbApplicationVersion -AppRoot $resolvedSourceRoot
    $productRoot = Split-Path -Parent $resolvedCacheRoot
    if ([string]::IsNullOrWhiteSpace($productRoot)) { throw 'ローカルキャッシュの親フォルダーが不正です。' }
    [void](New-Item -ItemType Directory -Path $productRoot -Force)

    if (Test-MbCachedApplication -CacheRoot $resolvedCacheRoot -ExpectedVersion $sourceVersion) {
        $localLauncher = Write-MbLocalLauncher -ProductRoot $productRoot
        return [pscustomobject]@{ Updated = $false; Version = $sourceVersion; CacheRoot = $resolvedCacheRoot; SourceIsCache = $false; LocalLauncher = $localLauncher }
    }

    $sourceFiles = @(Get-MbAppPackageFiles -AppRoot $resolvedSourceRoot)
    $stagingRoot = Join-Path $productRoot ('.app-update-' + [guid]::NewGuid().ToString('N'))
    $previousRoot = $resolvedCacheRoot + '.previous'
    $movedCurrent = $false
    $installedNew = $false
    try {
        [void](New-Item -ItemType Directory -Path $stagingRoot)
        foreach ($sourceFile in $sourceFiles) {
            $destinationPath = Join-Path $stagingRoot ([string]$sourceFile.RelativePath)
            $destinationDirectory = Split-Path -Parent $destinationPath
            if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
                [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
            }
            Copy-Item -LiteralPath $sourceFile.FullName -Destination $destinationPath -Force -ErrorAction Stop
        }

        $stagedFiles = @(Get-MbAppPackageFiles -AppRoot $stagingRoot)
        if ($sourceFiles.Count -ne $stagedFiles.Count) { throw 'ローカルキャッシュのファイル数が配布元と一致しません。' }
        $installFiles = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $sourceFiles.Count; $index++) {
            $sourceFile = $sourceFiles[$index]
            $stagedFile = $stagedFiles[$index]
            if ([string]$sourceFile.RelativePath -ne [string]$stagedFile.RelativePath -or
                [long]$sourceFile.Length -ne [long]$stagedFile.Length) {
                throw "ローカルキャッシュの内容が配布元と一致しません: $($sourceFile.RelativePath)"
            }
            $sourceHash = (Get-FileHash -LiteralPath $sourceFile.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            $stagedHash = (Get-FileHash -LiteralPath $stagedFile.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not $sourceHash.Equals($stagedHash, [StringComparison]::OrdinalIgnoreCase)) {
                throw "ローカルキャッシュのハッシュが配布元と一致しません: $($sourceFile.RelativePath)"
            }
            [void]$installFiles.Add([pscustomobject]@{ path = [string]$sourceFile.RelativePath; length = [long]$sourceFile.Length; sha256 = $sourceHash.ToLowerInvariant() })
        }
        if ((Get-MbApplicationVersion -AppRoot $resolvedSourceRoot) -ne $sourceVersion) {
            throw 'コピー中に配布元のバージョンが変更されたため、更新を中止しました。'
        }

        $installState = [pscustomobject]@{
            schemaVersion = 1
            appVersion = $sourceVersion
            installedAt = [DateTime]::UtcNow.ToString('o')
            files = @($installFiles)
        } | ConvertTo-Json -Depth 5
        [IO.File]::WriteAllText((Join-Path $stagingRoot '.install.json'), $installState, (New-Object Text.UTF8Encoding($false)))
        if (-not (Test-MbCachedApplication -CacheRoot $stagingRoot -ExpectedVersion $sourceVersion)) {
            throw '検証済みローカルキャッシュを作成できませんでした。'
        }

        if (Test-Path -LiteralPath $resolvedCacheRoot -PathType Container) {
            if (Test-Path -LiteralPath $previousRoot -PathType Container) {
                Remove-Item -LiteralPath $previousRoot -Recurse -Force -ErrorAction Stop
            }
            [IO.Directory]::Move($resolvedCacheRoot, $previousRoot)
            $movedCurrent = $true
        }
        [IO.Directory]::Move($stagingRoot, $resolvedCacheRoot)
        $installedNew = $true
        if (-not (Test-MbCachedApplication -CacheRoot $resolvedCacheRoot -ExpectedVersion $sourceVersion)) {
            throw '切替後のローカルキャッシュを検証できませんでした。'
        }
        $localLauncher = Write-MbLocalLauncher -ProductRoot $productRoot
        return [pscustomobject]@{ Updated = $true; Version = $sourceVersion; CacheRoot = $resolvedCacheRoot; SourceIsCache = $false; LocalLauncher = $localLauncher }
    } catch {
        if ($installedNew -and (Test-Path -LiteralPath $resolvedCacheRoot -PathType Container)) {
            Remove-Item -LiteralPath $resolvedCacheRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (($movedCurrent -or $installedNew) -and
            -not (Test-Path -LiteralPath $resolvedCacheRoot) -and
            (Test-Path -LiteralPath $previousRoot -PathType Container)) {
            [IO.Directory]::Move($previousRoot, $resolvedCacheRoot)
        }
        throw
    } finally {
        if (Test-Path -LiteralPath $stagingRoot -PathType Container) {
            Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Export-ModuleMember -Function @(
    'Get-MbApplicationVersion',
    'Get-MbDefaultAppCacheRoot',
    'Install-MbLocalApplication',
    'Test-MbCachedApplication'
)
