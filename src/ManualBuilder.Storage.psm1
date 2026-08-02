# ManualBuilder per-user storage and legacy migration.

Set-StrictMode -Version 2.0

function Get-MbDefaultDataRoot {
    param([AllowEmptyString()][string]$LocalApplicationData = '')

    if ([string]::IsNullOrWhiteSpace($LocalApplicationData)) {
        $LocalApplicationData = [Environment]::GetFolderPath('LocalApplicationData')
    }
    if ([string]::IsNullOrWhiteSpace($LocalApplicationData)) {
        throw 'ユーザーのLocalApplicationDataフォルダーを取得できません。'
    }
    return [IO.Path]::GetFullPath((Join-Path $LocalApplicationData 'ManualBuilder\data'))
}

function Get-MbStorageLayout {
    param(
        [Parameter(Mandatory = $true)][string]$AppRoot,
        [AllowEmptyString()][string]$DataRoot = '',
        [AllowEmptyString()][string]$ProjectPath = '',
        [AllowEmptyString()][string]$LegacyAppRoot = ''
    )

    $resolvedAppRoot = [IO.Path]::GetFullPath($AppRoot)
    $usesDefaultUserData = [string]::IsNullOrWhiteSpace($DataRoot) -and [string]::IsNullOrWhiteSpace($ProjectPath)
    if ([string]::IsNullOrWhiteSpace($DataRoot)) {
        if ([string]::IsNullOrWhiteSpace($ProjectPath)) {
            $DataRoot = Get-MbDefaultDataRoot
        } else {
            $DataRoot = Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))
        }
    }
    $resolvedDataRoot = [IO.Path]::GetFullPath($DataRoot)
    $resolvedLegacyAppRoot = if ([string]::IsNullOrWhiteSpace($LegacyAppRoot)) {
        $resolvedAppRoot
    } else {
        [IO.Path]::GetFullPath($LegacyAppRoot)
    }
    if ([string]::IsNullOrWhiteSpace($ProjectPath)) {
        $ProjectPath = Join-Path $resolvedDataRoot 'projects\default\project.json'
    }

    return [pscustomobject]@{
        AppRoot = $resolvedAppRoot
        DataRoot = $resolvedDataRoot
        ProjectPath = [IO.Path]::GetFullPath($ProjectPath)
        RuntimePath = Join-Path $resolvedDataRoot 'runtime.json'
        ExportJobsRoot = Join-Path $resolvedDataRoot 'export-jobs'
        LegacyProjectPath = Join-Path $resolvedLegacyAppRoot 'data\projects\default\project.json'
        UsesDefaultUserData = $usesDefaultUserData
    }
}

function Initialize-MbUserStorage {
    param(
        [Parameter(Mandatory = $true)][object]$Layout,
        [switch]$MigrateLegacy
    )

    [void](New-Item -ItemType Directory -Path ([string]$Layout.DataRoot) -Force)
    $projectsRoot = Join-Path ([string]$Layout.DataRoot) 'projects'
    [void](New-Item -ItemType Directory -Path $projectsRoot -Force)

    $destinationProjectPath = [IO.Path]::GetFullPath([string]$Layout.ProjectPath)
    $legacyProjectPath = [IO.Path]::GetFullPath([string]$Layout.LegacyProjectPath)
    $migrationMarkerPath = Join-Path ([string]$Layout.DataRoot) 'legacy-default-migration.completed'
    $migrationNotApplicable = -not $MigrateLegacy -or
        $destinationProjectPath.Equals($legacyProjectPath, [StringComparison]::OrdinalIgnoreCase)
    $destinationExists = Test-Path -LiteralPath $destinationProjectPath -PathType Leaf
    $legacyExists = Test-Path -LiteralPath $legacyProjectPath -PathType Leaf
    if ($destinationExists -and $MigrateLegacy -and $legacyExists -and
        -not (Test-Path -LiteralPath $migrationMarkerPath -PathType Leaf)) {
        [IO.File]::WriteAllText($migrationMarkerPath, [DateTime]::UtcNow.ToString('o'), (New-Object Text.UTF8Encoding($false)))
    }
    if ($migrationNotApplicable -or $destinationExists -or
        (Test-Path -LiteralPath $migrationMarkerPath -PathType Leaf) -or -not $legacyExists) {
        return [pscustomobject]@{
            ProjectPath = $destinationProjectPath
            Migrated = $false
            LegacyProjectPath = $legacyProjectPath
        }
    }

    $sourceDirectory = Split-Path -Parent $legacyProjectPath
    $destinationDirectory = Split-Path -Parent $destinationProjectPath
    $destinationParent = Split-Path -Parent $destinationDirectory
    [void](New-Item -ItemType Directory -Path $destinationParent -Force)

    if (Test-Path -LiteralPath $destinationDirectory -PathType Container) {
        $existingEntries = @(Get-ChildItem -LiteralPath $destinationDirectory -Force -ErrorAction Stop)
        if ($existingEntries.Count -gt 0) {
            throw "ローカル保存先に未完成のデータがあります。自動移行を中止しました: $destinationDirectory"
        }
        Remove-Item -LiteralPath $destinationDirectory -Force
    }

    $stagingDirectory = Join-Path $destinationParent ('.migrate-default-' + [guid]::NewGuid().ToString('N'))
    try {
        [void](New-Item -ItemType Directory -Path $stagingDirectory)
        foreach ($entry in @(Get-ChildItem -LiteralPath $sourceDirectory -Force -ErrorAction Stop)) {
            Copy-Item -LiteralPath $entry.FullName -Destination $stagingDirectory -Recurse -Force -ErrorAction Stop
        }

        $stagedProjectPath = Join-Path $stagingDirectory 'project.json'
        if (-not (Test-Path -LiteralPath $stagedProjectPath -PathType Leaf)) {
            throw '移行先にproject.jsonをコピーできませんでした。'
        }
        [void]([IO.File]::ReadAllText($stagedProjectPath, [Text.Encoding]::UTF8) | ConvertFrom-Json)

        $sourceFiles = @(Get-ChildItem -LiteralPath $sourceDirectory -Recurse -File -Force -ErrorAction Stop)
        $stagedFiles = @(Get-ChildItem -LiteralPath $stagingDirectory -Recurse -File -Force -ErrorAction Stop)
        if ($sourceFiles.Count -ne $stagedFiles.Count) {
            throw '既存データのファイル数が移行元と移行先で一致しません。'
        }
        foreach ($sourceFile in $sourceFiles) {
            $trimCharacters = [char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
            $relativePath = $sourceFile.FullName.Substring($sourceDirectory.Length).TrimStart($trimCharacters)
            $stagedFile = Join-Path $stagingDirectory $relativePath
            if (-not (Test-Path -LiteralPath $stagedFile -PathType Leaf) -or
                [long](Get-Item -LiteralPath $stagedFile).Length -ne [long]$sourceFile.Length) {
                throw "既存データを完全にコピーできませんでした: $relativePath"
            }
            $sourceHash = (Get-FileHash -LiteralPath $sourceFile.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            $stagedHash = (Get-FileHash -LiteralPath $stagedFile -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not $sourceHash.Equals($stagedHash, [StringComparison]::OrdinalIgnoreCase)) {
                throw "既存データのハッシュが移行元と移行先で一致しません: $relativePath"
            }
        }

        [IO.Directory]::Move($stagingDirectory, $destinationDirectory)
        [IO.File]::WriteAllText($migrationMarkerPath, [DateTime]::UtcNow.ToString('o'), (New-Object Text.UTF8Encoding($false)))
        return [pscustomobject]@{
            ProjectPath = $destinationProjectPath
            Migrated = $true
            LegacyProjectPath = $legacyProjectPath
        }
    } finally {
        if (Test-Path -LiteralPath $stagingDirectory -PathType Container) {
            Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Export-ModuleMember -Function @(
    'Get-MbDefaultDataRoot',
    'Get-MbStorageLayout',
    'Initialize-MbUserStorage'
)
