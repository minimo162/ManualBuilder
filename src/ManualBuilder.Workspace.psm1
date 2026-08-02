# ManualBuilder multi-project catalog and non-destructive archive operations.

Set-StrictMode -Version 2.0

# Do not use -Force here. The server imports Project into its own session first;
# forcing a nested re-import would remove those commands from the caller scope.
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1')

function Test-MbProjectKey {
    param([Parameter(Mandatory = $true)][string]$ProjectKey)
    return $ProjectKey -match '^(default|project-[a-f0-9]{32})$'
}

function Get-MbProjectCollectionRoot {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [switch]$Archived
    )
    $name = if ($Archived) { 'projects-archive' } else { 'projects' }
    return [IO.Path]::GetFullPath((Join-Path $DataRoot $name))
}

function Get-MbCatalogProjectPath {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$ProjectKey,
        [switch]$Archived
    )
    if (-not (Test-MbProjectKey -ProjectKey $ProjectKey)) { throw 'マニュアルIDが不正です。' }
    $root = Get-MbProjectCollectionRoot -DataRoot $DataRoot -Archived:$Archived
    return [IO.Path]::GetFullPath((Join-Path (Join-Path $root $ProjectKey) 'project.json'))
}

function Test-MbCatalogProjectFiles {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath
    )
    Test-MbProject -Project $Project
    $imageRoot = Join-Path (Split-Path -Parent $ProjectPath) 'images'
    foreach ($image in @($Project.images)) {
        $imagePath = Join-Path $imageRoot ([string]$image.fileName)
        if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) {
            throw "画像ファイルが見つかりません: $($image.fileName)"
        }
        $hash = (Get-FileHash -LiteralPath $imagePath -Algorithm SHA256 -ErrorAction Stop).Hash
        if (-not $hash.Equals([string]$image.sha256, [StringComparison]::OrdinalIgnoreCase)) {
            throw "画像ファイルの内容が一致しません: $($image.fileName)"
        }
    }
}

function Get-MbProjectCatalog {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [switch]$Archived
    )
    $root = Get-MbProjectCollectionRoot -DataRoot $DataRoot -Archived:$Archived
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return @() }
    $entries = New-Object System.Collections.ArrayList
    foreach ($directory in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction Stop)) {
        $key = [string]$directory.Name
        if (-not (Test-MbProjectKey -ProjectKey $key)) { continue }
        $projectPath = Join-Path $directory.FullName 'project.json'
        if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) { continue }
        try {
            $project = Get-MbProject -Path $projectPath
            $stepCount = 0
            foreach ($sheet in @($project.sheets)) { $stepCount += @($sheet.steps).Count }
            [void]$entries.Add([pscustomobject]@{
                key = $key
                title = [string]$project.title
                sheetCount = @($project.sheets).Count
                stepCount = $stepCount
                imageCount = @($project.images).Count
                updatedAt = [string]$project.updatedAt
                createdAt = [string]$project.createdAt
                archived = [bool]$Archived
                readable = $true
                error = ''
            })
        } catch {
            [void]$entries.Add([pscustomobject]@{
                key = $key
                title = '読み込みできないマニュアル'
                sheetCount = 0
                stepCount = 0
                imageCount = 0
                updatedAt = $directory.LastWriteTimeUtc.ToString('o')
                createdAt = $directory.CreationTimeUtc.ToString('o')
                archived = [bool]$Archived
                readable = $false
                error = $_.Exception.Message
            })
        }
    }
    return @($entries | Sort-Object @{ Expression = { [DateTime]$_.updatedAt }; Descending = $true }, @{ Expression = { $_.title }; Descending = $false })
}

function New-MbCatalogProject {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [AllowEmptyString()][string]$Title = ''
    )
    $key = 'project-' + [guid]::NewGuid().ToString('N')
    $projectPath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $key
    $project = New-MbProject
    if (-not [string]::IsNullOrWhiteSpace($Title)) { Set-MbProjectTitle -Project $project -Title $Title.Trim() }
    $project = Save-MbProject -Project $project -Path $projectPath
    return [pscustomobject]@{ Key = $key; Path = $projectPath; Project = $project }
}

function Copy-MbCatalogProject {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$ProjectKey
    )
    $sourcePath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw '複製元のマニュアルが見つかりません。' }
    $sourceProject = Get-MbProject -Path $sourcePath
    Test-MbCatalogProjectFiles -Project $sourceProject -ProjectPath $sourcePath

    $newKey = 'project-' + [guid]::NewGuid().ToString('N')
    $projectsRoot = Get-MbProjectCollectionRoot -DataRoot $DataRoot
    [void](New-Item -ItemType Directory -Path $projectsRoot -Force)
    $destinationDirectory = Join-Path $projectsRoot $newKey
    $stagingDirectory = Join-Path $projectsRoot ('.copy-' + [guid]::NewGuid().ToString('N'))
    try {
        [void](New-Item -ItemType Directory -Path $stagingDirectory)
        foreach ($entry in @(Get-ChildItem -LiteralPath (Split-Path -Parent $sourcePath) -Force -ErrorAction Stop)) {
            Copy-Item -LiteralPath $entry.FullName -Destination $stagingDirectory -Recurse -Force -ErrorAction Stop
        }
        $stagedPath = Join-Path $stagingDirectory 'project.json'
        $copy = Get-MbProject -Path $stagedPath
        $now = [DateTime]::UtcNow.ToString('o')
        $copy.id = 'project-' + [guid]::NewGuid().ToString('N')
        $copy.title = ([string]$copy.title + ' - コピー')
        if ($copy.title.Length -gt 100) { $copy.title = $copy.title.Substring(0, 100) }
        $copy.revision = 0
        $copy.createdAt = $now
        $copy.updatedAt = $now
        if (Test-Path -LiteralPath "$stagedPath.bak" -PathType Leaf) { Remove-Item -LiteralPath "$stagedPath.bak" -Force }
        $copy = Save-MbProject -Project $copy -Path $stagedPath
        Test-MbCatalogProjectFiles -Project $copy -ProjectPath $stagedPath
        [IO.Directory]::Move($stagingDirectory, $destinationDirectory)
        return [pscustomobject]@{ Key = $newKey; Path = (Join-Path $destinationDirectory 'project.json'); Project = $copy }
    } finally {
        if (Test-Path -LiteralPath $stagingDirectory -PathType Container) {
            Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Move-MbCatalogProjectToArchive {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$ProjectKey
    )
    $sourcePath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw 'アーカイブするマニュアルが見つかりません。' }
    $project = Get-MbProject -Path $sourcePath
    Test-MbCatalogProjectFiles -Project $project -ProjectPath $sourcePath
    $archivePath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey -Archived
    if (Test-Path -LiteralPath (Split-Path -Parent $archivePath)) { throw '同じIDのアーカイブがすでにあります。' }
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent (Split-Path -Parent $archivePath)) -Force)
    [IO.Directory]::Move((Split-Path -Parent $sourcePath), (Split-Path -Parent $archivePath))
    return $archivePath
}

function Restore-MbCatalogProject {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$ProjectKey
    )
    $sourcePath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey -Archived
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw '復元するマニュアルが見つかりません。' }
    $project = Get-MbProject -Path $sourcePath
    Test-MbCatalogProjectFiles -Project $project -ProjectPath $sourcePath
    $destinationPath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey
    if (Test-Path -LiteralPath (Split-Path -Parent $destinationPath)) { throw '同じIDのマニュアルがすでにあります。' }
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent (Split-Path -Parent $destinationPath)) -Force)
    [IO.Directory]::Move((Split-Path -Parent $sourcePath), (Split-Path -Parent $destinationPath))
    return $destinationPath
}

function Initialize-MbProjectPackageAssembly {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
}

function Get-MbProjectPackageFileName {
    param([Parameter(Mandatory = $true)][string]$Title)
    $safe = $Title.Trim()
    foreach ($character in [IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    $safe = $safe.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'マニュアル' }
    if ($safe.Length -gt 60) { $safe = $safe.Substring(0, 60).TrimEnd() }
    return $safe + '_ManualBuilder_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.zip'
}

function Export-MbCatalogProjectPackage {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$ProjectKey,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )
    Initialize-MbProjectPackageAssembly
    $projectPath = Get-MbCatalogProjectPath -DataRoot $DataRoot -ProjectKey $ProjectKey
    if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) { throw '書き出すマニュアルが見つかりません。' }
    $project = Get-MbProject -Path $projectPath
    Test-MbCatalogProjectFiles -Project $project -ProjectPath $projectPath

    $fullOutputPath = [IO.Path]::GetFullPath($OutputPath)
    $outputDirectory = Split-Path -Parent $fullOutputPath
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    $tempPath = Join-Path $outputDirectory ('.manual-package-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::Open($tempPath, [IO.Compression.ZipArchiveMode]::Create)
        $manifestEntry = $archive.CreateEntry('manifest.json', [IO.Compression.CompressionLevel]::Optimal)
        $manifestStream = $manifestEntry.Open()
        $manifestWriter = [IO.StreamWriter]::new($manifestStream, (New-Object Text.UTF8Encoding($false)))
        try {
            $manifest = [pscustomobject]@{
                packageType = 'ManualBuilder.ProjectPackage'
                schemaVersion = 1
                exportedAt = [DateTime]::UtcNow.ToString('o')
                title = [string]$project.title
            } | ConvertTo-Json
            $manifestWriter.Write($manifest)
        } finally {
            $manifestWriter.Dispose()
        }
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $archive, $projectPath, 'project.json', [IO.Compression.CompressionLevel]::Optimal
        )
        $imageRoot = Join-Path (Split-Path -Parent $projectPath) 'images'
        foreach ($image in @($project.images)) {
            $imagePath = Join-Path $imageRoot ([string]$image.fileName)
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive, $imagePath, ('images/' + [string]$image.fileName), [IO.Compression.CompressionLevel]::Optimal
            )
        }
        $archive.Dispose()
        $archive = $null
        if (Test-Path -LiteralPath $fullOutputPath) { throw '同じ名前の書き出しファイルがすでにあります。' }
        [IO.File]::Move($tempPath, $fullOutputPath)
        return [pscustomobject]@{
            Path = $fullOutputPath
            FileName = Get-MbProjectPackageFileName -Title ([string]$project.title)
            ProjectKey = $ProjectKey
            ImageCount = @($project.images).Count
        }
    } finally {
        if ($archive) { $archive.Dispose() }
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Import-MbCatalogProjectPackage {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$PackagePath
    )
    Initialize-MbProjectPackageAssembly
    $fullPackagePath = [IO.Path]::GetFullPath($PackagePath)
    if (-not (Test-Path -LiteralPath $fullPackagePath -PathType Leaf)) { throw '取り込むZIPファイルが見つかりません。' }
    if ((Get-Item -LiteralPath $fullPackagePath).Length -gt (250 * 1024 * 1024)) { throw '取り込めるZIPは250MBまでです。' }

    $projectsRoot = Get-MbProjectCollectionRoot -DataRoot $DataRoot
    [void](New-Item -ItemType Directory -Path $projectsRoot -Force)
    $newKey = 'project-' + [guid]::NewGuid().ToString('N')
    $destinationDirectory = Join-Path $projectsRoot $newKey
    $stagingDirectory = Join-Path $projectsRoot ('.import-' + [guid]::NewGuid().ToString('N'))
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($fullPackagePath)
        if ($archive.Entries.Count -lt 2 -or $archive.Entries.Count -gt 25002) { throw 'ZIP内のファイル数が不正です。' }
        $entryNames = New-Object 'System.Collections.Generic.HashSet[string]'
        $uncompressedLength = [long]0
        $manifestEntry = $null
        $projectEntry = $null
        $imageEntries = New-Object System.Collections.ArrayList
        foreach ($entry in @($archive.Entries)) {
            $entryName = [string]$entry.FullName
            if ([string]::IsNullOrWhiteSpace($entryName) -or $entryName.Contains('\')) { throw 'ZIP内のパスが不正です。' }
            if (-not $entryNames.Add($entryName.ToLowerInvariant())) { throw 'ZIP内のファイル名が重複しています。' }
            $uncompressedLength += [long]$entry.Length
            if ($uncompressedLength -gt (500 * 1024 * 1024)) { throw 'ZIP展開後のサイズが500MBを超えています。' }
            if ($entryName -eq 'manifest.json') {
                if ($entry.Length -gt 65536) { throw 'ZIPのマニフェストが大きすぎます。' }
                $manifestEntry = $entry
            } elseif ($entryName -eq 'project.json') {
                if ($entry.Length -gt (10 * 1024 * 1024)) { throw 'ZIPのプロジェクトデータが大きすぎます。' }
                $projectEntry = $entry
            } elseif ($entryName -match '^images/image-[a-f0-9]{32}\.(png|jpg|bmp)$') {
                if ($entry.Length -lt 1 -or $entry.Length -gt (20 * 1024 * 1024)) { throw 'ZIP内の画像サイズが不正です。' }
                [void]$imageEntries.Add($entry)
            } else {
                throw "ZIPに未対応のファイルが含まれています: $entryName"
            }
        }
        if (-not $manifestEntry -or -not $projectEntry) { throw 'ManualBuilder用ZIPではありません。' }

        $manifestReader = [IO.StreamReader]::new($manifestEntry.Open(), [Text.Encoding]::UTF8, $true)
        try { $manifest = $manifestReader.ReadToEnd() | ConvertFrom-Json } finally { $manifestReader.Dispose() }
        if ([string]$manifest.packageType -ne 'ManualBuilder.ProjectPackage' -or [int]$manifest.schemaVersion -ne 1) {
            throw '未対応のManualBuilder ZIPです。'
        }

        [void](New-Item -ItemType Directory -Path $stagingDirectory)
        $stagedProjectPath = Join-Path $stagingDirectory 'project.json'
        $projectStream = $projectEntry.Open()
        $projectFile = [IO.File]::Open($stagedProjectPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $projectStream.CopyTo($projectFile) } finally { $projectFile.Dispose(); $projectStream.Dispose() }

        if ($imageEntries.Count -gt 0) { [void](New-Item -ItemType Directory -Path (Join-Path $stagingDirectory 'images')) }
        foreach ($entry in @($imageEntries)) {
            $imageName = [IO.Path]::GetFileName([string]$entry.FullName)
            $destinationImagePath = Join-Path (Join-Path $stagingDirectory 'images') $imageName
            $sourceStream = $entry.Open()
            $destinationStream = [IO.File]::Open($destinationImagePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try { $sourceStream.CopyTo($destinationStream) } finally { $destinationStream.Dispose(); $sourceStream.Dispose() }
        }

        $packageProject = Get-MbProject -Path $stagedProjectPath
        Test-MbCatalogProjectFiles -Project $packageProject -ProjectPath $stagedProjectPath
        $expectedImageNames = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($image in @($packageProject.images)) { [void]$expectedImageNames.Add(([string]$image.fileName).ToLowerInvariant()) }
        if ($expectedImageNames.Count -ne $imageEntries.Count) { throw 'ZIP内にプロジェクト未登録の画像があります。' }
        foreach ($entry in @($imageEntries)) {
            if (-not $expectedImageNames.Contains([IO.Path]::GetFileName([string]$entry.FullName).ToLowerInvariant())) {
                throw 'ZIP内にプロジェクト未登録の画像があります。'
            }
        }

        $now = [DateTime]::UtcNow.ToString('o')
        $packageProject.id = 'project-' + [guid]::NewGuid().ToString('N')
        $packageProject.revision = 0
        $packageProject.createdAt = $now
        $packageProject.updatedAt = $now
        Remove-Item -LiteralPath $stagedProjectPath -Force
        $packageProject = Save-MbProject -Project $packageProject -Path $stagedProjectPath
        Test-MbCatalogProjectFiles -Project $packageProject -ProjectPath $stagedProjectPath
        [IO.Directory]::Move($stagingDirectory, $destinationDirectory)
        return [pscustomobject]@{
            Key = $newKey
            Path = Join-Path $destinationDirectory 'project.json'
            Project = $packageProject
        }
    } finally {
        if ($archive) { $archive.Dispose() }
        if (Test-Path -LiteralPath $stagingDirectory -PathType Container) {
            Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-MbWorkspaceSettings {
    param([Parameter(Mandatory = $true)][string]$DataRoot)
    $path = Join-Path $DataRoot 'settings.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [pscustomobject]@{ schemaVersion = 1; lastOpenedProjectKey = '' }
    }
    try {
        $settings = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
        if ([int]$settings.schemaVersion -ne 1) { throw '設定形式が不正です。' }
        $key = [string]$settings.lastOpenedProjectKey
        if ($key -and -not (Test-MbProjectKey -ProjectKey $key)) { $key = '' }
        return [pscustomobject]@{ schemaVersion = 1; lastOpenedProjectKey = $key }
    } catch {
        return [pscustomobject]@{ schemaVersion = 1; lastOpenedProjectKey = '' }
    }
}

function Set-MbLastOpenedProject {
    param(
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [Parameter(Mandatory = $true)][string]$ProjectKey
    )
    if (-not (Test-MbProjectKey -ProjectKey $ProjectKey)) { throw 'マニュアルIDが不正です。' }
    [void](New-Item -ItemType Directory -Path $DataRoot -Force)
    $path = Join-Path $DataRoot 'settings.json'
    $tempPath = Join-Path $DataRoot ('.settings-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = "$path.bak"
    $json = [pscustomobject]@{ schemaVersion = 1; lastOpenedProjectKey = $ProjectKey } | ConvertTo-Json
    try {
        [IO.File]::WriteAllText($tempPath, $json, (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            [IO.File]::Replace($tempPath, $path, $backupPath, $true)
        } else {
            [IO.File]::Move($tempPath, $path)
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
    }
}

Export-ModuleMember -Function @(
    'Test-MbProjectKey',
    'Get-MbCatalogProjectPath',
    'Get-MbProjectCatalog',
    'New-MbCatalogProject',
    'Copy-MbCatalogProject',
    'Move-MbCatalogProjectToArchive',
    'Restore-MbCatalogProject',
    'Get-MbProjectPackageFileName',
    'Export-MbCatalogProjectPackage',
    'Import-MbCatalogProjectPackage',
    'Get-MbWorkspaceSettings',
    'Set-MbLastOpenedProject'
)
