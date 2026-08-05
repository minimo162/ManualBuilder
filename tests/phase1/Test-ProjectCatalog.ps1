# ManualBuilder multi-project catalog tests.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$testRoot = Join-Path $env:TEMP ('ManualBuilder-ProjectCatalog-' + [guid]::NewGuid().ToString('N'))
$packagePath = Join-Path $env:TEMP ('ManualBuilder-ProjectPackage-' + [guid]::NewGuid().ToString('N') + '.zip')
$unsafePackagePath = Join-Path $env:TEMP ('ManualBuilder-UnsafePackage-' + [guid]::NewGuid().ToString('N') + '.zip')

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
    Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Workspace.psm1') -Force
    Assert-Mb ($null -ne (Get-Command New-MbProject -ErrorAction SilentlyContinue)) 'Workspace読込後もProjectコマンドを呼び出せる'
    [void](New-Item -ItemType Directory -Path $testRoot -Force)

    $defaultPath = Get-MbCatalogProjectPath -DataRoot $testRoot -ProjectKey 'default'
    $default = New-MbProject
    Set-MbProjectTitle -Project $default -Title '既存マニュアル'
    [void](Add-MbStep -Project $default -SheetId ([string]$default.selectedSheetId))
    [void](Save-MbProject -Project $default -Path $defaultPath)

    $initial = @(Get-MbProjectCatalog -DataRoot $testRoot)
    Assert-Mb ($initial.Count -eq 1) '既存のdefaultプロジェクトを一覧へ引き継ぐ'
    Assert-Mb ([string]$initial[0].title -eq '既存マニュアル') '一覧でマニュアル名を読み込む'
    Assert-Mb ([int]$initial[0].stepCount -eq 1) '一覧で手順数を集計する'

    $created = New-MbCatalogProject -DataRoot $testRoot -Title '経費精算マニュアル'
    Assert-Mb (Test-Path -LiteralPath ([string]$created.Path) -PathType Leaf) '新しいマニュアルを作成する'
    Assert-Mb ([string]$created.Project.title -eq '経費精算マニュアル') '新規作成時の名前を保存する'

    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $bitmap = New-Object Drawing.Bitmap 8, 6
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $imageStream = New-Object IO.MemoryStream
    try {
        $graphics.Clear([Drawing.Color]::CornflowerBlue)
        $bitmap.Save($imageStream, [Drawing.Imaging.ImageFormat]::Png)
        $imageBytes = $imageStream.ToArray()
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
        $imageStream.Dispose()
    }
    $imageResult = Add-MbImageStep -Project $created.Project -ProjectPath $created.Path -SheetId ([string]$created.Project.selectedSheetId) -Bytes $imageBytes -Source file
    $created.Project = Save-MbProject -Project $created.Project -Path $created.Path
    Assert-Mb ($imageResult.Status -eq 'added') '受け渡し対象へ画像付き手順を追加する'

    $copied = Copy-MbCatalogProject -DataRoot $testRoot -ProjectKey ([string]$created.Key)
    Assert-Mb ([string]$copied.Key -ne [string]$created.Key) '複製先に新しいフォルダーIDを付ける'
    Assert-Mb ([string]$copied.Project.id -ne [string]$created.Project.id) '複製先に新しいプロジェクトIDを付ける'
    Assert-Mb ([string]$copied.Project.title -eq '経費精算マニュアル - コピー') '複製したマニュアルを判別できる'

    $package = Export-MbCatalogProjectPackage -DataRoot $testRoot -ProjectKey ([string]$created.Key) -OutputPath $packagePath
    Assert-Mb (Test-Path -LiteralPath $packagePath -PathType Leaf) 'マニュアルをZIPへ書き出す'
    Assert-Mb ([string]$package.FileName -match '_ManualBuilder_\d{8}_\d{6}\.zip$') '受け渡し用ZIPに分かりやすい名前を付ける'
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $packageArchive = [IO.Compression.ZipFile]::OpenRead($packagePath)
    try {
        $packageEntries = @($packageArchive.Entries | ForEach-Object { $_.FullName })
        Assert-Mb ($packageEntries -contains 'manifest.json') 'ZIPへ形式情報を記録する'
        Assert-Mb ($packageEntries -contains 'project.json') 'ZIPへプロジェクトデータを格納する'
        Assert-Mb (@($packageEntries | Where-Object { $_ -match '^images/image-[a-f0-9]{32}\.png$' }).Count -eq 1) 'ZIPへ参照画像を格納する'
    } finally {
        $packageArchive.Dispose()
    }

    $imported = Import-MbCatalogProjectPackage -DataRoot $testRoot -PackagePath $packagePath
    Assert-Mb ([string]$imported.Key -ne [string]$created.Key) '取り込んだマニュアルへ新しいフォルダーIDを付ける'
    Assert-Mb ([string]$imported.Project.id -ne [string]$created.Project.id) '取り込んだマニュアルへ新しいプロジェクトIDを付ける'
    Assert-Mb ([string]$imported.Project.title -eq '経費精算マニュアル') '取り込み後もマニュアル名を維持する'
    Assert-Mb (@($imported.Project.sheets[0].steps).Count -eq @($created.Project.sheets[0].steps).Count) '取り込み後も手順を維持する'
    Assert-Mb (@($imported.Project.images).Count -eq 1) '取り込み後も画像メタデータを維持する'
    $importedImagePath = Join-Path (Split-Path -Parent ([string]$imported.Path)) ('images\' + [string]$imported.Project.images[0].fileName)
    Assert-Mb (Test-Path -LiteralPath $importedImagePath -PathType Leaf) '取り込み後も画像実体を維持する'

    $unsafeArchive = [IO.Compression.ZipFile]::Open($unsafePackagePath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $manifestEntry = $unsafeArchive.CreateEntry('manifest.json')
        $manifestWriter = [IO.StreamWriter]::new($manifestEntry.Open(), (New-Object Text.UTF8Encoding($false)))
        try { $manifestWriter.Write('{"packageType":"ManualBuilder.ProjectPackage","schemaVersion":1}') } finally { $manifestWriter.Dispose() }
        $unsafeEntry = $unsafeArchive.CreateEntry('../outside.txt')
        $unsafeWriter = [IO.StreamWriter]::new($unsafeEntry.Open(), (New-Object Text.UTF8Encoding($false)))
        try { $unsafeWriter.Write('unsafe') } finally { $unsafeWriter.Dispose() }
    } finally {
        $unsafeArchive.Dispose()
    }
    $unsafeRejected = $false
    try { [void](Import-MbCatalogProjectPackage -DataRoot $testRoot -PackagePath $unsafePackagePath) } catch { $unsafeRejected = $true }
    Assert-Mb $unsafeRejected 'ZIP内の不正なパスを拒否する'
    Assert-Mb (-not (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $testRoot) 'outside.txt'))) '不正なZIPを保存先外へ展開しない'

    [void](Move-MbCatalogProjectToArchive -DataRoot $testRoot -ProjectKey ([string]$created.Key))
    Assert-Mb (-not (Test-Path -LiteralPath ([string]$created.Path) -PathType Leaf)) 'アーカイブしたマニュアルを現役一覧から外す'
    $archived = @(Get-MbProjectCatalog -DataRoot $testRoot -Archived)
    Assert-Mb ($archived.Count -eq 1) 'アーカイブ一覧に表示する'

    $restoredPath = Restore-MbCatalogProject -DataRoot $testRoot -ProjectKey ([string]$created.Key)
    Assert-Mb (Test-Path -LiteralPath $restoredPath -PathType Leaf) 'アーカイブから非破壊で復元する'
    Assert-Mb (@(Get-MbProjectCatalog -DataRoot $testRoot -Archived).Count -eq 0) '復元後はアーカイブ一覧から外す'

    Set-MbLastOpenedProject -DataRoot $testRoot -ProjectKey ([string]$copied.Key)
    $settings = Get-MbWorkspaceSettings -DataRoot $testRoot
    Assert-Mb ([string]$settings.lastOpenedProjectKey -eq [string]$copied.Key) '前回開いたマニュアルを安全に記録する'
    Assert-Mb (@(Get-ChildItem -LiteralPath $testRoot -Filter '.settings-*.tmp' -File -Force).Count -eq 0) '設定保存後に一時ファイルを残さない'

    Set-MbLastOpenedProject -DataRoot $testRoot -ProjectKey ([string]$created.Key)
    $localSettings = Get-MbWorkspaceSettings -DataRoot $testRoot
    Assert-Mb ([string]$localSettings.lastOpenedProjectKey -eq [string]$created.Key) 'ローカルで最後に開いたマニュアルだけを記録する'
    Assert-Mb (-not ($localSettings.PSObject.Properties.Name -contains 'publishTargets')) '共有フォルダーの反映先を記録しない'

    Write-Host ''
    Write-Host 'Project catalog tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $packagePath) { Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $unsafePackagePath) { Remove-Item -LiteralPath $unsafePackagePath -Force -ErrorAction SilentlyContinue }
}
