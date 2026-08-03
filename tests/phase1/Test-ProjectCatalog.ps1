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

    # --- 反映先（共有フォルダー上の配布フォルダー）の記憶 ---
    Assert-Mb ((Get-MbPublishTarget -DataRoot $testRoot -ProjectKey ([string]$copied.Key)) -eq '') '反映先を知らないうちは空を返す'
    Set-MbPublishTarget -DataRoot $testRoot -ProjectKey ([string]$copied.Key) -Path '\\share\manuals\営業手順'
    Assert-Mb ((Get-MbPublishTarget -DataRoot $testRoot -ProjectKey ([string]$copied.Key)) -eq '\\share\manuals\営業手順') '反映先を記録する'
    # 反映先は前回開いたマニュアルの記録と同じファイルに入る。片方の保存でもう片方が消えてはいけない。
    Set-MbLastOpenedProject -DataRoot $testRoot -ProjectKey ([string]$created.Key)
    Assert-Mb ((Get-MbPublishTarget -DataRoot $testRoot -ProjectKey ([string]$copied.Key)) -eq '\\share\manuals\営業手順') '他の設定を保存しても反映先を消さない'
    Assert-Mb ([string](Get-MbWorkspaceSettings -DataRoot $testRoot).lastOpenedProjectKey -eq [string]$created.Key) '反映先を保存しても前回開いたマニュアルを消さない'
    Set-MbPublishTarget -DataRoot $testRoot -ProjectKey ([string]$copied.Key) -Path ''
    Assert-Mb ((Get-MbPublishTarget -DataRoot $testRoot -ProjectKey ([string]$copied.Key)) -eq '') '反映先を取り消せる'
    $badKeyRejected = $false
    try { Set-MbPublishTarget -DataRoot $testRoot -ProjectKey '../etc' -Path 'C:\temp' } catch { $badKeyRejected = $true }
    Assert-Mb $badKeyRejected '不正なマニュアルIDの反映先は受け付けない'

    # --- 配布フォルダーの元データ取り込み（何度取り込んでも増やさない） ---
    $sourceFolder = Join-Path $testRoot 'distributed-source'
    [void](New-Item -ItemType Directory -Path $sourceFolder -Force)
    # 別のPCで作られたマニュアルを配布フォルダーから受け取った状況にする。
    $distributed = Get-MbProject -Path ([string]$created.Path)
    $distributed.id = 'project-' + [guid]::NewGuid().ToString('N')
    $distributed.title = '配布されたマニュアル'
    $distributedImageRoot = Join-Path (Split-Path -Parent ([string]$created.Path)) 'images'
    if (Test-Path -LiteralPath $distributedImageRoot -PathType Container) {
        Copy-Item -LiteralPath $distributedImageRoot -Destination $sourceFolder -Recurse -Force
    }
    [void](Save-MbProject -Project $distributed -Path (Join-Path $sourceFolder 'project.json'))
    $beforeCount = @(Get-MbProjectCatalog -DataRoot $testRoot).Count
    $firstImport = Import-MbCatalogProjectFolder -DataRoot $testRoot -SourceFolder $sourceFolder
    Assert-Mb ([string]$firstImport.Status -eq 'imported') '知らないマニュアルは取り込む'
    Assert-Mb (@(Get-MbProjectCatalog -DataRoot $testRoot).Count -eq ($beforeCount + 1)) '取り込むと一覧が1件増える'
    $secondImport = Import-MbCatalogProjectFolder -DataRoot $testRoot -SourceFolder $sourceFolder
    Assert-Mb ([string]$secondImport.Status -eq 'existing') '同じマニュアルは取り込まず既存を返す'
    Assert-Mb ([string]$secondImport.Key -eq [string]$firstImport.Key) '2回目も同じマニュアルを開く'
    Assert-Mb (@(Get-MbProjectCatalog -DataRoot $testRoot).Count -eq ($beforeCount + 1)) '取り込み直しても一覧が増えない'

    # 別PCで共有版が更新された場合、古いローカル版をそのまま開くと再反映で巻き戻る。
    # 古い版を残したまま、共有版を識別できる別項目として開く。
    $sourceProjectPath = Join-Path $sourceFolder 'project.json'
    $newerDistributed = Get-MbProject -Path $sourceProjectPath
    $newerDistributed.title = '配布されたマニュアル 改訂'
    for ($revisionAttempt = 0; $revisionAttempt -lt 10 -and
        [int]$newerDistributed.revision -le [int]$secondImport.LocalRevision; $revisionAttempt++) {
        $newerDistributed = Save-MbProject -Project $newerDistributed -Path $sourceProjectPath
    }
    Assert-Mb ([int]$newerDistributed.revision -gt [int]$secondImport.LocalRevision) '共有側がローカルより新しい状態を再現する'
    $countBeforeNewerImport = @(Get-MbProjectCatalog -DataRoot $testRoot).Count
    $newerImport = Import-MbCatalogProjectFolder -DataRoot $testRoot -SourceFolder $sourceFolder
    Assert-Mb ([string]$newerImport.Status -eq 'imported-newer') '共有側が新しいときは古いローカル版を開かない'
    Assert-Mb ([string]$newerImport.Key -ne [string]$firstImport.Key) '新しい共有版を安全な別項目として取り込む'
    Assert-Mb ([string]$newerImport.Project.id -ne [string]$firstImport.Project.id) '共有版コピーへ衝突しないプロジェクトIDを付ける'
    Assert-Mb ([string]$newerImport.Project.title -eq '配布されたマニュアル 改訂 - 共有版') '共有版コピーを一覧で識別できる'
    Assert-Mb (Test-Path -LiteralPath ([string]$firstImport.Path) -PathType Leaf) '古いローカル版を削除しない'
    Assert-Mb (@(Get-MbProjectCatalog -DataRoot $testRoot).Count -eq ($countBeforeNewerImport + 1)) '共有版コピーだけを1件追加する'
    $newerImportAgain = Import-MbCatalogProjectFolder -DataRoot $testRoot -SourceFolder $sourceFolder
    Assert-Mb ([string]$newerImportAgain.Status -eq 'existing') '同じ共有版コピーを再度増やさない'
    Assert-Mb ([string]$newerImportAgain.Key -eq [string]$newerImport.Key) '再度開いた共有版を同じ項目へ戻す'
    Assert-Mb (@(Get-MbProjectCatalog -DataRoot $testRoot).Count -eq ($countBeforeNewerImport + 1)) '共有版を開き直しても一覧を増やさない'

    Assert-Mb ((Find-MbCatalogProjectById -DataRoot $testRoot -ProjectId 'project-00000000000000000000000000000000') -eq $null) '無いIDは見つからない'
    Assert-Mb ((Find-MbCatalogProjectById -DataRoot $testRoot -ProjectId '../etc') -eq $null) '不正なIDは探さない'
    $missingSourceRejected = $false
    try { [void](Import-MbCatalogProjectFolder -DataRoot $testRoot -SourceFolder (Join-Path $testRoot 'no-such-source')) }
    catch { $missingSourceRejected = $true }
    Assert-Mb $missingSourceRejected '元データが無ければ取り込まない'
    Assert-Mb (@(Get-ChildItem -LiteralPath (Join-Path $testRoot 'projects') -Directory -Force | Where-Object { $_.Name -like '.import-*' }).Count -eq 0) '取り込み作業用フォルダーを残さない'

    Write-Host ''
    Write-Host 'Project catalog tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $packagePath) { Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $unsafePackagePath) { Remove-Item -LiteralPath $unsafePackagePath -Force -ErrorAction SilentlyContinue }
}
