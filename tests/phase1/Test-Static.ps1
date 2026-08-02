# Phase 1 repository and PowerShell 5.1 compatibility checks.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

$scripts = @(
    Get-ChildItem -LiteralPath (Join-Path $repoRoot 'src') -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1') }
)
foreach ($script in $scripts) {
    $bytes = [IO.File]::ReadAllBytes($script.FullName)
    Add-Result ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) ("UTF-8 BOM: " + $script.Name)
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$tokens, [ref]$parseErrors)
    Add-Result (@($parseErrors).Count -eq 0) ("PowerShell構文: " + $script.Name)
    if (@($parseErrors).Count -gt 0) {
        foreach ($parseError in $parseErrors) { Write-Host ("     " + $parseError.Message) -ForegroundColor Red }
    }
}

$required = @(
    'run.cmd',
    'app-version.json',
    'src\ManualBuilder.Launcher.psm1',
    'src\Start-ManualBuilderLauncher.ps1',
    'src\ManualBuilder.Capture.psm1',
    'src\ManualBuilder.Storage.psm1',
    'src\ManualBuilder.Workspace.psm1',
    'src\ManualBuilder.Excel.psm1',
    'src\Export-ManualBuilderExcel.ps1',
    'src\ManualBuilder.Word.psm1',
    'src\Export-ManualBuilderWord.ps1',
    'web\index.html',
    'web\assets\css\app.css',
    'web\assets\js\app.js',
    'web\vendor\htmx-2.0.10.min.js',
    'web\vendor\HTMX-LICENSE.txt'
)
foreach ($relative in $required) {
    Add-Result (Test-Path -LiteralPath (Join-Path $repoRoot $relative) -PathType Leaf) ("必須ファイル: " + $relative)
}

$htmxPath = Join-Path $repoRoot 'web\vendor\htmx-2.0.10.min.js'
if (Test-Path -LiteralPath $htmxPath) {
    $htmxHash = (Get-FileHash -LiteralPath $htmxPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Add-Result ($htmxHash -eq '71ea67185bfa8c98c39d31717c6fce5d852370fcdfd129db4543774d3145c0de') '同梱htmx 2.0.10のSHA-256が一致する'
}

$serverText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\Start-ManualBuilder.ps1'), [Text.Encoding]::UTF8)
$launcherText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\Start-ManualBuilderLauncher.ps1'), [Text.Encoding]::UTF8)
$launcherModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Launcher.psm1'), [Text.Encoding]::UTF8)
$storageModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Storage.psm1'), [Text.Encoding]::UTF8)
$workspaceModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Workspace.psm1'), [Text.Encoding]::UTF8)
$captureModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1'), [Text.Encoding]::UTF8)
$projectModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Project.psm1'), [Text.Encoding]::UTF8)
$webModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Web.psm1'), [Text.Encoding]::UTF8)
$excelModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Excel.psm1'), [Text.Encoding]::UTF8)
$wordModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Word.psm1'), [Text.Encoding]::UTF8)
$cssText = [IO.File]::ReadAllText((Join-Path $repoRoot 'web\assets\css\app.css'), [Text.Encoding]::UTF8)
$jsText = [IO.File]::ReadAllText((Join-Path $repoRoot 'web\assets\js\app.js'), [Text.Encoding]::UTF8)
$indexText = [IO.File]::ReadAllText((Join-Path $repoRoot 'web\index.html'), [Text.Encoding]::UTF8)
$runCommandText = [IO.File]::ReadAllText((Join-Path $repoRoot 'run.cmd'), [Text.Encoding]::UTF8)
$appVersionText = [IO.File]::ReadAllText((Join-Path $repoRoot 'app-version.json'), [Text.Encoding]::UTF8)
$appVersionManifest = $appVersionText | ConvertFrom-Json
$expectedAppVersionPattern = [regex]::Escape([string]$appVersionManifest.appVersion)
Add-Result ($serverText -match 'Hostヘッダー') 'Host検証が実装されている'
Add-Result ($serverText -match 'X-Manual-Token') 'セッショントークン検証が実装されている'
Add-Result ($serverText -match 'Originが不正') 'Origin検証が実装されている'
Add-Result ($serverText -match 'HttpListener') 'HttpListenerをlocalhostで使用する'
Add-Result ($serverText -match 'ManualBuilder\.Storage\.psm1') 'ユーザー保存先モジュールを読み込む'
Add-Result ($serverText -match 'Get-MbStorageLayout') '実行時の保存先を一元解決する'
Add-Result ($serverText -match 'Initialize-MbUserStorage') '起動前にユーザー保存先を初期化する'
Add-Result ($storageModuleText -match "GetFolderPath\('LocalApplicationData'\)") '既定データをLocalApplicationData配下へ保存する'
Add-Result ($storageModuleText -match "'ManualBuilder\\data'") 'ユーザー別のManualBuilderデータフォルダーを使用する'
Add-Result (($storageModuleText -match 'Copy-Item') -and ($storageModuleText -match 'Get-FileHash') -and ($storageModuleText -match '\[IO\.Directory\]::Move')) '旧データを検証してから非破壊移行する'
Add-Result ($storageModuleText -match 'legacy-default-migration\.completed') 'defaultのアーカイブ後に旧データを再移行しない'
Add-Result (($serverText -match '\$projectReady = \$false') -and ($serverText -match 'if \(\$projectReady -and')) '移行失敗時に空の新規プロジェクトを作らない'
Add-Result ($serverText -match '\$storageLayout\.RuntimePath') '二重起動情報をユーザーデータ配下へ置く'
Add-Result ($serverText -match '\$storageLayout\.ExportJobsRoot') 'Office一時ジョブをユーザーデータ配下へ置く'
Add-Result (($runCommandText -match '%~dp0src\\Start-ManualBuilderLauncher\.ps1') -and ($runCommandText -notmatch '(?im)^cd /d')) 'UNC共有フォルダーから更新ランチャーを起動できる'
Add-Result ([string]$appVersionManifest.appVersion -eq '0.22.0') '配布用アプリバージョンを0.22.0へ更新する'
Add-Result ($workspaceModuleText -notmatch "ManualBuilder\.Project\.psm1'\) -Force") 'WorkspaceがProjectコマンドを強制再読込しない'
Add-Result ($launcherModuleText -match "'ManualBuilder\\app'") 'アプリ実行コードをLocalApplicationDataへキャッシュする'
Add-Result ($launcherModuleText -match "@\('src', 'web', 'run\.cmd', 'app-version\.json'\)") 'キャッシュ対象からプロジェクトデータを除外する'
Add-Result (($launcherModuleText -match '\.app-update-') -and ($launcherModuleText -match 'Get-FileHash') -and ($launcherModuleText -match '\[IO\.Directory\]::Move')) '検証後にローカル実行版を切り替える'
Add-Result ($launcherModuleText -match 'コピー中に配布元のバージョンが変更されたため') '配布元の更新途中を完成版として採用しない'
Add-Result ($launcherModuleText -match "\.previous'") '更新前のローカル実行版を1世代残す'
Add-Result ($launcherModuleText -match "'ManualBuilder\.cmd'") '共有フォルダーなしで使えるローカル起動ファイルを作る'
Add-Result (($launcherText -match 'Test-MbApplicationIsRunning') -and ($launcherText -match 'Show-MbAlreadyRunningNotice')) '起動中はローカル実行版を差し替えない'
Add-Result (($launcherText -match 'Test-MbCachedApplication -CacheRoot \$cacheRoot') -and ($launcherText -match "検証済みの既存ローカル版で起動します")) '共有更新失敗時は検証済みローカル版へフォールバックする'
Add-Result (($launcherText -match '-LegacyAppRoot \$sourceRoot') -and ($serverText -match '-LegacyAppRoot \$LegacyAppRoot')) 'ローカル起動後も共有元の旧データを移行できる'
Add-Result ($serverText -match 'FileSystemWatcher') 'スクリーンショット保存先の監視を実装する'
Add-Result ($serverText -match '/api/images/import') '生バイト画像取込みAPIを実装する'
Add-Result ($serverText -match '/api/images/replace') '画像差し替えAPIを実装する'
Add-Result ($serverText -match '/api/images/replace/undo') '画像差し替えの復元APIを実装する'
Add-Result ($serverText -match '/api/capture/heartbeat') '撮影対象タブのハートビートを実装する'
Add-Result ($serverText -match '/api/steps/reorder') '手順並べ替えAPIを実装する'
Add-Result ($serverText -match '/api/sheets/reorder') 'シート並べ替えAPIを実装する'
Add-Result ($serverText -match '/api/steps/move') '手順のシート移動APIを実装する'
Add-Result ($serverText -match '/api/steps/annotations') '注釈保存APIを実装する'
Add-Result ($serverText -match '/api/projects/create') '複数マニュアルの新規作成APIを実装する'
Add-Result ($serverText -match '/api/projects/open') 'マニュアル選択APIを実装する'
Add-Result ($serverText -match '/api/projects/duplicate') 'マニュアル複製APIを実装する'
Add-Result ($serverText -match '/api/projects/archive') '非破壊のアーカイブAPIを実装する'
Add-Result ($serverText -match '/api/projects/restore') 'アーカイブ復元APIを実装する'
Add-Result ($serverText -match '/api/projects/export') 'マニュアルZIP書き出しAPIを実装する'
Add-Result ($serverText -match '/api/projects/import') 'マニュアルZIP取込みAPIを実装する'
Add-Result ($workspaceModuleText -match 'projects-archive') 'アーカイブを別フォルダーへ保存する'
Add-Result ($workspaceModuleText -match 'lastOpenedProjectKey') '前回開いたマニュアルを記録する'
Add-Result (($workspaceModuleText -match 'Export-MbCatalogProjectPackage') -and ($workspaceModuleText -match 'Import-MbCatalogProjectPackage')) '画像を含むマニュアルをZIPで受け渡す'
Add-Result (($workspaceModuleText -match 'ZIP内のパスが不正') -and ($workspaceModuleText -match 'Test-MbCatalogProjectFiles')) '取込みZIPを展開前後に検証する'
Add-Result ($serverText -match '/api/export/excel/start') 'Excel出力開始APIを実装する'
Add-Result ($serverText -match '/api/export/excel/status') 'Excel出力進捗APIを実装する'
Add-Result ($serverText -match '/api/export/excel/cancel') 'Excel出力中止APIを実装する'
Add-Result ($serverText -match 'Export-ManualBuilderExcel\.ps1') 'Excel出力を別プロセスで実行する'
Add-Result ($serverText -match '/api/export/word/start') 'Word出力開始APIを実装する'
Add-Result ($serverText -match '/api/export/word/status') 'Word出力進捗APIを実装する'
Add-Result ($serverText -match '/api/export/word/cancel') 'Word出力中止APIを実装する'
Add-Result ($serverText -match 'Export-ManualBuilderWord\.ps1') 'Word出力を別プロセスで実行する'
Add-Result ($wordModuleText -match 'GetWindowThreadProcessId') 'Wordの所有PIDをHwndから確認する'
Add-Result ($wordModuleText -match '\$pidsBefore\.Count -gt 0') 'Word起動中はCOM生成前に安全停止する'
Add-Result ($wordModuleText -match 'New-MbAnnotatedImage') 'Word用画像へ注釈と切り抜きを反映する'
Add-Result ($wordModuleText -match 'TablesOfContents\.Add') 'Wordへ自動目次を作成する'
Add-Result ($wordModuleText -match '\.tmp\.docx') 'Wordを一時名で保存してから完成扱いにする'
Add-Result ($wordModuleText -match 'AlternativeText') 'Word画像へ代替テキストを設定する'
Add-Result ($wordModuleText -match 'Get-MbWordContentSheets') 'Word出力から空シートを除外する'
Add-Result ($wordModuleText -match '\[Math\]::Min\(2\.0, \[Math\]::Min\(450\.0 /') 'Wordで小さな切り抜き画像を最大2倍まで拡大する'
Add-Result ($wordModuleText -match '\$imageParagraph\.Alignment = 1') 'Word画像を本文中央へ配置する'
Add-Result ($wordModuleText -match '\$footerParagraph\.Alignment = 1') 'Wordページ番号の段落を中央揃えにする'
Add-Result ($excelModuleText -match 'GetWindowThreadProcessId') 'Excelの所有PIDをHwndから確認する'
Add-Result ($excelModuleText -match 'MB_CONNECTED_TO_EXISTING_EXCEL') '既存Excelへ接続した場合は安全停止する'
Add-Result ($excelModuleText -match 'New-MbAnnotatedImage') 'Excel用画像へ注釈を合成する'
Add-Result ($excelModuleText -match '\.tmp\.xlsx') 'Excelを一時名で保存してから完成扱いにする'
Add-Result ($excelModuleText -match '\$window\.Zoom = 100') 'Excel出力の標準ズームを100%にする'
Add-Result ($excelModuleText -match 'A\$\{contentStart\}:G\$\{contentEnd\}') 'Excel画像領域を読みやすい約半幅へ調整する'
Add-Result (($excelModuleText -match '\$textColumn = if \(\$hasImage\) \{ ''H'' \} else \{ ''A'' \}') -and ($excelModuleText -match '\$\{textColumn\}\$\{descriptionLabelRow\}:L\$\{descriptionLabelRow\}')) 'Excel説明領域を5列へ拡大する'
Add-Result (($excelModuleText -match '\$imageRows = if \(\$HasImage\) \{ 6 \} else \{ 0 \}') -and ($excelModuleText -notmatch "imageArea\.Value2 = '画像なし'")) '画像なし手順をExcelで全幅の文章カードにする'
Add-Result (($excelModuleText -match 'Get-MbExcelStepCardLayout') -and ($excelModuleText -match '\$contentEnd = \$contentStart \+ \[int\]\$layout\.ContentRows - 1') -and ($excelModuleText -match '\$startRow = Add-MbExcelStepCard')) 'Excelカードの高さを画像と文章に合わせて可変化する'
Add-Result (($excelModuleText -match 'Columns\.Item\(1\)\.ColumnWidth = 9') -and ($excelModuleText -match 'foreach \(\$column in 2\.\.7\).*ColumnWidth = 15') -and ($excelModuleText -match 'foreach \(\$column in 8\.\.12\).*ColumnWidth = 20')) 'Excelの画像と説明を約半幅ずつへ再配分する'
Add-Result (($excelModuleText -match '\$descriptionArea\.Font\.Size = 12') -and ($excelModuleText -match '\$noteArea\.Font\.Size = 11')) 'Excelの説明と補足を読みやすい文字サイズにする'
Add-Result (($excelModuleText -match 'TargetDisplayWidth') -and ($excelModuleText -match '\$annotationUnit = 0\.62') -and ($jsText -match 'const unit = 0\.62') -and ($excelModuleText -match '7\.0 \* \$annotationUnit') -and ($excelModuleText -match '24\.0 \* \$annotationUnit') -and ($excelModuleText -match 'New-MbRoundedRectanglePath') -and ($excelModuleText -notmatch '\$minimumSide')) 'アプリとOfficeの注釈を画像比率に依存しない共通寸法へ揃える'
Add-Result (($excelModuleText -match '\$descriptionLines \* 16\.0') -and ($excelModuleText -match '\$noteLines \* 15\.0') -and ($excelModuleText -match '\$visualLength') -and ($excelModuleText -match '0\.55')) '長文を文字幅と物理行高から見積もって余白を抑える'
Add-Result (($excelModuleText -match '\$descriptionEnd = \[Math\]::Min\(\$contentEnd, \$descriptionStart \+ \[int\]\$layout\.DescriptionBodyRows - 1\)') -and ($excelModuleText -match '\$noteLabelRow = if \(\$hasNote\) \{ \$descriptionEnd \+ 1 \}') -and ($excelModuleText -match 'NoteBodyRows')) 'Excelの説明と補足を必要な高さだけ表示する'
Add-Result (($excelModuleText -match '\$maximumImageScale = 1\.5') -and ($excelModuleText -match 'MaximumDisplayScale 1\.5')) 'Excelで小さな元画像の拡大を最大1.5倍に抑える'
Add-Result (($excelModuleText -match '\$compactImageWidthRatio = if .*?-ge 3\.0.*?0\.85') -and ($excelModuleText -match '\$compactImageHeightRatio = if .*?-ge 3\.0.*?0\.85') -and ($excelModuleText -match '\$renderTargetWidth = if .*?646.*?760') -and ($excelModuleText -match '\$renderTargetHeight = if .*?620.*?880')) 'Excelで極端に細長い画像を長辺方向85%へ抑える'
Add-Result ($excelModuleText -match 'if \(\$hasNote\)') '補足がある場合だけExcelへ補足欄を出す'
Add-Result ($excelModuleText -match '\$startRow = if \(\[string\]::IsNullOrWhiteSpace\(\$summaryText\)\) \{ 2 \} else \{ 3 \}') '空のシート概要で不要な行を残さない'
Add-Result (($excelModuleText -match '\$titleRange\.Interior\.Color = \$colorAccentDark') -and ($excelModuleText -match '\$sheetHeader\.Interior\.Color = \$colorWhite') -and ($excelModuleText -match '\$sheetHeader.*-Weight -4138')) 'Excelの濃紺を目次に限定して手順見出しを軽くする'
Add-Result (($excelModuleText -match 'NumberFormat = .*STEP.*00') -and ($excelModuleText -match '\$headerBand\.Interior\.Color = \$colorWhite')) 'Excel手順カードへ白地の見出し階層を付ける'
Add-Result ($excelModuleText -notmatch 'Weight 3') 'Excel罫線に未定義のWeight 3を使用しない'
Add-Result (($excelModuleText -match '\$indexSheet\.Tab\.Color') -and ($excelModuleText -match '\$worksheet\.Tab\.Color')) 'Excelのシートタブへ文書テーマ色を付ける'
Add-Result ($excelModuleText -notmatch '\[IO\.File\]::Replace\([^\r\n]*\$null') 'Excel進捗JSONを有効なバックアップパスで置換する'
Add-Result ($serverText -notmatch '\[IO\.File\]::Replace\([^\r\n]*\$null') 'サーバー進捗JSONを有効なバックアップパスで置換する'
Add-Result ($serverText -notmatch 'Start-Process \(\[string\]\$runtime\.url\)') '二重起動時に新しいブラウザータブを開かない'
Add-Result ($serverText -match '既存のブラウザータブへ戻ってください') '二重起動時の案内が実装されている'
Add-Result ($projectModuleText -match "'Remove-MbSheet'") '承認済み動詞のシート削除コマンドを公開する'
Add-Result ($projectModuleText -match "'Remove-MbStep'") '承認済み動詞の手順削除コマンドを公開する'
Add-Result ($projectModuleText -match "'Set-MbStepAnnotations'") '注釈保存コマンドを公開する'
Add-Result ($projectModuleText -match "'Set-MbSheetOrder'") 'シート並べ替えコマンドを公開する'
Add-Result ($projectModuleText -match "'Move-MbStepToSheet'") '手順のシート移動コマンドを公開する'
Add-Result ($projectModuleText -match "'Set-MbStepImageEdits'") '切り抜きと注釈の保存コマンドを公開する'
Add-Result ($captureModuleText -match "'Set-MbStepImage'") '画像差し替えコマンドを公開する'
Add-Result ($captureModuleText -match "'Restore-MbStepImage'") '元画像の復元コマンドを公開する'
Add-Result ($projectModuleText -notmatch "'Delete-Mb(?:Sheet|Step)'") '未承認動詞の削除コマンドを公開しない'
Add-Result ($serverText -match 'Remove-MbUnusedImage') '削除手順の未参照画像を整理する'
Add-Result ($webModuleText -notmatch "'Render-Mb") 'Webモジュールで未承認動詞のコマンドを公開しない'
Add-Result ($webModuleText -match 'class="action-menu') '破壊的操作をメニューへ整理する'
Add-Result (($webModuleText -notmatch 'class="capture-toolbar') -and ($webModuleText -match 'class="button button--primary editor-add-image"') -and ($webModuleText -match 'empty-state__button')) '画像追加を見出しと空状態へ整理する'
Add-Result (($webModuleText -notmatch 'editable-name__action|field__edit-action') -and ($cssText -match '\.name-field:hover::after') -and ($cssText -match '\.name-field:focus-within::after')) '名称入力欄自体で編集可能性を示す'
Add-Result (($webModuleText -match '>手順名<') -and ($webModuleText -notmatch '>タイトル</span>')) '手順名として編集対象を明示する'
Add-Result (($webModuleText -match '画像から追加</button>') -and ($webModuleText -match '>文字だけ追加</button>')) '手順追加を画像と文字の選択肢で表示する'
Add-Result (($webModuleText -match 'step-nav__title--fallback') -and ($jsText -match 'fallbackTitle')) '手順名が空なら説明の先頭をナビへ表示する'
Add-Result (($cssText -match '\.step-nav--sorting \.step-nav__guide') -and ($cssText -match '\.sheet-nav--sorting \.sheet-nav__guide') -and ($webModuleText -notmatch 'sidebar__hint')) '並べ替えガイドをドラッグ中だけ表示する'
Add-Result (($cssText -match '\.step-nav__item--complete \.step-nav__status\s*\{[^}]*display:\s*none') -and ($cssText -match 'box-shadow:\s*inset 2px 0 0 var\(--accent\)')) '未完了状態と選択中の手順を優先表示する'
Add-Result (($webModuleText -match 'step-card--no-image') -and ($cssText -match '\.step-card--no-image \.image-placeholder')) '画像なし手順の空白を縮小する'
Add-Result (($webModuleText -match 'data-add-image-to-step') -and ($jsText -match 'step-card--active \.image-placeholder')) '空の手順へ画像を直接追加できる'
Add-Result (($jsText -match "replaceStepImage\(file, emptyCard\.dataset\.stepId, 'paste'\)") -and ($jsText -match "replaceStepImage\(supported\[0\], emptyCard\.dataset\.stepId, 'drop'\)")) '空の手順へ貼り付けとドロップで画像を設定する'
Add-Result (($webModuleText -match 'aria-label="画像の操作"') -and ($cssText -match '\.image-edit-actions\s*\{[^}]*position:\s*absolute')) '画像上に編集操作を配置する'
Add-Result ($webModuleText -notmatch '次の工程で接続') '未実装を示す古い案内を表示しない'
Add-Result ($webModuleText -match 'data-step-nav-drag-handle') '左ナビへ手順のドラッグ操作を表示する'
Add-Result ($webModuleText -match 'data-sheet-nav-drag-handle') '左ナビへシートのドラッグ操作を表示する'
Add-Result ($webModuleText -match 'data-sheet-drop-target') '手順を移動できるシートを表示する'
Add-Result ($webModuleText -match 'data-sheet-sort-guide') 'シート並べ替えの操作案内を表示する'
Add-Result ($webModuleText -match 'data-step-sort-guide') '手順並べ替えの操作案内を表示する'
Add-Result ($webModuleText -match ('sidebar__version.*v' + $expectedAppVersionPattern)) '画面へ適用バージョンを表示する'
Add-Result ($webModuleText -match ('data-app-version="' + $expectedAppVersionPattern + '"')) 'サーバー側のアプリバージョンを画面へ埋め込む'
Add-Result (($indexText -match ('app\.css\?v=' + $expectedAppVersionPattern)) -and ($indexText -match ('app\.js\?v=' + $expectedAppVersionPattern))) 'CSSとJavaScriptの更新URLを切り替える'
Add-Result ($webModuleText -match 'ConvertTo-MbProjectLibraryHtml') '起動時のマニュアル一覧を実装する'
Add-Result ($jsText -match 'data-project-search') 'マニュアル名の検索を実装する'
Add-Result (($webModuleText -match 'data-project-export') -and ($webModuleText -match 'data-import-project-package')) '一覧からマニュアルZIPを書き出し・取り込める'
Add-Result (($jsText -match '/api/projects/export') -and ($jsText -match '/api/projects/import')) 'マニュアルZIPの画面処理を実装する'
Add-Result (($webModuleText -notmatch '共有フォルダー上のアプリ') -and ($webModuleText -notmatch '%LOCALAPPDATA%')) '一覧画面に開発者向けの保存・配布説明を表示しない'
Add-Result (($webModuleText -match 'class="brand brand--home"') -and ($webModuleText -match 'data-project-home') -and ($webModuleText -notmatch 'project-home-button')) '左上のアプリロゴを一覧へ戻るホーム操作にする'
Add-Result ($cssText -match '(?m)^\[hidden\]\s*\{[^}]*display:\s*none\s*!important') 'hidden属性を常に非表示として扱う'
Add-Result (($jsText -match 'ensureCurrentAssets') -and ($jsText -match 'window\.location\.replace')) '旧アセットを検出したタブを自動再読込する'
Add-Result ($webModuleText -match 'data-step-nav-delete') '左ナビへ手順の削除操作を常設する'
Add-Result (($webModuleText -notmatch 'data-step-nav-move') -and ($jsText -notmatch 'data-step-nav-move')) '左ナビの重複する矢印移動を表示しない'
Add-Result ($webModuleText -match 'data-step-jump') '手順ナビゲーションを表示する'
Add-Result ($webModuleText -match 'data-open-annotation') '画像へ注釈編集ボタンを表示する'
Add-Result ($webModuleText -match '切り抜き・赤枠・矢印・番号・黒塗り') '画像編集機能を見つけやすく表示する'
Add-Result (($webModuleText -match '>画像を編集<') -and ($webModuleText -match 'aria-label="画像の操作"')) '画像編集を明確な主操作として表示する'
Add-Result ($webModuleText -match 'data-replace-image') '画像差し替え操作を手順カードへ表示する'
Add-Result ($webModuleText -match 'data-undo-image-replace') '元画像へ戻す操作を手順カードへ表示する'
Add-Result ($webModuleText -match 'replacement-image-file-input') '差し替え画像の選択入力を実装する'
Add-Result ($jsText -match "addEventListener\('dragstart'") '手順のドラッグ並べ替えを実装する'
Add-Result ($jsText -match 'queueSheetOrderSave') 'シートのドラッグ並べ替えを保存する'
Add-Result ($jsText -match 'moveStepToSheet') '手順を別シートへドラッグ移動する'
Add-Result (($jsText -match 'data-sheet-drop-placeholder') -and ($jsText -match 'positionSheetDropPlaceholder') -and ($jsText -match 'drop\.position.*番目')) 'シートのドロップ位置と順位を表示する'
Add-Result ($jsText -match 'data-step-drop-placeholder') '手順のドロップ位置へ挿入線を表示する'
Add-Result (($jsText -match 'getStepDropPosition') -and ($jsText -match '番目へ移動')) '移動後の正確な手順位置を案内する'
Add-Result (($jsText -match 'positionStepDropPlaceholder') -and ($jsText -match 'clientY < bounds\.top \+ bounds\.height / 2') -and ($jsText -match 'drop\.atEnd')) '手順一覧の隙間でも正確な挿入位置を表示する'
Add-Result (($cssText -match '\.sheet-nav__drop-placeholder \{[\s\S]*?height: 0;[\s\S]*?pointer-events: none') -and ($cssText -match '\.step-nav__drop-placeholder \{[\s\S]*?height: 0;[\s\S]*?pointer-events: none')) 'シートと手順の挿入線で一覧を押し広げない'
Add-Result (($jsText -match "addEventListener\('dragenter'.*?[\s\S]*?preventDefault") -and ($jsText -match 'validDrop') -and ($jsText -match 'stepDragState\.validDrop && placeholder')) 'ドラッグ中の禁止カーソル点滅と無効位置へのドロップを防ぐ'
Add-Result ($jsText -match 'autoScrollStepNavigation') '長い手順一覧をドラッグ中に自動スクロールする'
Add-Result (($jsText -match 'targetSheetName') -and ($jsText -match '末尾へ移動しました')) '別シート移動の行き先と完了を通知する'
Add-Result ($cssText -match '\.toast--success') '手順移動の完了通知を表示する'
Add-Result ($jsText -match 'data-image-preview') 'スクリーンショット拡大表示を実装する'
Add-Result (($jsText -match 'maximumCardImageScale = 1\.25') -and ($jsText -match 'isExtremeWideImage = cropAspectRatio >= 3') -and ($jsText -match 'isExtremePortraitImage = cropAspectRatio <= \(1 / 3\)') -and ($jsText -match 'maximumRenderedHeight = height \* compactImageHeightRatio') -and ($cssText -match 'step-image-frame--extreme-wide')) '編集画面で小さく細長い画像の過拡大と余白を抑える'
$annotationTypesImplemented =
    ($jsText -match "annotation\.type === 'rect'") -and
    ($jsText -match "annotation\.type === 'arrow'") -and
    ($jsText -match "annotation\.type === 'number'") -and
    ($jsText -match "annotation\.type === 'blackout'")
Add-Result $annotationTypesImplemented '赤枠・赤矢印・番号・黒塗り注釈を実装する'
Add-Result (($jsText -match 'const nextNumberLabelInSheet') -and ($jsText -match 'const usedNumberLabelsInSheet') -and ($jsText -notmatch 'Math\.min\(99, Math\.max\(0')) '番号注釈をシート内の手順をまたいで連番にする'
Add-Result (($jsText -match 'data-annotation-number\b') -and ($jsText -match 'applySelectedNumberLabel')) '選択した番号注釈を任意の番号へ変更できる'
Add-Result ($jsText -match 'renderAnnotations') 'SVG注釈レイヤーを実装する'
Add-Result ($jsText -match 'data-annotation-tool="crop"') '切り抜きツールを実装する'
Add-Result ($jsText -match 'queueImageEditSave') '画像編集を自動保存する'
Add-Result ($jsText -notmatch 'data-annotation-save') '画像編集で手動保存ボタンを要求しない'
Add-Result ($jsText -match "annotationEditor\.tool = 'select'") '注釈作成後に選択へ切り替える'
Add-Result (($jsText -match 'replaceStepImage') -and ($jsText -match 'undoStepImageReplacement')) '画像差し替えと1段階復元の画面処理を実装する'
Add-Result ($webModuleText -match 'data-export-excel') 'Excel作成ボタンを有効にする'
Add-Result ($jsText -match 'excel-export-dialog') 'Excel出力の進捗画面を実装する'
Add-Result (($cssText -match 'export-status-spin') -and ($jsText -notmatch '>↗</span>') -and ($jsText -match "state === 'cancelled' \? '×' : ''")) 'Office出力の処理中・完了・失敗状態を明確な記号で表示する'
Add-Result ($jsText -match '/api/export/excel/open') '出力したExcelと保存先を開ける'
Add-Result ($webModuleText -match 'data-export-word') 'Word副出力をメニューから選べる'
Add-Result ($jsText -match 'word-export-dialog') 'Word出力の進捗画面を実装する'
Add-Result ($jsText -match '/api/export/word/open') '出力したWordと保存先を開ける'
Add-Result ($jsText -match 'data-word-export-fallback') 'Word安全停止時にExcel出力を案内する'
Add-Result (($jsText -match 'const safeStop') -and ($jsText -match '開いているWord文書とManualBuilderの入力内容には影響していません')) 'Word安全停止の要約と対処を重複なく表示する'
Add-Result (($cssText -match '\.step-card--active') -and ($webModuleText -notmatch 'data-view-mode')) '1手順への集中表示へ一本化する'
Add-Result ($excelModuleText -match '\$requiresCrop') 'Excel出力へ切り抜きを反映する'
Add-Result ($excelModuleText -match '\[AllowEmptyCollection\(\)\]\[object\[\]\]\$Annotations') '注釈なしの切り抜き画像を許可する'
Add-Result ($cssText -match '--accent: #3a5ba0') 'ミニマルUIのアクセントトークンを使用する'
Add-Result ($cssText -notmatch 'linear-gradient') 'グラデーションを使用しない'

if ($errors.Count -gt 0) {
    Write-Host ''
    Write-Host ("Static tests failed: " + $errors.Count) -ForegroundColor Red
    exit 1
}
Write-Host ''
Write-Host 'Static tests passed.' -ForegroundColor Cyan
