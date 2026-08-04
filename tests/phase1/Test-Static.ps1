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
    'src\ManualBuilder.Html.psm1',
    'src\ManualBuilder.Ocr.psm1',
    'src\ManualBuilder.Copilot.psm1',
    'src\ManualBuilder.CopilotJob.psm1',
    'src\ManualBuilder.CopilotServer.psm1',
    'src\Invoke-ManualBuilderCopilotJob.ps1',
    'src\ManualBuilder.Recorder.psm1',
    'src\ManualBuilder.RecorderServer.psm1',
    'src\Invoke-ManualBuilderRecorder.ps1',
    'src\Invoke-ManualBuilderUiaRecorder.ps1',
    'src\ManualBuilder.EdgeRecorder.psm1',
    'src\Invoke-ManualBuilderEdgeRecorder.ps1',
    'src\ManualBuilder.Dictation.psm1',
    'src\Invoke-ManualBuilderDictation.ps1',
    'web\index.html',
    'web\assets\css\app.css',
    'web\assets\js\app.js',
    'web\assets\js\video-scenes.js',
    'web\assets\js\heartbeat-worker.js',
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
$htmlModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Html.psm1'), [Text.Encoding]::UTF8)
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
Add-Result ([string]$appVersionManifest.appVersion -eq '0.35.0') '配布用アプリバージョンを0.35.0へ更新する'
Add-Result ($workspaceModuleText -notmatch "ManualBuilder\.Project\.psm1'\) -Force") 'WorkspaceがProjectコマンドを強制再読込しない'
Add-Result ($launcherModuleText -match "'ManualBuilder\\app'") 'アプリ実行コードをLocalApplicationDataへキャッシュする'
Add-Result ($launcherModuleText -match "@\('src', 'web', 'run\.cmd', 'app-version\.json'\)") 'キャッシュ対象からプロジェクトデータを除外する'
Add-Result (($launcherModuleText -match '\.app-update-') -and ($launcherModuleText -match 'Get-FileHash') -and ($launcherModuleText -match '\[IO\.Directory\]::Move')) '検証後にローカル実行版を切り替える'
Add-Result ($launcherModuleText -match 'コピー中に配布元のバージョンが変更されたため') '配布元の更新途中を完成版として採用しない'
Add-Result ($launcherModuleText -match "\.previous'") '更新前のローカル実行版を1世代残す'
Add-Result ($launcherModuleText -match "'ManualBuilder\.cmd'") '共有フォルダーなしで使えるローカル起動ファイルを作る'
Add-Result (($launcherText -match 'Test-MbApplicationIsRunning') -and ($launcherText -match 'Open-MbRunningApplication')) '起動中はローカル実行版を差し替えず既存画面を開く'
Add-Result (($launcherText -match 'Test-MbCachedApplication -CacheRoot \$cacheRoot') -and ($launcherText -match "検証済みの既存ローカル版で起動します")) '共有更新失敗時は検証済みローカル版へフォールバックする'
Add-Result (($launcherText -match '-LegacyAppRoot \$sourceRoot') -and ($serverText -match '-LegacyAppRoot \$LegacyAppRoot')) 'ローカル起動後も共有元の旧データを移行できる'
Add-Result ($serverText -match 'FileSystemWatcher') 'スクリーンショット保存先の監視を実装する'
Add-Result ($serverText -match '/api/images/import') '生バイト画像取込みAPIを実装する'
Add-Result ($serverText -match '/api/images/replace') '画像差し替えAPIを実装する'
Add-Result ($serverText -match '/api/images/replace/undo') '画像差し替えの復元APIを実装する'
Add-Result ($serverText -match '/api/capture/heartbeat') '撮影対象タブのハートビートを実装する'
Add-Result ($serverText -match '\$HeartbeatTimeoutSec = 90') '裏タブのタイマー間引きを見込んだハートビート猶予にする'
Add-Result (($serverText -match '\$CaptureStandbySec') -and ($serverText -match "'standby'")) 'ハートビート失効後も新着を保留する状態を持つ'
Add-Result ($serverText -match "-notin @\('active', 'standby'\)") '保留中も保存先の新着を取りこぼさない'
Add-Result ($serverText -match '/assets/js/heartbeat-worker\.js') 'ハートビート用Workerを配信する'
Add-Result ($jsText -match 'heartbeat-worker\.js') '画面がハートビートをWorkerタイマーで送る'
Add-Result ($jsText -match "addEventListener\('focus', wakeHeartbeat\)") '復帰時にハートビートを送り直す'
Add-Result ($webModuleText -match 'data-open-video-picker') '動画から手順を作る入口を画面へ置く'
Add-Result ($webModuleText -match 'accept="video/mp4,video/webm"') '取り込める動画をmp4とwebmに限る'
Add-Result ($jsText -match 'isSupportedVideo') '動画の判定を画面側で行う'
Add-Result ($jsText -match 'VIDEO_FRAME_MAX_EDGE = 1280') '動画のコマを既定で長辺1280pxへ縮小する'
Add-Result ($jsText -match "toBlob\(resolve, 'image/jpeg'") '動画のコマをJPEGで取り込む'
Add-Result (($serverText -match "'paste', 'drop', 'file', 'video'") -and ($captureModuleText -match "'file', 'video'")) '取り込み元としてvideoを受け付ける'
Add-Result ($serverText -notmatch 'video/mp4') '動画ファイルをブラウザーへ配信しない'
Add-Result ($serverText -match '/api/videos/attach') '手順へ動画を添付するAPIを実装する'
Add-Result ($serverText -match '/api/videos/detach') '手順から動画を外すAPIを実装する'
Add-Result (($captureModuleText -match "'videos'") -and ($projectModuleText -match "'videos'")) '動画を画像とは別に保存する'
Add-Result ($projectModuleText -match 'videoId') '手順に動画の紐づけを持つ'
Add-Result ($serverText -match "media-src 'self' blob:") '動画ダイアログのblob:再生をCSPで止めない'
Add-Result ($serverText -match "/api/export/html'") 'HTML出力APIを実装する'
Add-Result ($webModuleText -match 'data-export-html') 'HTML出力の入口を画面へ置く'
Add-Result ($htmlModuleText -notmatch '(?i)ComObject') 'HTML出力はCOMを使わない'
Add-Result ($htmlModuleText -match 'New-MbAnnotatedImage') 'HTMLも注釈を画像へ焼き込む'
Add-Result ($htmlModuleText -match '@media print') 'HTMLに印刷用の指定を入れる'
Add-Result ($htmlModuleText -match 'HtmlEncode') 'HTMLへ出す文字列をエスケープする'
Add-Result ($htmlModuleText -match "'_source'") 'HTML出力に元データを同梱する'
Add-Result ($htmlModuleText -match "IO\.FileAttributes\]::Hidden") '同梱する元データを隠しフォルダーにする'
Add-Result ($htmlModuleText -match "'_source/videos/'") '動画は元データの1本だけを参照する'
Add-Result (($htmlModuleText -match 'マニュアルを開く\.cmd') -and ($htmlModuleText -match '編集する\.cmd')) '配布フォルダーへ操作用の.cmdを入れる'
Add-Result ($htmlModuleText -match 'msedge\.exe') '共有フォルダーのHTMLをEdgeで開く（IEモードを避ける）'
Add-Result ($htmlModuleText -match '-ImportFrom "%~dp0_source"') '「編集する.cmd」から元データを取り込む'
Add-Result ($htmlModuleText -match 'GetEncoding\(932\)') '.cmdはcmdが読める文字コードで書く'
Add-Result ($htmlModuleText -notmatch "yyyyMMdd_HHmmss") '出力フォルダー名に日付を付けない'
Add-Result ($htmlModuleText -match 'Copy-MbHtmlManualFolder') '共有フォルダーへ反映する処理を持つ'
Add-Result ($htmlModuleText -match '\.mb-publish-') '反映は別名でコピーしてから差し替える'
Add-Result ($serverText -match "/api/export/html/publish'") '共有フォルダーへの反映APIを実装する'
Add-Result ($serverText -match "FOLDER_EXISTS") '同じ名前のフォルダーは確認してから作り直す'
Add-Result (($serverText -match '\[string\]\$ImportFrom') -and ($serverText -match '\[string\]\$PublishTo')) '「編集する.cmd」からの起動を受け取る'
Add-Result ($serverText -match 'Import-MbCatalogProjectFolder') '配布フォルダーの元データを起動時に取り込む'
Add-Result (($launcherText -match '\[string\]\$ImportFrom') -and ($launcherText -match '@startArguments')) 'ランチャーが取り込み指定を受け渡す'
Add-Result ($workspaceModuleText -match 'Find-MbCatalogProjectById') '同じマニュアルを二重に取り込まない'
Add-Result ($workspaceModuleText -match 'publishTargets') 'マニュアルごとの反映先を覚える'
Add-Result ($jsText -match 'data-html-publish') '完了画面から共有フォルダーへ反映できる'
Add-Result ($jsText -match 'data-export-video-note') '動画つきならExcelの完了画面で知らせる'
Add-Result ($excelModuleText -match 'Get-MbExcelVideoPlan') '動画つきの手順をExcel出力でも扱う'
Add-Result ($excelModuleText -match 'MbExcelVideoFolderName') 'Excelの動画を決まったフォルダーへまとめる'
# Hyperlinks.Addは保存時に絶対パスへ変換されるため、相対パスが保たれる=HYPERLINK()数式を使う。
Add-Result ($excelModuleText -match '\$videoCell\.Formula = ''=HYPERLINK\(') 'Excelの動画リンクは相対パスが保たれる数式で入れる'
# COMのRangeやShapesは列挙できるため、if式の値として受け取るとパイプラインで展開され、
# オブジェクト1個ではなく配列になる。配列にはプロパティを設定できず、出力全体が失敗する。
$comObjectIfAssignment = '\$\w+\s*=\s*if\s*\([^\r\n]*\)\s*\{[^\r\n]*\.(Range|Cells|Shapes|Slides|Paragraphs|Tables|Worksheets|Hyperlinks|Presentations|Documents)\('
Add-Result (($excelModuleText -notmatch $comObjectIfAssignment) -and
    ($wordModuleText -notmatch $comObjectIfAssignment)) 'COMオブジェクトをif式の値として受け取らない（配列へ展開されるため）'
Add-Result ($excelModuleText -match '\$usesFolderOutput = \[int\]\$videoPlan\.Count -gt 0') '動画つきのときだけExcelをフォルダー出力にする'
Add-Result ($excelModuleText -match '\.mb-excel-') 'Excelのフォルダー出力も組み立ててから差し替える'
Add-Result ($excelModuleText -match '出力した動画数の自己検査に失敗しました') '出力した動画数を自己検査する'
Add-Result ($serverText -match '\$snapshotVideoDirectory') 'Excel出力用に動画もスナップショットへ複製する'
Add-Result ($jsText -match 'outputFolderName') 'フォルダー出力になったことを完了画面へ出す'
# 文字リンクのままだと見出しの中で埋もれ、同じ行の「目次へ戻る」とも見分けが付かない。
Add-Result (($excelModuleText -match '\$videoCell\.Interior\.Color = \$colorAccent') -and
    ($excelModuleText -match '\$videoCell\.Font\.Color = \$colorWhite')) 'Excelの動画リンクは押せると分かる見た目にする'
# 「フォルダーごとコピー」とだけ書いても次の操作へつながらないため、案内文はボタン名で指す。
Add-Result (($jsText -match 'フォルダーを開く』?」から') -or ($jsText -match '「フォルダーを開く」から')) '配布の案内から次に押すボタンへつなぐ'
Add-Result ($jsText -match "folderButton\.textContent = outputFolderName \? 'フォルダーを開く'") 'フォルダー出力のときはボタン名も「フォルダーを開く」にする'
# New-MbAnnotatedImage は焼き込みが不要だと元画像のパスを返し、出力先へは書かない。
# 戻り値を捨てると、注釈を付けていない手順の画像がすべてリンク切れになる。
Add-Result ($htmlModuleText -notmatch '\[void\]\(New-MbAnnotatedImage') 'HTML出力で焼き込み結果の戻り値を捨てない'
Add-Result ($htmlModuleText -match '\$renderedPath = New-MbAnnotatedImage') 'HTML出力は焼き込み結果のパスを見て画像を置く'
Add-Result ($htmlModuleText -match '画像を出力できませんでした') '画像が出力できていなければ気付けるようにする'
Add-Result (($htmlModuleText -match "Import-Module \(Join-Path \`$PSScriptRoot 'ManualBuilder\.Excel\.psm1'\)") -and
    ($htmlModuleText -match "Import-Module \(Join-Path \`$PSScriptRoot 'ManualBuilder\.Capture\.psm1'\)")) 'HTMLモジュールが借りているコマンドを明示して読み込む'
# OneDriveやウイルス対策が書いたばかりのファイルを掴んでいると、移動がアクセス拒否で失敗する。
Add-Result ($excelModuleText -match 'function Move-MbDirectorySafely') 'フォルダーの移動を待って試し直せるようにする'
Add-Result ($htmlModuleText -notmatch '\[IO\.Directory\]::Move') 'HTML出力の差し替えは再試行つきの移動を使う'
Add-Result ($excelModuleText -notmatch '\[IO\.Directory\]::Move\(\$stagingDirectory') 'Excelのフォルダー出力も再試行つきの移動を使う'
# 差し替えに失敗して元へ戻せなかった場合、退避先が前の内容の唯一の実体になる。
Add-Result ($htmlModuleText -match '\$moveCompleted -and \$replacedMoved') '差し替えを終えたときだけ前のフォルダーを消す'
Add-Result ($excelModuleText -match 'Set-MbExcelEdgeBorder -Range \$videoCell') '動画ボタンを白い余白で囲んで帯に見せない'
Add-Result (($webModuleText -match 'button--primary" data-export-excel') -and ($webModuleText -match 'button--secondary" data-export-html')) 'ExcelとHTMLのボタンを並べる'
Add-Result ($serverText -match '/api/steps/reorder') '手順並べ替えAPIを実装する'
Add-Result ($serverText -match '/api/sheets/reorder') 'シート並べ替えAPIを実装する'
Add-Result ($serverText -match '/api/steps/move') '手順のシート移動APIを実装する'
Add-Result (($serverText -match '/api/steps/move-many') -and ($serverText -match '/api/steps/delete-many')) '複数手順の移動・削除APIを実装する'
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
Add-Result (($launcherText -match 'Start-Process \$url') -and
    ($serverText -match 'Start-Process \$existingUrl')) '二重起動時に既存のManualBuilderをブラウザーで開き直す'
Add-Result (($launcherText -match 'ManualBuilderを終了してから、もう一度「編集する」を実行してください') -and
    ($serverText -match 'ManualBuilderを終了してから、もう一度「編集する」を実行してください')) '共有HTMLの編集時は既存画面で終了してから再実行するよう案内する'
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
# 入力済みの手順は目印そのものを出さないため、CSSで隠す指定は持たない。
# 未入力は「説明未入力」と「画像なし」で直し方が違うので、塗りと輪郭で形でも分ける。
Add-Result (($webModuleText -match 'step-nav__status" role="img"') -and ($cssText -match '\.step-nav__item--empty \.step-nav__status') -and ($cssText -match 'box-shadow:\s*inset 2px 0 0 var\(--accent\)')) '未完了状態と選択中の手順を優先表示する'
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

$ocrModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Ocr.psm1'), [Text.Encoding]::UTF8)
$copilotModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Copilot.psm1'), [Text.Encoding]::UTF8)
$copilotServerText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.CopilotServer.psm1'), [Text.Encoding]::UTF8)
$copilotJobText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.CopilotJob.psm1'), [Text.Encoding]::UTF8)
$sceneText = [IO.File]::ReadAllText((Join-Path $repoRoot 'web\assets\js\video-scenes.js'), [Text.Encoding]::UTF8)
Add-Result ($indexText -match 'video-scenes\.js') '場面分割のスクリプトを読み込む'
Add-Result ($jsText -match 'data-video-auto') '録画を自動で手順へ分けるボタンがある'
Add-Result ($sceneText -match 'locateChangeRect') '遷移の入口から操作位置を求める'
Add-Result ($serverText -match '/api/videos/scenes/import') '場面の取り込み口がある'
Add-Result ($serverText -match '/api/copilot/draft/start') 'Copilot下書きの開始口がある'
Add-Result ($serverText -match '/api/copilot/draft/apply') '採用した下書きの反映口がある'
Add-Result ($webModuleText -match 'data-copilot-draft') 'Copilotでの下書きをメニューから選べる'
Add-Result ($jsText -match 'copilot-draft-dialog') 'Copilot下書きの確認画面を実装する'
Add-Result ($copilotModuleText -match 'm365\.cloud\.microsoft') '普段使うM365 Copilotの画面を操作する'
Add-Result ($copilotModuleText -notmatch '(?i)api[_-]?key') 'APIキーを持たない'
Add-Result ($copilotJobText -match '\$rendered = New-MbAnnotatedImage') '焼き込み結果の戻り値を捨てない'
Add-Result ($copilotServerText -match 'Resolve-MbOperationRect') '赤枠を読み取った文字へ寄せる'
Add-Result ($ocrModuleText -match 'return \$false') '文字認識が使えない環境では機能だけを止める'
Add-Result ($projectModuleText -match 'Add-MbPropertyIfMissing \$step ''capture''') '古い手順にも録画情報の入れ物を補う'

$recorderModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Recorder.psm1'), [Text.Encoding]::UTF8)
$recorderServerText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.RecorderServer.psm1'), [Text.Encoding]::UTF8)
$edgeRecorderModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.EdgeRecorder.psm1'), [Text.Encoding]::UTF8)
Add-Result (($copilotJobText -notmatch '(?m)^Import-Module .+ -Force$') -and
    ($copilotServerText -notmatch '(?m)^Import-Module .+ -Force$') -and
    ($recorderServerText -notmatch '(?m)^Import-Module .+ -Force$')) '入れ子のモジュールが共有コマンドを強制再読込しない'
Add-Result ($recorderModuleText -match 'AutomationElement\]::FromPoint') '押した位置のコントロールをUI Automationから取る'
Add-Result ($recorderModuleText -match 'SetProcessDpiAwarenessContext') '高DPIで座標がずれないようDPI認識にする'
Add-Result ($recorderModuleText -notmatch 'SetWindowsHookEx') '低レベルフックを使わない'
Add-Result ($recorderModuleText -match 'rightClicked') '右クリックも操作として記録する'
Add-Result ($recorderModuleText -match '\$capture = Copy-MbScreenBitmap') '押す直前の画面を先に確保する'
Add-Result ($recorderModuleText -match 'DWMWA_EXTENDED_FRAME_BOUNDS') '見た目どおりのウィンドウ範囲を使う'
Add-Result ($recorderServerText -match "Source 'recorder'") '記録した画面を専用の出所として取り込む'
Add-Result ($recorderServerText -match 'AllowDuplicateStep') '同じ画面でも別の操作はそれぞれ手順にする'
Add-Result (($recorderModuleText -match 'ConvertTo-MbRecorderTargetName') -and ($recorderServerText -match 'ConvertTo-MbRecorderTargetName.+-Suffix \$suffix')) '操作対象を補足込みで200文字へ収める'
Add-Result (($recorderServerText -notmatch 'ManualBuilder\.EdgeRecorder') -and
    ($recorderServerText -notmatch 'Invoke-ManualBuilderEdgeRecorder')) '記録用Edgeを必須経路から廃止する'
Add-Result (($jsText -notmatch 'data-recorder-mode') -and
    ($jsText -notmatch '記録用Edgeを使う')) '記録用Edgeの選択UIを廃止する'
Add-Result (($recorderModuleText -match 'event-\{0:d3\}-after\.jpg') -and
    ($recorderModuleText -match 'preClickCaptureMaxAgeMs') -and
    ($recorderModuleText -notmatch '-not \[string\]::IsNullOrWhiteSpace\(\$DomTargetPath\) -and\s*\(\(\[int\]\$watch\.ElapsedMilliseconds - \$preClickCaptureAttemptAtMs')) 'すべてのアプリでクリック前後の画像を保持する'
Add-Result (($projectModuleText -match 'afterImageId') -and
    ($projectModuleText -match 'targetCandidates') -and
    ($projectModuleText -match 'analysisState')) '再解析できる操作証跡をプロジェクトへ保存する'
Add-Result (($copilotJobText -match 'New-MbCopilotOperationPrompt') -and
    ($copilotJobText -match '1件の操作') -and
    ($copilotJobText -match '候補枠.+正解として扱わない')) 'Copilotが操作対象・意味・手順文を1操作ずつ解析する'
Add-Result (($copilotJobText -match 'New-MbCopilotOperationEvidence') -and
    ($copilotJobText -match 'operation-before\.jpg') -and
    ($copilotJobText -match 'operation-after\.jpg')) 'Copilotへ操作前・周辺・操作後の証跡を渡す'
Add-Result ($copilotJobText -match 'Test-MbCopilotOperationPreflight') '実画像の前に合成画像で証跡生成を自己診断する'
Add-Result (($recorderServerText -match 'Invoke-ManualBuilderUiaRecorder\.ps1') -and
    ($recorderServerText -match 'UiaTargetPath') -and
    ($recorderModuleText -match 'Get-MbUiaTargetFromCache')) 'Windows操作対象もクリック前に別プロセスで保持する'
Add-Result (($recorderServerText -match "AnalysisState 'pending'") -and
    ($recorderServerText -notmatch 'Set-MbStepImageEdits.+\$annotation')) '取り込み時は候補枠を確定赤枠にしない'
Add-Result (($copilotServerText -match "AnalysisState 'confirmed'") -and
    ($copilotServerText -match 'Where-Object \{ \[string\]\$_\.type -eq ''blackout'' \}')) '確認後に赤枠を確定し手動黒塗りを保持する'
Add-Result (($webModuleText -match 'data-copilot-operation') -and
    ($jsText -match "openCopilotDialog\('operation'\)")) '記録操作のCopilot解析を再開できる'
Add-Result ($serverText -match '/api/recorder/start') '操作記録の開始口がある'
Add-Result ($serverText -match '/api/recorder/import') '記録した操作の取り込み口がある'
Add-Result ($serverText -match '\^/images/recording/') '記録した画面をクエリのトークンで表示できる'
Add-Result ($webModuleText -match 'data-record-operations') '操作の記録をメニューから選べる'
Add-Result ($jsText -match 'recorder-dialog') '記録の確認画面を実装する'
Add-Result ($jsText -match 'const shouldDiscard = recorder\.active') '記録中に確認画面を閉じても記録プロセスを残さない'
Add-Result (($recorderModuleText -notmatch 'RedactTarget') -and
    ($jsText -match '画像を編集.*黒塗り')) '入力欄を自動マスクせず、必要な箇所だけ手動で黒塗りする'
Add-Result ($recorderModuleText -match 'Test-MbAsyncKeyStatePressed') '短いクリックやキー入力も押下履歴から検出する'
Add-Result (($recorderModuleText -match 'FindAll') -and
    ($recorderModuleText -match 'IsOffscreenProperty')) '長いページでも表示中の操作コントロールを条件検索する'
Add-Result (($recorderModuleText -match 'LegacyIAccessiblePattern') -and
    ($recorderModuleText -match 'ControlType\]::Group') -and
    ($recorderModuleText -match 'Select-MbUiaNamedTargetInfo')) 'Web独自要素と名前付き要素まで操作対象候補を広げる'
Add-Result (($recorderModuleText -match 'AutomationElement\]::FromHandle') -and
    ($recorderModuleText -match 'Get-MbUiaFocusedElement') -and
    ($recorderModuleText -match 'AccessibleObjectFromPoint') -and
    ($recorderModuleText -match "provider = 'MSAA'")) 'フォーカス・前面ウィンドウ全枝・MSAAで操作対象を再検索する'
Add-Result (($recorderModuleText -match '\[int\]\$MaxEdge = 2560') -and
    ($recorderModuleText -match '\[long\]\$Quality = 94') -and
    ($recorderModuleText -match 'HighQualityBicubic')) '対象周辺表示に十分な解像度と高品質縮小で記録する'
Add-Result (($recorderModuleText -match 'Select-MbUiaTargetInfo') -and
    ($recorderModuleText -match 'isActionable')) '操作可能な最小要素だけを赤枠候補にする'
Add-Result (($recorderModuleText -match 'RawViewWalker') -and
    ($recorderModuleText -match 'New-MbClickPointTargetInfo')) '近傍点とクリック位置フォールバックでUIA非対応画面も示す'
Add-Result (($projectModuleText -match 'Move-MbStepsToSheet') -and ($projectModuleText -match 'Remove-MbSteps') -and
    ($webModuleText -match 'data-step-select') -and ($jsText -match 'runBulkStepAction')) '手順を複数選択してまとめて移動・削除する'
Add-Result (($webModuleText -match 'data-step-select-all') -and
    ($jsText -match "event\.target\.matches\('\[data-step-select-all\]'\)") -and
    ($jsText -match 'selectAll\.indeterminate')) 'このシートの手順を一度に全選択・全解除する'
Add-Result (($projectModuleText -match '\$recoveryPrefix') -and
    ($projectModuleText -match 'foreach \(\$delay in @\(0, 25, 50, 100, 200, 400\)\)') -and
    ($projectModuleText -match '\$successfulRecoveryPath')) 'プロジェクト保存は固有バックアップと再試行で一時ロックを避ける'
Add-Result ($recorderModuleText -match '\$rectWidth -gt 0\.82') 'ページ全体に近い矩形を最終段でも除外する'
Add-Result (($recorderModuleText -match '\[IO\.File\]::Replace\(\$temporary, \$StatusPath') -and
    ($recorderModuleText -match '\$delaysMs')) '操作記録の進捗JSONを完成後に差し替え、短いロックは再試行する'
Add-Result ($recorderServerText -match '\[IO\.FileShare\]::ReadWrite -bor \[IO\.FileShare\]::Delete') '進捗を読む側はワーカーの原子的な差し替えを妨げない'
Add-Result ($cssText -match '\.copilot-dialog\[open\]') '閉じた記録・Copilotダイアログを画面に残さない'
$dictationModuleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.Dictation.psm1'), [Text.Encoding]::UTF8)
Add-Result ($dictationModuleText -match 'SpeechRecognitionScenario\]::Dictation') 'Win+Hと同じ口述筆記の仕組みを使う'
Add-Result ($dictationModuleText -match 'PhraseStartTime') '受信時刻ではなく発話の開始時刻で突き合わせる'
Add-Result ($recorderServerText -match 'Merge-MbNarrationIntoEvents') '話した内容を操作へ振り分ける'
Add-Result ($jsText -match 'data-recorder-narration') '音声を記録するかを選べる'
Add-Result ($jsText -match 'Microsoftのオンライン音声認識へ送られます') '音声が端末の外へ出ることを画面に明記する'
Add-Result ($jsText -match "narrationToggle.checked = false") '音声の記録は既定で行わない'
$copilotServerText2 = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.CopilotServer.psm1'), [Text.Encoding]::UTF8)
$sceneText2 = [IO.File]::ReadAllText((Join-Path $repoRoot 'web\assets\js\video-scenes.js'), [Text.Encoding]::UTF8)
Add-Result ($copilotServerText2 -notmatch 'System\.Speech') '精度の低いローカル音声認識を持たない'
Add-Result ($serverText -notmatch '/api/narration/transcribe') '録画からの文字起こしの口を持たない'
Add-Result ($sceneText2 -notmatch 'extractNarration') '録画から音声を取り出さない'
$copilotJobText2 = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.CopilotJob.psm1'), [Text.Encoding]::UTF8)
Add-Result ($copilotJobText2 -match 'New-MbCopilotReviewPrompt') '文章を整える依頼文を持つ'
Add-Result ($copilotJobText2 -match '表記ゆれ') '表記ゆれを見るよう依頼する'
Add-Result ($webModuleText -match 'data-copilot-review') '文章を整えるをメニューから選べる'
Add-Result ($jsText -match "copilotDraft.mode === 'review'") '下書きと校正で画面の出し分けをする'
Add-Result ($serverText -notmatch 'PowerPoint') 'PowerPoint出力を持たない'
Add-Result ($jsText -notmatch '(?i)powerpoint') '画面にPowerPoint出力が残っていない'

# 入口スクリプトが呼ぶ関数が、その場で解決できることを確かめる。
#
# PowerShell では、モジュールの中で Import-Module したものは、そのモジュールの中でしか
# 見えない。入口スクリプトから呼びたい関数は、入口スクリプト自身が読み込んだモジュールが
# 公開していなければならない。この取り違えは実行するまで気付けず、実際に2度作り込んだ。
$entryDefinitions = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($match in [regex]::Matches($serverText, '(?m)^\s*function\s+([A-Za-z]+-Mb[A-Za-z0-9]*)')) {
    [void]$entryDefinitions.Add($match.Groups[1].Value)
}
foreach ($match in [regex]::Matches($serverText, "Import-Module \(Join-Path \`$PSScriptRoot '([A-Za-z.]+\.psm1)'\)")) {
    $modulePath = Join-Path $repoRoot ('src\' + $match.Groups[1].Value)
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) { continue }
    $moduleSource = [IO.File]::ReadAllText($modulePath, [Text.Encoding]::UTF8)
    $exportBlock = [regex]::Match($moduleSource, 'Export-ModuleMember\s+-Function\s+@\(([\s\S]*?)\)')
    if (-not $exportBlock.Success) { continue }
    foreach ($exported in [regex]::Matches($exportBlock.Groups[1].Value, "'([A-Za-z]+-Mb[A-Za-z0-9]*)'")) {
        [void]$entryDefinitions.Add($exported.Groups[1].Value)
    }
}
# 文字列リテラルの中の名前は呼び出しではない。ヘッダー名などを拾わないよう外す。
$callSites = [regex]::Replace($serverText, "'[^'\r\n]*'", "''")
$callSites = [regex]::Replace($callSites, '"[^"\r\n]*"', '""')
$unresolved = New-Object 'System.Collections.Generic.List[string]'
foreach ($match in [regex]::Matches($callSites, '\b([A-Za-z]+-Mb[A-Za-z0-9]*)\b')) {
    $name = $match.Groups[1].Value
    if (-not $entryDefinitions.Contains($name) -and -not $unresolved.Contains($name)) {
        [void]$unresolved.Add($name)
    }
}
$unresolvedMessage = '本体が呼ぶ関数がすべて解決する'
if ($unresolved.Count -gt 0) { $unresolvedMessage += '（未解決: ' + ($unresolved -join ', ') + '）' }
Add-Result ($unresolved.Count -eq 0) $unresolvedMessage

# 同じ名前の関数を複数のモジュールが公開していると、読み込み順で後勝ちになる。
$exportOwners = @{}
foreach ($moduleFile in (Get-ChildItem -LiteralPath (Join-Path $repoRoot 'src') -Filter '*.psm1' -File)) {
    $moduleSource = [IO.File]::ReadAllText($moduleFile.FullName, [Text.Encoding]::UTF8)
    $exportBlock = [regex]::Match($moduleSource, 'Export-ModuleMember\s+-Function\s+@\(([\s\S]*?)\)')
    if (-not $exportBlock.Success) { continue }
    foreach ($exported in [regex]::Matches($exportBlock.Groups[1].Value, "'([A-Za-z]+-Mb[A-Za-z0-9]*)'")) {
        $name = $exported.Groups[1].Value
        if (-not $exportOwners.ContainsKey($name)) { $exportOwners[$name] = New-Object 'System.Collections.Generic.List[string]' }
        [void]$exportOwners[$name].Add($moduleFile.Name)
    }
}
$collisions = @($exportOwners.Keys | Where-Object { $exportOwners[$_].Count -gt 1 } | Sort-Object)
$collisionMessage = '同じ関数名を複数のモジュールが公開していない'
if ($collisions.Count -gt 0) { $collisionMessage += '（重複: ' + ($collisions -join ', ') + '）' }
Add-Result ($collisions.Count -eq 0) $collisionMessage

if ($errors.Count -gt 0) {
    Write-Host ''
    Write-Host ("Static tests failed: " + $errors.Count) -ForegroundColor Red
    exit 1
}
Write-Host ''
Write-Host 'Static tests passed.' -ForegroundColor Cyan
