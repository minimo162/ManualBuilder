# 記録モニター(RecorderCompanion)のC#が実行時にコンパイルできることだけを確かめる。
# 本体は Add-Type でランタイムコンパイルされるため csproj もビルド段階も無く、
# 構文エラー・型エラーは「利用者が記録を開始した瞬間」まで表面化しない。
# WebView2の実表示は Test-RecorderController.ps1 が担当する。ここは出荷前の門番。
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$failures = 0
function Add-Result {
    param([bool]$Ok, [string]$Name, [string]$Detail = '')
    if ($Ok) {
        Write-Host "[OK] $Name"
    } else {
        $script:failures++
        Write-Host "[NG] $Name $Detail" -ForegroundColor Red
    }
}

$companionRoot = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'src\RecorderCompanion'
$sourcePath = Join-Path $companionRoot 'ManualBuilder.RecorderCompanion.cs'
$vendorPath = Join-Path $companionRoot 'vendor\WebView2'
$webRoot = Join-Path $companionRoot 'web'

Add-Result (Test-Path -LiteralPath $sourcePath -PathType Leaf) 'C#ソースが存在する' $sourcePath
foreach ($name in @('Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.Wpf.dll', 'WebView2Loader.dll')) {
    Add-Result (Test-Path -LiteralPath (Join-Path $vendorPath $name) -PathType Leaf) "同梱DLLが存在する: $name"
}
foreach ($name in @('index.html', 'app.js', 'styles.css')) {
    Add-Result (Test-Path -LiteralPath (Join-Path $webRoot $name) -PathType Leaf) "Web資産が存在する: $name"
}

if ($failures -gt 0) {
    Write-Host 'RecorderCompanion compile checks failed.' -ForegroundColor Red
    exit 1
}

# 参照一覧は Invoke-RecorderCompanion.ps1 と同じ構成にする。
# ここがずれると「テストは通るのに実行時に落ちる」状態になる。
$frameworkPath = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
$wpfPath = Join-Path $frameworkPath 'WPF'
$references = @(
    (Join-Path $frameworkPath 'System.dll'),
    (Join-Path $frameworkPath 'System.Core.dll'),
    (Join-Path $frameworkPath 'System.Web.Extensions.dll'),
    (Join-Path $frameworkPath 'System.Xaml.dll'),
    (Join-Path $wpfPath 'WindowsBase.dll'),
    (Join-Path $wpfPath 'PresentationCore.dll'),
    (Join-Path $wpfPath 'PresentationFramework.dll'),
    (Join-Path $vendorPath 'Microsoft.Web.WebView2.Core.dll'),
    (Join-Path $vendorPath 'Microsoft.Web.WebView2.Wpf.dll')
)

$compiledTypes = @()
try {
    $source = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::UTF8)
    $indexSource = [IO.File]::ReadAllText((Join-Path $webRoot 'index.html'), [Text.Encoding]::UTF8)
    $appSource = [IO.File]::ReadAllText((Join-Path $webRoot 'app.js'), [Text.Encoding]::UTF8)
    $styleSource = [IO.File]::ReadAllText((Join-Path $webRoot 'styles.css'), [Text.Encoding]::UTF8)
    Add-Result (($source -match 'RegisterHotKey') -and ($source -match 'UnregisterHotKey') -and
        ($source -match 'ModNoRepeat') -and ($source -match 'WmHotKey')) `
        '重複入力を抑えたグローバルホットキーを登録・解除する'
    Add-Result (($source -match 'state == "ready" \? "記録を開始"') -and
        ($source -match 'state == "paused" \? "記録を再開"') -and ($source -match 'return "開始待ち"')) `
        '待機・一時停止の状態に応じた操作名を提供する'
    Add-Result (($source -match 'hotkeysAvailable') -and
        ($indexSource -match 'Control\+Alt\+Space') -and ($indexSource -match 'Control\+Alt\+Enter') -and
        ($appSource -match 'ショートカットは現在利用できません')) `
        'ショートカットと登録失敗時の代替操作を画面で案内する'
    Add-Result (($source -match 'manualBuilderWindow') -and
        ($source -match 'ReturnWindowHandle') -and
        ($source -match 'SetForegroundWindow\(manualBuilderWindow\)')) `
        '記録完了後は候補を確認するManualBuilderへ戻す'
    Add-Result (($source -match 'pendingPauseState') -and
        ($source -match 'pendingPauseAtUtc') -and
        ($source -match 'AddSeconds\(4\)')) `
        '開始・一時停止・再開は状態反映まで連打を受け付けない'
    Add-Result (($source -match 'if \(pauseHotKeyRegistered\) UnregisterHotKey') -and
        ($source -match 'if \(finishHotKeyRegistered\) UnregisterHotKey')) `
        'ショートカットの一部だけ登録できた場合は両方を無効化する'
    Add-Result (($appSource -match '記録の準備をやめますか') -and
        ($appSource -match '準備をやめる') -and
        ($indexSource -match '直前の操作が表示されます')) `
        '開始待ちの終了確認と直前操作の案内を状態に合わせる'
    Add-Result (($appSource -match "finishButton\.addEventListener\('click', \(\) => send\(\{ type: 'close' \}\)\)") -and
        ($appSource -match "compactFinishButton\.addEventListener\('click', \(\) => send\(\{ type: 'close' \}\)\)")) `
        '大小どちらの終了ボタンも確認要求を送る'
    Add-Result (($appSource -match 'confirmCloseButton\.addEventListener') -and
        ($appSource -match 'closeConfirm\.close\(\)') -and ($appSource -match "command\('finish'\)")) `
        '終了確定後は確認画面を閉じて処理結果を隠さない'
    Add-Result (($styleSource -match '\.compact-actions button \{[^}]*font-size: 14px') -and
        ($styleSource -match '\.shortcut-guide \{[^}]*font-size: 14px') -and
        ($source -match 'Width = 480') -and ($source -match 'MinWidth = 440')) `
        '小型画面でも操作・案内を14px以上にして約480×220に収める'
    [void][Reflection.Assembly]::LoadFrom((Join-Path $vendorPath 'Microsoft.Web.WebView2.Core.dll'))
    [void][Reflection.Assembly]::LoadFrom((Join-Path $vendorPath 'Microsoft.Web.WebView2.Wpf.dll'))
    # -PassThru で型を受け取る。動的アセンブリなので [Type]::GetType() では引けない。
    $compiledTypes = @(Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies $references -PassThru -ErrorAction Stop)
    Add-Result $true 'C#がコンパイルできる'
} catch {
    Add-Result $false 'C#がコンパイルできる' $_.Exception.Message
}

# 起動側が呼ぶ入口が、コンパイル結果に実在することまで確かめる。
$entryPoint = @($compiledTypes | Where-Object { $_.FullName -eq 'ManualBuilder.RecorderCompanion.Program' })[0]
Add-Result ($null -ne $entryPoint) 'Program 型が公開されている'
if ($null -ne $entryPoint) {
    Add-Result ($null -ne $entryPoint.GetMethod('Run')) 'Invoke-RecorderCompanion.ps1 が呼ぶ Run が存在する'
}

if ($failures -gt 0) {
    Write-Host 'RecorderCompanion compile checks failed.' -ForegroundColor Red
    exit 1
}
Write-Host 'RecorderCompanion compile checks passed.'
