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

Add-Result (Test-Path -LiteralPath $sourcePath -PathType Leaf) 'C#ソースが存在する' $sourcePath
foreach ($name in @('Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.Wpf.dll', 'WebView2Loader.dll')) {
    Add-Result (Test-Path -LiteralPath (Join-Path $vendorPath $name) -PathType Leaf) "同梱DLLが存在する: $name"
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
