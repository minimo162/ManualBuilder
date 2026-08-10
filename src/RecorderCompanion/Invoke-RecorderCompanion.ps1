[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][string]$EventsDirectory,
    [Parameter(Mandatory = $true)][string]$PausePath,
    [Parameter(Mandatory = $true)][string]$UndoPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [Parameter(Mandatory = $true)][string]$StopPath,
    [Parameter(Mandatory = $true)][string]$JobId,
    [Parameter(Mandatory = $true)][string]$WebRoot,
    [long]$ReturnWindowHandle = 0,
    [switch]$TestMode
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $root 'ManualBuilder.RecorderCompanion.cs'
$vendorPath = Join-Path $root 'vendor\WebView2'
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

foreach ($requiredPath in @($sourcePath, $WebRoot, (Join-Path $vendorPath 'WebView2Loader.dll')) + $references[-2..-1]) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "記録モニターの実行ファイルが不足しています: $requiredPath"
    }
}

# Smart App Controlに拒否される未署名EXEを生成せず、Microsoft署名済みの
# PowerShellプロセス内へWPFホストをコンパイルして実行する。
$source = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::UTF8)
[void][Reflection.Assembly]::LoadFrom((Join-Path $vendorPath 'Microsoft.Web.WebView2.Core.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $vendorPath 'Microsoft.Web.WebView2.Wpf.dll'))
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies $references -ErrorAction Stop

$arguments = @(
    '--status', [IO.Path]::GetFullPath($StatusPath),
    '--events', [IO.Path]::GetFullPath($EventsDirectory),
    '--pause', [IO.Path]::GetFullPath($PausePath),
    '--undo', [IO.Path]::GetFullPath($UndoPath),
    '--result', [IO.Path]::GetFullPath($ResultPath),
    '--stop', [IO.Path]::GetFullPath($StopPath),
    '--job', $JobId,
    '--web-root', [IO.Path]::GetFullPath($WebRoot),
    '--return-window', ([string]$ReturnWindowHandle)
)
if ($TestMode) { $arguments += @('--test-mode', 'true') }

$exitCode = [ManualBuilder.RecorderCompanion.Program]::Run([string[]]$arguments)
exit $exitCode
