# Phase 1 Word naming and status tests (Word COM is not started).

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Word.psm1') -Force
$testRoot = Join-Path $env:TEMP ('ManualBuilder-WordUtilityTest-' + [guid]::NewGuid().ToString('N'))

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)
    Assert-Mb ((Get-MbSafeWordFileName -Name '経費/申請:*?' -Directory $testRoot) -eq '経費_申請___.docx') 'Wordファイル名の禁止文字を置換する'
    Assert-Mb ((Get-MbSafeWordFileName -Name 'CON' -Directory $testRoot) -eq '_CON.docx') 'Windows予約名を安全なWordファイル名へ変換する'
    $wordModule = Get-Module -Name 'ManualBuilder.Word'
    $projectWithEmptySheets = [pscustomobject]@{ sheets = @(
        [pscustomobject]@{ name = '空シート1'; steps = @() },
        [pscustomobject]@{ name = '出力対象'; steps = @([pscustomobject]@{ id = 'step-1' }) },
        [pscustomobject]@{ name = '空シート2'; steps = @() }
    ) }
    $contentSheets = @(& $wordModule { param($Project) Get-MbWordContentSheets -Project $Project } $projectWithEmptySheets)
    Assert-Mb ($contentSheets.Count -eq 1 -and [string]$contentSheets[0].name -eq '出力対象') 'Word出力から空シートを除外する'
    $statusPath = Join-Path $testRoot 'status.json'
    $status = [pscustomobject]@{ state = 'queued'; message = '開始'; updatedAt = '' }
    & $wordModule { param($Path, $Value) Write-MbWordStatus -StatusPath $Path -Status $Value } $statusPath $status
    $status.state = 'running'; $status.message = '更新'
    & $wordModule { param($Path, $Value) Write-MbWordStatus -StatusPath $Path -Status $Value } $statusPath $status
    $saved = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$saved.state -eq 'running') '同じWord進捗JSONをWindows互換方式で連続更新する'
    Assert-Mb (@(Get-ChildItem -LiteralPath $testRoot -Filter '.word-status-*.tmp' -File -ErrorAction SilentlyContinue).Count -eq 0) 'Word進捗更新後に一時・バックアップファイルを残さない'
    Write-Host ''
    Write-Host 'Word utility tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
