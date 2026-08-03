# Phase 1 PowerPoint exporter helper test (COMを使わない部分だけ).
# スライド生成そのものはPowerPointが必要なため、実機での確認手順は
# docs/RETEST-PHASE1-v0.24.0.md にまとめている。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.PowerPoint.psm1') -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ManualBuilder-PptTest-' + [guid]::NewGuid().ToString('N'))

# Word出力と同じく、モジュールが公開するのは出力の入口だけ。
# 内部の補助関数はここで取り出して確かめる。
$moduleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'src\ManualBuilder.PowerPoint.psm1'), [Text.Encoding]::UTF8)
$tokens = $null
$parseErrors = $null
$moduleAst = [System.Management.Automation.Language.Parser]::ParseInput($moduleText, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw "ManualBuilder.PowerPoint.psm1 を解析できません: $($parseErrors[0].Message)" }
foreach ($functionName in @('ConvertTo-MbPowerPointRgb', 'Write-MbPowerPointStatus', 'Set-MbPowerPointStatusProgress', 'Test-MbPowerPointCancellation', 'Get-MbPowerPointContentSheets')) {
    $found = @($moduleAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    }, $true))
    if ($found.Count -ne 1) { throw "モジュールに関数が見つかりません: $functionName" }
    Invoke-Expression $found[0].Extent.Text
}

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

try {
    [void](New-Item -ItemType Directory -Path $testRoot -Force)

    # --- 出力ファイル名 ---
    Assert-Mb ((Get-MbSafePowerPointFileName -Name '営業手順' -Directory $testRoot) -eq '営業手順.pptx') '日本語のファイル名をそのまま使う'
    Assert-Mb ((Get-MbSafePowerPointFileName -Name 'a/b:c*d?' -Directory $testRoot) -eq 'a_b_c_d_.pptx') '使えない文字を置き換える'
    Assert-Mb ((Get-MbSafePowerPointFileName -Name 'CON' -Directory $testRoot) -eq '_CON.pptx') '予約された名前を避ける'
    Assert-Mb ((Get-MbSafePowerPointFileName -Name '   ' -Directory $testRoot) -eq 'manual.pptx') '空の名前でも出力できる'
    Assert-Mb ((Get-MbSafePowerPointFileName -Name 'ろんぐ' -Directory $testRoot -Extension '.pptx').EndsWith('.pptx')) '拡張子はpptxになる'

    $existing = Join-Path $testRoot '重複.pptx'
    [IO.File]::WriteAllText($existing, 'x')
    Assert-Mb ((Get-MbSafePowerPointFileName -Name '重複' -Directory $testRoot) -eq '重複_2.pptx') '同名ファイルがあれば連番を付ける'

    # --- 進捗ファイル ---
    $statusPath = Join-Path $testRoot 'status.json'
    $status = [pscustomobject]@{
        jobId = 'ppt-export-test'; state = 'running'; phase = 'building-steps'; message = '作成中'; percent = 40
        currentStep = 2; totalSteps = 5; slideCount = 0; videoCount = 0; updatedAt = ''
    }
    Write-MbPowerPointStatus -StatusPath $statusPath -Status $status
    $written = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([string]$written.jobId -eq 'ppt-export-test') '進捗ファイルを書き出せる'
    Assert-Mb (-not [string]::IsNullOrWhiteSpace([string]$written.updatedAt)) '進捗ファイルに更新時刻が入る'

    Set-MbPowerPointStatusProgress $status $statusPath 'saving' '保存中' 5 5 94
    $written = [IO.File]::ReadAllText($statusPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Mb ([int]$written.percent -eq 94 -and [string]$written.phase -eq 'saving') '進捗を更新できる'

    # --- 中止の検知 ---
    $cancelPath = Join-Path $testRoot 'cancel.requested'
    Test-MbPowerPointCancellation $cancelPath
    Assert-Mb $true '中止ファイルが無ければ続行する'
    [IO.File]::WriteAllText($cancelPath, 'cancel')
    $cancelled = $false
    try { Test-MbPowerPointCancellation $cancelPath } catch { $cancelled = ([string]$_.Exception.Message -eq 'MB_EXPORT_CANCELLED') }
    Assert-Mb $cancelled '中止ファイルがあれば中止として扱う'

    # --- 出力対象のシート ---
    $project = New-MbProject
    $emptySheet = Add-MbSheet -Project $project
    [void](Add-MbStep -Project $project -SheetId ([string]$project.sheets[0].id))
    $sheets = @(Get-MbPowerPointContentSheets -Project $project)
    Assert-Mb ($sheets.Count -eq 1) '手順の無いシートは出力しない'
    Assert-Mb ([string]$sheets[0].id -eq [string]$project.sheets[0].id) '手順のあるシートだけを出力する'

    # --- 色の変換 ---
    Assert-Mb ((ConvertTo-MbPowerPointRgb 255 0 0) -eq 255) '赤をRGB値へ変換できる'
    Assert-Mb ((ConvertTo-MbPowerPointRgb 0 0 255) -eq 16711680) '青をRGB値へ変換できる（PowerPointはBGR順）'

    Write-Host ''
    Write-Host 'PowerPoint utility tests passed.' -ForegroundColor Cyan
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
