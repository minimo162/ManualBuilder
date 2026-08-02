# Phase 1 web rendering tests (no server, no browser).

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Web.psm1') -Force

function Assert-Mb {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "NG: $Message" }
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function New-MbTestStep {
    param([int]$AnnotationCount)
    $annotations = @()
    for ($i = 1; $i -le $AnnotationCount; $i++) {
        $annotations += [pscustomobject]@{
            id    = 'annotation-' + ([guid]::NewGuid().ToString('N'))
            type  = 'number'
            x1    = 0.5; y1 = 0.5; x2 = 0.5; y2 = 0.5
            label = $i
        }
    }
    return [pscustomobject]@{
        id          = 'step-' + ([guid]::NewGuid().ToString('N'))
        title       = 'テスト手順'
        description = 'テスト説明'
        note        = ''
        imageId     = 'image-' + ([guid]::NewGuid().ToString('N'))
        annotations = @($annotations)
        crop        = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
    }
}

function Get-MbAnnotationsJson {
    param([string]$Html)
    $match = [regex]::Match($Html, '<textarea class="step-annotations-data" hidden>(.*?)</textarea>')
    if (-not $match.Success) { throw '注釈データのtextareaが見つかりません。' }
    return [Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
}

Write-Host '--- 注釈データのJSON形式 ---' -ForegroundColor Cyan

# 1件のときにPowerShellのパイプが配列を解いてしまうと、画面側のArray.isArray判定で
# 注釈が空扱いになり、表示が消えたうえ画像編集の保存で失われる。
foreach ($count in @(1, 2, 3)) {
    $step = New-MbTestStep -AnnotationCount $count
    $html = ConvertTo-MbStepCardHtml -Step $step -Number 1 -Total 1 -Token 'testtoken'
    $json = Get-MbAnnotationsJson -Html $html
    Assert-Mb ($json.StartsWith('[')) "注釈が${count}件でもJSON配列で出力する"
    $parsed = $json | ConvertFrom-Json
    Assert-Mb (@($parsed).Count -eq $count) "注釈${count}件が欠落せずJSONへ入る"
}

$emptyStep = New-MbTestStep -AnnotationCount 0
$emptyHtml = ConvertTo-MbStepCardHtml -Step $emptyStep -Number 1 -Total 1 -Token 'testtoken'
Assert-Mb ((Get-MbAnnotationsJson -Html $emptyHtml) -eq '[]') '注釈が0件のときは空配列を出力する'

Write-Host ''
Write-Host 'Web rendering tests passed.' -ForegroundColor Green
exit 0
