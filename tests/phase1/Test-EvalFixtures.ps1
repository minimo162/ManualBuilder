# 人が採点するときの基準として使う、正解付き録画セットの構造と実体を検査する。

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$manifestPath = Join-Path $repoRoot 'samples\evals\gold\manifest.json'
$manifestDirectory = Split-Path -Parent $manifestPath
$errors = New-Object 'System.Collections.Generic.List[string]'

function Add-Result {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "[OK] $Message" -ForegroundColor Green }
    else { Write-Host "[NG] $Message" -ForegroundColor Red; [void]$errors.Add($Message) }
}

try {
    Add-Result (Test-Path -LiteralPath $manifestPath -PathType Leaf) '正解マニフェストが存在する'
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json

    Add-Result ([int]$manifest.schemaVersion -eq 1) '対応する正解データ形式である'
    Add-Result (@($manifest.scenarios).Count -eq 3) '評価用シナリオが3件ある'
    Add-Result ([bool]$manifest.evaluationDefaults.sceneCountMustMatch) '場面数を厳密に採点する設定である'

    $expectedCounts = @{
        'expense-application' = 5
        'settings-roundtrip' = 3
        'terms-slow-scroll' = 4
    }
    $totalScenes = 0

    foreach ($scenario in @($manifest.scenarios)) {
        $scenarioId = [string]$scenario.id
        $scenes = @($scenario.scenes)
        $totalScenes += $scenes.Count
        Add-Result ($expectedCounts.ContainsKey($scenarioId)) "既知のシナリオである: $scenarioId"
        if ($expectedCounts.ContainsKey($scenarioId)) {
            Add-Result ($scenes.Count -eq $expectedCounts[$scenarioId]) "正解場面数が一致する: $scenarioId"
        }
        Add-Result ([int]$scenario.expectedSceneCount -eq $scenes.Count) "宣言した場面数と正解が一致する: $scenarioId"
        Add-Result ([int]$scenario.durationMs -gt 0) "録画時間が定義されている: $scenarioId"

        $videoPath = [IO.Path]::GetFullPath((Join-Path $manifestDirectory ([string]$scenario.videoFile)))
        $videoExists = Test-Path -LiteralPath $videoPath -PathType Leaf
        Add-Result $videoExists "録画ファイルが存在する: $scenarioId"
        if ($videoExists) {
            $videoInfo = Get-Item -LiteralPath $videoPath
            Add-Result ($videoInfo.Length -gt 10000) "録画ファイルに十分なデータがある: $scenarioId"
            Add-Result ($videoInfo.Length -eq [long]$scenario.byteLength) "録画ファイルの長さが正解データと一致する: $scenarioId"
            $actualHash = (Get-FileHash -LiteralPath $videoPath -Algorithm SHA256).Hash.ToLowerInvariant()
            Add-Result ($actualHash -eq [string]$scenario.sha256) "録画ファイルのSHA-256が一致する: $scenarioId"
            $stream = [IO.File]::OpenRead($videoPath)
            try {
                $signature = New-Object byte[] 4
                [void]$stream.Read($signature, 0, 4)
                $isWebM = $signature[0] -eq 0x1A -and $signature[1] -eq 0x45 -and
                    $signature[2] -eq 0xDF -and $signature[3] -eq 0xA3
                Add-Result $isWebM "WebMのヘッダーを持つ: $scenarioId"
            } finally {
                $stream.Dispose()
            }
        }

        for ($index = 0; $index -lt $scenes.Count; $index++) {
            $scene = $scenes[$index]
            $sceneLabel = "$scenarioId / scene $($index + 1)"
            Add-Result ([int]$scene.order -eq ($index + 1)) "場面順が連番である: $sceneLabel"

            $timeRange = @($scene.representativeTimeRangeMs)
            $validTimeRange = $timeRange.Count -eq 2 -and [int]$timeRange[0] -ge 0 -and
                [int]$timeRange[0] -lt [int]$timeRange[1] -and [int]$timeRange[1] -le [int]$scenario.durationMs
            Add-Result $validTimeRange "代表時刻が録画範囲内である: $sceneLabel"
            Add-Result (@($scene.requiredConcepts).Count -gt 0) "必須概念が定義されている: $sceneLabel"
            Add-Result (@($scene.forbiddenClaims).Count -gt 0) "禁止する誤記が定義されている: $sceneLabel"
            Add-Result (-not [string]::IsNullOrWhiteSpace([string]$scene.expectedTitle)) "期待タイトルがある: $sceneLabel"
            Add-Result (-not [string]::IsNullOrWhiteSpace([string]$scene.expectedDescription)) "期待説明文がある: $sceneLabel"

            if ($null -ne $scene.expectedRect) {
                $rect = $scene.expectedRect
                $validRect = [double]$rect.x1 -ge 0 -and [double]$rect.y1 -ge 0 -and
                    [double]$rect.x2 -le 1 -and [double]$rect.y2 -le 1 -and
                    [double]$rect.x1 -lt [double]$rect.x2 -and [double]$rect.y1 -lt [double]$rect.y2
                Add-Result $validRect "操作矩形が正規化座標内である: $sceneLabel"
            }
        }
    }

    Add-Result ($totalScenes -eq 12) '全12場面の正解が揃っている'
} catch {
    Add-Result $false ("評価用データの検査中に例外: " + $_.Exception.Message)
}

if ($errors.Count -gt 0) {
    Write-Host "`n$($errors.Count) checks failed." -ForegroundColor Red
    exit 1
}

Write-Host "`nEvaluation fixture checks passed." -ForegroundColor Green
exit 0
