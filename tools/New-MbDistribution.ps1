# 共有フォルダーへ置く配布フォルダーを作る。
# ランチャー（ManualBuilder.Launcher.psm1）がコピー・照合する4項目だけを取り出し、
# 配布前に見落としやすい不整合を止める。
#
#   .\tools\make-dist.cmd
#   .\tools\New-MbDistribution.ps1 -Destination '\\server\share\ManualBuilder'

[CmdletBinding()]
param(
    # 出力先。既定はリポジトリの隣の ManualBuilder-dist。共有フォルダーを直接指定してもよい。
    [string]$Destination,
    # src\README.md と web\README.md（開発者向け）も含める
    [switch]$IncludeDeveloperNotes,
    # 出力先に既にある配布物を確認なしで置き換える
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
# ランチャーが必要とする4項目。ここを増減するとローカルキャッシュの照合が壊れる。
$packageEntries = @('src', 'web', 'run.cmd', 'app-version.json')
$developerNotes = @('src\README.md', 'web\README.md')

function Write-MbStep {
    param([string]$Text)
    Write-Host ''
    Write-Host "== $Text" -ForegroundColor Cyan
}

function Get-MbFileHash256 {
    param([string]$Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $stream = [IO.File]::OpenRead($Path)
        try { return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
        finally { $stream.Dispose() }
    } finally { $sha.Dispose() }
}

# --- 1. 必須項目がそろっているか ---
Write-MbStep '必須項目を確認します'
foreach ($entry in $packageEntries) {
    $path = Join-Path $repoRoot $entry
    if (-not (Test-Path -LiteralPath $path)) { throw "必須項目が見つかりません: $entry" }
    Write-Host "   OK  $entry"
}

# --- 2. バージョン表記がそろっているか ---
# バージョンは5ファイルに散在する。1つでも古いと、共有フォルダーへ置いても
# 各PCが更新を検知しない（またはブラウザーが古い画面を使い続ける）。
Write-MbStep 'バージョン表記の整合を確認します'
$manifest = [IO.File]::ReadAllText((Join-Path $repoRoot 'app-version.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$appVersion = [string]$manifest.appVersion
if ($appVersion -notmatch '^\d+\.\d+\.\d+$') { throw "app-version.json のバージョンが不正です: $appVersion" }
Write-Host "   app-version.json : $appVersion"

$versionChecks = @(
    @{ Path = 'web\assets\js\app.js';        Pattern = "const appVersion = '([0-9.]+)'"; Label = 'app.js の定数' },
    @{ Path = 'web\index.html';              Pattern = '\?v=([0-9.]+)';                  Label = 'index.html のキャッシュ更新URL' },
    @{ Path = 'src\ManualBuilder.Web.psm1';  Pattern = 'data-app-version="([0-9.]+)"';   Label = '画面へ埋め込むバージョン' },
    @{ Path = 'src\ManualBuilder.Web.psm1';  Pattern = '>v([0-9.]+)<';                   Label = '画面に表示するバージョン' }
)
$versionProblems = @()
foreach ($check in $versionChecks) {
    $text = [IO.File]::ReadAllText((Join-Path $repoRoot $check.Path), [Text.Encoding]::UTF8)
    $matches = [regex]::Matches($text, $check.Pattern)
    if ($matches.Count -eq 0) {
        $versionProblems += "$($check.Label)（$($check.Path)）にバージョン表記が見つかりません"
        continue
    }
    $found = @($matches | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    $wrong = @($found | Where-Object { $_ -ne $appVersion })
    if ($wrong.Count -gt 0) {
        $versionProblems += "$($check.Label)（$($check.Path)）が $($wrong -join ', ') のままです"
    } else {
        Write-Host "   OK  $($check.Label)（$($matches.Count)箇所）"
    }
}
if ($versionProblems.Count -gt 0) {
    Write-Host ''
    foreach ($problem in $versionProblems) { Write-Host "   NG  $problem" -ForegroundColor Red }
    throw "バージョン表記が $appVersion に揃っていません。配布しても各PCが更新を検知しません。"
}

# --- 3. リンクが混ざっていないか ---
# ランチャーはリパースポイントを拒否するため、配布時点で弾く。
Write-MbStep 'リンク（シンボリックリンク・ジャンクション）が無いか確認します'
$sourceFiles = New-Object System.Collections.ArrayList
foreach ($entry in $packageEntries) {
    $item = Get-Item -LiteralPath (Join-Path $repoRoot $entry) -Force
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "リンクは配布できません: $entry" }
    $children = if ($item.PSIsContainer) { @(Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force) } else { @($item) }
    foreach ($child in $children) {
        if (($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "リンクは配布できません: $($child.FullName)" }
        $relative = $child.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
        if (-not $IncludeDeveloperNotes -and ($developerNotes -contains $relative)) { continue }
        [void]$sourceFiles.Add([pscustomobject]@{ Relative = $relative; FullName = $child.FullName; Length = [long]$child.Length })
    }
}
Write-Host "   OK  リンクなし（$($sourceFiles.Count) ファイル）"

# --- 4. 出力先を決める ---
if ([string]::IsNullOrWhiteSpace($Destination)) {
    $Destination = Join-Path (Split-Path -Parent $repoRoot) 'ManualBuilder-dist'
}
$destinationRoot = [IO.Path]::GetFullPath($Destination)
$repoPrefix = $repoRoot.TrimEnd('\') + '\'
if ($destinationRoot.TrimEnd('\') -eq $repoRoot.TrimEnd('\')) { throw '出力先にリポジトリ自身は指定できません。' }
if ($destinationRoot.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "出力先をリポジトリの内側にはできません: $destinationRoot"
}

Write-MbStep "出力先: $destinationRoot"
$existing = @($packageEntries | Where-Object { Test-Path -LiteralPath (Join-Path $destinationRoot $_) })
if ($existing.Count -gt 0 -and -not $Force) {
    Write-Host "   既存の配布物があります: $($existing -join ', ')" -ForegroundColor Yellow
    $answer = Read-Host '   置き換えますか？ (y/N)'
    if ($answer -ne 'y' -and $answer -ne 'Y') { Write-Host '中止しました。'; exit 1 }
}
if (-not (Test-Path -LiteralPath $destinationRoot)) { [void](New-Item -ItemType Directory -Path $destinationRoot -Force) }

# 出力先の他のファイルは触らない。管理対象の4項目だけ入れ替える。
foreach ($entry in $packageEntries) {
    $target = Join-Path $destinationRoot $entry
    if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
}

# --- 5. コピー ---
Write-MbStep 'コピーします'
foreach ($file in $sourceFiles) {
    $target = Join-Path $destinationRoot $file.Relative
    $targetDir = Split-Path -Parent $target
    if (-not (Test-Path -LiteralPath $targetDir)) { [void](New-Item -ItemType Directory -Path $targetDir -Force) }
    Copy-Item -LiteralPath $file.FullName -Destination $target -Force
}
Write-Host "   $($sourceFiles.Count) ファイルをコピーしました"

# --- 6. SHA-256で照合 ---
Write-MbStep 'コピー結果をSHA-256で照合します'
$mismatch = @()
foreach ($file in $sourceFiles) {
    $target = Join-Path $destinationRoot $file.Relative
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { $mismatch += "$($file.Relative)（コピーされていません）"; continue }
    if ((Get-MbFileHash256 -Path $file.FullName) -ne (Get-MbFileHash256 -Path $target)) { $mismatch += "$($file.Relative)（内容が一致しません）" }
}
# 配布対象の4項目の中に、想定外のファイルが残っていないか見る（出力先の他の物には触れない）
$expected = @{}
foreach ($file in $sourceFiles) { $expected[$file.Relative.ToLowerInvariant()] = $true }
$destinationPrefixLength = $destinationRoot.TrimEnd('\').Length
foreach ($entry in $packageEntries) {
    $entryPath = Join-Path $destinationRoot $entry
    if (-not (Test-Path -LiteralPath $entryPath)) { continue }
    $item = Get-Item -LiteralPath $entryPath -Force
    $children = if ($item.PSIsContainer) { @(Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force) } else { @($item) }
    foreach ($child in $children) {
        $relative = $child.FullName.Substring($destinationPrefixLength).TrimStart('\')
        if (-not $expected.ContainsKey($relative.ToLowerInvariant())) {
            $mismatch += "$relative（配布対象外のファイルが残っています）"
        }
    }
}

if ($mismatch.Count -gt 0) {
    Write-Host ''
    foreach ($problem in $mismatch) { Write-Host "   NG  $problem" -ForegroundColor Red }
    throw '照合に失敗しました。出力先を確認してください。'
}
$totalBytes = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
$totalSizeText = '{0:N1} MB' -f ($totalBytes / 1MB)
Write-Host "   OK  全 $($sourceFiles.Count) ファイル一致（$totalSizeText）"

# --- 7. 結果 ---
Write-Host ''
Write-Host '----------------------------------------------------------------------' -ForegroundColor Green
Write-Host " 配布フォルダーを作成しました" -ForegroundColor Green
Write-Host '----------------------------------------------------------------------' -ForegroundColor Green
Write-Host " バージョン : $appVersion"
Write-Host " 出力先     : $destinationRoot"
Write-Host " 内容       : $($packageEntries -join ' / ')"
if (-not $IncludeDeveloperNotes) { Write-Host " 除外       : $($developerNotes -join ' / ')" }
Write-Host ''
Write-Host " このフォルダーの中身を共有フォルダーへ置いてください。"
Write-Host " 利用者は run.cmd をダブルクリックします。"
Write-Host ''
Write-Host " docs / tests / .git などは意図的に含めていません。"
Write-Host " docs にはレビューで見つけた弱点の記録が含まれるため、共有フォルダーへは置かないでください。"
Write-Host ''
exit 0
