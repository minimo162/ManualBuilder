# ManualBuilder dictation worker.
#
# 操作記録と並走して、話した内容を文字にする専用プロセス。
# RecognizeAsync は発話が終わるまで戻らないため、記録ループとは必ず分ける。
# 結果は narration.jsonl へ追記し、取り込み時に操作と時刻で突き合わせる。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [Parameter(Mandatory = $true)][string]$StopPath,
    [Parameter(Mandatory = $true)][string]$PausePath,
    [Parameter(Mandatory = $true)][string]$StatusPath,
    [Parameter(Mandatory = $true)][long]$StartedAtUtcTicks
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Dictation.psm1') -Force

$encoding = New-Object Text.UTF8Encoding($false)

function Write-MbDictationStatus {
    param([string]$State, [int]$Count = 0, [string]$Message = '')
    $status = [pscustomobject]@{
        state = $State; count = $Count; message = $Message; updatedAt = [DateTime]::UtcNow.ToString('o')
    }
    $temporary = $StatusPath + '.tmp'
    [IO.File]::WriteAllText($temporary, ($status | ConvertTo-Json -Depth 4), $encoding)
    [IO.File]::Copy($temporary, $StatusPath, $true)
    Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
}

try {
    Write-MbDictationStatus -State 'listening' -Message '音声を聞き取っています'
    $count = Invoke-MbDictationLoop -OutputPath $OutputPath -StopPath $StopPath -PausePath $PausePath `
        -StartedAtUtcTicks $StartedAtUtcTicks
    Write-MbDictationStatus -State 'completed' -Count ([int]$count) -Message ("$count 件の発話を文字にしました")
    exit 0
} catch {
    try {
        # 記録そのものは続けられるので、失敗しても操作の記録は止めない。
        Write-MbDictationStatus -State 'failed' -Message ('音声を文字にできませんでした: ' + $_.Exception.Message)
    } catch { }
    exit 1
}
