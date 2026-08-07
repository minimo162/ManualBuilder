[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProfileRoot,
    [AllowEmptyString()][string]$ConfigPath = '',
    [AllowEmptyString()][string]$StatusPath = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1') -Force

function Write-WarmupStatus {
    param([string]$State, [string]$Message)
    if ([string]::IsNullOrWhiteSpace($StatusPath)) { return }
    try {
        $directory = Split-Path -Parent $StatusPath
        if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
        $body = [pscustomobject]@{
            state = $State
            message = $Message
            updatedAt = [DateTime]::UtcNow.ToString('o')
        } | ConvertTo-Json -Compress
        [IO.File]::WriteAllText($StatusPath, $body, (New-Object Text.UTF8Encoding($false)))
    } catch { }
}

try {
    Write-WarmupStatus -State 'starting' -Message 'Copilot用Edgeを準備しています。'
    $settings = Get-MbCopilotSettings -ConfigPath $ConfigPath
    Start-MbCopilotEdge -Settings $settings -ProfileDirectory $ProfileRoot
    # CDPが応答するだけでは、Edgeが画面外・最小化・ウィンドウなしの状態を
    # 起動成功と誤認できる。Copilotタブを持つEdgeを画面上へ出し、表示結果まで確認する。
    if (-not (Show-MbCopilotWindow -Settings $settings -ProfileDirectory $ProfileRoot)) {
        throw 'Copilot用Edgeの画面を確認できませんでした。'
    }
    $page = Get-MbCopilotPage -Settings $settings
    $gate = Wait-MbCopilotScreenReady -WsUrl ([string]$page.webSocketDebuggerUrl) -Settings $settings -TimeoutSeconds 60
    if ($gate.ok) {
        Write-WarmupStatus -State 'ready' -Message 'Copilotを利用できます。'
        exit 0
    }
    Write-WarmupStatus -State 'attention' -Message ([string]$gate.message)
    exit 2
} catch {
    Write-WarmupStatus -State 'error' -Message $_.Exception.Message
    exit 1
}
