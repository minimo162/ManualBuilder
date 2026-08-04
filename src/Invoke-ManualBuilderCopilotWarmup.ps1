# ManualBuilder起動時に、Copilot制御用Edgeを通常表示で準備する。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProfileDirectory,
    [AllowEmptyString()][string]$ConfigPath = '',
    [Parameter(Mandatory = $true)][string]$StatusPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Copilot.psm1') -Force
$script:Utf8NoBom = New-Object Text.UTF8Encoding($false)

function Write-MbWarmupStatus {
    param([string]$State, [string]$Message, [string]$ErrorCode = '')

    $status = [pscustomobject]@{
        state = $State; message = $Message; errorCode = $ErrorCode
        updatedAt = [DateTime]::UtcNow.ToString('o'); visible = $true; canOpen = $true
    }
    $directory = Split-Path -Parent $StatusPath
    if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory -Force) }
    # 短い状態JSONなので直接書く。置換APIはウイルス対策ソフト等が読み取り中だと
    # 「置換されるファイルを削除できません」になるため使用しない。
    $json = $status | ConvertTo-Json -Depth 5
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        try { [IO.File]::WriteAllText($StatusPath, $json, $script:Utf8NoBom); return }
        catch {
            if ($attempt -ge 4) { throw }
            Start-Sleep -Milliseconds 80
        }
    }
}

try {
    Write-MbWarmupStatus -State 'starting' -Message 'Copilot用Edgeを起動しています。'
    $settings = Get-MbCopilotSettings -ConfigPath $ConfigPath
    # 起動時ウォームアップは、サインイン切れや停止を利用者が確認できる通常表示に固定する。
    $settings.browser_display_mode = 'foreground'
    Start-MbCopilotEdge -Settings $settings -ProfileDirectory $ProfileDirectory

    Write-MbWarmupStatus -State 'loading' -Message 'Copilot画面を読み込んでいます。'
    $page = Get-MbCopilotPage -Settings $settings
    $wsUrl = [string]$page.webSocketDebuggerUrl
    $null = Show-MbCopilotWindow -Settings $settings -ProfileDirectory $ProfileDirectory
    $gate = Wait-MbCopilotScreenReady -WsUrl $wsUrl -Settings $settings -TimeoutSeconds 90
    if ($gate.ok) {
        $null = Set-MbCopilotWindowMarker -WsUrl $wsUrl
        Write-MbWarmupStatus -State 'ready' -Message 'Copilotの準備ができました。'
        exit 0
    }
    if ($gate.signinRequired) {
        Write-MbWarmupStatus -State 'signin-required' -Message 'Copilot画面でサインインし、完了後にManualBuilderの表示をクリックしてください。' -ErrorCode 'COPILOT_SIGNIN_REQUIRED'
        exit 0
    }
    $code = if ($gate.PSObject.Properties.Name -contains 'errorCode') { [string]$gate.errorCode } else { 'COPILOT_CHAT_NOT_READY' }
    Write-MbWarmupStatus -State 'failed' -Message ([string]$gate.message) -ErrorCode $code
    exit 1
} catch {
    Write-MbWarmupStatus -State 'failed' -Message ('Copilot画面を準備できませんでした: ' + $_.Exception.Message) -ErrorCode 'COPILOT_EDGE_START_FAILED'
    exit 1
}
