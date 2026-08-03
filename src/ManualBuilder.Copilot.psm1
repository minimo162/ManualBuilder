# Microsoft 365 Copilot をブラウザー越しに操作する。
#
# PdfKoseiAssist の CopilotClient.ps1 を土台に、ManualBuilder 向けへ作り替えたもの。
# 変更点は、添付するのがPDFではなく手順の画面画像であることと、受け取るJSONが
# 校正指摘ではなく手順の下書きであること。CDPの扱い方そのものは同じ。
#
# APIキーもGraphの権限も使わない。利用者が普段サインインしているMicrosoft 365の
# Copilotを、専用のEdgeプロファイルで開いて操作する。管理者権限は要らない。
#
# PowerShell 5.1 での約束事:
#   1. 非同期の .GetAwaiter().GetResult() は必ず $null = か変数で受ける。
#      戻り値がパイプラインを汚し、関数の戻り値が配列になる事故を防ぐ。
#   2. Invoke-RestMethod がJSON配列を返すときは平坦化してから使う。
#   3. CDPのnodeIdは接続ごとの値。DOM操作の連なりは同じWebSocket上で行う。
#   4. サインイン中は login.microsoftonline.com へ飛ぶため、ターゲット選択は
#      Copilotのホスト一致 → 通常のhttp(s)ページ、の順に緩める。

$script:MbCdpNextId = 52000
$script:MbCopilotLogger = $null

function Set-MbCopilotLogger {
    param([AllowNull()][scriptblock]$Logger)
    $script:MbCopilotLogger = $Logger
}

function Write-MbCopilotLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ($null -eq $script:MbCopilotLogger) { return }
    try { & $script:MbCopilotLogger $Message $Level } catch { }
}

# ---------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------
function Get-MbCopilotDefaultSettings {
    return [ordered]@{
        copilot_url          = 'https://m365.cloud.microsoft/chat/'
        cdp_port             = 9455
        request_timeout      = 600
        max_prompt_chars     = 60000
        attach_wait_seconds  = 90
        # 1回の依頼で渡す手順の数。画像はこの数だけ添付される。
        # 多いほど前後の文脈が効くが、添付とトークンの上限に当たりやすくなる。
        steps_per_packet     = 6
        copilot_model        = 'GPT 5.6 Think deeper,Opus,Think Deeper'
        browser_display_mode = 'minimized'
        poll_interval_ms     = 2000
        response_end_marker  = 'MB_END'
        selectors            = [ordered]@{
            file_input          = '#upload-file-button'
            file_input_fallback = 'input[type="file"]'
            attachment_item_any = @('.fai-BebopAttachment', '.fai-Attachment', '[class*="Attachment"][data-overflow-item]')
            attachment_name_any = @('.fai-BebopAttachment__content > span:first-child', '.fai-Attachment__content span')
            upload_done_pattern = '完了しました|upload(ed)?\s*(complete|finished)'
            upload_fail_pattern = '失敗|エラー|failed|error'
            model_switcher      = '#gptModeSwitcher'
            chat_input_any      = @('#m365-chat-editor-target-element', '[data-lexical-editor="true"][contenteditable]', '[role="textbox"][contenteditable]')
        }
    }
}

function Get-MbCopilotSettings {
    param([AllowEmptyString()][string]$ConfigPath = '')

    $defaults = Get-MbCopilotDefaultSettings
    if (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        $loaded = $null
        try {
            $raw = [IO.File]::ReadAllText($ConfigPath, [Text.Encoding]::UTF8)
            $loaded = $raw | ConvertFrom-Json
        } catch {
            Write-MbCopilotLog ('Copilot設定の読み込みに失敗しました（既定値を使います）: ' + $_.Exception.Message) 'WARN'
            $loaded = $null
        }
        if ($loaded) {
            foreach ($property in $loaded.PSObject.Properties) {
                if ($property.Name -eq 'selectors' -and $property.Value) {
                    foreach ($selector in $property.Value.PSObject.Properties) {
                        $defaults.selectors[$selector.Name] = $selector.Value
                    }
                } else {
                    $defaults[$property.Name] = $property.Value
                }
            }
        }
    }
    return [pscustomobject]$defaults
}

function Get-MbCopilotSelector {
    param([Parameter(Mandatory = $true)]$Settings, [Parameter(Mandatory = $true)][string]$Name)
    $selectors = $Settings.selectors
    if ($selectors -is [System.Collections.IDictionary]) { return $selectors[$Name] }
    return $selectors.PSObject.Properties[$Name].Value
}

# ---------------------------------------------------------------------
# Edge / DevTools
# ---------------------------------------------------------------------
function Get-MbEdgePath {
    $roots = @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA) | Where-Object { $_ }
    foreach ($root in $roots) {
        $candidate = Join-Path $root 'Microsoft\Edge\Application\msedge.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    $command = Get-Command 'msedge.exe' -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw 'Microsoft Edgeが見つかりません。Edgeをインストールしてから、もう一度実行してください。'
}

function Test-MbDevTools {
    param([int]$Port)
    try {
        $null = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$Port/json/version" -TimeoutSec 2
        return $true
    } catch { return $false }
}

function Wait-MbDevTools {
    param([int]$Port, [int]$TimeoutSeconds = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-MbDevTools -Port $Port) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

function Start-MbCopilotEdge {
    param(
        [Parameter(Mandatory = $true)]$Settings,
        [Parameter(Mandatory = $true)][string]$ProfileDirectory
    )

    $port = [int]$Settings.cdp_port
    if (Test-MbDevTools -Port $port) { return }
    $edge = Get-MbEdgePath
    if (-not (Test-Path -LiteralPath $ProfileDirectory)) {
        [void](New-Item -ItemType Directory -Path $ProfileDirectory -Force)
    }

    $arguments = @(
        "--remote-debugging-port=$port",
        '--remote-debugging-address=127.0.0.1',
        '--remote-allow-origins=*',
        "--user-data-dir=$ProfileDirectory",
        '--no-first-run',
        # 画面外・最小化のままでもページのタイマーを止めない。応答待機が凍る事故を防ぐ。
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-renderer-backgrounding',
        '--disable-features=CalculateNativeWinOcclusion,msEdgeTranslate'
    )
    $display = 'minimized'
    try { $display = [string]$Settings.browser_display_mode } catch { }
    if ($display -ne 'foreground') { $display = 'minimized' }
    if ($display -eq 'minimized') {
        $arguments += '--window-position=-32000,-32000'
        $arguments += '--window-size=1280,900'
    }
    $arguments += [string]$Settings.copilot_url

    Write-MbCopilotLog "Edgeを起動します: port=$port display=$display" 'INFO'
    try {
        if ($display -eq 'minimized') {
            Start-Process -FilePath $edge -ArgumentList $arguments -WindowStyle Minimized | Out-Null
        } else {
            Start-Process -FilePath $edge -ArgumentList $arguments | Out-Null
        }
    } catch {
        if ($display -ne 'minimized') { throw }
        Write-MbCopilotLog ('画面外での起動に失敗したため最小化で起動し直します: ' + $_.Exception.Message) 'WARN'
        $fallback = @($arguments | Where-Object { $_ -notlike '--window-position=*' -and $_ -notlike '--window-size=*' })
        Start-Process -FilePath $edge -ArgumentList $fallback -WindowStyle Minimized | Out-Null
    }

    if (-not (Wait-MbDevTools -Port $port -TimeoutSeconds 30)) {
        throw "Edgeを操作できる状態にできませんでした（ポート $port）。このアプリ専用のEdgeウィンドウをすべて閉じてから、もう一度実行してください。"
    }
}

# ---------------------------------------------------------------------
# CDPターゲット
# ---------------------------------------------------------------------
function Get-MbCdpTargets {
    param([int]$Port)
    $raw = $null
    try { $raw = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$Port/json" -TimeoutSec 5 } catch { return @() }
    # 配列の配列で返ることがあるため平坦化する（約束事2）。
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($raw)) {
        if ($null -eq $item) { continue }
        if ($item -is [System.Array]) {
            foreach ($inner in $item) { if ($null -ne $inner) { $targets.Add($inner) } }
        } else {
            $targets.Add($item)
        }
    }
    return $targets.ToArray()
}

function Get-MbCopilotPage {
    param([Parameter(Mandatory = $true)]$Settings)

    $port = [int]$Settings.cdp_port
    $url = [string]$Settings.copilot_url
    $copilotHost = ([Uri]$url).Host
    for ($attempt = 0; $attempt -lt 3; $attempt++) {
        $targets = @(Get-MbCdpTargets -Port $port)
        $pages = @($targets | Where-Object {
            $_ -and ([string]$_.type) -eq 'page' -and
            (-not [string]::IsNullOrWhiteSpace([string]$_.webSocketDebuggerUrl)) -and
            ((([string]$_.url) -like ('*' + $copilotHost + '*')) -or (([string]$_.url) -like '*copilot*'))
        })
        if ($pages.Count -eq 0) {
            # サインインへ飛んでいる最中の受け皿（約束事4）。
            $pages = @($targets | Where-Object {
                $_ -and ([string]$_.type) -eq 'page' -and
                (-not [string]::IsNullOrWhiteSpace([string]$_.webSocketDebuggerUrl)) -and
                (([string]$_.url) -like 'http*')
            })
        }
        if ($pages.Count -gt 0) { return $pages[0] }

        foreach ($method in @('Put', 'Get')) {
            try {
                $null = Invoke-WebRequest -UseBasicParsing -Method $method -Uri ("http://127.0.0.1:$port/json/new?" + [Uri]::EscapeUriString($url)) -TimeoutSec 5
                break
            } catch { }
        }
        Start-Sleep -Seconds 2
    }
    throw 'Copilotの画面を開けませんでした。Edgeを閉じてから、もう一度実行してください。'
}

# ---------------------------------------------------------------------
# WebSocket / CDP
# ---------------------------------------------------------------------
function Receive-MbWsMessage {
    param([Parameter(Mandatory = $true)]$WebSocket, [int]$TimeoutSeconds = 30)

    $buffer = New-Object byte[] 65536
    $memory = New-Object System.IO.MemoryStream
    $cancellation = [System.Threading.CancellationTokenSource]::new()
    $cancellation.CancelAfter([TimeSpan]::FromSeconds([Math]::Max(1, $TimeoutSeconds)))
    try {
        do {
            $segment = [ArraySegment[byte]]::new($buffer)
            $received = $WebSocket.ReceiveAsync($segment, $cancellation.Token).GetAwaiter().GetResult()
            if ($received.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { return $null }
            if ($received.Count -gt 0) { $memory.Write($buffer, 0, $received.Count) }
        } while (-not $received.EndOfMessage)
        return [System.Text.Encoding]::UTF8.GetString($memory.ToArray())
    } catch {
        return $null
    } finally {
        try { $cancellation.Dispose() } catch { }
        try { $memory.Dispose() } catch { }
    }
}

function Connect-MbWebSocket {
    param([Parameter(Mandatory = $true)][string]$WebSocketUrl, [int]$TimeoutSeconds = 15)
    $socket = [System.Net.WebSockets.ClientWebSocket]::new()
    $cancellation = [System.Threading.CancellationTokenSource]::new()
    $cancellation.CancelAfter([TimeSpan]::FromSeconds($TimeoutSeconds))
    try {
        $null = $socket.ConnectAsync([Uri]$WebSocketUrl, $cancellation.Token).GetAwaiter().GetResult()
    } finally {
        $cancellation.Dispose()
    }
    return $socket
}

function Invoke-MbCdpOnSocket {
    param(
        [Parameter(Mandatory = $true)]$WebSocket,
        [Parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{},
        [int]$TimeoutSeconds = 30
    )

    $requestId = $script:MbCdpNextId
    $script:MbCdpNextId = $script:MbCdpNextId + 1
    $payload = @{ id = $requestId; method = $Method; params = $Params } | ConvertTo-Json -Depth 30 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $segment = [ArraySegment[byte]]::new($bytes)
    $sendCancellation = [System.Threading.CancellationTokenSource]::new()
    $sendCancellation.CancelAfter([TimeSpan]::FromSeconds(15))
    try {
        $null = $WebSocket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $sendCancellation.Token).GetAwaiter().GetResult()
    } finally {
        $sendCancellation.Dispose()
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $remaining = [int][Math]::Max(1, [Math]::Ceiling(($deadline - (Get-Date)).TotalSeconds))
        $message = Receive-MbWsMessage -WebSocket $WebSocket -TimeoutSeconds $remaining
        if ([string]::IsNullOrWhiteSpace($message)) { continue }
        $parsed = $null
        try { $parsed = $message | ConvertFrom-Json } catch { continue }
        if ($parsed -and ($parsed.PSObject.Properties.Name -contains 'id') -and $parsed.id -eq $requestId) { return $parsed }
    }
    throw "Copilot画面の応答がありませんでした: $Method"
}

function Invoke-MbCdpMethod {
    param(
        [Parameter(Mandatory = $true)][string]$WebSocketUrl,
        [Parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{},
        [int]$TimeoutSeconds = 30
    )
    $socket = Connect-MbWebSocket -WebSocketUrl $WebSocketUrl
    try {
        return (Invoke-MbCdpOnSocket -WebSocket $socket -Method $Method -Params $Params -TimeoutSeconds $TimeoutSeconds)
    } finally {
        try { $socket.Dispose() } catch { }
    }
}

function Invoke-MbCdpEval {
    param(
        [Parameter(Mandatory = $true)][string]$WebSocketUrl,
        [Parameter(Mandatory = $true)][string]$Expression,
        [int]$TimeoutSeconds = 30
    )
    $response = Invoke-MbCdpMethod -WebSocketUrl $WebSocketUrl -Method 'Runtime.evaluate' -Params @{
        expression    = $Expression
        awaitPromise  = $true
        returnByValue = $true
        userGesture   = $true
    } -TimeoutSeconds $TimeoutSeconds
    if ($response.error) { throw ($response.error | ConvertTo-Json -Compress) }
    if ($response.result.exceptionDetails) {
        throw ('Copilot画面の操作に失敗しました: ' + ($response.result.exceptionDetails | ConvertTo-Json -Depth 20 -Compress))
    }
    return $response.result.result.value
}

function ConvertTo-MbJsString {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { $Value = '' }
    return (ConvertTo-Json ([string]$Value))
}

function Get-MbChatInputSelectorsJson {
    param([Parameter(Mandatory = $true)]$Settings)
    $selectors = @([string[]](Get-MbCopilotSelector -Settings $Settings -Name 'chat_input_any'))
    return (ConvertTo-Json $selectors -Compress)
}

# ---------------------------------------------------------------------
# 画面の状態
# ---------------------------------------------------------------------
function Get-MbCopilotScreenState {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)

    $fileSelectors = @(
        [string](Get-MbCopilotSelector -Settings $Settings -Name 'file_input'),
        [string](Get-MbCopilotSelector -Settings $Settings -Name 'file_input_fallback')
    ) | Where-Object { $_ }
    $fileSelectorsJson = ConvertTo-Json -InputObject @($fileSelectors) -Compress

    $template = @'
(() => {
  const sels = __INPUT_SELS__;
  const fileSels = __FILE_SELS__;
  const visible = el => {if(!el)return false;const r=el.getBoundingClientRect(),cs=el.ownerDocument.defaultView.getComputedStyle(el);return r.width>0&&r.height>0&&cs.display!=='none'&&cs.visibility!=='hidden';};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}}
  let input=null;for(const d of docs){input=sels.map(s=>({s,el:d.querySelector(s)})).find(x=>visible(x.el));if(input)break;}
  const buttons = docs.flatMap(d=>Array.from(d.querySelectorAll('button,a,[role="button"],[tabindex]'))).filter(visible);
  const sendButton=buttons.find(el=>/送信|send/i.test((el.getAttribute('aria-label')||el.title||el.textContent||'').trim())&&!el.disabled&&el.getAttribute('aria-disabled')!=='true');
  let attachElement=null;for(const d of docs){for(const s of fileSels){const e=d.querySelector(s);if(e){attachElement=e;break;}}if(attachElement)break;}
  const attachButton=buttons.find(el=>/添付|attach|ファイルを追加|add file/i.test((el.getAttribute('aria-label')||el.title||el.textContent||'').trim()));
  const signIn=buttons.find(el=>/sign\s*in|log\s*in|サインイン|ログイン/i.test((el.innerText||el.textContent||el.getAttribute('aria-label')||el.title||'').trim()));
  const url=String(location.href||'');
  const title=String(document.title||'');
  const surface=(/\/conversation\//i.test(url)||/チャット|chat/i.test(title))?'chat':(/copilot/i.test(title)?'home':'unknown');
  const bodyPreview=String((document.body&&document.body.innerText)||'').replace(/\s+/g,' ').trim().slice(0,200);
  const signinRequired=/(?:login|signin|sign-in|auth)/i.test(url)||(!input&&!!signIn);
  const editor=document.querySelector('#m365-chat-editor-target-element');
  return JSON.stringify({ ready:!!input&&!!(sendButton||attachElement||attachButton)&&surface==='chat',
    inputVisible:!!input, sendExists:!!sendButton, attachExists:!!attachElement, attachButtonExists:!!attachButton,
    surface, selector:input?input.s:'', signin_required:signinRequired, url, title, bodyPreview, editorExists:!!editor });
})()
'@
    $js = $template.Replace('__INPUT_SELS__', (Get-MbChatInputSelectorsJson -Settings $Settings)).Replace('__FILE_SELS__', $fileSelectorsJson)
    try {
        $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 15
        return ($raw | ConvertFrom-Json)
    } catch {
        return [pscustomobject]@{
            ready = $false; selector = ''; signin_required = $false; url = ''; title = ''; surface = 'unknown'
            editorExists = $false; bodyPreview = ('画面の状態を取得できませんでした: ' + $_.Exception.Message)
        }
    }
}

function Format-MbCopilotScreenDiagnostic {
    param($State)
    if ($null -eq $State) { return 'url= title=' }
    $preview = ([string]$State.bodyPreview -replace '[\r\n]+', ' ')
    if ($preview.Length -gt 200) { $preview = $preview.Substring(0, 200) }
    return ('surface={0} url={1} title={2} inputVisible={3} sendExists={4} attachExists={5} editorExists={6} bodyPreview={7}' -f
        [string]$State.surface, [string]$State.url, [string]$State.title, [string]$State.inputVisible,
        [string]$State.sendExists, [string]$State.attachExists, [string]$State.editorExists, $preview)
}

function Wait-MbCopilotScreenReady {
    param(
        [Parameter(Mandatory = $true)][string]$WsUrl,
        [Parameter(Mandatory = $true)]$Settings,
        [int]$TimeoutSeconds = 60,
        [scriptblock]$ShouldCancel = $null
    )

    $deadline = (Get-Date).AddSeconds([Math]::Max(10, $TimeoutSeconds))
    $last = $null
    $homeTransitionTried = $false
    while ((Get-Date) -lt $deadline) {
        if ($ShouldCancel -and (& $ShouldCancel)) {
            return [pscustomobject]@{ ok = $false; cancelled = $true; signinRequired = $false; message = '中止しました。' }
        }
        $last = Get-MbCopilotScreenState -WsUrl $WsUrl -Settings $Settings
        if ($last.ready -eq $true) {
            return [pscustomobject]@{ ok = $true; cancelled = $false; signinRequired = $false; message = '' }
        }
        if ($last.signin_required -eq $true) {
            Write-MbCopilotLog ('Copilotの準備ができません（サインインが必要）: ' + (Format-MbCopilotScreenDiagnostic -State $last)) 'ERROR'
            return [pscustomobject]@{
                ok = $false; cancelled = $false; signinRequired = $true
                message = 'Microsoft 365 Copilotへのサインインが必要です。［Copilotの画面を開く］からサインインして、もう一度実行してください。'
            }
        }
        # ホーム画面で止まっている場合は自分でチャットへ入る。
        if ([string]$last.surface -eq 'home' -and -not $homeTransitionTried) {
            $homeTransitionTried = $true
            try { $null = Invoke-MbFreshChat -WsUrl $WsUrl -Settings $Settings } catch { }
        }
        Start-Sleep -Milliseconds 500
    }
    if ($null -eq $last) { $last = Get-MbCopilotScreenState -WsUrl $WsUrl -Settings $Settings }
    Write-MbCopilotLog ('Copilotの準備ができません: ' + (Format-MbCopilotScreenDiagnostic -State $last)) 'ERROR'
    return [pscustomobject]@{
        ok = $false; cancelled = $false; signinRequired = $false
        message = 'Copilotの画面が開きませんでした。［Copilotの画面を開く］で様子を確認して、もう一度実行してください。'
    }
}

function Get-MbCopilotMainText {
    param([Parameter(Mandatory = $true)][string]$WsUrl)
    $js = "(() => ((document.querySelector('main') || document.body).innerText || ''))()"
    $text = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 20
    if ($null -eq $text) { return '' }
    return [string]$text
}

# ---------------------------------------------------------------------
# 新しいチャット
# ---------------------------------------------------------------------
function Invoke-MbFreshChat {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)

    $js = @'
(() => {
  const visible=e=>{if(!e)return false;const r=e.getBoundingClientRect(),s=e.ownerDocument.defaultView.getComputedStyle(e);return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden';};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}}
  const buttons=docs.flatMap(d=>Array.from(d.querySelectorAll('button, [role="button"], a, [tabindex]')));
  const candidates=[];
  for (const b of buttons) {
    const label=(b.getAttribute('aria-label')||b.title||b.textContent||'').trim();
    if(!label)continue;
    let score=0;
    if(/^(新しいチャット|New chat)$/i.test(label))score+=1000;
    else if(/新しいチャット|New chat/i.test(label))score+=400;
    else if(/チャット|chat/i.test(label))score+=80;
    if(/その他|履歴|検索|ライブラリ|more|history|search|library/i.test(label))score-=300;
    if(score<=0)continue;
    if(b.disabled||b.getAttribute('aria-disabled')==='true')continue;
    if(!visible(b))continue;
    candidates.push({el:b,label,score});
  }
  candidates.sort((a,b)=>b.score-a.score);
  const best=candidates[0];
  if(best){best.el.click();return JSON.stringify({clicked:true,label:best.label.slice(0,80)});}
  return JSON.stringify({clicked:false});
})()
'@
    $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 20
    $result = $raw | ConvertFrom-Json
    if (-not $result.clicked) {
        Write-MbCopilotLog '新しいチャットのボタンが見つからないため、URLを開き直します。' 'WARN'
        $null = Invoke-MbCdpMethod -WebSocketUrl $WsUrl -Method 'Page.navigate' -Params @{ url = [string]$Settings.copilot_url } -TimeoutSeconds 20
    }
    Start-Sleep -Milliseconds 800
    return $result
}

# ---------------------------------------------------------------------
# モデルの選択
# ---------------------------------------------------------------------
function Set-MbCopilotModel {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)

    $rawSetting = ''
    try { $rawSetting = [string]$Settings.copilot_model } catch { }
    $priority = @($rawSetting -split '[,、\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($priority.Count -eq 0) { return $null }
    $switcher = [string](Get-MbCopilotSelector -Settings $Settings -Name 'model_switcher')
    if ([string]::IsNullOrWhiteSpace($switcher)) { $switcher = '#gptModeSwitcher' }
    $candidatesJson = ConvertTo-Json -InputObject @($priority) -Compress
    if ($priority.Count -eq 1) { $candidatesJson = '[' + (ConvertTo-MbJsString $priority[0]) + ']' }

    # 上から順に試し、どれも選べなければ今のモデルのまま続ける。
    # モデルが変わらなくても手順の下書きは作れるので、ここで止めない。
    $template = @'
(async () => {
  const candidates=__CANDIDATES__;
  const switcherSelector=__SWITCHER__;
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}}
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  const norm=s=>(s||'').replace(/\s+/g,' ').trim();
  const stripTail=s=>norm(s).replace(/[…‥]|\.{3}$/g,'');
  const eq=(a,b)=>a.toLowerCase()===b.toLowerCase();
  const has=(a,b)=>a.toLowerCase().indexOf(b.toLowerCase())!==-1;
  const matchesModel=(shown,cand,picked)=>{const a=stripTail(shown);if(!a)return false;if(eq(a,cand)||has(a,cand))return true;if(picked&&(eq(a,picked)||has(a,picked)))return true;return a.length>=6&&(has(cand,a)||(picked&&has(picked,a)));};
  const visible=el=>{if(!el)return false;const r=el.getBoundingClientRect(),s=getComputedStyle(el);return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden';};
  const primaryLabel=el=>{const p=el.querySelector('.fai-CapabilityPickerMenuItem__primaryContentWrapper');if(p)return norm(p.innerText);const c=el.querySelector('.fui-MenuItem__content > span:first-child');if(c)return norm(c.innerText);return norm((el.innerText||'').split('\n')[0]);};
  const itemSelector='[role="menuitem"],[role="menuitemradio"],[role="menuitemcheckbox"],[role="option"]';
  const menuRoot=()=>{for(const d of docs){const r=d.querySelector('.fui-MenuPopover')||d.querySelector('[data-portal-node] [role="menu"]');if(r)return r;}return null;};
  const collectItems=()=>{const r=menuRoot();return r?Array.from(r.querySelectorAll(itemSelector)).filter(visible):[];};
  const collectItemsAll=()=>{const roots=docs.flatMap(d=>Array.from(d.querySelectorAll('.fui-MenuPopover, [data-portal-node] [role="menu"]'))).filter(visible);return Array.from(new Set(roots.flatMap(r=>Array.from(r.querySelectorAll(itemSelector)).filter(visible))));};
  const pressEscape=()=>{try{const o={key:'Escape',code:'Escape',keyCode:27,which:27,bubbles:true,cancelable:true};const t=document.activeElement||document.body;t.dispatchEvent(new KeyboardEvent('keydown',o));t.dispatchEvent(new KeyboardEvent('keyup',o));}catch(e){}};
  const fireMenuClick=async el=>{try{const r=el.getBoundingClientRect(),cx=r.x+r.width/2,cy=r.y+r.height/2,base={bubbles:true,cancelable:true,view:window,clientX:cx,clientY:cy};el.dispatchEvent(new PointerEvent('pointerover',{...base,pointerType:'mouse'}));el.dispatchEvent(new MouseEvent('mouseover',base));try{el.focus();}catch(e){}await sleep(60);el.dispatchEvent(new PointerEvent('pointerdown',{...base,pointerType:'mouse',button:0}));el.dispatchEvent(new MouseEvent('mousedown',{...base,button:0}));el.dispatchEvent(new PointerEvent('pointerup',{...base,pointerType:'mouse',button:0}));el.dispatchEvent(new MouseEvent('mouseup',{...base,button:0}));el.dispatchEvent(new MouseEvent('click',{...base,button:0}));try{el.click();}catch(e){}return true;}catch(e){return false;}};
  const findSwitcher=()=>{for(const d of docs){let b=d.querySelector(switcherSelector);if(b&&visible(b))return b;b=Array.from(d.querySelectorAll('button[aria-haspopup="menu"]')).find(x=>visible(x)&&/モデル|model/i.test(x.getAttribute('aria-label')||''));if(b)return b;}return null;};
  const btn=findSwitcher();
  if(!btn)return JSON.stringify({changed:false,reason:'switcher_not_found'});
  const current=norm(btn.innerText);
  if(candidates.length&&matchesModel(current,candidates[0],''))return JSON.stringify({changed:false,reason:'already_selected',current});
  await fireMenuClick(btn);
  let items=[];for(let i=0;i<30;i++){items=collectItems();if(items.length)break;await sleep(100);}
  if(!items.length){pressEscape();return JSON.stringify({changed:false,reason:'menu_not_found',current});}
  let labeled=items.map(el=>({el,label:primaryLabel(el),submenu:el.getAttribute('aria-haspopup')==='menu',checked:el.getAttribute('aria-checked')==='true'})).filter(x=>x.label);
  for(let pi=0;pi<candidates.length;pi++){
    const cand=candidates[pi];
    const hit=labeled.find(x=>eq(x.label,cand))||labeled.find(x=>has(x.label,cand));
    if(!hit)continue;
    if(hit.checked){pressEscape();return JSON.stringify({changed:false,reason:'already_selected',current,picked:hit.label});}
    const before=new Set(collectItemsAll());
    await fireMenuClick(hit.el);
    let picked=hit.label;
    if(hit.submenu){
      let fresh=[];for(let i=0;i<20;i++){fresh=collectItemsAll().filter(x=>!before.has(x));if(fresh.length)break;await sleep(100);}
      if(fresh.length){const sub=fresh.map(el=>({el,label:primaryLabel(el)})).filter(x=>x.label);const suffix=cand.replace(/^GPT[\s-]*[\d.]*\s*/i,'');const h=sub.find(x=>eq(x.label,cand))||sub.find(x=>has(x.label,cand))||sub.find(x=>eq(x.label,suffix))||sub.find(x=>suffix&&has(x.label,suffix));if(h){picked=h.label;await fireMenuClick(h.el);}}
    }
    const timeout=hit.submenu?5000:2000,t0=Date.now();
    while(Date.now()-t0<timeout){await sleep(100);const after=norm((findSwitcher()||{innerText:''}).innerText);if(matchesModel(after,cand,picked))return JSON.stringify({changed:true,reason:'selected',before:current,after,picked});}
    pressEscape();await sleep(400);
    const late=norm((findSwitcher()||{innerText:''}).innerText);
    if(matchesModel(late,cand,picked))return JSON.stringify({changed:true,reason:'selected_late',before:current,after:late,picked});
    const again=findSwitcher();if(!again)break;
    await fireMenuClick(again);await sleep(300);
    items=collectItems();if(!items.length)break;
    labeled=items.map(el=>({el,label:primaryLabel(el),submenu:el.getAttribute('aria-haspopup')==='menu',checked:el.getAttribute('aria-checked')==='true'})).filter(x=>x.label);
  }
  pressEscape();
  return JSON.stringify({changed:false,reason:'model_not_in_menu',current,tried:candidates});
})()
'@
    try {
        $js = $template.Replace('__CANDIDATES__', $candidatesJson).Replace('__SWITCHER__', (ConvertTo-MbJsString $switcher))
        $raw = [string](Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 25)
        Write-MbCopilotLog ('モデル選択: ' + $raw) 'INFO'
        return ($raw | ConvertFrom-Json)
    } catch {
        Write-MbCopilotLog ('モデルを切り替えられませんでした（そのまま続けます）: ' + $_.Exception.Message) 'WARN'
        return $null
    }
}

# ---------------------------------------------------------------------
# 添付
# ---------------------------------------------------------------------
function Get-MbAttachmentSnapshot {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)

    $itemSelectors = @([string[]](Get-MbCopilotSelector -Settings $Settings -Name 'attachment_item_any'))
    $nameSelectors = @([string[]](Get-MbCopilotSelector -Settings $Settings -Name 'attachment_name_any'))
    $template = @'
(() => {
  const itemSels=__ITEM_SELS__, nameSels=__NAME_SELS__;
  const visible=e=>{if(!e)return false;const r=e.getBoundingClientRect();return r.width>0&&r.height>0;};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}}
  let nodes=[],used='';
  for(const s of itemSels){const found=docs.flatMap(d=>Array.from(d.querySelectorAll(s))).filter(visible);if(found.length){nodes=found;used=s;break;}}
  const items=nodes.map(n=>{
    let name='';
    for(const s of nameSels){const e=n.querySelector(s);if(e&&(e.innerText||'').trim()){name=(e.innerText||'').trim();break;}}
    if(!name)name=(n.getAttribute('aria-label')||n.title||(n.innerText||'').trim().split('\n')[0]||'').trim();
    const live=(n.getAttribute('aria-label')||'')+' '+((n.querySelector('[aria-live]')||{innerText:''}).innerText||'');
    const busy=n.getAttribute('aria-busy')==='true'||!!n.querySelector('[role="progressbar"]');
    return {name,live:live.trim(),busy};
  });
  return JSON.stringify({count:items.length,items,usedItemSelector:used});
})()
'@
    $js = $template.Replace('__ITEM_SELS__', (ConvertTo-Json @($itemSelectors) -Compress)).Replace('__NAME_SELS__', (ConvertTo-Json @($nameSelectors) -Compress))
    try {
        $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 15
        return ($raw | ConvertFrom-Json)
    } catch {
        return [pscustomobject]@{ count = 0; items = @(); usedItemSelector = '' }
    }
}

function Test-MbAttachmentNameMatch {
    param([string]$Actual, [string]$Expected)
    if ([string]::IsNullOrWhiteSpace($Actual) -or [string]::IsNullOrWhiteSpace($Expected)) { return $false }
    $a = $Actual.Trim()
    $e = $Expected.Trim()
    if ($a -eq $e) { return $true }
    # 画面側で末尾が省略されることがあるため、拡張子を外した先頭一致も許す。
    $bare = [IO.Path]::GetFileNameWithoutExtension($e)
    if ($bare.Length -ge 8 -and $a -like ($bare + '*')) { return $true }
    return $false
}

function Invoke-MbCopilotAttachFiles {
    param(
        [Parameter(Mandatory = $true)][string]$WsUrl,
        [Parameter(Mandatory = $true)]$Settings,
        [Parameter(Mandatory = $true)][string[]]$Files,
        [scriptblock]$ShouldCancel = $null
    )

    foreach ($file in $Files) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "添付する画像が見つかりません: $file" }
    }
    $expected = @($Files | ForEach-Object { [IO.Path]::GetFileName($_) })
    $selector = [string](Get-MbCopilotSelector -Settings $Settings -Name 'file_input')
    $fallback = [string](Get-MbCopilotSelector -Settings $Settings -Name 'file_input_fallback')

    # nodeIdは接続ごとの値なので、探索と設定を同じ接続で行う（約束事3）。
    $socket = $null
    try {
        $socket = Connect-MbWebSocket -WebSocketUrl $WsUrl
        $null = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'DOM.enable'
        $document = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'DOM.getDocument' -Params @{ depth = 1 }
        if ($document.error) { throw ('画面の構造を取得できませんでした: ' + ($document.error | ConvertTo-Json -Compress)) }
        $rootId = [int]$document.result.root.nodeId

        $nodeId = 0
        foreach ($candidate in @($selector, $fallback)) {
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            $found = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'DOM.querySelector' -Params @{ nodeId = $rootId; selector = $candidate }
            if (-not $found.error -and $found.result -and $found.result.nodeId) {
                $nodeId = [int]$found.result.nodeId
                if ($nodeId -gt 0) { break }
            }
        }
        if ($nodeId -le 0) {
            # 同一オリジンのiframeへ移動している場合の受け皿。
            $selectorsJson = ConvertTo-Json -InputObject @($selector, $fallback) -Compress
            $expression = "(() => { const sels=$selectorsJson,docs=[document]; for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument)}catch(e){}} for(const d of docs)for(const s of sels){const e=d.querySelector(s);if(e)return e;} return null; })()"
            $evaluated = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'Runtime.evaluate' -Params @{ expression = $expression; returnByValue = $false; userGesture = $true } -TimeoutSeconds 15
            $objectId = ''
            if (-not $evaluated.error -and $evaluated.result -and $evaluated.result.result) { $objectId = [string]$evaluated.result.result.objectId }
            if (-not [string]::IsNullOrWhiteSpace($objectId)) {
                $requested = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'DOM.requestNode' -Params @{ objectId = $objectId } -TimeoutSeconds 15
                if (-not $requested.error) { $nodeId = [int]$requested.result.nodeId }
            }
        }
        if ($nodeId -le 0) {
            $state = Get-MbCopilotScreenState -WsUrl $WsUrl -Settings $Settings
            Write-MbCopilotLog ('添付欄が見つかりません ' + (Format-MbCopilotScreenDiagnostic -State $state)) 'ERROR'
            if ($state.signin_required -eq $true) {
                throw 'Microsoft 365 Copilotへのサインインが必要です。［Copilotの画面を開く］からサインインして、もう一度実行してください。'
            }
            throw 'Copilotの添付欄が見つかりませんでした。［Copilotの画面を開く］で様子を確認して、もう一度実行してください。'
        }

        $set = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'DOM.setFileInputFiles' -Params @{ nodeId = $nodeId; files = @($Files) }
        if ($set.error) { throw ('画像を添付できませんでした: ' + ($set.error | ConvertTo-Json -Compress)) }
    } finally {
        if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
    }

    # 添付が終わるまで待つ。終わらないまま送ると画像なしで回答されてしまう。
    $donePattern = [regex]::new([string](Get-MbCopilotSelector -Settings $Settings -Name 'upload_done_pattern'), 'IgnoreCase')
    $failPattern = [regex]::new([string](Get-MbCopilotSelector -Settings $Settings -Name 'upload_fail_pattern'), 'IgnoreCase')
    $waitSeconds = [Math]::Max(15, [int]$Settings.attach_wait_seconds)
    $deadline = (Get-Date).AddSeconds($waitSeconds)
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $stable = @{}

    while ((Get-Date) -lt $deadline) {
        if ($ShouldCancel -and (& $ShouldCancel)) {
            $null = Invoke-MbClickStop -WsUrl $WsUrl
            return [pscustomobject]@{ ok = $false; cancelled = $true; elapsedMs = [int]$watch.ElapsedMilliseconds }
        }
        Start-Sleep -Milliseconds 500
        $snapshot = Get-MbAttachmentSnapshot -WsUrl $WsUrl -Settings $Settings
        $mine = @($snapshot.items | Where-Object {
            $actual = [string]$_.name
            @($expected | Where-Object { Test-MbAttachmentNameMatch -Actual $actual -Expected $_ }).Count -gt 0
        })
        $failed = @($mine | Where-Object { $_.live -and $failPattern.IsMatch([string]$_.live) })
        if ($failed.Count -gt 0) {
            throw ('画像の添付に失敗しました: ' + (($failed | ForEach-Object { [string]$_.name }) -join ', '))
        }

        $done = @()
        foreach ($name in $expected) {
            $match = @($mine | Where-Object { Test-MbAttachmentNameMatch -Actual ([string]$_.name) -Expected $name } | Select-Object -First 1)
            if ($match.Count -eq 0) { continue }
            $item = $match[0]
            if ($item.live -and $donePattern.IsMatch([string]$item.live)) {
                $done += $name
                $stable[$name] = 0
            } elseif (-not $item.busy) {
                # 完了の文言が出ない画面もあるため、処理中でない状態が続いたら完了とみなす。
                $stable[$name] = 1 + [int]$stable[$name]
                if ([int]$stable[$name] -ge 2) { $done += $name }
            } else {
                $stable[$name] = 0
            }
        }

        $allDone = $true
        foreach ($name in $expected) { if ($done -notcontains $name) { $allDone = $false } }
        if ($mine.Count -ge $expected.Count -and $allDone) {
            Write-MbCopilotLog ("画像を添付しました count=$($expected.Count) elapsedMs=$([int]$watch.ElapsedMilliseconds)") 'INFO'
            return [pscustomobject]@{ ok = $true; cancelled = $false; elapsedMs = [int]$watch.ElapsedMilliseconds }
        }
    }
    throw ("画像の添付が {0} 秒以内に終わりませんでした。" -f $waitSeconds)
}

# ---------------------------------------------------------------------
# 依頼文の入力と送信
# ---------------------------------------------------------------------
function Invoke-MbFocusChatInput {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)
    $template = @'
(() => {
  const sels=__INPUT_SELS__;
  const visible=e=>{if(!e)return false;const r=e.getBoundingClientRect(),cs=e.ownerDocument.defaultView.getComputedStyle(e);return r.width>0&&r.height>0&&cs.display!=='none'&&cs.visibility!=='hidden';};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument)}catch(e){}}
  for(const d of docs)for(const s of sels){const el=d.querySelector(s);if(visible(el)){el.focus();return JSON.stringify({ok:true,sel:s});}}
  return JSON.stringify({ok:false});
})()
'@
    $js = $template.Replace('__INPUT_SELS__', (Get-MbChatInputSelectorsJson -Settings $Settings))
    $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 15
    return ($raw | ConvertFrom-Json)
}

function Get-MbChatInputTextLength {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)
    $template = @'
(() => {
  const sels=__INPUT_SELS__;
  const visible=e=>{if(!e)return false;const r=e.getBoundingClientRect(),cs=e.ownerDocument.defaultView.getComputedStyle(e);return r.width>0&&r.height>0&&cs.display!=='none'&&cs.visibility!=='hidden';};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument)}catch(e){}}
  for(const d of docs)for(const s of sels){const el=d.querySelector(s);if(visible(el))return JSON.stringify({len:(el.innerText||el.value||'').length});}
  return JSON.stringify({len:-1});
})()
'@
    $js = $template.Replace('__INPUT_SELS__', (Get-MbChatInputSelectorsJson -Settings $Settings))
    $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 15
    return ([int](($raw | ConvertFrom-Json).len))
}

function Invoke-MbKeyEvent {
    param(
        [Parameter(Mandatory = $true)][string]$WsUrl,
        [Parameter(Mandatory = $true)][string]$Type,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][int]$KeyCode,
        [int]$Modifiers = 0
    )
    $null = Invoke-MbCdpMethod -WebSocketUrl $WsUrl -Method 'Input.dispatchKeyEvent' -Params @{
        type = $Type; key = $Key; code = $Code
        windowsVirtualKeyCode = $KeyCode; nativeVirtualKeyCode = $KeyCode
        modifiers = $Modifiers
    } -TimeoutSeconds 15
}

function Clear-MbChatInput {
    param([Parameter(Mandatory = $true)][string]$WsUrl, [Parameter(Mandatory = $true)]$Settings)
    $length = Get-MbChatInputTextLength -WsUrl $WsUrl -Settings $Settings
    if ($length -le 0) { return }
    $null = Invoke-MbFocusChatInput -WsUrl $WsUrl -Settings $Settings
    # クリップボードは使わない。利用者のコピー内容を壊さないため。
    Invoke-MbKeyEvent -WsUrl $WsUrl -Type 'rawKeyDown' -Key 'a' -Code 'KeyA' -KeyCode 65 -Modifiers 2
    Invoke-MbKeyEvent -WsUrl $WsUrl -Type 'keyUp' -Key 'a' -Code 'KeyA' -KeyCode 65 -Modifiers 2
    Start-Sleep -Milliseconds 120
    Invoke-MbKeyEvent -WsUrl $WsUrl -Type 'rawKeyDown' -Key 'Backspace' -Code 'Backspace' -KeyCode 8
    Invoke-MbKeyEvent -WsUrl $WsUrl -Type 'keyUp' -Key 'Backspace' -Code 'Backspace' -KeyCode 8
    Start-Sleep -Milliseconds 200
}

function Invoke-MbInsertPrompt {
    param(
        [Parameter(Mandatory = $true)][string]$WsUrl,
        [Parameter(Mandatory = $true)]$Settings,
        [Parameter(Mandatory = $true)][string]$Prompt
    )

    $maxChars = [int]$Settings.max_prompt_chars
    if ($Prompt.Length -gt $maxChars) {
        throw ("依頼文が上限の {0} 文字を超えました（{1} 文字）。1回に渡す手順の数を減らしてください。" -f $maxChars, $Prompt.Length)
    }
    Clear-MbChatInput -WsUrl $WsUrl -Settings $Settings
    $null = Invoke-MbFocusChatInput -WsUrl $WsUrl -Settings $Settings

    # 添付直後は画面が再描画され、入力が落ちることがある。
    # 塊ごとに文字数の増加を確かめ、増えていなければ入れ直す。
    $chunkSize = 3000
    for ($offset = 0; $offset -lt $Prompt.Length; $offset += $chunkSize) {
        $length = [Math]::Min($chunkSize, $Prompt.Length - $offset)
        $chunk = $Prompt.Substring($offset, $length)
        $expectedGrowth = [int][Math]::Floor($chunk.Length * 0.9)
        $ok = $false
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $before = Get-MbChatInputTextLength -WsUrl $WsUrl -Settings $Settings
            if ($before -lt 0) { $before = 0 }
            $null = Invoke-MbFocusChatInput -WsUrl $WsUrl -Settings $Settings
            $null = Invoke-MbCdpMethod -WebSocketUrl $WsUrl -Method 'Input.insertText' -Params @{ text = $chunk } -TimeoutSeconds 30
            Start-Sleep -Milliseconds 300
            $after = Get-MbChatInputTextLength -WsUrl $WsUrl -Settings $Settings
            if (($after - $before) -ge $expectedGrowth) { $ok = $true; break }
            Start-Sleep -Milliseconds 500
        }
        if (-not $ok) {
            throw ('依頼文をCopilotの入力欄へ入れられませんでした（{0} 文字目）。' -f $offset)
        }
    }
    Start-Sleep -Milliseconds 300
    $inputLength = Get-MbChatInputTextLength -WsUrl $WsUrl -Settings $Settings
    if ($inputLength -lt [int]($Prompt.Length * 0.9)) {
        throw ('依頼文の入力を確認できませんでした（期待 {0} 文字 / 実際 {1} 文字）。' -f $Prompt.Length, $inputLength)
    }
}

function Invoke-MbClickSend {
    param([Parameter(Mandatory = $true)][string]$WsUrl)
    $js = @'
(() => {
  const visible=e=>{if(!e)return false;const r=e.getBoundingClientRect(),s=e.ownerDocument.defaultView.getComputedStyle(e);return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden';};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument)}catch(e){}}
  const buttons=docs.flatMap(d=>Array.from(d.querySelectorAll('button, [role="button"]')));
  const exclude=/stop|cancel|停止|キャンセル|regenerate|再生成|attach|添付|microphone|voice|ボイス|音声|マイク|new chat|新しいチャット|clear|クリア|close|閉じる|search|検索|library|ライブラリ|file|ファイル|mail|メール|delete|削除/;
  const clickable=[];
  for(const b of buttons){
    const label=(b.getAttribute('aria-label')||b.title||b.textContent||'').trim();
    if(!label)continue;
    const lower=label.toLowerCase();
    let score=0;
    if(/^(送信|send)$/i.test(label))score+=1000;
    else if(/送信|send/i.test(lower))score+=400;
    if(score<=0)continue;
    if(exclude.test(lower))continue;
    if(b.disabled||b.getAttribute('aria-disabled')==='true')continue;
    if(!visible(b))continue;
    clickable.push({el:b,label:label.slice(0,80),score});
  }
  clickable.sort((a,c)=>c.score-a.score);
  if(clickable.length>0){clickable[0].el.click();return JSON.stringify({clicked:true,label:clickable[0].label});}
  return JSON.stringify({clicked:false});
})()
'@
    $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 20
    $result = $raw | ConvertFrom-Json
    if (-not $result.clicked) { throw 'Copilotの送信ボタンを押せませんでした。' }
    return $result
}

function Invoke-MbClickStop {
    param([Parameter(Mandatory = $true)][string]$WsUrl)
    $js = @'
(() => {
  const visible=e=>{if(!e)return false;const r=e.getBoundingClientRect(),s=e.ownerDocument.defaultView.getComputedStyle(e);return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden';};
  const docs=[document];for(const f of document.querySelectorAll('iframe')){try{if(f.contentDocument)docs.push(f.contentDocument)}catch(e){}}
  const buttons=docs.flatMap(d=>Array.from(d.querySelectorAll('button, [role="button"]')));
  for(const b of buttons){
    const label=(b.getAttribute('aria-label')||b.title||b.textContent||'').trim();
    if(!label)continue;
    if(!/停止|stop/i.test(label))continue;
    if(b.disabled||!visible(b))continue;
    b.click();
    return JSON.stringify({clicked:true});
  }
  return JSON.stringify({clicked:false});
})()
'@
    try {
        $raw = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 15
        return ($raw | ConvertFrom-Json)
    } catch { return $null }
}

function Test-MbCopilotGenerating {
    param([Parameter(Mandatory = $true)][string]$WsUrl)
    $js = @'
(() => {
  const visible=el=>!!(el&&(el.offsetWidth||el.offsetHeight||el.getClientRects().length));
  const buttons=[...document.querySelectorAll('button,[role="button"]')].filter(visible);
  const stop=buttons.some(el=>/^(stop|停止|応答を停止|生成を停止)$/i.test((el.innerText||el.getAttribute('aria-label')||el.title||'').trim()));
  const streaming=[...document.querySelectorAll('[aria-busy="true"],[data-state="streaming"],[data-status="streaming"],[class*="streaming" i]')].some(visible);
  return JSON.stringify({generating:stop||streaming});
})()
'@
    try {
        $result = Invoke-MbCdpEval -WebSocketUrl $WsUrl -Expression $js -TimeoutSeconds 20
        if ($null -eq $result) { return $true }
        return [bool](($result | ConvertFrom-Json).generating)
    } catch {
        # 取れないときは生成中として扱う。早すぎる打ち切りより待つほうが安全。
        return $true
    }
}

# ---------------------------------------------------------------------
# 回答からJSONを取り出す
# ---------------------------------------------------------------------

# 波括弧の対応を数えて、文字列とエスケープを考慮しながらJSONらしき塊を切り出す。
# Copilotは前後に説明文を付けることがあるため、素直なパースはできない。
function Get-MbJsonObjectCandidates {
    param([AllowNull()][string]$Text)

    $found = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrEmpty($Text)) { return $found.ToArray() }
    $length = $Text.Length
    $i = 0
    while ($i -lt $length) {
        if ($Text[$i] -ne '{') { $i++; continue }
        $depth = 0
        $inString = $false
        $escaped = $false
        $start = $i
        $j = $i
        while ($j -lt $length) {
            $c = $Text[$j]
            if ($inString) {
                if ($escaped) { $escaped = $false }
                elseif ($c -eq '\') { $escaped = $true }
                elseif ($c -eq '"') { $inString = $false }
            } else {
                if ($c -eq '"') { $inString = $true }
                elseif ($c -eq '{') { $depth++ }
                elseif ($c -eq '}') {
                    $depth--
                    if ($depth -eq 0) {
                        $found.Add($Text.Substring($start, $j - $start + 1))
                        break
                    }
                }
            }
            $j++
        }
        if ($depth -eq 0 -and $j -lt $length) { $i = $j + 1 } else { $i = $start + 1 }
    }
    return $found.ToArray()
}

function Repair-MbJsonText {
    param([AllowNull()][string]$Text)
    $source = [string]$Text
    $fixed = $source
    # 末尾のカンマは生成物によく出る。ここだけは直しても意味が変わらない。
    $fixed = [regex]::Replace($fixed, ',\s*([}\]])', '$1')
    return [pscustomobject]@{ text = $fixed; changed = ($fixed -ne $source) }
}

# 手順の下書きとして使える形かどうかを見る。
function Get-MbStepAnswerJson {
    param([AllowNull()][string]$Text)

    $candidates = @(Get-MbJsonObjectCandidates -Text $Text)
    if ($candidates.Count -eq 0) { return $null }
    # 後ろにあるものほど最新の回答なので、末尾から見る。
    [array]::Reverse($candidates)
    foreach ($candidate in $candidates) {
        foreach ($attempt in @($candidate, (Repair-MbJsonText -Text $candidate).text)) {
            $parsed = $null
            try { $parsed = $attempt | ConvertFrom-Json } catch { continue }
            if ($null -eq $parsed) { continue }
            if ($parsed.PSObject.Properties.Name -notcontains 'steps') { continue }
            if (@($parsed.steps).Count -lt 1) { continue }
            return $parsed
        }
    }
    return $null
}

# 依頼文の末尾に置く決まり文句。回答の始まりを機械的に見つけるための目印。
$script:MbPromptTailAnchor = 'この後には何も書かないでください。'

function Get-MbCopilotResponseRegion {
    param([Parameter(Mandatory = $true)][string]$WsUrl)
    $text = Get-MbCopilotMainText -WsUrl $WsUrl
    # 依頼文そのものが画面に残っているため、その末尾より後ろだけを回答として見る。
    # これをしないと依頼文の中の終了マーカーを回答と誤認する。
    $position = $text.LastIndexOf($script:MbPromptTailAnchor)
    if ($position -ge 0) { return $text.Substring($position + $script:MbPromptTailAnchor.Length).TrimStart() }
    return ''
}

function Wait-MbCopilotResponse {
    param(
        [Parameter(Mandatory = $true)][string]$WsUrl,
        [Parameter(Mandatory = $true)]$Settings,
        [string]$Marker = '',
        [int]$TimeoutSeconds = 600,
        [scriptblock]$ShouldCancel = $null,
        [scriptblock]$OnProgress = $null
    )

    if ([string]::IsNullOrWhiteSpace($Marker)) { $Marker = [string]$Settings.response_end_marker }
    $pollMs = [Math]::Max(500, [int]$Settings.poll_interval_ms)
    $deadline = (Get-Date).AddSeconds([Math]::Max(30, $TimeoutSeconds))
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $idleChecks = 0
    $lastLength = -1

    while ((Get-Date) -lt $deadline) {
        if ($ShouldCancel -and (& $ShouldCancel)) {
            $null = Invoke-MbClickStop -WsUrl $WsUrl
            return [pscustomobject]@{ ok = $false; cancelled = $true; completedBy = 'cancelled'; answer = $null; tail = '' }
        }
        Start-Sleep -Milliseconds $pollMs

        $region = ''
        try { $region = Get-MbCopilotResponseRegion -WsUrl $WsUrl } catch { $region = '' }
        if ($OnProgress) {
            try { & $OnProgress ([pscustomobject]@{ elapsedMs = [int]$watch.ElapsedMilliseconds; length = $region.Length }) } catch { }
        }

        $hasMarker = (-not [string]::IsNullOrEmpty($Marker)) -and ($region -like ('*' + $Marker + '*'))
        if ($hasMarker) {
            $answer = Get-MbStepAnswerJson -Text $region
            if ($null -ne $answer) {
                return [pscustomobject]@{ ok = $true; cancelled = $false; completedBy = 'marker'; answer = $answer; tail = '' }
            }
        }

        # マーカーが出ない画面もあるため、生成が止まってから読めるJSONがあれば採用する。
        $generating = Test-MbCopilotGenerating -WsUrl $WsUrl
        if (-not $generating -and $region.Length -gt 0 -and $region.Length -eq $lastLength) {
            $idleChecks++
            if ($idleChecks -ge 2) {
                $answer = Get-MbStepAnswerJson -Text $region
                if ($null -ne $answer) {
                    return [pscustomobject]@{ ok = $true; cancelled = $false; completedBy = 'idle'; answer = $answer; tail = '' }
                }
                # 生成が終わっているのにJSONが無い。断って説明文だけ返された場合。
                if ($idleChecks -ge 4) {
                    $tail = $region
                    if ($tail.Length -gt 400) { $tail = $tail.Substring($tail.Length - 400) }
                    return [pscustomobject]@{ ok = $false; cancelled = $false; completedBy = 'no-json'; answer = $null; tail = $tail }
                }
            }
        } else {
            $idleChecks = 0
        }
        $lastLength = $region.Length
    }

    $tail = ''
    try {
        $tail = Get-MbCopilotResponseRegion -WsUrl $WsUrl
        if ($tail.Length -gt 400) { $tail = $tail.Substring($tail.Length - 400) }
    } catch { $tail = '' }
    return [pscustomobject]@{ ok = $false; cancelled = $false; completedBy = 'timeout'; answer = $null; tail = $tail }
}

# ---------------------------------------------------------------------
# 1回のやり取り
# ---------------------------------------------------------------------
function Invoke-MbCopilotRequest {
    param(
        [Parameter(Mandatory = $true)]$Settings,
        [Parameter(Mandatory = $true)][string]$ProfileDirectory,
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string[]]$AttachPaths = @(),
        [string]$Marker = '',
        [scriptblock]$OnPhase = $null,
        [scriptblock]$ShouldCancel = $null,
        [scriptblock]$OnWaitProgress = $null
    )

    $report = {
        param([string]$Phase)
        if ($OnPhase) { try { & $OnPhase $Phase } catch { } }
    }

    & $report 'preparing'
    Start-MbCopilotEdge -Settings $Settings -ProfileDirectory $ProfileDirectory
    $page = Get-MbCopilotPage -Settings $Settings
    $wsUrl = [string]$page.webSocketDebuggerUrl

    $gate = Wait-MbCopilotScreenReady -WsUrl $wsUrl -Settings $Settings -TimeoutSeconds 60 -ShouldCancel $ShouldCancel
    if ($gate.cancelled) { return [pscustomobject]@{ ok = $false; cancelled = $true; completedBy = 'cancelled'; answer = $null; tail = '' } }
    if (-not $gate.ok) { throw ([string]$gate.message) }

    # 手順のまとまりごとに新しいチャットで始める。前の依頼の添付や文脈を引きずらない。
    $null = Invoke-MbFreshChat -WsUrl $wsUrl -Settings $Settings
    $gate = Wait-MbCopilotScreenReady -WsUrl $wsUrl -Settings $Settings -TimeoutSeconds 60 -ShouldCancel $ShouldCancel
    if ($gate.cancelled) { return [pscustomobject]@{ ok = $false; cancelled = $true; completedBy = 'cancelled'; answer = $null; tail = '' } }
    if (-not $gate.ok) { throw ([string]$gate.message) }

    $null = Set-MbCopilotModel -WsUrl $wsUrl -Settings $Settings

    if (@($AttachPaths).Count -gt 0) {
        & $report 'attaching'
        $attached = Invoke-MbCopilotAttachFiles -WsUrl $wsUrl -Settings $Settings -Files $AttachPaths -ShouldCancel $ShouldCancel
        if ($attached.cancelled) { return [pscustomobject]@{ ok = $false; cancelled = $true; completedBy = 'cancelled'; answer = $null; tail = '' } }
    }

    & $report 'sending'
    Invoke-MbInsertPrompt -WsUrl $wsUrl -Settings $Settings -Prompt $Prompt
    # 添付の後始末の最中は送信ボタンが一時的に無効になる。少しの間は押し直す。
    $sendDeadline = (Get-Date).AddSeconds(20)
    $sent = $false
    $lastError = ''
    while ((Get-Date) -lt $sendDeadline) {
        try { $null = Invoke-MbClickSend -WsUrl $wsUrl; $sent = $true; break }
        catch { $lastError = $_.Exception.Message; Start-Sleep -Milliseconds 1000 }
    }
    if (-not $sent) { throw ('Copilotへ送信できませんでした: ' + $lastError) }

    & $report 'waiting'
    return (Wait-MbCopilotResponse -WsUrl $wsUrl -Settings $Settings -Marker $Marker `
        -TimeoutSeconds ([int]$Settings.request_timeout) -ShouldCancel $ShouldCancel -OnProgress $OnWaitProgress)
}

# 利用者にCopilotの画面を見せる。サインインや様子の確認に使う。
function Show-MbCopilotWindow {
    param(
        [Parameter(Mandatory = $true)]$Settings,
        [Parameter(Mandatory = $true)][string]$ProfileDirectory
    )

    $port = [int]$Settings.cdp_port
    if (-not (Test-MbDevTools -Port $port)) {
        Start-MbCopilotEdge -Settings $Settings -ProfileDirectory $ProfileDirectory
    }
    $page = Get-MbCopilotPage -Settings $Settings
    $socket = $null
    try {
        $version = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$port/json/version" -TimeoutSec 5
        $socket = Connect-MbWebSocket -WebSocketUrl ([string]$version.webSocketDebuggerUrl)
        $window = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'Browser.getWindowForTarget' -Params @{ targetId = [string]$page.id } -TimeoutSeconds 10
        if ($window.error) { throw ($window.error | ConvertTo-Json -Compress) }
        $windowId = [int]$window.result.windowId
        # 状態と位置は別々に指定する。まとめて渡すと実装差で失敗することがある。
        $null = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'Browser.setWindowBounds' -Params @{ windowId = $windowId; bounds = @{ windowState = 'normal' } } -TimeoutSeconds 10
        $null = Invoke-MbCdpOnSocket -WebSocket $socket -Method 'Browser.setWindowBounds' -Params @{ windowId = $windowId; bounds = @{ left = 120; top = 120; width = 1280; height = 900 } } -TimeoutSeconds 10
        return $true
    } finally {
        if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
    }
}

function Get-MbCopilotPromptTailAnchor { return $script:MbPromptTailAnchor }

Export-ModuleMember -Function @(
    'Set-MbCopilotLogger',
    'Get-MbCopilotDefaultSettings',
    'Get-MbCopilotSettings',
    'Get-MbCopilotSelector',
    'Start-MbCopilotEdge',
    'Get-MbCopilotPage',
    'Get-MbCopilotScreenState',
    'Wait-MbCopilotScreenReady',
    'Invoke-MbFreshChat',
    'Set-MbCopilotModel',
    'Invoke-MbCopilotAttachFiles',
    'Invoke-MbInsertPrompt',
    'Invoke-MbClickSend',
    'Invoke-MbClickStop',
    'Wait-MbCopilotResponse',
    'Invoke-MbCopilotRequest',
    'Show-MbCopilotWindow',
    'Get-MbJsonObjectCandidates',
    'Repair-MbJsonText',
    'Get-MbStepAnswerJson',
    'Get-MbCopilotPromptTailAnchor'
)

