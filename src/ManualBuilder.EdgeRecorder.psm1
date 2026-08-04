# 記録専用Edgeを起動し、表示中ページのDOM要素をクリック前にキャッシュする。
#
# Windows UI Automationはクリック後の画面へ問い合わせるため、画面遷移やメニュー消滅に
# 間に合わないことがある。Edge側でpointermove/pointerdownを先に保持し、別プロセスから
# CDPで読み出しておけば、画面記録の60Hzループを止めずに実要素の名前と矩形を使える。

Set-StrictMode -Version 2.0

$script:MbEdgeRecorderCdpNextId = 73000
$script:MbEdgeRecorderNativeReady = $false
$script:MbEdgeRecorderPendingPointer = $null

function Initialize-MbEdgeRecorderNative {
    if ($script:MbEdgeRecorderNativeReady) { return }
    if ($null -ne ('MbEdgeRecorderNative' -as [type])) {
        $script:MbEdgeRecorderNativeReady = $true
        return
    }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class MbEdgeRecorderNative
{
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetCursorPos(out POINT point);
}
'@ -ErrorAction Stop
    $script:MbEdgeRecorderNativeReady = $true
}

function Get-MbRecorderEdgePath {
    $roots = @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA) | Where-Object { $_ }
    foreach ($root in $roots) {
        $candidate = Join-Path $root 'Microsoft\Edge\Application\msedge.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    $command = Get-Command 'msedge.exe' -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw 'Microsoft Edgeが見つかりません。Edgeをインストールしてから、もう一度実行してください。'
}

function Test-MbRecorderDevTools {
    param([Parameter(Mandatory = $true)][int]$Port)
    try {
        $null = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$Port/json/version" -TimeoutSec 2
        return $true
    } catch { return $false }
}

function Wait-MbRecorderDevTools {
    param([Parameter(Mandatory = $true)][int]$Port, [int]$TimeoutSeconds = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-MbRecorderDevTools -Port $Port) { return $true }
        Start-Sleep -Milliseconds 300
    }
    return $false
}

function Get-MbEdgeRecorderCapability {
    try {
        $path = Get-MbRecorderEdgePath
        return [pscustomobject]@{ available = $true; reason = ''; path = $path }
    } catch {
        return [pscustomobject]@{ available = $false; reason = [string]$_.Exception.Message; path = '' }
    }
}

function Start-MbRecorderEdge {
    param(
        [Parameter(Mandatory = $true)][string]$ProfileDirectory,
        [int]$Port = 9465
    )

    $edge = Get-MbRecorderEdgePath
    if (-not (Test-Path -LiteralPath $ProfileDirectory)) {
        [void](New-Item -ItemType Directory -Path $ProfileDirectory -Force)
    }

    if (Test-MbRecorderDevTools -Port $Port) {
        # 前回の専用Edgeが残っていれば、新しいタブを前面へ出して再利用する。
        Start-Process -FilePath $edge -ArgumentList @("--user-data-dir=$ProfileDirectory", 'about:blank') | Out-Null
        return
    }

    $arguments = @(
        "--remote-debugging-port=$Port",
        '--remote-debugging-address=127.0.0.1',
        '--remote-allow-origins=*',
        "--user-data-dir=$ProfileDirectory",
        '--no-first-run',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-renderer-backgrounding',
        'about:blank'
    )
    Start-Process -FilePath $edge -ArgumentList $arguments | Out-Null
    if (-not (Wait-MbRecorderDevTools -Port $Port -TimeoutSeconds 30)) {
        throw '記録用Edgeを操作できる状態にできませんでした。専用Edgeを閉じてから、もう一度実行してください。'
    }
}

function Get-MbEdgeRecorderTargets {
    param([Parameter(Mandatory = $true)][int]$Port)
    $raw = $null
    try { $raw = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$Port/json" -TimeoutSec 3 } catch { return @() }
    $targets = New-Object System.Collections.ArrayList
    foreach ($item in @($raw)) {
        if ($null -eq $item) { continue }
        if ($item -is [System.Array]) {
            foreach ($inner in $item) { if ($null -ne $inner) { [void]$targets.Add($inner) } }
        } else {
            [void]$targets.Add($item)
        }
    }
    return @($targets | Where-Object {
        [string]$_.type -eq 'page' -and
        -not [string]::IsNullOrWhiteSpace([string]$_.webSocketDebuggerUrl) -and
        ([string]$_.url -notlike 'devtools://*')
    })
}

function Receive-MbEdgeRecorderWsMessage {
    param([Parameter(Mandatory = $true)]$WebSocket, [int]$TimeoutMilliseconds = 3000)

    $buffer = New-Object byte[] 65536
    $memory = New-Object IO.MemoryStream
    $cancellation = [Threading.CancellationTokenSource]::new()
    $cancellation.CancelAfter([Math]::Max(50, $TimeoutMilliseconds))
    try {
        do {
            $segment = [ArraySegment[byte]]::new($buffer)
            $received = $WebSocket.ReceiveAsync($segment, $cancellation.Token).GetAwaiter().GetResult()
            if ($received.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) { return $null }
            if ($received.Count -gt 0) { $memory.Write($buffer, 0, $received.Count) }
        } while (-not $received.EndOfMessage)
        return [Text.Encoding]::UTF8.GetString($memory.ToArray())
    } catch {
        return $null
    } finally {
        try { $cancellation.Dispose() } catch { }
        try { $memory.Dispose() } catch { }
    }
}

function Connect-MbEdgeRecorderWebSocket {
    param([Parameter(Mandatory = $true)][string]$WebSocketUrl)
    $socket = [Net.WebSockets.ClientWebSocket]::new()
    $cancellation = [Threading.CancellationTokenSource]::new()
    $cancellation.CancelAfter([TimeSpan]::FromSeconds(10))
    try {
        $null = $socket.ConnectAsync([Uri]$WebSocketUrl, $cancellation.Token).GetAwaiter().GetResult()
    } finally {
        $cancellation.Dispose()
    }
    return $socket
}

function Invoke-MbEdgeRecorderCdpOnSocket {
    param(
        [Parameter(Mandatory = $true)]$WebSocket,
        [Parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{},
        [int]$TimeoutMilliseconds = 3000
    )

    $requestId = $script:MbEdgeRecorderCdpNextId
    $script:MbEdgeRecorderCdpNextId++
    $payload = @{ id = $requestId; method = $Method; params = $Params } | ConvertTo-Json -Depth 30 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
    $segment = [ArraySegment[byte]]::new($bytes)
    $cancellation = [Threading.CancellationTokenSource]::new()
    $cancellation.CancelAfter([TimeSpan]::FromSeconds(5))
    try {
        $null = $WebSocket.SendAsync($segment, [Net.WebSockets.WebSocketMessageType]::Text, $true, $cancellation.Token).GetAwaiter().GetResult()
    } finally {
        $cancellation.Dispose()
    }

    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $remaining = [int][Math]::Max(50, ($deadline - [DateTime]::UtcNow).TotalMilliseconds)
        $message = Receive-MbEdgeRecorderWsMessage -WebSocket $WebSocket -TimeoutMilliseconds $remaining
        if ([string]::IsNullOrWhiteSpace($message)) { continue }
        $parsed = $null
        try { $parsed = $message | ConvertFrom-Json } catch { continue }
        if ($parsed -and $parsed.PSObject.Properties.Name -contains 'method' -and
            [string]$parsed.method -eq 'Runtime.bindingCalled' -and
            $parsed.params -and [string]$parsed.params.name -eq '__manualBuilderRecorderEmit') {
            try { $script:MbEdgeRecorderPendingPointer = ([string]$parsed.params.payload | ConvertFrom-Json) } catch { }
            continue
        }
        if ($parsed -and $parsed.PSObject.Properties.Name -contains 'id' -and [int]$parsed.id -eq $requestId) {
            if ($parsed.PSObject.Properties.Name -contains 'error') {
                throw ($parsed.error | ConvertTo-Json -Depth 10 -Compress)
            }
            return $parsed
        }
    }
    throw "記録用Edgeの応答がありませんでした: $Method"
}

function Invoke-MbEdgeRecorderEvalOnSocket {
    param(
        [Parameter(Mandatory = $true)]$WebSocket,
        [Parameter(Mandatory = $true)][string]$Expression,
        [int]$TimeoutMilliseconds = 3000
    )
    $response = Invoke-MbEdgeRecorderCdpOnSocket -WebSocket $WebSocket -Method 'Runtime.evaluate' -Params @{
        expression = $Expression; returnByValue = $true; awaitPromise = $false
    } -TimeoutMilliseconds $TimeoutMilliseconds
    if ($response.result.PSObject.Properties.Name -contains 'exceptionDetails') { return $null }
    if ($null -eq $response.result.result -or
        -not ($response.result.result.PSObject.Properties.Name -contains 'value')) { return $null }
    return $response.result.result.value
}

function Get-MbEdgeRecorderVisiblePage {
    param([Parameter(Mandatory = $true)][int]$Port)
    $targets = @(Get-MbEdgeRecorderTargets -Port $Port)
    if ($targets.Count -eq 0) { return $null }
    foreach ($target in $targets) {
        $socket = $null
        try {
            $socket = Connect-MbEdgeRecorderWebSocket -WebSocketUrl ([string]$target.webSocketDebuggerUrl)
            $visibility = Invoke-MbEdgeRecorderEvalOnSocket -WebSocket $socket -Expression 'document.visibilityState' -TimeoutMilliseconds 800
            if ([string]$visibility -eq 'visible') { return $target }
        } catch { } finally {
            if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
        }
    }
    return $targets[0]
}

function Get-MbEdgeRecorderInjectionScript {
    return @'
(() => {
  if (window.__manualBuilderRecorderInstalled) return true;
  window.__manualBuilderRecorderInstalled = true;
  const semantic = 'button,a[href],input,select,textarea,summary,label,[contenteditable]:not([contenteditable="false"]),[role="button"],[role="link"],[role="menuitem"],[role="tab"],[role="checkbox"],[role="radio"],[role="option"],[role="combobox"],[role="textbox"],[onclick],[tabindex]:not([tabindex="-1"])';
  const text = (value) => String(value || '').replace(/\s+/g, ' ').trim().slice(0, 400);
  const labelledBy = (node) => text((node.getAttribute('aria-labelledby') || '').split(/\s+/).map((id) => document.getElementById(id)?.innerText || '').join(' '));
  const describe = (start, clientX, clientY, source, eventPath = []) => {
    if (!(start instanceof Element)) return null;
    const path = eventPath.filter((item) => item instanceof Element);
    if (!path.length) for (let node = start; node && node instanceof Element; node = node.parentElement) path.push(node);
    let node = path.find((item) => { try { return item.matches(semantic); } catch { return false; } });
    if (!node) node = path.find((item) => { try { return getComputedStyle(item).cursor === 'pointer'; } catch { return false; } });
    if (!node) node = start;
    const rect = node.getBoundingClientRect();
    if (!(rect.width > 1 && rect.height > 1)) return null;
    const role = text(node.getAttribute('role'));
    const tag = node.tagName.toLowerCase();
    const name = text(node.getAttribute('aria-label')) || labelledBy(node) || text(node.getAttribute('alt')) ||
      text(node.getAttribute('title')) || text(node.getAttribute('placeholder')) || text(node.innerText) || text(node.value);
    return {
      source, at: Date.now(), name, role, tag, type: text(node.getAttribute('type')), editable: Boolean(node.isContentEditable),
      pageTitle: document.title || '', pageUrl: location.href,
      clientX, clientY, dpr: Number(window.devicePixelRatio) || 1,
      rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height }
    };
  };
  let hover = null;
  let pointer = null;
  document.addEventListener('pointermove', (event) => {
    const path = event.composedPath ? event.composedPath() : [];
    hover = describe(path.find((item) => item instanceof Element) || event.target, event.clientX, event.clientY, 'hover', path);
  }, true);
  document.addEventListener('pointerdown', (event) => {
    const path = event.composedPath ? event.composedPath() : [];
    pointer = describe(path.find((item) => item instanceof Element) || event.target, event.clientX, event.clientY, 'pointerdown', path);
    try { if (pointer && window.__manualBuilderRecorderEmit) window.__manualBuilderRecorderEmit(JSON.stringify(pointer)); } catch { }
  }, true);
  window.__manualBuilderRecorderRead = () => {
    let active = null;
    const element = document.activeElement;
    if (element && element !== document.body && element !== document.documentElement) {
      const rect = element.getBoundingClientRect();
      active = describe(element, rect.left + rect.width / 2, rect.top + rect.height / 2, 'active');
    }
    return { pointer, hover, active, title: document.title || '', url: location.href };
  };
  return true;
})()
'@
}

function Initialize-MbEdgeRecorderPage {
    param([Parameter(Mandatory = $true)]$WebSocket)
    $source = Get-MbEdgeRecorderInjectionScript
    $null = Invoke-MbEdgeRecorderCdpOnSocket -WebSocket $WebSocket -Method 'Runtime.enable' -TimeoutMilliseconds 3000
    $null = Invoke-MbEdgeRecorderCdpOnSocket -WebSocket $WebSocket -Method 'Page.enable' -TimeoutMilliseconds 3000
    try {
        $null = Invoke-MbEdgeRecorderCdpOnSocket -WebSocket $WebSocket -Method 'Runtime.addBinding' `
            -Params @{ name = '__manualBuilderRecorderEmit' } -TimeoutMilliseconds 3000
    } catch {
        # 同じページへ再接続するとbindingは残っている。重複エラーでもポーリング経路は使える。
    }
    $null = Invoke-MbEdgeRecorderCdpOnSocket -WebSocket $WebSocket -Method 'Page.addScriptToEvaluateOnNewDocument' `
        -Params @{ source = $source } -TimeoutMilliseconds 3000
    $null = Invoke-MbEdgeRecorderEvalOnSocket -WebSocket $WebSocket -Expression $source -TimeoutMilliseconds 3000
}

function Write-MbEdgeRecorderCache {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)]$Value)
    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $backup = $Path + '.' + [guid]::NewGuid().ToString('N') + '.bak'
    try {
        [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 20 -Compress), (New-Object Text.UTF8Encoding($false)))
        $delaysMs = @(0, 25, 50, 100, 200, 400)
        for ($attempt = 0; $attempt -lt $delaysMs.Count; $attempt++) {
            if ([int]$delaysMs[$attempt] -gt 0) { Start-Sleep -Milliseconds ([int]$delaysMs[$attempt]) }
            try {
                if ([IO.File]::Exists($Path)) {
                    [IO.File]::Replace($temporary, $Path, $backup, $true)
                } else {
                    [IO.File]::Move($temporary, $Path)
                }
                return
            } catch [IO.IOException] {
                if ($attempt -eq ($delaysMs.Count - 1)) { throw }
            } catch [UnauthorizedAccessException] {
                if ($attempt -eq ($delaysMs.Count - 1)) { throw }
            }
        }
    } finally {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue
    }
}

function Stop-MbRecorderEdge {
    param([int]$Port = 9465)
    $version = $null
    try { $version = Invoke-RestMethod -UseBasicParsing -Uri "http://127.0.0.1:$Port/json/version" -TimeoutSec 2 } catch { return }
    $url = [string]$version.webSocketDebuggerUrl
    if ([string]::IsNullOrWhiteSpace($url)) { return }
    $socket = $null
    try {
        $socket = Connect-MbEdgeRecorderWebSocket -WebSocketUrl $url
        $requestId = $script:MbEdgeRecorderCdpNextId
        $script:MbEdgeRecorderCdpNextId++
        $bytes = [Text.Encoding]::UTF8.GetBytes((@{ id = $requestId; method = 'Browser.close'; params = @{} } | ConvertTo-Json -Compress))
        $segment = [ArraySegment[byte]]::new($bytes)
        $null = $socket.SendAsync($segment, [Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
    } catch { } finally {
        if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
    }
}

function Invoke-MbEdgeRecorderCacheLoop {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$StopPath,
        [AllowEmptyString()][string]$LogPath = '',
        [int]$Port = 9465,
        [int]$PollIntervalMs = 90
    )

    Initialize-MbEdgeRecorderNative
    $socket = $null
    $targetId = ''
    $nextTargetRefresh = [DateTime]::MinValue
    try {
        while (-not (Test-Path -LiteralPath $StopPath -PathType Leaf)) {
            try {
                if ($null -eq $socket -or [DateTime]::UtcNow -ge $nextTargetRefresh) {
                    $page = Get-MbEdgeRecorderVisiblePage -Port $Port
                    $nextTargetRefresh = [DateTime]::UtcNow.AddMilliseconds(700)
                    if ($null -eq $page) { Start-Sleep -Milliseconds 200; continue }
                    if ($null -eq $socket -or [string]$page.id -ne $targetId) {
                        if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
                        $socket = Connect-MbEdgeRecorderWebSocket -WebSocketUrl ([string]$page.webSocketDebuggerUrl)
                        $targetId = [string]$page.id
                        Initialize-MbEdgeRecorderPage -WebSocket $socket
                    }
                }

                $value = Invoke-MbEdgeRecorderEvalOnSocket -WebSocket $socket `
                    -Expression '(window.__manualBuilderRecorderRead ? window.__manualBuilderRecorderRead() : null)' `
                    -TimeoutMilliseconds 1200
                if ($null -ne $value) {
                    if ($null -ne $script:MbEdgeRecorderPendingPointer) {
                        $value.pointer = $script:MbEdgeRecorderPendingPointer
                    }
                    $point = New-Object 'MbEdgeRecorderNative+POINT'
                    [void][MbEdgeRecorderNative]::GetCursorPos([ref]$point)
                    Write-MbEdgeRecorderCache -Path $CachePath -Value ([pscustomobject]@{
                        updatedAtUtc = [DateTime]::UtcNow.ToString('o')
                        cursorX = [int]$point.X
                        cursorY = [int]$point.Y
                        targetId = $targetId
                        page = $value
                    })
                }
            } catch {
                if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
                    try {
                        $line = [DateTime]::UtcNow.ToString('o') + "`t" + $_.Exception.Message + [Environment]::NewLine
                        [IO.File]::AppendAllText($LogPath, $line, (New-Object Text.UTF8Encoding($false)))
                    } catch { }
                }
                if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
                $socket = $null
                $targetId = ''
                $nextTargetRefresh = [DateTime]::MinValue
                Start-Sleep -Milliseconds 250
                continue
            }
            Start-Sleep -Milliseconds $PollIntervalMs
        }
    } finally {
        if ($null -ne $socket) { try { $socket.Dispose() } catch { } }
        # 監視エラーでワーカーが落ちてもEdgeを巻き添えで閉じない。
        # 明示的な停止要求が届いた場合だけ、専用Edgeを閉じる。
        if (Test-Path -LiteralPath $StopPath -PathType Leaf) { Stop-MbRecorderEdge -Port $Port }
    }
}

Export-ModuleMember -Function @(
    'Get-MbEdgeRecorderCapability',
    'Start-MbRecorderEdge',
    'Stop-MbRecorderEdge',
    'Invoke-MbEdgeRecorderCacheLoop'
)
