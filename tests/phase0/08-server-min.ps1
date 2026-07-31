# =====================================================================
# 08-server-min.ps1  —  サーバの最小検証（V-1 / V-2 / V-7 / V-10）
#
#   V-1   HttpListener が localhost で非管理者起動できるか
#   V-2   生バイトPOSTで画像を受け取れるか（multipart を使わない）
#   V-7   FileSystemWatcher が Win+Shift+S を確実に拾うか
#   V-10  ハートビートによる監視の自動停止とバックフィル防止
#
# ★ このスクリプトは Word を一切起動しません。
#   （旧キットの 02b-export-worker.ps1 は PID差分で WINWORD を一括終了する
#     危険な実装だったため、このキットからは削除しました）
#
# 使い方:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\08-server-min.ps1
#   停止: Ctrl+C または画面の「終了」ボタン
# =====================================================================
[CmdletBinding()]
param(
    [int]$Port = 8765,
    [int]$HeartbeatTimeoutSec = 30,
    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------
# DPI（System.Windows.Forms を読む「前」に実行）
# ---------------------------------------------------------------------
$script:DpiResult = 'スキップ'
try {
    Add-Type -Namespace MBS -Name Dpi -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
'@ -ErrorAction Stop
    $script:DpiResult = if ([MBS.Dpi]::SetProcessDpiAwarenessContext([IntPtr](-4))) {
        '成功 (PER_MONITOR_AWARE_V2)'
    } else {
        "失敗 (Win32Error=$([Runtime.InteropServices.Marshal]::GetLastWin32Error()))"
    }
} catch { $script:DpiResult = '失敗: ' + $_.Exception.Message }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# ---------------------------------------------------------------------
# 状態
# ---------------------------------------------------------------------
$script:Root = Join-Path $PSScriptRoot 'out\server'
$script:ImgDir = Join-Path $script:Root 'images'
New-Item -ItemType Directory -Force -Path $script:ImgDir | Out-Null

$script:Steps = New-Object System.Collections.ArrayList
$script:Seq = 0
$script:Version = 0
$script:Token = [guid]::NewGuid().ToString('N')
$script:Pending = New-Object System.Collections.ArrayList
$script:SeenHashes = New-Object System.Collections.ArrayList
$script:Running = $true

# V-10: 監視のライフサイクル
$script:ImportWatermark = Get-Date
$script:LastHeartbeat = $null
$script:CaptureOwnerTab = $null
$script:CaptureOwnerLastHeartbeat = $null
$script:WatcherState = 'disabled'
$script:SuspendCount = 0
$script:SkippedByWatermark = 0

$script:Stats = [ordered]@{
    受信_貼り付け = 0; 受信_ドロップ = 0; 受信_ファイル選択 = 0
    受信_ボタン撮影 = 0; 受信_監視 = 0; 重複スキップ = 0
    基準時刻でスキップ = 0; リクエスト数 = 0; ハートビート = 0
}
function Bump { param([string]$K, [int]$By = 1) $script:Stats[$K] = [int]$script:Stats[$K] + $By }

function Write-MBLog {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss.fff'), $Level, $Msg
    $c = switch ($Level) { 'ERR' { 'Red' } 'WARN' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $c
    Add-Content -LiteralPath (Join-Path $script:Root 'server.log') -Value $line -Encoding UTF8
}
function HtmlEnc { param([string]$s) return [System.Web.HttpUtility]::HtmlEncode([string]$s) }

# ---------------------------------------------------------------------
# 保存先の解決（02-known-folders.ps1 と同じロジックの簡易版）
# ---------------------------------------------------------------------
function Resolve-ScreenshotDir {
    $guid = '{B7BEDE81-DF94-4682-A7D8-57A52620B86F}'
    foreach ($key in @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')) {
        try {
            $v = (Get-ItemProperty -Path $key -Name $guid -ErrorAction Stop).$guid
            if ($v) {
                $p = [Environment]::ExpandEnvironmentVariables($v)
                if (Test-Path -LiteralPath $p) { return $p }
            }
        } catch { }
    }
    $pics = [Environment]::GetFolderPath('MyPictures')
    foreach ($n in @('Screenshots', 'スクリーンショット')) {
        $p = Join-Path $pics $n
        if (Test-Path -LiteralPath $p) { return $p }
    }
    foreach ($base in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
        if (-not $base) { continue }
        foreach ($n in @('Pictures\Screenshots', 'ピクチャ\スクリーンショット')) {
            $p = Join-Path $base $n
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    return $null
}
$script:WatchDir = Resolve-ScreenshotDir

# ---------------------------------------------------------------------
# 画像
# ---------------------------------------------------------------------
function Test-ImageMagic {
    param([byte[]]$Bytes)
    if ($Bytes.Length -lt 8) { return $null }
    if ($Bytes[0] -eq 0x89 -and $Bytes[1] -eq 0x50) { return 'png' }
    if ($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xD8) { return 'jpg' }
    if ($Bytes[0] -eq 0x42 -and $Bytes[1] -eq 0x4D) { return 'bmp' }
    return $null
}
function Get-ImageContentType {
    param([string]$Path)
    switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.jpg'  { return 'image/jpeg' }
        '.jpeg' { return 'image/jpeg' }
        '.bmp'  { return 'image/bmp' }
        default { return 'image/png' }
    }
}
function Get-FileSha256 {
    param([string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $fs = [IO.File]::OpenRead($Path)
        try { return [BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '') }
        finally { $fs.Close() }
    } finally { $sha.Dispose() }
}
function Add-Step {
    param([string]$SourcePath, [string]$Source)
    $hash = Get-FileSha256 -Path $SourcePath
    if ($hash -and $script:SeenHashes -contains $hash) {
        Bump '重複スキップ'
        Write-MBLog "重複のためスキップ: $([IO.Path]::GetFileName($SourcePath))" 'INFO'
        return $null
    }

    $w = 0; $h = 0
    $im = $null
    try {
        $im = [System.Drawing.Image]::FromFile($SourcePath)
        $w = $im.Width; $h = $im.Height
        if ($w -le 0 -or $h -le 0 -or $w -gt 12000 -or $h -gt 12000) {
            throw "画像サイズ ${w}x${h}px は範囲外です（長辺12000pxまで）"
        }
    } catch {
        throw "画像として読み込めません: $($_.Exception.Message)"
    } finally {
        if ($im) { $im.Dispose() }
    }

    $script:Seq++
    $id = 's-{0:d3}' -f $script:Seq
    $ext = [IO.Path]::GetExtension($SourcePath)
    if (-not $ext) { $ext = '.png' }
    $dest = Join-Path $script:ImgDir ($id + $ext)
    Copy-Item -LiteralPath $SourcePath -Destination $dest -Force
    if ($hash) { [void]$script:SeenHashes.Add($hash) }
    [void]$script:Steps.Add([pscustomobject]@{
        id = $id; file = $dest; w = $w; h = $h; source = $Source; at = (Get-Date)
    })
    $script:Version++
    Write-MBLog "ステップ追加 $id ($Source) ${w}x${h}" 'OK'
    return $id
}
function Get-Thumb {
    param([string]$Path, [int]$MaxW = 240)
    $src = [System.Drawing.Image]::FromFile($Path)
    try {
        $sc = [Math]::Min([double]$MaxW / $src.Width, 1.0)
        $tw = [Math]::Max(1, [int]($src.Width * $sc)); $th = [Math]::Max(1, [int]($src.Height * $sc))
        $bmp = New-Object System.Drawing.Bitmap $tw, $th
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.DrawImage($src, 0, 0, $tw, $th)
            $ms = New-Object System.IO.MemoryStream
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            return $ms.ToArray()
        } finally { $g.Dispose(); $bmp.Dispose() }
    } finally { $src.Dispose() }
}
function Invoke-DelayedCapture {
    param([int]$Delay = 3)
    Start-Sleep -Seconds $Delay
    $pt = [System.Windows.Forms.Cursor]::Position
    $t = [System.Windows.Forms.Screen]::FromPoint($pt)
    $b = $t.Bounds
    $tmp = Join-Path $env:TEMP ('mbs-' + [guid]::NewGuid().ToString('N') + '.png')
    $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.CopyFromScreen($b.X, $b.Y, 0, 0, (New-Object System.Drawing.Size $b.Width, $b.Height))
        $bmp.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally { $g.Dispose(); $bmp.Dispose() }
    try {
        $added = Add-Step -SourcePath $tmp -Source 'button'
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
    if ($added) { Bump '受信_ボタン撮影' }
}

# ---------------------------------------------------------------------
# 監視（Created + Renamed）
# ---------------------------------------------------------------------
$script:Watcher = $null
if ($script:WatchDir) {
    try {
        $script:Watcher = New-Object System.IO.FileSystemWatcher
        $script:Watcher.Path = $script:WatchDir
        $script:Watcher.Filter = '*.*'
        $script:Watcher.IncludeSubdirectories = $false
        $script:Watcher.EnableRaisingEvents = $true
        Register-ObjectEvent -InputObject $script:Watcher -EventName Created -SourceIdentifier 'MBS.C' | Out-Null
        Register-ObjectEvent -InputObject $script:Watcher -EventName Renamed -SourceIdentifier 'MBS.R' | Out-Null
        $script:WatcherState = 'suspended'   # ハートビートが来るまでは止めておく
        Write-MBLog "監視の準備完了: $script:WatchDir" 'OK'
    } catch {
        Write-MBLog "監視の開始に失敗: $($_.Exception.Message)" 'ERR'
    }
} else {
    Write-MBLog '保存先が検出できませんでした（貼り付け経路のみ利用可）' 'WARN'
}

function Test-Heartbeat {
    if (-not $script:Watcher) { return }
    $alive = ($script:CaptureOwnerLastHeartbeat -and
              ((Get-Date) - $script:CaptureOwnerLastHeartbeat).TotalSeconds -le $HeartbeatTimeoutSec)
    if ($alive -and $script:CaptureOwnerTab) {
        if ($script:WatcherState -ne 'active') {
            $script:WatcherState = 'active'
            $script:ImportWatermark = Get-Date      # ★ バックフィル防止
            Write-MBLog "監視を再開しました（取り込み基準時刻を更新: $($script:ImportWatermark.ToString('HH:mm:ss'))）" 'OK'
        }
    } else {
        if ($script:WatcherState -eq 'active') {
            $script:WatcherState = 'suspended'
            $script:SuspendCount++
            Write-MBLog 'ハートビートが途絶えたため監視を一時停止しました' 'WARN'
        }
    }
}

function Test-WatchAcceptable {
    param([string]$Path)
    $name = [IO.Path]::GetFileName($Path)
    if ($name -like '~$*') { return $false }
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -notin @('.png', '.jpg', '.jpeg', '.bmp')) { return $false }
    return $true
}

function Invoke-WatcherFlush {
    foreach ($sid in @('MBS.C', 'MBS.R')) {
        foreach ($e in @(Get-Event -SourceIdentifier $sid -ErrorAction SilentlyContinue)) {
            $p = $null
            try { $p = $e.SourceEventArgs.FullPath } catch { }
            Remove-Event -EventIdentifier $e.EventIdentifier -ErrorAction SilentlyContinue
            if (-not $p) { continue }
            if ($script:WatcherState -ne 'active') {
                Write-MBLog "監視停止中のため無視: $([IO.Path]::GetFileName($p))" 'INFO'
                continue
            }
            if (-not (Test-WatchAcceptable $p)) { continue }
            [void]$script:Pending.Add([pscustomobject]@{ Path = $p; Size = -1; Tries = 0 })
            Write-MBLog "検出: $([IO.Path]::GetFileName($p))" 'INFO'
        }
    }
    if ($script:Pending.Count -eq 0) { return }

    $done = New-Object System.Collections.ArrayList
    foreach ($item in $script:Pending) {
        $item.Tries = $item.Tries + 1
        if (-not (Test-Path -LiteralPath $item.Path)) {
            if ($item.Tries -gt 30) { [void]$done.Add($item) }
            continue
        }
        $fi = $null
        try { $fi = Get-Item -LiteralPath $item.Path } catch { }
        if (-not $fi) { if ($item.Tries -gt 30) { [void]$done.Add($item) }; continue }

        # OneDrive プレースホルダ
        if ("$($fi.Attributes)" -match 'Offline|RecallOnDataAccess') {
            if ($item.Tries -gt 60) {
                [void]$done.Add($item); Write-MBLog "実体が来ません: $($fi.Name)" 'WARN'
            }
            continue
        }
        # 取り込み基準時刻（V-10）
        if ($fi.CreationTime -lt $script:ImportWatermark) {
            [void]$done.Add($item)
            $script:SkippedByWatermark++
            Bump '基準時刻でスキップ'
            Write-MBLog "基準時刻より前のため取り込みません: $($fi.Name)" 'INFO'
            continue
        }
        # 書き込み完了待ち
        if ($fi.Length -le 0 -or $fi.Length -ne $item.Size) {
            $item.Size = $fi.Length
            if ($item.Tries -gt 30) { [void]$done.Add($item); Write-MBLog "サイズが安定しません: $($fi.Name)" 'WARN' }
            continue
        }
        try { $fs = [IO.File]::Open($item.Path, 'Open', 'Read', 'None'); $fs.Close() }
        catch { if ($item.Tries -gt 30) { [void]$done.Add($item) }; continue }

        try {
            $added = Add-Step -SourcePath $item.Path -Source 'watcher'
            if ($added) { Bump '受信_監視' }
        } catch {
            Write-MBLog "画像を取り込めません: $($fi.Name) / $($_.Exception.Message)" 'WARN'
        }
        [void]$done.Add($item)
    }
    foreach ($d in $done) { $script:Pending.Remove($d) }
}

# ---------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------
function Render-StepList {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div id=""steplist"" data-ver=""$($script:Version)"">")
    if ($script:Steps.Count -eq 0) {
        [void]$sb.Append('<p class="empty">まだステップがありません。</p>')
    }
    $n = 0
    foreach ($s in $script:Steps) {
        $n++
        $lbl = switch ($s.source) {
            'watcher' { '監視' } 'paste' { '貼り付け' } 'drop' { 'ドロップ' }
            'file' { 'ファイル' } 'button' { 'ボタン撮影' } default { $s.source }
        }
        [void]$sb.Append("<div class=""step""><div class=""no"">$n</div>")
        [void]$sb.Append("<img class=""thumb"" src=""/img/$($s.id)?thumb=1&v=$($script:Version)"" alt="""">")
        [void]$sb.Append("<div class=""meta"">$($s.id)<br>$($s.w)x$($s.h)px<br>$lbl<br>$($s.at.ToString('HH:mm:ss'))</div></div>")
    }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Render-WatchState {
    param([string]$TabId = '')
    $t = switch ($script:WatcherState) {
        'active' { "● 監視中（$(Split-Path $script:WatchDir -Leaf)）" }
        'suspended' { '⏸ 一時停止中（ハートビート待ち）' }
        default { '⚠ 保存先が検出できません' }
    }
    $c = switch ($script:WatcherState) { 'active' { '#16a34a' } 'suspended' { '#b45309' } default { '#b91c1c' } }
    $wm = $script:ImportWatermark.ToString('HH:mm:ss')
    $role = if ($TabId -and $script:CaptureOwnerTab -and $TabId -ne $script:CaptureOwnerTab) {
        '<span style="color:#b45309;font-size:12px;margin-left:10px">このタブは閲覧専用</span>'
    } elseif ($TabId -and $TabId -eq $script:CaptureOwnerTab) {
        '<span style="color:#166534;font-size:12px;margin-left:10px">このタブが撮影対象</span>'
    } else { '' }
    return "<span id=""watchstate"" style=""color:$c;font-weight:600"">$(HtmlEnc $t)</span>" +
           $role +
           "<span style=""color:#666;font-size:12px;margin-left:10px"">取り込み基準時刻 $wm / 一時停止 $($script:SuspendCount) 回</span>"
}

function Render-Stats {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table>')
    foreach ($k in $script:Stats.Keys) {
        [void]$sb.Append("<tr><td>$($k.Replace('_',': '))</td><td><b>$($script:Stats[$k])</b></td></tr>")
    }
    [void]$sb.Append("<tr><td>ステップ数</td><td><b>$($script:Steps.Count)</b></td></tr>")
    [void]$sb.Append("<tr><td>保存先</td><td>$(HtmlEnc $(if ($script:WatchDir) { $script:WatchDir } else { '未検出' }))</td></tr>")
    [void]$sb.Append("<tr><td>DPI設定</td><td>$(HtmlEnc $script:DpiResult)</td></tr>")
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

$PageTemplate = @'
<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>ManualBuilder サーバ検証</title><style>
*{box-sizing:border-box}
body{font-family:"BIZ UDPGothic","Meiryo","Yu Gothic UI",sans-serif;font-size:15px;margin:0;background:#f6f7f9;color:#222}
header{background:#1f3864;color:#fff;padding:12px 20px;display:flex;align-items:center;gap:16px}
header h1{font-size:16px;margin:0}
header .sp{flex:1}
button{font:inherit;font-size:13px;padding:7px 14px;border-radius:6px;border:1px solid #c8ccd2;background:#fff;cursor:pointer}
button.pri{background:#2563eb;color:#fff;border-color:#2563eb}
main{max-width:1000px;margin:0 auto;padding:20px}
.bar{background:#fff;border:1px solid #e3e6ea;border-radius:10px;padding:14px 16px;margin-bottom:16px}
.bar h2{font-size:13px;margin:0 0 10px;color:#555}
#drop{border:2px dashed #b9c0c9;border-radius:10px;padding:20px;text-align:center;color:#6b7280;background:#fbfcfd;margin-top:10px;font-size:13px}
#drop.hot{border-color:#2563eb;background:#eff6ff;color:#1d4ed8}
.step{display:flex;gap:12px;align-items:flex-start;background:#fff;border:1px solid #e3e6ea;border-radius:10px;padding:10px;margin-bottom:8px}
.no{width:26px;height:26px;border-radius:50%;background:#1f3864;color:#fff;display:flex;align-items:center;justify-content:center;font-size:13px;flex:0 0 auto}
.thumb{width:240px;border:1px solid #dde1e6;border-radius:6px;flex:0 0 auto}
.meta{font-size:12px;color:#666;line-height:1.6}
.empty{color:#8b909a;font-size:13px}
.stats td{padding:2px 10px 2px 0;font-size:12px}
kbd{background:#eef1f5;border:1px solid #d5d9df;border-radius:4px;padding:1px 5px;font-size:12px}
</style></head><body>
<header><h1>ManualBuilder サーバ検証 (V-1/V-2/V-7/V-10)</h1><div class="sp"></div>
<button onclick="shutdown()">検証を終了</button></header>
<main>
  <div class="bar"><h2>監視の状態</h2><div id="ws">__WATCHSTATE__</div>
    <div style="font-size:12px;color:#666;margin-top:8px">
      このタブを閉じるか放置すると、__TIMEOUT__ 秒後に監視が止まります。<br>
      停止中に撮ったスクショは、再開しても取り込まれません（バックフィル防止）。
    </div>
  </div>
  <div class="bar"><h2>ステップの追加</h2>
    <button class="pri" onclick="capture(3)">3秒後にキャプチャ</button>
    <button onclick="document.getElementById('fi').click()">画像を選択</button>
    <input type="file" id="fi" accept="image/*" multiple style="display:none">
    <span style="font-size:13px;color:#555;margin-left:10px">
      <kbd>Win</kbd>+<kbd>Shift</kbd>+<kbd>S</kbd> で自動追加 ／ <kbd>Ctrl</kbd>+<kbd>V</kbd> で貼り付け
    </span>
    <div id="drop">ここに画像ファイルをドロップ</div>
  </div>
  <div class="bar stats"><h2>受信状況</h2><div id="stats">-</div></div>
  <div id="list">__STEPLIST__</div>
</main>
<script>
const TOKEN='__TOKEN__';
const TAB=(sessionStorage.getItem('mbtab')||(()=>{const v=(crypto.randomUUID?crypto.randomUUID():String(Math.random()));sessionStorage.setItem('mbtab',v);return v})());
let curVer=-1, pendingList=null;
function hdr(x){return Object.assign({'X-Manual-Token':TOKEN,'X-Tab-Id':TAB},x||{})}

async function postImage(blob,source){
  const r=await fetch('/api/image',{method:'POST',
    headers:hdr({'Content-Type':blob.type||'image/png','X-Source':source}),body:blob});
  if(!r.ok){alert('失敗: '+r.status+' '+await r.text());return}
  applyList(await r.text()); refreshStats();
}
document.addEventListener('paste',e=>{
  const it=[...(e.clipboardData?e.clipboardData.items:[])].find(i=>i.type&&i.type.indexOf('image/')===0);
  if(it){e.preventDefault();postImage(it.getAsFile(),'paste')}
});
const dz=document.getElementById('drop');
['dragenter','dragover'].forEach(v=>dz.addEventListener(v,e=>{e.preventDefault();dz.classList.add('hot')}));
['dragleave','drop'].forEach(v=>dz.addEventListener(v,e=>{e.preventDefault();dz.classList.remove('hot')}));
dz.addEventListener('drop',async e=>{
  const fs=[...e.dataTransfer.files].filter(f=>f.type.indexOf('image/')===0);
  for(const f of fs){await postImage(f,'drop')}   // 順序を保つため直列
});
document.getElementById('fi').addEventListener('change',async e=>{
  for(const f of e.target.files){await postImage(f,'file')} e.target.value='';
});
async function capture(sec){
  const r=await fetch('/api/capture?delay='+sec,{method:'POST',headers:hdr({'Content-Length':'0'})});
  if(!r.ok){alert(await r.text());return}
  applyList(await r.text()); refreshStats();
}
function applyList(html){
  document.getElementById('list').innerHTML=html;
  const sl=document.getElementById('steplist');
  if(sl) curVer=parseInt(sl.dataset.ver,10);
}
// ハートビート（10秒間隔）
async function beat(){
  try{
    const r=await fetch('/api/heartbeat',{method:'POST',headers:hdr({'Content-Length':'0'})});
    if(r.ok) document.getElementById('ws').innerHTML=await r.text();
  }catch(e){}
}
setInterval(beat,10000); beat();
// 一覧のポーリング（入力中は保留 — 今回は入力欄が無いので即適用）
setInterval(async()=>{
  const r=await fetch('/api/poll?ver='+curVer,{headers:hdr()});
  if(r.status===204) return;
  applyList(await r.text()); refreshStats();
},1500);
async function refreshStats(){
  const r=await fetch('/api/stats',{headers:hdr()});
  document.getElementById('stats').innerHTML=await r.text();
}
async function shutdown(){
  if(!confirm('必要なテストはすべて終わりましたか？ サーバを終了します。'))return;
  await fetch('/api/shutdown',{method:'POST',headers:hdr({'Content-Length':'0'})});
  document.body.innerHTML='<main><h2>終了しました。このタブを閉じてください。</h2></main>';
}
refreshStats();
</script></body></html>
'@

function Render-Page {
    $h = $PageTemplate
    $h = $h.Replace('__TOKEN__', $script:Token)
    $h = $h.Replace('__TIMEOUT__', "$HeartbeatTimeoutSec")
    $h = $h.Replace('__WATCHSTATE__', (Render-WatchState))
    $h = $h.Replace('__STEPLIST__', (Render-StepList))
    return $h
}

# ---------------------------------------------------------------------
# レスポンス
# ---------------------------------------------------------------------
function Write-Text {
    param($Ctx, [string]$Body, [int]$Status = 200, [string]$Type = 'text/html; charset=utf-8')
    $b = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Ctx.Response.StatusCode = $Status
    $Ctx.Response.ContentType = $Type
    $Ctx.Response.Headers.Add('Cache-Control', 'no-store')
    $Ctx.Response.Headers.Add('X-Content-Type-Options', 'nosniff')
    $Ctx.Response.ContentLength64 = $b.Length
    $Ctx.Response.OutputStream.Write($b, 0, $b.Length)
}
function Write-Bytes {
    param($Ctx, [byte[]]$Bytes, [string]$Type)
    $Ctx.Response.StatusCode = 200
    $Ctx.Response.ContentType = $Type
    $Ctx.Response.Headers.Add('Cache-Control', 'no-store')
    $Ctx.Response.Headers.Add('X-Content-Type-Options', 'nosniff')
    $Ctx.Response.ContentLength64 = $Bytes.Length
    $Ctx.Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
}

# ---------------------------------------------------------------------
# ルーティング
# ---------------------------------------------------------------------
function Invoke-Route {
    param($Ctx)
    $req = $Ctx.Request
    $path = [Uri]::UnescapeDataString($req.Url.LocalPath)
    Bump 'リクエスト数'

    # Host 検証（DNSリバインディング対策）
    $expected = @("localhost:$Port", "127.0.0.1:$Port")
    if ($req.Headers['Host'] -and ($expected -notcontains $req.Headers['Host'])) {
        Write-MBLog "Host ヘッダー不正: $($req.Headers['Host'])" 'WARN'
        Write-Text $Ctx 'bad host' 400 'text/plain; charset=utf-8'; return
    }
    if ($path -ne '/') {
        if ($req.Headers['X-Manual-Token'] -ne $script:Token) {
            Write-Text $Ctx 'トークンが一致しません' 403 'text/plain; charset=utf-8'; return
        }
        # Origin 検証（状態変更のみ）
        if ($req.HttpMethod -ne 'GET') {
            $o = $req.Headers['Origin']
            if ($o -and $o -notmatch "^https?://(localhost|127\.0\.0\.1):$Port$") {
                Write-MBLog "Origin 不正: $o" 'WARN'
                Write-Text $Ctx 'bad origin' 403 'text/plain; charset=utf-8'; return
            }
        }
    }

    switch -Regex ($path) {
        '^/$' { Write-Text $Ctx (Render-Page); return }

        '^/api/image$' {
            if ($req.HttpMethod -ne 'POST') { Write-Text $Ctx 'POST only' 405 'text/plain; charset=utf-8'; return }
            $len = [long]$req.ContentLength64
            if ($len -le 0) { Write-Text $Ctx 'ボディが空です' 400 'text/plain; charset=utf-8'; return }
            if ($len -gt 20MB) { Write-Text $Ctx '20MB以下にしてください' 400 'text/plain; charset=utf-8'; return }
            $ms = New-Object System.IO.MemoryStream
            $req.InputStream.CopyTo($ms)
            $bytes = $ms.ToArray(); $ms.Dispose()
            $kind = Test-ImageMagic -Bytes $bytes
            if (-not $kind) { Write-Text $Ctx 'PNG / JPEG / BMP に対応しています' 400 'text/plain; charset=utf-8'; return }
            $tmp = Join-Path $env:TEMP ('mbs-up-' + [guid]::NewGuid().ToString('N') + '.' + $kind)
            [IO.File]::WriteAllBytes($tmp, $bytes)
            $src = $req.Headers['X-Source']
            if ($src -notin @('paste', 'drop', 'file')) { $src = 'paste' }
            try {
                $added = Add-Step -SourcePath $tmp -Source $src
            } finally {
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            }
            if ($added) {
                switch ($src) { 'paste' { Bump '受信_貼り付け' } 'drop' { Bump '受信_ドロップ' } 'file' { Bump '受信_ファイル選択' } }
            }
            Write-Text $Ctx (Render-StepList); return
        }

        '^/api/capture$' {
            if ($req.HttpMethod -ne 'POST') { Write-Text $Ctx 'POST only' 405 'text/plain; charset=utf-8'; return }
            $tab = $req.Headers['X-Tab-Id']
            if (-not $tab -or $tab -ne $script:CaptureOwnerTab) {
                Write-Text $Ctx 'このタブは閲覧専用です' 409 'text/plain; charset=utf-8'; return
            }
            $d = 3
            if ($req.QueryString['delay']) {
                $parsed = 0
                if (-not [int]::TryParse($req.QueryString['delay'], [ref]$parsed)) {
                    Write-Text $Ctx 'delay は整数で指定してください' 400 'text/plain; charset=utf-8'; return
                }
                $d = [Math]::Max(0, [Math]::Min(10, $parsed))
            }
            Invoke-DelayedCapture -Delay $d
            Write-Text $Ctx (Render-StepList); return
        }

        '^/api/heartbeat$' {
            if ($req.HttpMethod -ne 'POST') { Write-Text $Ctx 'POST only' 405 'text/plain; charset=utf-8'; return }
            Bump 'ハートビート'
            $script:LastHeartbeat = Get-Date
            $tab = $req.Headers['X-Tab-Id']
            if (-not $tab) { Write-Text $Ctx 'tab id required' 400 'text/plain; charset=utf-8'; return }
            $ownerExpired = (-not $script:CaptureOwnerLastHeartbeat -or
                ((Get-Date) - $script:CaptureOwnerLastHeartbeat).TotalSeconds -gt $HeartbeatTimeoutSec)
            if (-not $script:CaptureOwnerTab -or $ownerExpired) {
                $script:CaptureOwnerTab = $tab
                $script:CaptureOwnerLastHeartbeat = Get-Date
                Write-MBLog "撮影対象タブを設定: $tab" 'OK'
            } elseif ($script:CaptureOwnerTab -eq $tab) {
                $script:CaptureOwnerLastHeartbeat = Get-Date
            }
            Test-Heartbeat
            Write-Text $Ctx (Render-WatchState -TabId $tab); return
        }

        '^/api/poll$' {
            $cv = -1
            if ($req.QueryString['ver']) { $cv = [int]$req.QueryString['ver'] }
            if ($cv -eq $script:Version) { $Ctx.Response.StatusCode = 204; return }
            Write-Text $Ctx (Render-StepList); return
        }

        '^/api/stats$' { Write-Text $Ctx (Render-Stats); return }

        '^/img/(?<id>[a-zA-Z0-9\-]+)$' {
            $id = $Matches['id']
            $st = $script:Steps | Where-Object { $_.id -eq $id } | Select-Object -First 1
            if (-not $st -or -not (Test-Path -LiteralPath $st.file)) {
                Write-Text $Ctx 'not found' 404 'text/plain; charset=utf-8'; return
            }
            if ($req.QueryString['thumb'] -eq '1') { Write-Bytes $Ctx (Get-Thumb -Path $st.file) 'image/jpeg' }
            else { Write-Bytes $Ctx ([IO.File]::ReadAllBytes($st.file)) (Get-ImageContentType -Path $st.file) }
            return
        }

        '^/api/shutdown$' {
            Write-Text $Ctx 'bye' 200 'text/plain; charset=utf-8'
            $script:Running = $false; return
        }

        default { Write-Text $Ctx 'not found' 404 'text/plain; charset=utf-8'; return }
    }
}

# ---------------------------------------------------------------------
# 起動（V-1: localhost → 失敗コードを記録）
# ---------------------------------------------------------------------
$listener = $null
$bound = $false
$lastErr = ''
for ($i = 0; $i -lt 5; $i++) {
    try {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://localhost:$Port/")
        $listener.Start()
        $bound = $true; break
    } catch [System.Net.HttpListenerException] {
        $lastErr = "ErrorCode=$($_.Exception.ErrorCode) $($_.Exception.Message)"
        Write-MBLog "ポート $Port で失敗: $lastErr" 'WARN'
        if ($_.Exception.ErrorCode -eq 5) {
            Write-MBLog 'ERROR_ACCESS_DENIED = 名前空間予約が必要。Plan B（TcpListener）が必要です' 'ERR'
            break
        }
        $Port++
    } catch {
        $lastErr = $_.Exception.Message; $Port++
    }
}
if (-not $bound) {
    Write-Host ''
    Write-Host '  V-1 は失敗しました。' -ForegroundColor Red
    Write-Host "  最後のエラー: $lastErr" -ForegroundColor Red
    Write-Host '  → TECHNICAL_DESIGN.md の Plan B（TcpListener による内製HTTPサーバ）に切り替えます。' -ForegroundColor Yellow
    exit 1
}

$url = "http://localhost:$Port/"
Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host '  V-1 / V-2 / V-7 / V-10: サーバ検証（Word は起動しません）' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-MBLog "起動: $url" 'OK'
Write-MBLog "DPI設定: $script:DpiResult" 'INFO'
Write-MBLog "保存先: $(if ($script:WatchDir) { $script:WatchDir } else { '未検出' })" 'INFO'
Write-Host ''
Write-Host '  試してください:' -ForegroundColor Yellow
Write-Host '   1. Win+Shift+S で撮る → 2秒以内に増えるか（V-7）'
Write-Host '   2. Ctrl+V で貼り付け → 増えるか（V-2）'
Write-Host '   3. 同じ画像をもう一度貼り付け → 重複スキップされるか'
Write-Host '   4. 画像を複数まとめてドロップ → 順序が保たれるか'
Write-Host '   5. 「3秒後にキャプチャ」→ 対象アプリに切り替えて撮れるか'
Write-Host '   6. ★ブラウザのタブを閉じて30秒待つ → スクショを3枚撮る'
Write-Host "      → もう一度 $url を開く"
Write-Host '      → その3枚が「取り込まれていない」ことを確認（V-10）'
Write-Host ''
Write-Host '  停止: Ctrl+C または画面の「終了」ボタン' -ForegroundColor Yellow
Write-Host ''

if (-not $NoBrowser) { Start-Process $url }

try {
    while ($script:Running -and $listener.IsListening) {
        $task = $listener.GetContextAsync()
        while (-not $task.AsyncWaitHandle.WaitOne(200)) {
            Invoke-WatcherFlush
            Test-Heartbeat
            if (-not $script:Running) { break }
        }
        if (-not $script:Running) { break }
        if (-not $task.IsCompleted) { continue }
        $ctx = $task.GetAwaiter().GetResult()
        try { Invoke-Route -Ctx $ctx }
        catch {
            Write-MBLog "リクエスト処理でエラー: $($_.Exception.Message)" 'ERR'
            try { Write-Text $ctx ('500: ' + (HtmlEnc $_.Exception.Message)) 500 'text/plain; charset=utf-8' } catch { }
        }
        finally { try { $ctx.Response.Close() } catch { } }
    }
} finally {
    Write-Host ''
    Write-MBLog '停止処理' 'INFO'
    foreach ($sid in @('MBS.C', 'MBS.R')) {
        try { Unregister-Event -SourceIdentifier $sid -ErrorAction SilentlyContinue } catch { }
    }
    if ($script:Watcher) { try { $script:Watcher.EnableRaisingEvents = $false; $script:Watcher.Dispose() } catch { } }
    try { $listener.Stop(); $listener.Close() } catch { }

    $rep = New-Object System.Collections.ArrayList
    [void]$rep.Add([pscustomobject]@{ 検証 = 'V-1 HttpListener(localhost,非管理者)'; 判定 = 'OK'; 詳細 = $url })
    $raw = $script:Stats['受信_貼り付け'] + $script:Stats['受信_ドロップ'] + $script:Stats['受信_ファイル選択']
    [void]$rep.Add([pscustomobject]@{ 検証 = 'V-2 生バイトPOST'
        判定 = $(if ($raw -gt 0) { 'OK' } else { '未検証' })
        詳細 = "貼付 $($script:Stats['受信_貼り付け']) / ドロップ $($script:Stats['受信_ドロップ']) / ファイル $($script:Stats['受信_ファイル選択'])" })
    [void]$rep.Add([pscustomobject]@{ 検証 = 'V-7 監視による追加'
        判定 = $(if ($script:Stats['受信_監視'] -gt 0) { 'OK' } elseif ($script:WatchDir) { '未検証' } else { 'NG' })
        詳細 = "$($script:Stats['受信_監視']) 件" })
    [void]$rep.Add([pscustomobject]@{ 検証 = 'V-10 監視の自動停止'
        判定 = $(if ($script:SuspendCount -gt 0) { 'OK' } else { '未検証' })
        詳細 = "一時停止 $($script:SuspendCount) 回" })
    [void]$rep.Add([pscustomobject]@{ 検証 = 'V-10 バックフィル防止'
        判定 = $(if ($script:SkippedByWatermark -gt 0) { 'OK' } else { '未検証' })
        詳細 = "基準時刻でスキップ $($script:SkippedByWatermark) 件" })
    [void]$rep.Add([pscustomobject]@{ 検証 = '重複除去(SHA-256)'
        判定 = $(if ($script:Stats['重複スキップ'] -gt 0) { 'OK' } else { '未検証' })
        詳細 = "$($script:Stats['重複スキップ']) 件" })
    [void]$rep.Add([pscustomobject]@{ 検証 = '遅延キャプチャ'
        判定 = $(if ($script:Stats['受信_ボタン撮影'] -gt 0) { 'OK' } else { '未検証' })
        詳細 = "$($script:Stats['受信_ボタン撮影']) 件" })
    [void]$rep.Add([pscustomobject]@{ 検証 = 'DPI awareness'; 判定 = 'INFO'; 詳細 = $script:DpiResult })

    Write-Host ''
    Write-Host '=== 結果 ===' -ForegroundColor Cyan
    $rep | Format-Table -AutoSize -Wrap
    $rp = Join-Path $script:Root 'result-08-server.txt'
    $rep | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Out-File -LiteralPath $rp -Encoding UTF8
    Write-Host "レポート: $rp" -ForegroundColor Cyan
    Write-Host ''
}
