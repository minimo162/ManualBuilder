# 操作を記録して手順にする。
#
# 録画から場面を切り出す方式は、画面の変化から「押された場所」を推定する。
# こちらは推定しない。Windows の UI Automation が、クリックした先のコントロールの
# 名前・種類・矩形をそのまま持っているので、それを読む。赤枠の位置も操作対象の名前も
# 確定値になり、文字認識も要らなくなる。
#
# 実装上の判断:
#
#   低レベルマウスフック（WH_MOUSE_LL）は座標の受け取りだけに使う。
#     スクリーンショットや UI Automation を同じコールバックで実行すると
#     LowLevelHooksTimeout を超えるため、コールバックは時刻・座標・ウィンドウを
#     スレッドセーフなキューへ積んですぐ戻る。重い処理は記録ループが後から行う。
#     これにより画像保存中に続いたクリックも1件へ潰れない。フックを開始できない
#     環境だけは GetAsyncKeyState のポーリングへ安全にフォールバックする。
#
#   キーの文字は記録しない。
#     GetAsyncKeyStateからは「入力があった」ことだけを見て、文字列へ変換しない。
#     ただし画面に表示された文字はスクリーンショットへ残し、必要な箇所だけ利用者が
#     取り込み後の画像編集で黒塗りする。
#
#   プロセスを DPI 認識にする。
#     これをしないと、高DPI環境で GetCursorPos や CopyFromScreen が仮想化された座標を
#     返す一方、UI Automation は物理座標を返す。赤枠が実際の位置からずれる。

Set-StrictMode -Version 2.0

$script:MbRecorderNativeReady = $false
$script:MbRecorderUiaReady = $false
$script:MbRecorderJobsRoot = ''
$script:MbRecorderScriptRoot = ''
$script:MbRecorderJob = $null
$script:MbRecorderLastUndoRequestId = ''
$script:MbRecorderLastResultRequestId = ''
$script:MbRecorderCaptureCompleteness = 'unknown'
$script:MbRecorderCaptureWarning = ''

# 記録から除くウィンドウ。タスクバーとデスクトップを押しただけの操作は手順にしない。
$script:MbRecorderIgnoredClasses = @('Shell_TrayWnd', 'Shell_SecondaryTrayWnd', 'WorkerW', 'Progman', 'NotifyIconOverflowWindow')

# 操作対象として妥当なコントロール。葉の要素が文字や画像だったとき、
# ここに挙げた種類の親まで遡って赤枠を寄せる。
$script:MbRecorderInteractiveTypes = @(
    'ControlType.Button', 'ControlType.MenuItem', 'ControlType.TabItem', 'ControlType.ListItem',
    'ControlType.TreeItem', 'ControlType.CheckBox', 'ControlType.RadioButton', 'ControlType.ComboBox',
    'ControlType.Hyperlink', 'ControlType.SplitButton', 'ControlType.Edit', 'ControlType.DataItem'
)

function Initialize-MbRecorderNative {
    if ($script:MbRecorderNativeReady) { return }

    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName Accessibility -ErrorAction Stop
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class MbRecorderNative
{
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT Point;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint VirtualKey;
        public uint ScanCode;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr HWnd;
        public uint Message;
        public UIntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public POINT Point;
    }

    public sealed class MouseClick
    {
        public int X;
        public int Y;
        public int Message;
        public long WindowHandle;
        public long Timestamp;
    }

    // キーの内容は保存せず、「文字が変わる操作」か「入力確定」かという事実だけを
    // 一時キューへ積む。重い画面取得中の短いキー押下も取りこぼさないために使う。
    public sealed class KeyboardActivity
    {
        public int Kind; // 1=text-changing, 2=commit
        public long WindowHandle;
        public long Timestamp;
    }

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
    private static readonly ConcurrentQueue<MouseClick> MouseClicks = new ConcurrentQueue<MouseClick>();
    private static readonly ConcurrentQueue<KeyboardActivity> KeyboardActivities = new ConcurrentQueue<KeyboardActivity>();
    private static readonly object MouseHookSync = new object();
    private static LowLevelMouseProc MouseHookCallback;
    private static Thread MouseHookThread;
    private static ManualResetEventSlim MouseHookReady;
    private static IntPtr MouseHookHandle = IntPtr.Zero;
    private static LowLevelKeyboardProc KeyboardHookCallback;
    private static IntPtr KeyboardHookHandle = IntPtr.Zero;
    private static bool KeyboardHookStarted;
    private static uint MouseHookThreadId;
    private static bool MouseHookStarted;
    private static long DroppedMouseClickCount;
    private static long DroppedKeyboardActivityCount;

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc callback, IntPtr module, uint threadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc callback, IntPtr module, uint threadId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern int GetMessage(out MSG message, IntPtr window, uint minimum, uint maximum);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostThreadMessage(uint threadId, uint message, UIntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string moduleName);

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(POINT point);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr window, uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

    // UI Automationを公開しない旧式/独自アプリの操作対象をMSAAから取得する。
    // 戻り値のchildはCHILDID_SELFまたは親IAccessible内の子IDになる。
    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromPoint(
        POINT point,
        [MarshalAs(UnmanagedType.Interface)] out object accessible,
        [MarshalAs(UnmanagedType.Struct)] out object child);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr value);

    // 影や不可視の枠を含まない、見た目どおりのウィンドウ矩形を返す。
    // GetWindowRect は Windows 10 以降だと左右に数ピクセルの余白が付く。
    public static RECT GetVisualWindowRect(IntPtr hWnd)
    {
        RECT rect;
        const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
        int size = Marshal.SizeOf(typeof(RECT));
        if (DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, out rect, size) == 0)
        {
            if (rect.Right > rect.Left && rect.Bottom > rect.Top) { return rect; }
        }
        GetWindowRect(hWnd, out rect);
        return rect;
    }

    public static string GetWindowClass(IntPtr hWnd)
    {
        StringBuilder builder = new StringBuilder(256);
        GetClassName(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    public static string GetWindowTitle(IntPtr hWnd)
    {
        StringBuilder builder = new StringBuilder(512);
        GetWindowTextW(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private static IntPtr HandleMouseHook(int code, IntPtr wParam, IntPtr lParam)
    {
        const int WM_LBUTTONDOWN = 0x0201;
        const int WM_RBUTTONDOWN = 0x0204;
        if (code >= 0)
        {
            int message = unchecked((int)wParam.ToInt64());
            if (message == WM_LBUTTONDOWN || message == WM_RBUTTONDOWN)
            {
                MSLLHOOKSTRUCT data = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                IntPtr target = WindowFromPoint(data.Point);
                IntPtr root = target == IntPtr.Zero ? IntPtr.Zero : GetAncestor(target, 2); // GA_ROOT
                MouseClicks.Enqueue(new MouseClick {
                    X = data.Point.X,
                    Y = data.Point.Y,
                    Message = message,
                    WindowHandle = root.ToInt64(),
                    Timestamp = Stopwatch.GetTimestamp()
                });
                while (MouseClicks.Count > 512)
                {
                    MouseClick ignored;
                    if (!MouseClicks.TryDequeue(out ignored)) { break; }
                    Interlocked.Increment(ref DroppedMouseClickCount);
                }
            }
        }
        return CallNextHookEx(MouseHookHandle, code, wParam, lParam);
    }

    private static bool IsTextChangingKey(uint key)
    {
        if ((key >= 0x30 && key <= 0x5A) || (key >= 0x60 && key <= 0x6F) ||
            (key >= 0xBA && key <= 0xE2)) { return true; }
        return key == 0x08 || key == 0x20 || key == 0x2E; // Backspace, Space, Delete
    }

    private static IntPtr HandleKeyboardHook(int code, IntPtr wParam, IntPtr lParam)
    {
        const int WM_KEYDOWN = 0x0100;
        const int WM_SYSKEYDOWN = 0x0104;
        if (code >= 0)
        {
            int message = unchecked((int)wParam.ToInt64());
            if (message == WM_KEYDOWN || message == WM_SYSKEYDOWN)
            {
                KBDLLHOOKSTRUCT data = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                int kind = (data.VirtualKey == 0x0D || data.VirtualKey == 0x09) ? 2 :
                    (IsTextChangingKey(data.VirtualKey) ? 1 : 0);
                bool control = (GetAsyncKeyState(0x11) & 0x8000) != 0;
                bool commandModifier = (GetAsyncKeyState(0x12) & 0x8000) != 0 ||
                    (GetAsyncKeyState(0x5B) & 0x8000) != 0 || (GetAsyncKeyState(0x5C) & 0x8000) != 0;
                if (commandModifier || (control && data.VirtualKey != 0x56 &&
                    data.VirtualKey != 0x58 && data.VirtualKey != 0x5A)) { kind = 0; }
                if (kind != 0)
                {
                    IntPtr foreground = GetForegroundWindow();
                    KeyboardActivities.Enqueue(new KeyboardActivity {
                        Kind = kind,
                        WindowHandle = foreground.ToInt64(),
                        Timestamp = Stopwatch.GetTimestamp()
                    });
                    while (KeyboardActivities.Count > 2048)
                    {
                        KeyboardActivity ignored;
                        if (!KeyboardActivities.TryDequeue(out ignored)) { break; }
                        Interlocked.Increment(ref DroppedKeyboardActivityCount);
                    }
                }
            }
        }
        return CallNextHookEx(KeyboardHookHandle, code, wParam, lParam);
    }

    private static void RunMouseHook()
    {
        const int WH_MOUSE_LL = 14;
        MouseHookThreadId = GetCurrentThreadId();
        MouseHookCallback = HandleMouseHook;
        MouseHookHandle = SetWindowsHookEx(WH_MOUSE_LL, MouseHookCallback, GetModuleHandle(null), 0);
        KeyboardHookCallback = HandleKeyboardHook;
        KeyboardHookHandle = SetWindowsHookEx(13, KeyboardHookCallback, GetModuleHandle(null), 0); // WH_KEYBOARD_LL
        KeyboardHookStarted = KeyboardHookHandle != IntPtr.Zero;
        MouseHookStarted = MouseHookHandle != IntPtr.Zero;
        MouseHookReady.Set();
        if (!MouseHookStarted)
        {
            // StartMouseHook() reports failure when the mouse half cannot start. Do not
            // leave a keyboard-only hook alive behind that failed recorder session.
            if (KeyboardHookHandle != IntPtr.Zero) { UnhookWindowsHookEx(KeyboardHookHandle); }
            KeyboardHookHandle = IntPtr.Zero;
            KeyboardHookStarted = false;
            MouseHookThreadId = 0;
            return;
        }
        MSG message;
        while (GetMessage(out message, IntPtr.Zero, 0, 0) > 0) { }
        UnhookWindowsHookEx(MouseHookHandle);
        if (KeyboardHookHandle != IntPtr.Zero) { UnhookWindowsHookEx(KeyboardHookHandle); }
        MouseHookHandle = IntPtr.Zero;
        KeyboardHookHandle = IntPtr.Zero;
        KeyboardHookStarted = false;
        MouseHookStarted = false;
    }

    public static bool StartMouseHook()
    {
        lock (MouseHookSync)
        {
            if (MouseHookStarted) { return true; }
            ClearMouseClicks();
            ClearKeyboardActivities();
            Interlocked.Exchange(ref DroppedMouseClickCount, 0);
            Interlocked.Exchange(ref DroppedKeyboardActivityCount, 0);
            MouseHookReady = new ManualResetEventSlim(false);
            MouseHookThread = new Thread(RunMouseHook);
            MouseHookThread.IsBackground = true;
            MouseHookThread.Name = "ManualBuilder mouse capture";
            MouseHookThread.Start();
            if (!MouseHookReady.Wait(1500)) { return false; }
            return MouseHookStarted;
        }
    }

    public static void StopMouseHook()
    {
        lock (MouseHookSync)
        {
            if (MouseHookThreadId != 0) { PostThreadMessage(MouseHookThreadId, 0x0012, UIntPtr.Zero, IntPtr.Zero); }
            if (MouseHookThread != null && MouseHookThread.IsAlive) { MouseHookThread.Join(1000); }
            MouseHookThread = null;
            MouseHookThreadId = 0;
            MouseHookStarted = false;
            KeyboardHookStarted = false;
            ClearMouseClicks();
            ClearKeyboardActivities();
        }
    }

    public static MouseClick DequeueMouseClick()
    {
        MouseClick click;
        return MouseClicks.TryDequeue(out click) ? click : null;
    }

    public static MouseClick PeekMouseClick()
    {
        MouseClick click;
        return MouseClicks.TryPeek(out click) ? click : null;
    }

    public static void ClearMouseClicks()
    {
        MouseClick ignored;
        while (MouseClicks.TryDequeue(out ignored)) { }
    }

    public static KeyboardActivity DequeueKeyboardActivity()
    {
        KeyboardActivity activity;
        return KeyboardActivities.TryDequeue(out activity) ? activity : null;
    }

    public static KeyboardActivity PeekKeyboardActivity()
    {
        KeyboardActivity activity;
        return KeyboardActivities.TryPeek(out activity) ? activity : null;
    }

    public static void ClearKeyboardActivities()
    {
        KeyboardActivity ignored;
        while (KeyboardActivities.TryDequeue(out ignored)) { }
    }

    public static long GetTimestamp() { return Stopwatch.GetTimestamp(); }
    public static long TimestampFrequency { get { return Stopwatch.Frequency; } }
    public static bool KeyboardHookAvailable { get { return KeyboardHookStarted; } }
    public static long DroppedMouseClicks { get { return Interlocked.Read(ref DroppedMouseClickCount); } }
    public static long DroppedKeyboardActivities { get { return Interlocked.Read(ref DroppedKeyboardActivityCount); } }
}
'@ -ErrorAction Stop

    $script:MbRecorderNativeReady = $true
}

# 座標系を物理ピクセルへ揃える。画面取得・カーソル位置・UI Automation の3つが
# 同じ物差しを使うようにするため、記録を始める前に必ず呼ぶ。
function Set-MbProcessDpiAware {
    Initialize-MbRecorderNative
    # Windows 10 1703 以降はモニターごとのDPIに追従できる。使えなければ全体DPIで妥協する。
    $perMonitorV2 = [IntPtr]::new(-4)
    try {
        if ([MbRecorderNative]::SetProcessDpiAwarenessContext($perMonitorV2) -ne [IntPtr]::Zero) { return 'per-monitor' }
    } catch { }
    try {
        if ([MbRecorderNative]::SetProcessDPIAware()) { return 'system' }
    } catch { }
    return 'none'
}

function Initialize-MbRecorderUia {
    if ($script:MbRecorderUiaReady) { return $true }
    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop
        Add-Type -AssemblyName WindowsBase -ErrorAction Stop
        $script:MbRecorderUiaReady = $true
        return $true
    } catch {
        return $false
    }
}

function Get-MbRecorderCapability {
    $native = $true
    try { Initialize-MbRecorderNative } catch { $native = $false }
    $uia = Initialize-MbRecorderUia
    $reason = ''
    if (-not $native) { $reason = 'この環境では画面と入力の状態を取得できません。' }
    elseif (-not $uia) { $reason = 'この環境ではUI Automationを利用できないため、押したボタンの名前を取得できません。' }
    return [pscustomobject]@{ available = ($native -and $uia); native = $native; uia = $uia; reason = $reason }
}

# ---------------------------------------------------------------------
# UI Automation
# ---------------------------------------------------------------------

function Get-MbAutomationElementInfo {
    param([Parameter(Mandatory = $true)]$Element)

    $current = $Element.Current
    $rect = $current.BoundingRectangle
    $controlType = [string]$current.ControlType.ProgrammaticName
    $isActionable = $script:MbRecorderInteractiveTypes -contains $controlType

    # Chromium系ブラウザーや独自業務アプリでは、リンクやボタンがText/Pane/Customとして
    # 公開されても InvokePattern などを持つことがある。種類だけでなく操作パターンも見る。
    # ただし全要素への個別プロパティ照会は遅いため、候補になり得る種類だけを1回で調べる。
    $patternCandidateTypes = @('ControlType.Text', 'ControlType.Pane', 'ControlType.Custom', 'ControlType.Image', 'ControlType.Group')
    if (-not $isActionable -and $patternCandidateTypes -contains $controlType) {
        $actionPatternNames = @(
            'InvokePatternIdentifiers.Pattern',
            'SelectionItemPatternIdentifiers.Pattern',
            'TogglePatternIdentifiers.Pattern',
            'ExpandCollapsePatternIdentifiers.Pattern',
            'RangeValuePatternIdentifiers.Pattern',
            'ValuePatternIdentifiers.Pattern'
        )
        try {
            foreach ($pattern in @($Element.GetSupportedPatterns())) {
                if ($actionPatternNames -contains [string]$pattern.ProgrammaticName) {
                    $isActionable = $true
                    break
                }
            }
        } catch { }
        # Chromiumの一部要素はInvokePatternを出さず、LegacyIAccessibleの既定動作だけを
        # 公開する。既定動作が空でない要素だけを操作可能とみなし、単なる文章は除く。
        if (-not $isActionable) {
            try {
                $legacyObject = $null
                if ($Element.TryGetCurrentPattern(
                    [System.Windows.Automation.LegacyIAccessiblePattern]::Pattern,
                    [ref]$legacyObject
                )) {
                    $legacy = [System.Windows.Automation.LegacyIAccessiblePattern]$legacyObject
                    $isActionable = -not [string]::IsNullOrWhiteSpace([string]$legacy.Current.DefaultAction)
                }
            } catch { }
        }
    }

    return [pscustomobject]@{
        name         = [string]$current.Name
        controlType  = $controlType
        automationId = [string]$current.AutomationId
        className    = [string]$current.ClassName
        left         = [double]$rect.Left
        top          = [double]$rect.Top
        width        = [double]$rect.Width
        height       = [double]$rect.Height
        isActionable = $isActionable
    }
}

function Test-MbUsableElementInfo {
    param([AllowNull()]$Info)
    if ($null -eq $Info) { return $false }
    foreach ($value in @([double]$Info.left, [double]$Info.top, [double]$Info.width, [double]$Info.height)) {
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { return $false }
    }
    if ([double]$Info.width -le 0 -or [double]$Info.height -le 0) { return $false }
    return $true
}

function Test-MbRecorderInteractiveElementInfo {
    param([AllowNull()]$Info)
    if (-not (Test-MbUsableElementInfo -Info $Info)) { return $false }
    if ($Info.PSObject.Properties.Name -contains 'isActionable') {
        return [bool]$Info.isActionable
    }
    return $script:MbRecorderInteractiveTypes -contains [string]$Info.controlType
}

function Test-MbPointWithinElementInfo {
    param(
        [AllowNull()]$Info,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y,
        [double]$Tolerance = 1.0
    )
    if (-not (Test-MbUsableElementInfo -Info $Info)) { return $false }
    $right = [double]$Info.left + [double]$Info.width
    $bottom = [double]$Info.top + [double]$Info.height
    return ($X -ge ([double]$Info.left - $Tolerance) -and $X -le ($right + $Tolerance) -and
        $Y -ge ([double]$Info.top - $Tolerance) -and $Y -le ($bottom + $Tolerance))
}

# 同じ点に重なる候補から、実際に操作できる最小の要素を選ぶ。
# 名前は操作説明の手がかりだが、名前が空でもボタンの矩形は赤枠として有用なので除外しない。
function Select-MbUiaTargetInfo {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Candidates,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y
    )

    $exact = New-Object System.Collections.ArrayList
    $nearby = New-Object System.Collections.ArrayList
    foreach ($candidate in @($Candidates)) {
        if (-not (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y -Tolerance 12.0)) { continue }
        if (-not (Test-MbRecorderInteractiveElementInfo -Info $candidate)) { continue }
        if (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y -Tolerance 1.0) {
            [void]$exact.Add($candidate)
        } else {
            [void]$nearby.Add($candidate)
        }
    }

    # 小ささより、実クリックを矩形内に含むという直接証拠を優先する。
    [object[]]$pool = @()
    if ($exact.Count -gt 0) { $pool = @($exact) }
    elseif ($nearby.Count -gt 0) { $pool = @($nearby) }
    if (@($pool).Count -eq 0) { return $null }
    $best = $pool | Sort-Object `
        @{ Expression = {
            $right = [double]$_.left + [double]$_.width
            $bottom = [double]$_.top + [double]$_.height
            $dx = [Math]::Max(0.0, [Math]::Max([double]$_.left - $X, $X - $right))
            $dy = [Math]::Max(0.0, [Math]::Max([double]$_.top - $Y, $Y - $bottom))
            [Math]::Sqrt(($dx * $dx) + ($dy * $dy))
        } },
        @{ Expression = { [double]$_.width * [double]$_.height } },
        @{ Expression = { [string]::IsNullOrWhiteSpace([string]$_.name) } } | Select-Object -First 1
    if ($exact.Count -eq 0) {
        $best | Add-Member -NotePropertyName 'isNearby' -NotePropertyValue $true -Force
        $best | Add-Member -NotePropertyName 'confidence' -NotePropertyValue 'low' -Force
    }
    return $best
}

# 操作パターンを公開しないWebアプリでも、クリック点の下に名前付きText/Group等があれば
# その最小要素を対象として使う。ウィンドウの大半を占めるコンテナは採用しない。
function Select-MbUiaNamedTargetInfo {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Candidates,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y,
        [AllowNull()]$Window
    )

    $exact = New-Object System.Collections.ArrayList
    $nearby = New-Object System.Collections.ArrayList
    $windowArea = if ($null -ne $Window) { [double]$Window.width * [double]$Window.height } else { 0.0 }
    foreach ($candidate in @($Candidates)) {
        if (-not (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y -Tolerance 12.0)) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$candidate.name)) { continue }
        if ([string]$candidate.controlType -in @('ControlType.Window', 'ControlType.Document')) { continue }
        $area = [double]$candidate.width * [double]$candidate.height
        if ($windowArea -gt 0 -and $area -gt ($windowArea * 0.4)) { continue }
        if (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y -Tolerance 1.0) {
            [void]$exact.Add($candidate)
        } else {
            [void]$nearby.Add($candidate)
        }
    }
    [object[]]$pool = @()
    if ($exact.Count -gt 0) { $pool = @($exact) }
    elseif ($nearby.Count -gt 0) { $pool = @($nearby) }
    $best = $pool | Sort-Object @{ Expression = { [double]$_.width * [double]$_.height } } | Select-Object -First 1
    if ($null -ne $best) {
        $best | Add-Member -NotePropertyName 'isInferred' -NotePropertyValue $true -Force
        if ($exact.Count -eq 0) {
            $best | Add-Member -NotePropertyName 'isNearby' -NotePropertyValue $true -Force
            $best | Add-Member -NotePropertyName 'confidence' -NotePropertyValue 'low' -Force
        }
    }
    return $best
}

# MSAAの数値ロールをUI Automationと同じ名称へ寄せる。保存形式と後段の説明生成を
# UIA経路と共通にでき、どのAPIから取得したかを利用者が意識せずに済む。
function ConvertTo-MbMsaaControlType {
    param([Parameter(Mandatory = $true)][int]$Role)

    switch ($Role) {
        12 { return 'ControlType.MenuItem' }
        29 { return 'ControlType.DataItem' }
        30 { return 'ControlType.Hyperlink' }
        34 { return 'ControlType.ListItem' }
        36 { return 'ControlType.TreeItem' }
        37 { return 'ControlType.TabItem' }
        42 { return 'ControlType.Edit' }
        43 { return 'ControlType.Button' }
        44 { return 'ControlType.CheckBox' }
        45 { return 'ControlType.RadioButton' }
        { $_ -in @(46, 47) } { return 'ControlType.ComboBox' }
        { $_ -in @(56, 57, 58, 62) } { return 'ControlType.SplitButton' }
        default { return 'ControlType.Custom' }
    }
}

function New-MbMsaaElementInfo {
    param(
        [AllowEmptyString()][string]$Name,
        [Parameter(Mandatory = $true)][int]$Role,
        [AllowEmptyString()][string]$DefaultAction,
        [Parameter(Mandatory = $true)][double]$Left,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$Width,
        [Parameter(Mandatory = $true)][double]$Height
    )

    $actionableRoles = @(12, 29, 30, 34, 36, 37, 42, 43, 44, 45, 46, 47, 50, 51, 52, 56, 57, 58, 62, 64)
    return [pscustomobject]@{
        name = $Name
        controlType = ConvertTo-MbMsaaControlType -Role $Role
        automationId = ''
        className = ''
        left = $Left
        top = $Top
        width = $Width
        height = $Height
        isActionable = (($actionableRoles -contains $Role) -or -not [string]::IsNullOrWhiteSpace($DefaultAction))
        provider = 'MSAA'
    }
}

# UIAが対象を返さないアプリでは、Windows標準のアクセシビリティAPI (MSAA) を
# クリック点へ直接問い合わせる。画面全体に近いオブジェクトはここでも除外する。
function Get-MbMsaaTargetAtPoint {
    param(
        [Parameter(Mandatory = $true)][int]$X,
        [Parameter(Mandatory = $true)][int]$Y,
        [AllowNull()]$Window = $null
    )

    $accessible = $null
    try {
        Initialize-MbRecorderNative
        $point = New-Object 'MbRecorderNative+POINT'
        $point.X = $X
        $point.Y = $Y
        $child = $null
        $result = [MbRecorderNative]::AccessibleObjectFromPoint($point, [ref]$accessible, [ref]$child)
        if ($result -lt 0 -or $null -eq $accessible) { return $null }

        $accessibleObject = [Accessibility.IAccessible]$accessible
        $childId = if ($null -eq $child) { [object]0 } else { [object]$child }
        $name = ''
        $defaultAction = ''
        $role = 0
        try { $name = [string]$accessibleObject.get_accName($childId) } catch { }
        try { $defaultAction = [string]$accessibleObject.get_accDefaultAction($childId) } catch { }
        try { $role = [int]$accessibleObject.get_accRole($childId) } catch { }

        $left = 0
        $top = 0
        $width = 0
        $height = 0
        try {
            $accessibleObject.accLocation([ref]$left, [ref]$top, [ref]$width, [ref]$height, $childId)
        } catch {
            return $null
        }

        $info = New-MbMsaaElementInfo -Name $name -Role $role -DefaultAction $defaultAction `
            -Left $left -Top $top -Width $width -Height $height

        # 点の直下がボタン内の静的なラベルだった場合、pvarChildを返したIAccessible本体が
        # 操作可能な親であることがある。子を静的文字として採用せず、親自身を1段だけ確認する。
        if (-not $info.isActionable -and [int]$childId -ne 0) {
            $parentName = ''
            $parentAction = ''
            $parentRole = 0
            $self = [object]0
            try { $parentName = [string]$accessibleObject.get_accName($self) } catch { }
            try { $parentAction = [string]$accessibleObject.get_accDefaultAction($self) } catch { }
            try { $parentRole = [int]$accessibleObject.get_accRole($self) } catch { }
            $parentLeft = 0
            $parentTop = 0
            $parentWidth = 0
            $parentHeight = 0
            try {
                $accessibleObject.accLocation([ref]$parentLeft, [ref]$parentTop, [ref]$parentWidth, [ref]$parentHeight, $self)
                $parentInfo = New-MbMsaaElementInfo -Name $parentName -Role $parentRole -DefaultAction $parentAction `
                    -Left $parentLeft -Top $parentTop -Width $parentWidth -Height $parentHeight
                if ($parentInfo.isActionable) { $info = $parentInfo }
            } catch { }
        }

        if (-not (Test-MbPointWithinElementInfo -Info $info -X $X -Y $Y -Tolerance 12.0)) { return $null }
        if (-not $info.isActionable) { return $null }

        if ($null -ne $Window) {
            $windowArea = [double]$Window.width * [double]$Window.height
            $targetArea = [double]$info.width * [double]$info.height
            if ($windowArea -gt 0 -and $targetArea -gt ($windowArea * 0.4)) { return $null }
        }
        return $info
    } catch {
        return $null
    } finally {
        if ($null -ne $accessible -and [Runtime.InteropServices.Marshal]::IsComObject($accessible)) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($accessible) } catch { }
        }
    }
}

# FromPoint は境界線や文字の隙間では親の Document/Pane を返すことがある。
# 同じ操作対象の内側を数ピクセルずらして再照会し、各要素の親も Raw/Control の
# 両方のビューで確認する。DOM全体を1件ずつ歩くより軽く、入れ子の深いWeb UIにも強い。
function Add-MbUiaPointCandidates {
    param(
        [Parameter(Mandatory = $true)]$Candidates,
        [Parameter(Mandatory = $true)]$Element,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y
    )

    foreach ($walker in @(
        [System.Windows.Automation.TreeWalker]::ControlViewWalker,
        [System.Windows.Automation.TreeWalker]::RawViewWalker
    )) {
        $node = $Element
        for ($depth = 0; $depth -lt 12 -and $null -ne $node; $depth++) {
            $info = $null
            try { $info = Get-MbAutomationElementInfo -Element $node } catch { $info = $null }
            if (Test-MbPointWithinElementInfo -Info $info -X $X -Y $Y -Tolerance 12.0) {
                [void]$Candidates.Add($info)
            }
            if ($null -ne $info -and [string]$info.controlType -eq 'ControlType.Window') { break }
            try { $node = $walker.GetParent($node) } catch { $node = $null }
        }
    }
}

# FromPoint がページ全体を返した場合は、標準の操作コントロールだけをUIAプロバイダー側で
# 絞り込む。前版の「各階層の先頭256兄弟を順に調べる」方式では、長いページの後半や
# 重なった別枝にあるリンクを見落としていた。非表示要素を除くことで照会量も抑える。
function Add-MbUiaInteractiveDescendantCandidates {
    param(
        [Parameter(Mandatory = $true)]$Candidates,
        [Parameter(Mandatory = $true)]$Root,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y
    )

    try {
        $types = @(
            [System.Windows.Automation.ControlType]::Button,
            [System.Windows.Automation.ControlType]::MenuItem,
            [System.Windows.Automation.ControlType]::TabItem,
            [System.Windows.Automation.ControlType]::ListItem,
            [System.Windows.Automation.ControlType]::TreeItem,
            [System.Windows.Automation.ControlType]::CheckBox,
            [System.Windows.Automation.ControlType]::RadioButton,
            [System.Windows.Automation.ControlType]::ComboBox,
            [System.Windows.Automation.ControlType]::Hyperlink,
            [System.Windows.Automation.ControlType]::SplitButton,
            [System.Windows.Automation.ControlType]::Edit,
            [System.Windows.Automation.ControlType]::DataItem,
            [System.Windows.Automation.ControlType]::Text,
            [System.Windows.Automation.ControlType]::Pane,
            [System.Windows.Automation.ControlType]::Custom,
            [System.Windows.Automation.ControlType]::Image,
            [System.Windows.Automation.ControlType]::Group
        )
        $conditions = New-Object 'System.Collections.Generic.List[System.Windows.Automation.Condition]'
        foreach ($type in $types) {
            [void]$conditions.Add((New-Object System.Windows.Automation.PropertyCondition -ArgumentList @(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $type
            )))
        }
        $typeCondition = [System.Windows.Automation.OrCondition]::new(
            [System.Windows.Automation.Condition[]]$conditions.ToArray()
        )
        $visibleCondition = New-Object System.Windows.Automation.PropertyCondition -ArgumentList @(
            [System.Windows.Automation.AutomationElement]::IsOffscreenProperty, $false
        )
        $condition = [System.Windows.Automation.AndCondition]::new($typeCondition, $visibleCondition)
        $elements = $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
        foreach ($element in @($elements)) {
            # Group/Textまで検索対象を広げても、クリック点と無関係な全要素へ操作パターンを
            # 問い合わせない。まず軽い矩形だけで絞り、重い照会は点を含む数件に限定する。
            $boundsInfo = $null
            try {
                $bounds = $element.Current.BoundingRectangle
                $boundsInfo = [pscustomobject]@{
                    left = [double]$bounds.Left; top = [double]$bounds.Top
                    width = [double]$bounds.Width; height = [double]$bounds.Height
                }
            } catch { $boundsInfo = $null }
            if (-not (Test-MbPointWithinElementInfo -Info $boundsInfo -X $X -Y $Y -Tolerance 12.0)) { continue }
            $info = $null
            try { $info = Get-MbAutomationElementInfo -Element $element } catch { $info = $null }
            if (Test-MbPointWithinElementInfo -Info $info -X $X -Y $Y -Tolerance 12.0) {
                [void]$Candidates.Add($info)
            }
        }
    } catch {
        # UIAプロバイダーが子孫検索へ応答しないアプリでも、近傍点とフォールバックは使える。
    }
}

# UIAを公開しないキャンバス、リモートデスクトップ、独自描画アプリでは正確な矩形を
# 取得できない。その場合もクリック場所が分かるよう、小さな枠だけを残す。
# ウィンドウ全体へフォールバックしないため、以前の全画面赤枠は再発しない。
function New-MbClickPointTargetInfo {
    param(
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y,
        [AllowNull()]$Window
    )

    $width = 56.0
    $height = 36.0
    $left = $X - ($width / 2.0)
    $top = $Y - ($height / 2.0)
    if ($null -ne $Window -and [double]$Window.width -gt 0 -and [double]$Window.height -gt 0) {
        $minLeft = [double]$Window.left
        $minTop = [double]$Window.top
        $maxLeft = $minLeft + [double]$Window.width - $width
        $maxTop = $minTop + [double]$Window.height - $height
        $left = [Math]::Max($minLeft, [Math]::Min($maxLeft, $left))
        $top = [Math]::Max($minTop, [Math]::Min($maxTop, $top))
    }
    return [pscustomobject]@{
        name = ''; controlType = 'ControlType.ClickPoint'; automationId = ''; className = ''
        left = $left; top = $top; width = $width; height = $height
        anchorX = $X; anchorY = $Y
        isActionable = $true; isFallback = $true
    }
}

# UIAがクリック点を含む横長の親要素を返すことがある。候補自体はCopilotへ残すが、
# 既定の赤枠はクリック位置へ寄せ、誤った大枠を利用者へ確定表示しない。
function Select-MbRecordingPrimaryTarget {
    param(
        [AllowNull()]$Target,
        [AllowNull()]$PointTarget,
        [AllowNull()]$Window
    )

    if ($null -eq $Target) { return $PointTarget }
    if ($null -eq $PointTarget -or $null -eq $Window -or
        [double]$Window.width -le 0 -or [double]$Window.height -le 0) { return $Target }

    $targetType = if ($Target.PSObject.Properties.Name -contains 'controlType') { [string]$Target.controlType } else { '' }
    $wideParentTypes = @(
        'ControlType.SplitButton', 'ControlType.Header', 'ControlType.HeaderItem',
        'ControlType.Pane', 'ControlType.Group', 'ControlType.Custom', 'ControlType.Text'
    )
    $widthRatio = [double]$Target.width / [double]$Window.width
    $heightRatio = [double]$Target.height / [double]$Window.height
    $pointX = if ($PointTarget.PSObject.Properties.Name -contains 'anchorX') {
        [double]$PointTarget.anchorX
    } else { [double]$PointTarget.left + ([double]$PointTarget.width / 2.0) }
    $pointY = if ($PointTarget.PSObject.Properties.Name -contains 'anchorY') {
        [double]$PointTarget.anchorY
    } else { [double]$PointTarget.top + ([double]$PointTarget.height / 2.0) }
    # UIA-CACHE が隣のセルを返した実例がある。名前がもっともらしくても、実際の
    # クリック座標を含まない矩形は既定赤枠にしない。DOMも座標変換がずれた場合は同様。
    $nearbyOnly = $Target.PSObject.Properties.Name -contains 'isNearby' -and [bool]$Target.isNearby
    $missesClickPoint = -not (Test-MbPointWithinElementInfo -Info $Target -X $pointX -Y $pointY -Tolerance 1.0)
    # 数式バーの Edit のように、型に関係なく画面幅の大半を覆う細長い矩形も
    # 操作箇所を伝えないため、クリック点アンカーへ降格する。
    $oversizedFlatTarget = $widthRatio -ge 0.45 -and $heightRatio -le 0.16
    $wideParentTarget = $targetType -in $wideParentTypes -and $widthRatio -ge 0.20 -and $heightRatio -le 0.16
    $usePointAnchor = $nearbyOnly -or $missesClickPoint -or $oversizedFlatTarget -or $wideParentTarget
    if (-not $usePointAnchor) { return $Target }

    # 名前は文章化の手掛かりとして維持し、矩形と取得元だけをクリック点へ替える。
    # 直接観測した座標は確かなので中信頼、横長の親要素は比較用の低信頼候補へ降格する。
    $Target | Add-Member -NotePropertyName 'confidence' -NotePropertyValue 'low' -Force
    $Target | Add-Member -NotePropertyName 'isSuspicious' -NotePropertyValue $true -Force
    $pointConfidence = if ($nearbyOnly -or $missesClickPoint) { 'low' } else { 'medium' }
    $PointTarget | Add-Member -NotePropertyName 'confidence' -NotePropertyValue $pointConfidence -Force
    if (-not $nearbyOnly -and -not $missesClickPoint -and
        $Target.PSObject.Properties.Name -contains 'name') { $PointTarget.name = [string]$Target.name }
    return $PointTarget
}

function ConvertTo-MbNormalizedClickPoint {
    param(
        [Parameter(Mandatory = $true)]$Region,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y
    )

    if ([double]$Region.width -le 0 -or [double]$Region.height -le 0) { return $null }
    $normalizedX = ($X - [double]$Region.left) / [double]$Region.width
    $normalizedY = ($Y - [double]$Region.top) / [double]$Region.height
    if ($normalizedX -lt -0.02 -or $normalizedX -gt 1.02 -or $normalizedY -lt -0.02 -or $normalizedY -gt 1.02) { return $null }
    return [pscustomobject]@{
        x = [Math]::Round([Math]::Max(0.0, [Math]::Min(1.0, $normalizedX)), 6)
        y = [Math]::Round([Math]::Max(0.0, [Math]::Min(1.0, $normalizedY)), 6)
    }
}

# クリックした点にあるコントロールを返す。
#
# UIAのFromPointは、多くのアプリでは最も内側の要素を返すが、Chromium系ブラウザーでは
# ページ全体のDocument/Paneを返すことがある。その場合は点を含む子を下へ掘り、
# 文字要素が返った場合は親も調べる。最後は操作可能な最小要素だけを採用する。
function Get-MbUiaTargetAtPoint {
    param(
        [Parameter(Mandatory = $true)][int]$X,
        [Parameter(Mandatory = $true)][int]$Y,
        [AllowNull()]$Window = $null
    )

    if (-not (Initialize-MbRecorderUia)) {
        $msaaTarget = Get-MbMsaaTargetAtPoint -X $X -Y $Y -Window $Window
        if ($null -ne $msaaTarget) { return $msaaTarget }
        return (New-MbClickPointTargetInfo -X $X -Y $Y -Window $Window)
    }
    try {
        $candidates = New-Object System.Collections.ArrayList
        $root = $null
        $offsets = @(
            @(0, 0), @(-3, 0), @(3, 0), @(0, -3), @(0, 3),
            @(-6, -6), @(6, -6), @(-6, 6), @(6, 6),
            @(-12, 0), @(12, 0), @(0, -12), @(0, 12),
            @(-16, -16), @(16, -16), @(-16, 16), @(16, 16)
        )
        foreach ($offset in $offsets) {
            $samplePoint = New-Object System.Windows.Point -ArgumentList @(
                [double]($X + [int]$offset[0]), [double]($Y + [int]$offset[1])
            )
            $element = $null
            try { $element = [System.Windows.Automation.AutomationElement]::FromPoint($samplePoint) } catch { $element = $null }
            if ($null -eq $element) { continue }
            if ($null -eq $root) { $root = $element }
            Add-MbUiaPointCandidates -Candidates $candidates -Element $element -X $X -Y $Y
            $selected = Select-MbUiaTargetInfo -Candidates @($candidates) -X $X -Y $Y
            if ($null -ne $selected) { return $selected }
        }

        if ($null -ne $root) {
            # 最初の点がTextなどの葉でも、その配下だけを検索して終わらないよう、
            # クリック点を含むDocument/Paneまで検索根を引き上げる。
            $searchRoot = $root
            $node = $root
            $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
            for ($depth = 0; $depth -lt 12 -and $null -ne $node; $depth++) {
                $rootInfo = $null
                try { $rootInfo = Get-MbAutomationElementInfo -Element $node } catch { $rootInfo = $null }
                if ($null -ne $rootInfo -and [string]$rootInfo.controlType -in @(
                    'ControlType.Document', 'ControlType.Pane', 'ControlType.Group'
                )) { $searchRoot = $node }
                if ($null -ne $rootInfo -and [string]$rootInfo.controlType -eq 'ControlType.Window') { break }
                try { $node = $walker.GetParent($node) } catch { $node = $null }
            }
            Add-MbUiaInteractiveDescendantCandidates -Candidates $candidates -Root $searchRoot -X $X -Y $Y
            $selected = Select-MbUiaTargetInfo -Candidates @($candidates) -X $X -Y $Y
            if ($null -ne $selected) { return $selected }
        }

        # クリック後にフォーカスが移った要素は、FromPointがDocumentを返す画面でも
        # 実際の操作対象である可能性が高い。クリック点を含む場合だけ候補へ加える。
        $focused = Get-MbUiaFocusedElement
        if (Test-MbPointWithinElementInfo -Info $focused -X $X -Y $Y -Tolerance 12.0) {
            [void]$candidates.Add($focused)
            $selected = Select-MbUiaTargetInfo -Candidates @($candidates) -X $X -Y $Y
            if ($null -ne $selected) { return $selected }
        }

        # FromPoint由来の枝にポップアップや別アクセシビリティ枝が含まれない場合がある。
        # 最後に前面ウィンドウのHWNDからUIAルートを取り直して、全枝を1回だけ検索する。
        if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'handle' -and [long]$Window.handle -ne 0) {
            $windowRoot = $null
            try {
                $windowRoot = [System.Windows.Automation.AutomationElement]::FromHandle(
                    [IntPtr]::new([long]$Window.handle)
                )
            } catch { $windowRoot = $null }
            if ($null -ne $windowRoot) {
                Add-MbUiaInteractiveDescendantCandidates -Candidates $candidates -Root $windowRoot -X $X -Y $Y
                $selected = Select-MbUiaTargetInfo -Candidates @($candidates) -X $X -Y $Y
                if ($null -ne $selected) { return $selected }
            }
        }

        # UIAよりMSAAの方が正確なボタン/リンク矩形を公開する旧式アプリを補う。
        $msaaTarget = Get-MbMsaaTargetAtPoint -X $X -Y $Y -Window $Window
        if ($null -ne $msaaTarget) { return $msaaTarget }

        $namedTarget = Select-MbUiaNamedTargetInfo -Candidates @($candidates) -X $X -Y $Y -Window $Window
        if ($null -ne $namedTarget) { return $namedTarget }

        return (New-MbClickPointTargetInfo -X $X -Y $Y -Window $Window)
    } catch {
        # 応答しないアプリでも記録は止めず、クリック位置だけは小さな枠で残す。
        return (New-MbClickPointTargetInfo -X $X -Y $Y -Window $Window)
    }
}

# クリック前のカーソル下を問い合わせる。FromPointは論理ツリーの根に近い要素を返すことが
# あるため、親だけでなく必要ならウィンドウ全枝も調べる。重い検索は別プロセスへ隔離され、
# 画面取得とクリック検知の60Hzループを止めない。
function Get-MbUiaHoverTargetAtPoint {
    param(
        [Parameter(Mandatory = $true)][int]$X,
        [Parameter(Mandatory = $true)][int]$Y,
        [AllowNull()]$Window = $null
    )

    try {
        $target = Get-MbUiaTargetAtPoint -X $X -Y $Y -Window $Window
        if ($null -eq $target -or [string]$target.controlType -eq 'ControlType.ClickPoint') { return $null }
        return $target
    } catch {
        return $null
    }
}

function Write-MbUiaTargetCache {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)]$Value)

    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $backup = $Path + '.' + [guid]::NewGuid().ToString('N') + '.bak'
    try {
        [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 12 -Compress), (New-Object Text.UTF8Encoding($false)))
        foreach ($delay in @(0, 25, 50, 100, 200)) {
            if ([int]$delay -gt 0) { Start-Sleep -Milliseconds ([int]$delay) }
            try {
                if ([IO.File]::Exists($Path)) {
                    [IO.File]::Replace($temporary, $Path, $backup, $true)
                } else {
                    [IO.File]::Move($temporary, $Path)
                }
                return
            } catch [IO.IOException] {
                if ([int]$delay -eq 200) { throw }
            } catch [UnauthorizedAccessException] {
                if ([int]$delay -eq 200) { throw }
            }
        }
    } finally {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue
    }
}

function Get-MbUiaTargetFromCache {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$X,
        [Parameter(Mandatory = $true)][int]$Y,
        [AllowNull()]$Window,
        [int]$MaxAgeMs = 1800
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
        try {
            $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true, 4096, $false)
            try { $raw = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $stream.Dispose() }
        $cache = $raw | ConvertFrom-Json
        if ($null -eq $cache -or $null -eq $cache.target -or $null -eq $Window) { return $null }
        $updated = [DateTime]::Parse([string]$cache.updatedAtUtc).ToUniversalTime()
        $ageMs = ([DateTime]::UtcNow - $updated).TotalMilliseconds
        if ($ageMs -gt $MaxAgeMs) { return $null }
        $sameWindow = [long]$cache.windowHandle -eq [long]$Window.handle
        # 別HWNDの古い対象を新しい画面へ重ねない。操作前キャッシュを使う場合は、
        # 呼び出し側が操作前Windowと操作前画像を一組で渡す経路に限定する。
        if (-not $sameWindow) { return $null }
        if ($cache.PSObject.Properties.Name -contains 'window' -and $null -ne $cache.window) {
            $cachedTitle = if ($cache.window.PSObject.Properties.Name -contains 'title') { [string]$cache.window.title } else { '' }
            $currentTitle = if ($Window.PSObject.Properties.Name -contains 'title') { [string]$Window.title } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($cachedTitle) -and
                -not [string]::IsNullOrWhiteSpace($currentTitle) -and
                -not [string]::Equals($cachedTitle, $currentTitle, [StringComparison]::OrdinalIgnoreCase)) {
                return $null
            }
        }
        $distance = [Math]::Sqrt(
            [Math]::Pow(([double]$cache.cursorX - $X), 2.0) +
            [Math]::Pow(([double]$cache.cursorY - $Y), 2.0)
        )
        if ($distance -gt 36.0) { return $null }
        $target = $cache.target
        if (-not (Test-MbUsableElementInfo -Info $target)) { return $null }
        if ([string]$target.controlType -eq 'ControlType.ClickPoint') { return $null }
        if (-not (Test-MbPointWithinElementInfo -Info $target -X $X -Y $Y -Tolerance 1.0)) { return $null }
        $target | Add-Member -NotePropertyName 'provider' -NotePropertyValue 'UIA-CACHE' -Force
        if ($cache.PSObject.Properties.Name -contains 'window' -and $null -ne $cache.window) {
            $target | Add-Member -NotePropertyName 'captureWindow' -NotePropertyValue $cache.window -Force
        }
        return $target
    } catch {
        return $null
    }
}

# UI Automationはクリック後だと、画面遷移やフォルダー移動で元の要素が消えていることがある。
# 画面記録とは別プロセスでカーソル下を先に保持し、記録の60Hzループを止めない。
function Invoke-MbUiaTargetCacheLoop {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$StopPath,
        [AllowEmptyString()][string]$LogPath = '',
        [string[]]$IgnoreTitlePatterns = @(),
        [int]$PollIntervalMs = 55
    )

    Initialize-MbRecorderNative
    [void](Set-MbProcessDpiAware)
    [void](Initialize-MbRecorderUia)
    $lastX = [int]::MinValue
    $lastY = [int]::MinValue
    $lastHandle = 0L
    $lastChecked = [DateTime]::MinValue
    while (-not (Test-Path -LiteralPath $StopPath -PathType Leaf)) {
        try {
            $point = New-Object 'MbRecorderNative+POINT'
            [void][MbRecorderNative]::GetCursorPos([ref]$point)
            $window = Get-MbForegroundWindowInfo
            if (-not (Test-MbIgnoredWindow -Window $window -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                $elapsed = ([DateTime]::UtcNow - $lastChecked).TotalMilliseconds
                $moved = [Math]::Abs([int]$point.X - $lastX) -gt 2 -or [Math]::Abs([int]$point.Y - $lastY) -gt 2
                $windowChanged = [long]$window.handle -ne $lastHandle
                # UIAプロバイダーをマウス移動のたびに連打しない。動いている間は最大約10Hz、
                # 静止中も画面内容の変化に備えて450msごとに更新する。
                $movementRefresh = ($moved -or $windowChanged) -and $elapsed -ge 90
                if ($movementRefresh -or $elapsed -ge 450) {
                    $target = Get-MbUiaHoverTargetAtPoint -X ([int]$point.X) -Y ([int]$point.Y) -Window $window
                    Write-MbUiaTargetCache -Path $CachePath -Value ([pscustomobject]@{
                        updatedAtUtc = [DateTime]::UtcNow.ToString('o')
                        cursorX = [int]$point.X; cursorY = [int]$point.Y
                        windowHandle = [long]$window.handle; window = $window; target = $target
                    })
                    $lastX = [int]$point.X; $lastY = [int]$point.Y
                    $lastHandle = [long]$window.handle; $lastChecked = [DateTime]::UtcNow
                }
            }
        } catch {
            if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
                try {
                    $line = [DateTime]::UtcNow.ToString('o') + "`t" + $_.Exception.Message + [Environment]::NewLine
                    [IO.File]::AppendAllText($LogPath, $line, (New-Object Text.UTF8Encoding($false)))
                } catch { }
            }
        }
        Start-Sleep -Milliseconds $PollIntervalMs
    }
}

function Get-MbUiaFocusedElement {
    if (-not (Initialize-MbRecorderUia)) { return $null }
    try {
        $element = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($null -eq $element) { return $null }

        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $candidates = New-Object System.Collections.ArrayList
        $node = $element
        for ($depth = 0; $depth -lt 6 -and $null -ne $node; $depth++) {
            $info = $null
            try { $info = Get-MbAutomationElementInfo -Element $node } catch { $info = $null }
            if (Test-MbUsableElementInfo -Info $info) {
                [void]$candidates.Add($info)
            }
            if ($null -ne $info -and [string]$info.controlType -eq 'ControlType.Window') { break }
            try { $node = $walker.GetParent($node) } catch { $node = $null }
        }

        # 入力欄を特定できないときにDocument全体を操作対象にしない。
        foreach ($candidate in @($candidates)) {
            if (Test-MbRecorderInteractiveElementInfo -Info $candidate) { return $candidate }
        }
        return $null
    } catch {
        return $null
    }
}

# ---------------------------------------------------------------------
# 画面の取得
# ---------------------------------------------------------------------

function Get-MbWindowInfo {
    param([IntPtr]$Handle = [IntPtr]::Zero)

    Initialize-MbRecorderNative
    $handle = $Handle
    if ($handle -eq [IntPtr]::Zero) { return $null }
    $rect = [MbRecorderNative]::GetVisualWindowRect($handle)
    [uint32]$processId = 0
    $processName = ''
    try {
        [void][MbRecorderNative]::GetWindowThreadProcessId($handle, [ref]$processId)
        if ($processId -gt 0) { $processName = [Diagnostics.Process]::GetProcessById([int]$processId).ProcessName }
    } catch { $processName = '' }
    return [pscustomobject]@{
        handle = [long]$handle.ToInt64()
        processId = [int]$processId
        processName = [string]$processName
        title  = [MbRecorderNative]::GetWindowTitle($handle)
        class  = [MbRecorderNative]::GetWindowClass($handle)
        left   = [int]$rect.Left
        top    = [int]$rect.Top
        width  = [int]($rect.Right - $rect.Left)
        height = [int]($rect.Bottom - $rect.Top)
    }
}

function Get-MbForegroundWindowInfo {
    Initialize-MbRecorderNative
    return Get-MbWindowInfo -Handle ([MbRecorderNative]::GetForegroundWindow())
}

function Get-MbVirtualScreenBounds {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    try {
        $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
        return [pscustomobject]@{ left = [int]$bounds.X; top = [int]$bounds.Y; width = [int]$bounds.Width; height = [int]$bounds.Height }
    } catch {
        $primary = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        return [pscustomobject]@{ left = [int]$primary.X; top = [int]$primary.Y; width = [int]$primary.Width; height = [int]$primary.Height }
    }
}

# 取り込む範囲を決める。
# 前面ウィンドウだけだと、そこから外へ出るメニューやドロップダウンが切れる。
# 操作対象の矩形と合わせた範囲を取り、画面の外へははみ出さないようにする。
function Get-MbCaptureRegion {
    param([AllowNull()]$Window, [AllowNull()]$Target)

    $screen = Get-MbVirtualScreenBounds
    $left = $screen.left
    $top = $screen.top
    $right = $screen.left + $screen.width
    $bottom = $screen.top + $screen.height

    if ($null -ne $Window -and [int]$Window.width -gt 0 -and [int]$Window.height -gt 0) {
        $left = [int]$Window.left
        $top = [int]$Window.top
        $right = [int]$Window.left + [int]$Window.width
        $bottom = [int]$Window.top + [int]$Window.height
        if ($null -ne $Target) {
            $left = [Math]::Min($left, [int][Math]::Floor([double]$Target.left))
            $top = [Math]::Min($top, [int][Math]::Floor([double]$Target.top))
            $right = [Math]::Max($right, [int][Math]::Ceiling([double]$Target.left + [double]$Target.width))
            $bottom = [Math]::Max($bottom, [int][Math]::Ceiling([double]$Target.top + [double]$Target.height))
        }
        $left = [Math]::Max($left, $screen.left)
        $top = [Math]::Max($top, $screen.top)
        $right = [Math]::Min($right, $screen.left + $screen.width)
        $bottom = [Math]::Min($bottom, $screen.top + $screen.height)
    }

    $width = $right - $left
    $height = $bottom - $top
    if ($width -lt 16 -or $height -lt 16) {
        return [pscustomobject]@{ left = $screen.left; top = $screen.top; width = $screen.width; height = $screen.height }
    }
    return [pscustomobject]@{ left = $left; top = $top; width = $width; height = $height }
}

# 画面全体をメモリへ取る。
#
# 呼ぶ順番が重要。UI Automation の問い合わせは数十〜数百ミリ秒かかるので、
# 先にUIAを引いてから撮ると、押した結果すでに変化した画面が写ってしまう。
# マニュアルに要るのは「押す直前の画面」なので、まずここで画面を確保し、
# 切り出しはあとから行う。
function Copy-MbScreenBitmap {
    Initialize-MbRecorderNative
    $screen = Get-MbVirtualScreenBounds
    $bitmap = New-Object Drawing.Bitmap -ArgumentList @([int]$screen.width, [int]$screen.height)
    $graphics = $null
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $size = New-Object Drawing.Size -ArgumentList @([int]$screen.width, [int]$screen.height)
        $graphics.CopyFromScreen([int]$screen.left, [int]$screen.top, 0, 0, $size)
    } catch {
        $bitmap.Dispose()
        throw
    } finally {
        if ($null -ne $graphics) { try { $graphics.Dispose() } catch { } }
    }
    return [pscustomobject]@{ bitmap = $bitmap; origin = $screen }
}

function Save-MbBitmapRegion {
    param(
        [Parameter(Mandatory = $true)]$Capture,
        [Parameter(Mandatory = $true)]$Region,
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$MaxEdge = 1600,
        [long]$Quality = 88
    )

    Initialize-MbRecorderNative
    $source = $Capture.bitmap
    $origin = $Capture.origin
    # 画面の外を切り出そうとすると例外になる。取り込んだ画像の中へ収める。
    $left = [Math]::Max(0, [int]$Region.left - [int]$origin.left)
    $top = [Math]::Max(0, [int]$Region.top - [int]$origin.top)
    $width = [Math]::Min([int]$Region.width, $source.Width - $left)
    $height = [Math]::Min([int]$Region.height, $source.Height - $top)
    if ($width -lt 16 -or $height -lt 16) {
        $left = 0; $top = 0; $width = $source.Width; $height = $source.Height
    }

    $cropped = $null
    $scaled = $null
    try {
        $rectangle = New-Object Drawing.Rectangle -ArgumentList @($left, $top, $width, $height)
        $cropped = $source.Clone($rectangle, $source.PixelFormat)

        $output = $cropped
        $longest = [Math]::Max($width, $height)
        if ($MaxEdge -gt 0 -and $longest -gt $MaxEdge) {
            $scale = $MaxEdge / [double]$longest
            $scaledWidth = [int][Math]::Round($width * $scale)
            $scaledHeight = [int][Math]::Round($height * $scale)
            $scaled = New-Object Drawing.Bitmap -ArgumentList @($scaledWidth, $scaledHeight, [Drawing.Imaging.PixelFormat]::Format24bppRgb)
            $resizeGraphics = $null
            try {
                $resizeGraphics = [Drawing.Graphics]::FromImage($scaled)
                $resizeGraphics.CompositingQuality = [Drawing.Drawing2D.CompositingQuality]::HighQuality
                $resizeGraphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $resizeGraphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $resizeGraphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::HighQuality
                $resizeGraphics.DrawImage($cropped, 0, 0, $scaledWidth, $scaledHeight)
            } finally {
                if ($null -ne $resizeGraphics) { $resizeGraphics.Dispose() }
            }
            $output = $scaled
        }

        $encoder = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
        $parameters = New-Object Drawing.Imaging.EncoderParameters -ArgumentList 1
        $parameters.Param[0] = New-Object Drawing.Imaging.EncoderParameter -ArgumentList @([Drawing.Imaging.Encoder]::Quality, $Quality)
        try {
            $output.Save($Path, $encoder, $parameters)
        } finally {
            $parameters.Dispose()
        }
        return [pscustomobject]@{ left = ($left + [int]$origin.left); top = ($top + [int]$origin.top); width = $width; height = $height }
    } finally {
        foreach ($disposable in @($scaled, $cropped)) {
            if ($null -ne $disposable) { try { $disposable.Dispose() } catch { } }
        }
    }
}

# 操作対象の矩形を、取り込んだ画像の中での0〜1の位置へ直す。
function ConvertTo-MbRegionRect {
    param([Parameter(Mandatory = $true)]$Region, [AllowNull()]$Target)

    if ($null -eq $Target) { return $null }
    $width = [double]$Region.width
    $height = [double]$Region.height
    if ($width -le 0 -or $height -le 0) { return $null }

    # double のオーバーロードを明示し、0〜1の座標を整数へ丸めない。
    $clamp = { param([double]$Value) [Math]::Round([Math]::Min(1.0, [Math]::Max(0.0, $Value)), 6) }
    $x1 = & $clamp ((([double]$Target.left) - [double]$Region.left) / $width)
    $y1 = & $clamp ((([double]$Target.top) - [double]$Region.top) / $height)
    $x2 = & $clamp ((([double]$Target.left + [double]$Target.width) - [double]$Region.left) / $width)
    $y2 = & $clamp ((([double]$Target.top + [double]$Target.height) - [double]$Region.top) / $height)
    if (($x2 - $x1) -le 0.0 -or ($y2 - $y1) -le 0.0) { return $null }
    # 正規化値だけで判定すると、4K画面では6px程度ある正しい細線ボタンまで消える。
    # 物理サイズが実用的なら高解像度でも残し、実体が2px未満の潰れた矩形だけを除く。
    if ((($x2 - $x1) -lt 0.002 -and [double]$Target.width -lt 2.0) -or
        (($y2 - $y1) -lt 0.002 -and [double]$Target.height -lt 2.0)) { return $null }
    # ページ全体のDocument/Paneを掴んだ場合は赤枠にしない。
    # ブラウザーの枠やサイドバーを除くと96%未満になるため、従来の閾値では防げなかった。
    $rectWidth = $x2 - $x1
    $rectHeight = $y2 - $y1
    if (($rectWidth -gt 0.82 -and $rectHeight -gt 0.82) -or
        (($rectWidth * $rectHeight) -gt 0.72)) { return $null }
    return [pscustomobject]@{ x1 = $x1; y1 = $y1; x2 = $x2; y2 = $y2 }
}

# ---------------------------------------------------------------------
# 記録の本体
# ---------------------------------------------------------------------

$script:MbVkLeftButton = 0x01
$script:MbVkRightButton = 0x02

# 入力があったことだけを知るために見るキー。文字そのものは読まない。
function Get-MbWatchedTypingKeys {
    $keys = New-Object System.Collections.Generic.List[int]
    for ($vk = 0x30; $vk -le 0x39; $vk++) { [void]$keys.Add($vk) }   # 0-9
    for ($vk = 0x41; $vk -le 0x5A; $vk++) { [void]$keys.Add($vk) }   # A-Z
    for ($vk = 0x60; $vk -le 0x69; $vk++) { [void]$keys.Add($vk) }   # テンキー
    foreach ($vk in @(0x08, 0x20, 0xBA, 0xBB, 0xBC, 0xBD, 0xBE, 0xBF, 0xC0, 0xDB, 0xDC, 0xDD, 0xDE)) { [void]$keys.Add($vk) }
    return $keys.ToArray()
}

# Enter/Tabは文字入力ではなく「現在の入力を確定するキー」として別に監視する。
# 確定後の計算結果で入力中の最終画面を上書きしないため、typingKeysへは混ぜない。
function Get-MbWatchedCommitKeys {
    return [int[]]@(0x0D, 0x09) # Enter, Tab
}

# GetAsyncKeyState の上位ビットは「現在押されている」、下位ビットは
# 「前回の確認後に一度でも押された」を表す。タッチパッドのタップなどは
# 16ms の巡回間隔より短く、上位ビットだけでは押下を丸ごと見失うことがある。
function Test-MbAsyncKeyStateDown {
    param([int]$State)
    return (($State -band 0x8000) -ne 0)
}

function Test-MbAsyncKeyStatePressed {
    param([int]$State)
    return (($State -band 0x0001) -ne 0)
}

# Ctrl/Alt/Windows キーを伴うショートカットは、通常は文字入力ではなく
# フォーカス移動やコマンド実行である。ただし貼り付け・切り取り・Undo/Redo は
# 入力欄の表示内容を変えるため、入力手順として残す。
function Test-MbTextChangingShortcutKey {
    param([int]$VirtualKey)
    return ($VirtualKey -in @(0x08, 0x56, 0x58, 0x59, 0x5A)) # BackSpace, V, X, Y, Z
}

function Test-MbIgnoredWindow {
    param([AllowNull()]$Window, [string[]]$IgnoreTitlePatterns = @())
    if ($null -eq $Window) { return $true }
    if ($script:MbRecorderIgnoredClasses -contains [string]$Window.class) { return $true }
    $title = [string]$Window.title
    $pageTitle = [regex]::Replace(
        $title,
        '\s+[-—]\s+(?:Microsoft\s+Edge|Edge|Google\s+Chrome|Chrome)$',
        '',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase
    ).Trim()
    foreach ($pattern in $IgnoreTitlePatterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        if ([string]::Equals($pageTitle, $pattern.Trim(), [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Write-MbRecordingEvent {
    param(
        [Parameter(Mandatory = $true)][string]$EventsPath,
        [Parameter(Mandatory = $true)][hashtable]$Record
    )
    $line = ([pscustomobject]$Record | ConvertTo-Json -Depth 8 -Compress)
    [IO.File]::AppendAllText($EventsPath, $line + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
}

# 確認画面用の events.jsonl は取消時に書き直されるため、監査用の証拠は
# 別の追記専用台帳へ保存する。取り消した事実も新しい判断レコードとして追記する。
function Write-MbRecordingLedgerRecord {
    param(
        [AllowEmptyString()][string]$LedgerPath = '',
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Record
    )
    if ([string]::IsNullOrWhiteSpace($LedgerPath)) { return }
    $line = ([pscustomobject]$Record | ConvertTo-Json -Depth 10 -Compress)
    [IO.File]::AppendAllText($LedgerPath, $line + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
}

# 記録ワーカー自身が安全なタイミングで直前の1件を取り消す。
# events.jsonl と対応画像を同時に戻し、次の操作では同じ連番を再利用する。
function Remove-MbLastRecordingEvent {
    param(
        [Parameter(Mandatory = $true)][string]$EventsPath,
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [AllowEmptyString()][string]$LedgerPath = '',
        [AllowEmptyString()][string]$JobId = ''
    )

    if (-not (Test-Path -LiteralPath $EventsPath -PathType Leaf)) {
        return [pscustomobject]@{ removed = $false; count = 0; lastTarget = '' }
    }
    [string[]]$lines = @([IO.File]::ReadAllLines($EventsPath, [Text.Encoding]::UTF8) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) { return [pscustomobject]@{ removed = $false; count = 0; lastTarget = '' } }

    $removedIndex = $lines.Count
    $removedRecord = $null
    try {
        $removedRecord = $lines[$lines.Count - 1] | ConvertFrom-Json
        if ($removedRecord.PSObject.Properties.Name -contains 'index') { $removedIndex = [int]$removedRecord.index }
    } catch { }
    [string[]]$remaining = if ($lines.Count -gt 1) { @($lines[0..($lines.Count - 2)]) } else { @() }
    $temporary = $EventsPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $backup = $EventsPath + '.' + [guid]::NewGuid().ToString('N') + '.bak'
    try {
        [IO.File]::WriteAllLines($temporary, $remaining, (New-Object Text.UTF8Encoding($false)))

        # ここで例外が記録ループまで抜けると capture-end が書かれず、取り込み側は
        # 「証拠が揃っていない」として記録セッションを丸ごと捨てる。ウイルス対策ソフトが
        # 一瞬 events.jsonl を開いただけで全損しないよう、Write-MbRecordingStatus と
        # 同じ再試行で差し替える。
        $delaysMs = @(0, 25, 50, 100, 200, 400, 800)
        for ($attempt = 0; $attempt -lt $delaysMs.Count; $attempt++) {
            if ([int]$delaysMs[$attempt] -gt 0) {
                Start-Sleep -Milliseconds ([int]$delaysMs[$attempt])
            }
            try {
                [IO.File]::Replace($temporary, $EventsPath, $backup, $true)
                break
            } catch [IO.IOException] {
                if ($attempt -eq ($delaysMs.Count - 1)) { throw }
            } catch [UnauthorizedAccessException] {
                if ($attempt -eq ($delaysMs.Count - 1)) { throw }
            }
        }
    } finally {
        foreach ($temporaryFile in @($temporary, $backup)) {
            Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
        }
    }

    foreach ($suffix in @('.jpg', '-result.jpg')) {
        $imagePath = Join-Path $EventsDirectory (('event-{0:d3}{1}' -f $removedIndex, $suffix))
        Remove-Item -LiteralPath $imagePath -Force -ErrorAction SilentlyContinue
    }
    $removedEvidenceId = if ($null -ne $removedRecord -and
        $removedRecord.PSObject.Properties.Name -contains 'evidenceId') { [string]$removedRecord.evidenceId } else { '' }
    Write-MbRecordingLedgerRecord -LedgerPath $LedgerPath -Record ([ordered]@{
        recordType = 'decision'
        id = 'decision-' + [guid]::NewGuid().ToString('N')
        sessionId = $JobId
        action = 'undo'
        evidenceIds = $(if ([string]::IsNullOrWhiteSpace($removedEvidenceId)) { @() } else { @($removedEvidenceId) })
        workingEventIndex = $removedIndex
        recordedAt = [DateTime]::UtcNow.ToString('o')
        reason = '利用者が記録中に直前の操作を取り消しました。元の操作証拠は保持します。'
    })
    $lastTarget = ''
    if ($remaining.Count -gt 0) {
        try { $lastTarget = [string](($remaining[$remaining.Count - 1] | ConvertFrom-Json).targetName) } catch { }
    }
    return [pscustomobject]@{ removed = $true; count = $remaining.Count; lastTarget = $lastTarget }
}

# UI Automation の Name は、Webページのコンテナなどで画面中の文章をまとめて返すことがある。
# 取り込み側の操作対象は200文字までなので、記録時点から表示・進捗・保存のすべてを同じ長さへ
# 揃える。入力・右クリックの補足も上限の内側に収める。
function ConvertTo-MbRecorderTargetName {
    param(
        [AllowNull()][object]$Value,
        [string]$Suffix = '',
        [int]$MaxLength = 200
    )

    if ($MaxLength -le 0) { return '' }
    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    $text = ($text.Replace("`r`n", ' ').Replace("`r", ' ').Replace("`n", ' ') -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    if ($Suffix.Length -gt $MaxLength) { $Suffix = $Suffix.Substring(0, $MaxLength) }
    $available = $MaxLength - $Suffix.Length
    if ($text.Length -gt $available) {
        if ($available -le 1) {
            $text = if ($available -eq 1) { '…' } else { '' }
        } else {
            $text = $text.Substring(0, $available - 1).TrimEnd() + '…'
        }
    }
    return $text + $Suffix
}

function Write-MbRecordingStatus {
    param(
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$JobId,
        [Parameter(Mandatory = $true)][string]$State,
        [int]$Count = 0,
        [string]$Message = '',
        [string]$LastTarget = '',
        [string]$UndoRequestId = '',
        [string]$ResultRequestId = ''
    )
    if (-not [string]::IsNullOrWhiteSpace($UndoRequestId)) {
        $script:MbRecorderLastUndoRequestId = $UndoRequestId
    }
    if (-not [string]::IsNullOrWhiteSpace($ResultRequestId)) {
        $script:MbRecorderLastResultRequestId = $ResultRequestId
    }
    $status = [pscustomobject]@{
        jobId = $JobId; state = $State; count = $Count; message = $Message
        lastTarget = $LastTarget; updatedAt = [DateTime]::UtcNow.ToString('o')
        undoRequestId = [string]$script:MbRecorderLastUndoRequestId
        resultRequestId = [string]$script:MbRecorderLastResultRequestId
        captureCompleteness = [string]$script:MbRecorderCaptureCompleteness
        captureWarning = [string]$script:MbRecorderCaptureWarning
    }
    $temporary = $StatusPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $backup = $StatusPath + '.' + [guid]::NewGuid().ToString('N') + '.bak'
    try {
        [IO.File]::WriteAllText($temporary, ($status | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))

        # 画面側が進捗を読んでいる瞬間や、ウイルス対策ソフトが短時間ファイルを
        # 開いた瞬間でも記録全体を止めない。まず同一フォルダー内で書き終え、
        # 完成したファイルを原子的に差し替える。
        $delaysMs = @(0, 25, 50, 100, 200, 400, 800)
        for ($attempt = 0; $attempt -lt $delaysMs.Count; $attempt++) {
            if ([int]$delaysMs[$attempt] -gt 0) {
                Start-Sleep -Milliseconds ([int]$delaysMs[$attempt])
            }
            try {
                if ([IO.File]::Exists($StatusPath)) {
                    [IO.File]::Replace($temporary, $StatusPath, $backup, $true)
                } else {
                    [IO.File]::Move($temporary, $StatusPath)
                }
                return
            } catch [IO.IOException] {
                if ($attempt -eq ($delaysMs.Count - 1)) { throw }
            } catch [UnauthorizedAccessException] {
                if ($attempt -eq ($delaysMs.Count - 1)) { throw }
            }
        }
    } finally {
        foreach ($temporaryFile in @($temporary, $backup)) {
            Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# EdgeのDOMスナップショットを、画面キャプチャと同じ物理ピクセル座標へ変換する。
# ブラウザー枠の幅やマルチモニターの原点を推測せず、DOMイベント時のclient座標と
# 実際のマウス座標との差だけを使うため、表示倍率やモニター位置が変わってもずれにくい。
function ConvertFrom-MbDomSnapshotTarget {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y
    )

    try {
        if ($null -eq $Snapshot.rect) { return $null }
        $dpr = [double]$Snapshot.dpr
        if ([double]::IsNaN($dpr) -or [double]::IsInfinity($dpr) -or $dpr -lt 0.5 -or $dpr -gt 4.0) { return $null }
        $clientX = [double]$Snapshot.clientX
        $clientY = [double]$Snapshot.clientY
        $anchorX = $X
        $anchorY = $Y
        if ($Snapshot.PSObject.Properties.Name -contains 'screenX' -and
            $Snapshot.PSObject.Properties.Name -contains 'screenY') {
            $snapshotScreenX = [double]$Snapshot.screenX
            $snapshotScreenY = [double]$Snapshot.screenY
            $anchorDistance = [Math]::Sqrt(
                [Math]::Pow(($snapshotScreenX - $X), 2.0) + [Math]::Pow(($snapshotScreenY - $Y), 2.0)
            )
            if (-not [double]::IsNaN($anchorDistance) -and -not [double]::IsInfinity($anchorDistance) -and
                $anchorDistance -le 64.0) {
                $anchorX = $snapshotScreenX
                $anchorY = $snapshotScreenY
            }
        }
        $width = [double]$Snapshot.rect.width * $dpr
        $height = [double]$Snapshot.rect.height * $dpr
        $left = $anchorX + (([double]$Snapshot.rect.left - $clientX) * $dpr)
        $top = $anchorY + (([double]$Snapshot.rect.top - $clientY) * $dpr)
        foreach ($number in @($clientX, $clientY, $width, $height, $left, $top)) {
            if ([double]::IsNaN([double]$number) -or [double]::IsInfinity([double]$number)) { return $null }
        }
        if ($width -le 1.0 -or $height -le 1.0 -or $width -gt 12000 -or $height -gt 12000) { return $null }
        if ($anchorX -lt ($left - 16.0) -or $anchorX -gt ($left + $width + 16.0) -or
            $anchorY -lt ($top - 16.0) -or $anchorY -gt ($top + $height + 16.0)) { return $null }

        $role = ([string]$Snapshot.role).ToLowerInvariant()
        $tag = ([string]$Snapshot.tag).ToLowerInvariant()
        $type = ([string]$Snapshot.type).ToLowerInvariant()
        $controlType = switch ($role) {
            'button'    { 'ControlType.Button'; break }
            'link'      { 'ControlType.Hyperlink'; break }
            'menuitem'  { 'ControlType.MenuItem'; break }
            'tab'       { 'ControlType.TabItem'; break }
            'checkbox'  { 'ControlType.CheckBox'; break }
            'radio'     { 'ControlType.RadioButton'; break }
            'option'    { 'ControlType.ListItem'; break }
            'combobox'  { 'ControlType.ComboBox'; break }
            'textbox'   { 'ControlType.Edit'; break }
            default {
                switch ($tag) {
                    'button'   { 'ControlType.Button'; break }
                    'a'        { 'ControlType.Hyperlink'; break }
                    'select'   { 'ControlType.ComboBox'; break }
                    'textarea' { 'ControlType.Edit'; break }
                    'summary'  { 'ControlType.Button'; break }
                    'input' {
                        switch ($type) {
                            'checkbox' { 'ControlType.CheckBox'; break }
                            'radio'    { 'ControlType.RadioButton'; break }
                            { $_ -in @('button', 'submit', 'reset', 'image') } { 'ControlType.Button'; break }
                            default { 'ControlType.Edit' }
                        }
                        break
                    }
                    default { 'ControlType.Custom' }
                }
            }
        }
        if ($Snapshot.PSObject.Properties.Name -contains 'editable' -and [bool]$Snapshot.editable) {
            $controlType = 'ControlType.Edit'
        }
        return [pscustomobject]@{
            name = ConvertTo-MbRecorderTargetName -Value $Snapshot.name
            controlType = $controlType
            automationId = ''
            className = $tag
            inputType = $type
            left = $left; top = $top; width = $width; height = $height
            isActionable = $true; isFallback = $false
            # DOMも座標変換・イベント経路・入れ子要素の選択を誤ることがある。
            # 異種ソースとの一致を確認するまでは確定扱いにしない。
            provider = 'DOM'; confidence = 'medium'
        }
    } catch { return $null }
}

function Get-MbDomTargetFromCache {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$X,
        [Parameter(Mandatory = $true)][int]$Y,
        [AllowNull()]$Window
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        # 専用Edge以外のウィンドウを操作したとき、別ウィンドウのDOMを誤適用しない。
        if ($null -eq $Window -or [string]$Window.class -notlike 'Chrome_WidgetWin*') { return $null }
        $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
        try {
            $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true, 4096, $false)
            try { $raw = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $stream.Dispose() }
        $cache = $raw | ConvertFrom-Json
        if ($null -eq $cache -or $null -eq $cache.page) { return $null }
        $updated = [DateTime]::Parse([string]$cache.updatedAtUtc).ToUniversalTime()
        if (([DateTime]::UtcNow - $updated).TotalMilliseconds -gt 2500) { return $null }
        $pageTitle = ([string]$cache.page.title).Trim()
        $windowTitle = if ($null -ne $Window) { [string]$Window.title } else { '' }
        $snapshotTitle = ''

        $snapshot = $null
        $epoch = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
        $nowMs = [long](([DateTime]::UtcNow - $epoch).TotalMilliseconds)
        if ($cache.page.PSObject.Properties.Name -contains 'pointer' -and $null -ne $cache.page.pointer) {
            # pointerdownは画面遷移前の対象を保持する。ただし昔のクリックは再利用しない。
            $pointerMs = [long]$cache.page.pointer.at
            if ($pointerMs -gt 0 -and ($nowMs - $pointerMs) -ge -500 -and ($nowMs - $pointerMs) -le 2500) {
                $snapshot = $cache.page.pointer
                if ($snapshot.PSObject.Properties.Name -contains 'pageTitle' -and
                    -not [string]::IsNullOrWhiteSpace([string]$snapshot.pageTitle)) {
                    $snapshotTitle = ([string]$snapshot.pageTitle).Trim()
                }
            }
        }
        if ($null -eq $snapshot -and $cache.page.PSObject.Properties.Name -contains 'hover' -and $null -ne $cache.page.hover) {
            $hoverMs = [long]$cache.page.hover.at
            if ($hoverMs -gt 0 -and ($nowMs - $hoverMs) -ge -500 -and ($nowMs - $hoverMs) -le 1500) {
                $snapshot = $cache.page.hover
            }
        }
        if ($null -eq $snapshot) { return $null }
        # pointerdown時の物理座標があれば、監視ワーカーが後で読んだカーソル位置より優先する。
        $expectedX = [double]$cache.cursorX
        $expectedY = [double]$cache.cursorY
        if ($snapshot.PSObject.Properties.Name -contains 'screenX' -and
            $snapshot.PSObject.Properties.Name -contains 'screenY') {
            $expectedX = [double]$snapshot.screenX
            $expectedY = [double]$snapshot.screenY
        }
        $cursorDistance = [Math]::Sqrt(
            [Math]::Pow(($expectedX - $X), 2.0) + [Math]::Pow(($expectedY - $Y), 2.0)
        )
        if ($cursorDistance -gt 64.0) { return $null }
        # pointerdownの直後にページが遷移すると、cache.page.titleは遷移後、
        # pointer.pageTitleは遷移前になる。現在ページのタイトルだけで許可すると、
        # 遷移前の矩形を遷移後の画像へ描いてしまうため、snapshot側を必ず優先する。
        $titleMatches = $false
        if (-not [string]::IsNullOrWhiteSpace($snapshotTitle)) {
            $titleMatches = $windowTitle.IndexOf($snapshotTitle, [StringComparison]::OrdinalIgnoreCase) -ge 0
        } elseif (-not [string]::IsNullOrWhiteSpace($pageTitle)) {
            $titleMatches = $windowTitle.IndexOf($pageTitle, [StringComparison]::OrdinalIgnoreCase) -ge 0
        } else {
            $titleMatches = $true
        }
        if (-not $titleMatches) {
            return $null
        }
        return (ConvertFrom-MbDomSnapshotTarget -Snapshot $snapshot -X $X -Y $Y)
    } catch { return $null }
}

# DOM、UI Automation、クリック位置を同じ形式へ揃える。採用済みの対象も
# 「観測候補の1つ」に過ぎないため、Copilotが別候補か none を選べるように残す。
function Get-MbRecorderTargetEvidence {
    param([AllowNull()]$Target)

    if ($null -eq $Target) { return $null }
    $source = if ($Target.PSObject.Properties.Name -contains 'provider' -and
        -not [string]::IsNullOrWhiteSpace([string]$Target.provider)) {
        [string]$Target.provider
    } elseif ($Target.PSObject.Properties.Name -contains 'isFallback' -and [bool]$Target.isFallback) {
        'click-point'
    } else { 'UIA' }
    $confidence = if ($Target.PSObject.Properties.Name -contains 'confidence' -and
        [string]$Target.confidence -in @('high', 'medium', 'low')) {
        [string]$Target.confidence
    } elseif (($Target.PSObject.Properties.Name -contains 'isFallback' -and [bool]$Target.isFallback) -or
        ($Target.PSObject.Properties.Name -contains 'isInferred' -and [bool]$Target.isInferred)) {
        'low'
    } else { 'medium' }
    return [pscustomobject]@{ source = $source; confidence = $confidence }
}

function ConvertTo-MbRecordingTargetCandidates {
    param(
        [Parameter(Mandatory = $true)]$Region,
        [AllowNull()]$SelectedTarget,
        [AllowEmptyCollection()][object[]]$Targets = @(),
        [int]$Maximum = 4
    )

    $ordered = New-Object System.Collections.ArrayList
    if ($null -ne $SelectedTarget) { [void]$ordered.Add($SelectedTarget) }
    foreach ($target in @($Targets)) {
        if ($null -ne $target) { [void]$ordered.Add($target) }
    }

    $result = New-Object System.Collections.ArrayList
    foreach ($target in @($ordered)) {
        if ($result.Count -ge [Math]::Max(1, $Maximum)) { break }
        $rect = ConvertTo-MbRegionRect -Region $Region -Target $target
        if ($null -eq $rect) { continue }
        $evidence = Get-MbRecorderTargetEvidence -Target $target
        if ($null -eq $evidence) { continue }
        $label = ConvertTo-MbRecorderTargetName -Value $(if ($target.PSObject.Properties.Name -contains 'name') { $target.name } else { '' })
        $targetType = if ($target.PSObject.Properties.Name -contains 'controlType') { [string]$target.controlType } else { '' }

        # 同じ取得元が同じ名前・矩形を返しただけなら候補を水増ししない。
        $duplicate = $false
        foreach ($existing in @($result)) {
            if ([string]$existing.source -ne [string]$evidence.source -or
                [string]$existing.label -ne $label -or [string]$existing.targetType -ne $targetType) { continue }
            if ([Math]::Abs([double]$existing.rect.x1 - [double]$rect.x1) -lt 0.003 -and
                [Math]::Abs([double]$existing.rect.y1 - [double]$rect.y1) -lt 0.003 -and
                [Math]::Abs([double]$existing.rect.x2 - [double]$rect.x2) -lt 0.003 -and
                [Math]::Abs([double]$existing.rect.y2 - [double]$rect.y2) -lt 0.003) {
                $duplicate = $true
                break
            }
        }
        if ($duplicate) { continue }

        $safeSource = ([string]$evidence.source).ToLowerInvariant() -replace '[^a-z0-9]+', '-'
        $safeSource = $safeSource.Trim('-')
        if ([string]::IsNullOrWhiteSpace($safeSource)) { $safeSource = 'observed' }
        [void]$result.Add([pscustomobject]@{
            id = ($safeSource + '-' + ($result.Count + 1))
            source = [string]$evidence.source
            confidence = [string]$evidence.confidence
            label = $label
            targetType = $targetType
            rect = $rect
        })
    }
    return @($result)
}

# 1件ぶんの操作を記録する。押す直前の画面を先に確保してからUIAを引く。
function Save-MbRecordingEvent {
    param(
        [Parameter(Mandatory = $true)]$Capture,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][int]$ElapsedMs,
        [Parameter(Mandatory = $true)][string]$Kind,
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [Parameter(Mandatory = $true)][string]$EventsPath,
        [AllowEmptyString()][string]$EvidenceDirectory = '',
        [AllowEmptyString()][string]$LedgerPath = '',
        [AllowEmptyString()][string]$JobId = '',
        [AllowNull()]$Target,
        [AllowEmptyCollection()][object[]]$TargetCandidates = @(),
        [AllowNull()]$Window,
        [int]$ScreenX = [int]::MinValue,
        [int]$ScreenY = [int]::MinValue,
        [int]$MaxEdge = 2560,
        [long]$Quality = 94
    )

    $region = Get-MbCaptureRegion -Window $Window -Target $Target
    $fileName = ('event-{0:d3}.jpg' -f $Index)
    $saved = Save-MbBitmapRegion -Capture $Capture -Region $region -Path (Join-Path $EventsDirectory $fileName) `
        -MaxEdge $MaxEdge -Quality $Quality
    # 実際に切り出せた範囲で正規化する。画面の端では要求した範囲より狭くなる。
    $rect = ConvertTo-MbRegionRect -Region $saved -Target $Target

    $windowTitle = ''
    if ($null -ne $Window) { $windowTitle = [string]$Window.title }
    $targetName = ''
    $targetType = ''
    if ($null -ne $Target) {
        $targetName = ConvertTo-MbRecorderTargetName -Value $Target.name
        $targetType = [string]$Target.controlType
    }
    $record = @{
        evidenceId  = 'evidence-' + [guid]::NewGuid().ToString('N')
        index       = $Index
        kind        = $Kind
        timeMs      = $ElapsedMs
        image       = $fileName
        windowTitle = $windowTitle
        processName = $(if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'processName') { [string]$Window.processName } else { '' })
        windowClass = $(if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'class') { [string]$Window.class } else { '' })
        targetName  = $targetName
        targetType  = $targetType
        rect        = $rect
    }
    if ($ScreenX -ne [int]::MinValue -and $ScreenY -ne [int]::MinValue) {
        $clickPoint = ConvertTo-MbNormalizedClickPoint -Region $saved -X $ScreenX -Y $ScreenY
        if ($null -ne $clickPoint) {
            $record.clickPoint = [pscustomobject]@{
                x = [double]$clickPoint.x; y = [double]$clickPoint.y
                screenX = $ScreenX; screenY = $ScreenY
            }
            $record.captureRegion = [pscustomobject]@{
                left = [int]$saved.left; top = [int]$saved.top
                width = [int]$saved.width; height = [int]$saved.height
            }
        }
    }
    if ($null -ne $Target) {
        $evidence = Get-MbRecorderTargetEvidence -Target $Target
        $record.targetSource = [string]$evidence.source
        $record.confidence = [string]$evidence.confidence
    }
    $candidates = @(ConvertTo-MbRecordingTargetCandidates -Region $saved -SelectedTarget $Target -Targets $TargetCandidates -Maximum 4)
    if ($candidates.Count -gt 0) {
        $record.targetCandidates = $candidates
        $record.targetCandidateId = [string]$candidates[0].id
    }
    Write-MbRecordingEvent -EventsPath $EventsPath -Record $record
    if (-not [string]::IsNullOrWhiteSpace($EvidenceDirectory)) {
        if (-not (Test-Path -LiteralPath $EvidenceDirectory -PathType Container)) {
            [void](New-Item -ItemType Directory -Path $EvidenceDirectory -Force)
        }
        $evidenceFileName = ([string]$record.evidenceId + '.jpg')
        [IO.File]::Copy((Join-Path $EventsDirectory $fileName), (Join-Path $EvidenceDirectory $evidenceFileName), $false)
        $record.evidenceImage = $evidenceFileName
    }
    Write-MbRecordingLedgerRecord -LedgerPath $LedgerPath -Record ([ordered]@{
        recordType = 'operation'
        id = [string]$record.evidenceId
        sessionId = $JobId
        workingEventIndex = $Index
        kind = $Kind
        timeMs = $ElapsedMs
        recordedAt = [DateTime]::UtcNow.ToString('o')
        image = $(if ($record.ContainsKey('evidenceImage')) { [string]$record.evidenceImage } else { '' })
        windowTitle = $windowTitle
        processName = [string]$record.processName
        targetName = $targetName
        targetType = $targetType
        targetSource = $(if ($record.ContainsKey('targetSource')) { [string]$record.targetSource } else { '' })
        confidence = $(if ($record.ContainsKey('confidence')) { [string]$record.confidence } else { '' })
        clickPoint = $(if ($record.ContainsKey('clickPoint')) { $record.clickPoint } else { $null })
        rect = $rect
    })
    return $record
}

# クリック後の安定した画面を、クリック前画像とは別に保存する。
# JSONLを書き直さず、同じイベント番号の -result 画像を置くことで、記録中の
# クラッシュでも既存イベントを壊さず、読み出し側が後から関連付けられる。
function Save-MbRecordingResultImage {
    param(
        [Parameter(Mandatory = $true)]$Capture,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [AllowNull()]$Window,
        [int]$MaxEdge = 2560,
        [long]$Quality = 94
    )

    $region = Get-MbCaptureRegion -Window $Window -Target $null
    $fileName = ('event-{0:d3}-result.jpg' -f $Index)
    [void](Save-MbBitmapRegion -Capture $Capture -Region $region `
        -Path (Join-Path $EventsDirectory $fileName) -MaxEdge $MaxEdge -Quality $Quality)
    return $fileName
}

# AIが手順を構成できるよう、クリック検出とは独立した時系列フレームを残す。
# 画像とメタデータは同じ連番にし、Copilotが返した F00001 を元画像へ戻せるようにする。
function Save-MbRecordingTimelineFrame {
    param(
        [Parameter(Mandatory = $true)]$Capture,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][int]$ElapsedMs,
        [Parameter(Mandatory = $true)][string]$FramesDirectory,
        [Parameter(Mandatory = $true)][string]$FramesPath,
        [AllowNull()]$Window,
        [AllowEmptyString()][string]$Role = '',
        [AllowEmptyString()][string]$EvidenceKind = '',
        [int]$MaxEdge = 1600,
        [long]$Quality = 88
    )

    $frameId = ('F{0:d5}' -f $Index)
    $fileName = ('frame-{0:d5}.jpg' -f $Index)
    $region = Get-MbCaptureRegion -Window $Window -Target $null
    [void](Save-MbBitmapRegion -Capture $Capture -Region $region -Path (Join-Path $FramesDirectory $fileName) `
        -MaxEdge $MaxEdge -Quality $Quality)
    $record = [ordered]@{
        id = $frameId
        index = $Index
        timeMs = $ElapsedMs
        image = $fileName
        windowTitle = $(if ($null -ne $Window) { [string]$Window.title } else { '' })
        processName = $(if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'processName') { [string]$Window.processName } else { '' })
        windowClass = $(if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'class') { [string]$Window.class } else { '' })
    }
    if (-not [string]::IsNullOrWhiteSpace($Role)) { $record['role'] = $Role }
    if (-not [string]::IsNullOrWhiteSpace($EvidenceKind)) { $record['evidenceKind'] = $EvidenceKind }
    Write-MbRecordingEvent -EventsPath $FramesPath -Record $record
    return [pscustomobject]$record
}

# クリック用に保存済みの操作直前画像を、AI時系列にも無劣化で参照できるよう複製する。
# 画面を二度撮影・圧縮しないため、記録ループを止めずに次操作直前の安定状態を残せる。
function Save-MbRecordingTimelineEventFrame {
    param(
        [Parameter(Mandatory = $true)]$EventRecord,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][int]$ElapsedMs,
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [Parameter(Mandatory = $true)][string]$FramesDirectory,
        [Parameter(Mandatory = $true)][string]$FramesPath,
        [AllowNull()]$Window
    )
    $sourceName = [string]$EventRecord.image
    if ($sourceName -notmatch '^event-\d{3}\.jpg$' -or [IO.Path]::GetFileName($sourceName) -ne $sourceName) {
        throw 'Invalid event image name.'
    }
    $sourcePath = Join-Path $EventsDirectory $sourceName
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw 'Event image was not saved.' }
    $frameId = ('F{0:d5}' -f $Index)
    $fileName = ('frame-{0:d5}.jpg' -f $Index)
    [IO.File]::Copy($sourcePath, (Join-Path $FramesDirectory $fileName), $false)
    $record = [ordered]@{
        id = $frameId
        index = $Index
        timeMs = $ElapsedMs
        image = $fileName
        windowTitle = $(if ($null -ne $Window) { [string]$Window.title } else { '' })
        processName = $(if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'processName') { [string]$Window.processName } else { '' })
        windowClass = $(if ($null -ne $Window -and $Window.PSObject.Properties.Name -contains 'class') { [string]$Window.class } else { '' })
        role = 'click-evidence'
        evidenceEventId = [int]$EventRecord.index
    }
    Write-MbRecordingEvent -EventsPath $FramesPath -Record $record
    return [pscustomobject]$record
}

function Resolve-MbTypingEventElapsedMs {
    param(
        [int]$CurrentElapsedMs,
        [bool]$Clicked,
        [int]$ClickElapsedMs
    )
    if ($Clicked -and $ClickElapsedMs -gt 0) { return [Math]::Max(0, $ClickElapsedMs - 1) }
    return [Math]::Max(0, $CurrentElapsedMs)
}

function Test-MbQueuedKeyboardContinuation {
    param(
        [AllowNull()]$CurrentActivity,
        [AllowNull()]$NextActivity,
        [AllowNull()]$NextMouseClick,
        [int]$TypingIdleMs,
        [long]$TimestampFrequency
    )
    if ($null -eq $CurrentActivity -or $TimestampFrequency -le 0) {
        return $false
    }
    if ([int]$CurrentActivity.Kind -ne 1) { return $false }
    $nextTimestamp = [long]::MaxValue
    if ($null -ne $NextActivity -and [long]$NextActivity.Timestamp -lt $nextTimestamp) {
        $nextTimestamp = [long]$NextActivity.Timestamp
    }
    if ($null -ne $NextMouseClick -and [long]$NextMouseClick.Timestamp -lt $nextTimestamp) {
        $nextTimestamp = [long]$NextMouseClick.Timestamp
    }
    if ($nextTimestamp -eq [long]::MaxValue) { return $false }
    $gapTicks = $nextTimestamp - [long]$CurrentActivity.Timestamp
    if ($gapTicks -lt 0) { return $false }
    $gapMs = ($gapTicks * 1000.0) / [double]$TimestampFrequency
    return $gapMs -le [double]$TypingIdleMs
}

function Invoke-MbManualResultRequest {
    param(
        [Parameter(Mandatory = $true)][string]$RequestPath,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [string[]]$IgnoreTitlePatterns = @(),
        [Parameter(Mandatory = $true)][hashtable]$ProcessedRequests
    )

    if (-not (Test-Path -LiteralPath $RequestPath -PathType Leaf)) { return $null }
    $requestId = ''
    $saved = $false
    $capture = $null
    try {
        $request = Get-Content -Raw -LiteralPath $RequestPath -ErrorAction Stop | ConvertFrom-Json
        $requestId = [string]$request.requestId
        $windowHandle = [long]$request.windowHandle
        if (-not [string]::IsNullOrWhiteSpace($requestId) -and $ProcessedRequests.ContainsKey($requestId)) {
            $saved = [bool]$ProcessedRequests[$requestId]
        } elseif ($Index -gt 0 -and $windowHandle -gt 0) {
            $requestedWindow = Get-MbWindowInfo -Handle ([IntPtr]::new($windowHandle))
            $foregroundWindow = Get-MbForegroundWindowInfo
            $sameForegroundWindow = $null -ne $requestedWindow -and $null -ne $foregroundWindow -and
                [long]$requestedWindow.handle -eq [long]$foregroundWindow.handle
            if ($sameForegroundWindow -and
                -not (Test-MbIgnoredWindow -Window $requestedWindow -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                $capture = Copy-MbScreenBitmap
                $windowAfterCapture = Get-MbForegroundWindowInfo
                if ($null -ne $windowAfterCapture -and
                    [long]$windowAfterCapture.handle -eq [long]$requestedWindow.handle) {
                    [void](Save-MbRecordingResultImage -Capture $capture -Index $Index `
                        -EventsDirectory $EventsDirectory -Window $requestedWindow)
                    $saved = $true
                }
            }
        }
    } catch { }
    finally {
        if ($null -ne $capture) { try { $capture.bitmap.Dispose() } catch { } }
        Remove-Item -LiteralPath $RequestPath -Force -ErrorAction SilentlyContinue
    }
    if ([string]::IsNullOrWhiteSpace($requestId)) { return $null }
    if (-not $ProcessedRequests.ContainsKey($requestId)) { $ProcessedRequests[$requestId] = $saved }
    return [pscustomobject]@{ requestId = $requestId; saved = $saved }
}

function Invoke-MbRecordingLoop {
    param(
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [Parameter(Mandatory = $true)][string]$EventsPath,
        [AllowEmptyString()][string]$EvidenceDirectory = '',
        [AllowEmptyString()][string]$LedgerPath = '',
        [AllowEmptyString()][string]$FramesDirectory = '',
        [AllowEmptyString()][string]$FramesPath = '',
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$StopPath,
        [AllowEmptyString()][string]$PausePath = '',
        [AllowEmptyString()][string]$UndoPath = '',
        [AllowEmptyString()][string]$ManualResultPath = '',
        [Parameter(Mandatory = $true)][string]$JobId,
        [AllowEmptyString()][string]$DomTargetPath = '',
        [AllowEmptyString()][string]$UiaTargetPath = '',
        [string[]]$IgnoreTitlePatterns = @(),
        [int]$PollIntervalMs = 16,
        [int]$TypingIdleMs = 1200,
        [ValidateRange(200, 3000)][int]$ResultCaptureDelayMs = 700,
        [int]$MaxEvents = 300,
        [int]$MaxMinutes = 60
    )

    Initialize-MbRecorderNative
    [void](Set-MbProcessDpiAware)
    [void](Initialize-MbRecorderUia)

    $typingKeys = Get-MbWatchedTypingKeys
    $typingCommitKeys = Get-MbWatchedCommitKeys
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $diagnosticPath = $StatusPath + '.diagnostic.log'
    $recordingStartedTimestamp = [MbRecorderNative]::GetTimestamp()
    # 画像取得や対象解析から独立した専用スレッドでクリックを受ける。
    # 開始できない制限環境では、従来の GetAsyncKeyState へ自動的に戻る。
    $mouseHookActive = $false
    try { $mouseHookActive = [MbRecorderNative]::StartMouseHook() } catch { $mouseHookActive = $false }
    $keyboardHookActive = $false
    if ($mouseHookActive) {
        try { $keyboardHookActive = [bool][MbRecorderNative]::KeyboardHookAvailable } catch { $keyboardHookActive = $false }
    }
    $captureCompleteness = if ($mouseHookActive -and $keyboardHookActive) { 'no-known-gaps' } else { 'known-gaps' }
    $captureWarning = if (-not $mouseHookActive) {
        'クリックを確実に受け取る機能を開始できませんでした。抜けた操作がないか確認してください。'
    } elseif (-not $keyboardHookActive) {
        '入力活動を確実に受け取る機能を開始できませんでした。入力手順に抜けがないか確認してください。'
    } else { '' }
    $script:MbRecorderCaptureCompleteness = $captureCompleteness
    $script:MbRecorderCaptureWarning = $captureWarning
    $lastDroppedMouseClicks = 0L
    $lastDroppedKeyboardActivities = 0L
    Write-MbRecordingLedgerRecord -LedgerPath $LedgerPath -Record ([ordered]@{
        recordType = 'capture-start'; formatVersion = 2; sessionId = $JobId
        recordedAt = [DateTime]::UtcNow.ToString('o')
        mouseHook = [bool]$mouseHookActive; keyboardHook = [bool]$keyboardHookActive
        completeness = $captureCompleteness; warning = $captureWarning
    })
    try {
        [IO.File]::AppendAllText($diagnosticPath,
            ([DateTime]::UtcNow.ToString('o') + ' mouse-hook=' + [string]$mouseHookActive +
                ' keyboard-hook=' + [string]$keyboardHookActive + [Environment]::NewLine),
            [Text.UTF8Encoding]::new($false))
    } catch { }
    $index = 0
    $frameIndex = 0
    $lastFrameAtMs = -1000
    $timelineFrameIntervalMs = 500
    # 入力中だけは最終文字を落とさない頻度で保持する。キーの内容は読まず、
    # 画面に描画された状態だけを一時的に更新する。
    $typingEvidenceIntervalMs = 20
    $lastStatusMs = -1000
    $lastTarget = ''

    $leftState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkLeftButton)
    $rightState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkRightButton)
    $leftWasDown = Test-MbAsyncKeyStateDown -State $leftState
    $rightWasDown = Test-MbAsyncKeyStateDown -State $rightState
    $typingActive = $false
    $typingField = $null
    $typingWindow = $null
    $typingCapture = $null
    $typingCaptureAtMs = 0
    $typingCaptureDueAtMs = 0
    $typingEvidenceKind = ''
    $lastTypingMs = 0
    $lastDomTarget = $null
    $lastDomWindowHandle = 0L
    # Excel cells do not always expose the focused element while typing. Keep the
    # latest clicked editable target so the input event can retain its cell name
    # and rectangle when UI Automation only returns a generic Group.
    $lastInteractionTarget = $null
    $lastInteractionWindowHandle = 0L
    $lastInteractionAtMs = -10000
    # Edgeはpointerdownだけ先に取得できても、直後の画面遷移がスクリーンショットより
    # 速いことがある。専用Edgeだけ直前の安定画面を保持し、画像・タイトル・DOMを
    # 同じ時点へ揃える。デスクトップ記録は従来どおりクリック時だけ撮影する。
    $preClickCapture = $null
    $preClickWindow = $null
    $preClickCaptureAtMs = -1000
    $preClickCaptureAttemptAtMs = -1000
    $preClickCaptureIntervalMs = 80
    $preClickCaptureMaxAgeMs = 180
    $pendingLeftClick = $false
    $pendingRightClick = $false
    $pendingResultIndex = 0
    $pendingResultWindowHandle = 0L
    $pendingResultDueAtMs = 0
    $suppressedTypingKeys = New-Object 'System.Collections.Generic.HashSet[int]'
    $paused = $false
    # 常駐パネルは応答が遅いと同じIDを再送する。同一IDの破壊的操作は1回だけ実行する。
    $processedUndoRequests = @{}
    $processedResultRequests = @{}

    Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count 0 -Message '操作を記録しています'

    while ($true) {
        if ($mouseHookActive) {
            $droppedMouseClicks = try { [long][MbRecorderNative]::DroppedMouseClicks } catch { 0L }
            $droppedKeyboardActivities = try { [long][MbRecorderNative]::DroppedKeyboardActivities } catch { 0L }
            if ($droppedMouseClicks -gt $lastDroppedMouseClicks -or
                $droppedKeyboardActivities -gt $lastDroppedKeyboardActivities) {
                $script:MbRecorderCaptureCompleteness = 'known-gaps'
                $script:MbRecorderCaptureWarning = '操作が短時間に集中し、一部を記録できなかった可能性があります。手順の抜けを確認してください。'
                Write-MbRecordingLedgerRecord -LedgerPath $LedgerPath -Record ([ordered]@{
                    recordType = 'capture-gap'; formatVersion = 2; sessionId = $JobId
                    recordedAt = [DateTime]::UtcNow.ToString('o')
                    droppedMouseClicks = [Math]::Max(0L, $droppedMouseClicks - $lastDroppedMouseClicks)
                    droppedKeyboardActivities = [Math]::Max(0L, $droppedKeyboardActivities - $lastDroppedKeyboardActivities)
                    reason = 'capture-queue-overflow'
                })
                $lastDroppedMouseClicks = $droppedMouseClicks
                $lastDroppedKeyboardActivities = $droppedKeyboardActivities
            }
        }
        $pauseRequested = -not [string]::IsNullOrWhiteSpace($PausePath) -and
            (Test-Path -LiteralPath $PausePath -PathType Leaf)
        if ($pauseRequested) {
            if (-not $paused) {
                $paused = $true
                if ($mouseHookActive) {
                    [MbRecorderNative]::ClearMouseClicks()
                    [MbRecorderNative]::ClearKeyboardActivities()
                }
                $pendingResultIndex = 0
                $pendingResultWindowHandle = 0L
                $typingActive = $false
                if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
                $typingCapture = $null; $typingField = $null; $typingWindow = $null
                $typingCaptureDueAtMs = 0
                $typingEvidenceKind = ''
                $lastInteractionTarget = $null; $lastInteractionWindowHandle = 0L; $lastInteractionAtMs = -10000
                if ($null -ne $preClickCapture) { try { $preClickCapture.bitmap.Dispose() } catch { } }
                $preClickCapture = $null; $preClickWindow = $null; $preClickCaptureAtMs = -1000
            }
            if (-not [string]::IsNullOrWhiteSpace($UndoPath) -and (Test-Path -LiteralPath $UndoPath -PathType Leaf)) {
                $undoRequestId = try { [string](Get-Content -Raw -LiteralPath $UndoPath -ErrorAction Stop) } catch { '' }
                Remove-Item -LiteralPath $UndoPath -Force -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($undoRequestId)) {
                    $undoRemoved = $false
                    if ($processedUndoRequests.ContainsKey($undoRequestId)) {
                        $undoRemoved = [bool]$processedUndoRequests[$undoRequestId]
                    } else {
                        $undo = Remove-MbLastRecordingEvent -EventsPath $EventsPath -EventsDirectory $EventsDirectory `
                            -LedgerPath $LedgerPath -JobId $JobId
                        $undoRemoved = [bool]$undo.removed
                        $processedUndoRequests[$undoRequestId] = $undoRemoved
                        if ($undoRemoved) {
                            $index = [int]$undo.count; $lastTarget = [string]$undo.lastTarget
                        }
                    }
                    Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'paused' -Count $index `
                        -Message $(if ($undoRemoved) { '直前の操作を取り消しました' } else { '取り消せる操作がありませんでした' }) `
                        -LastTarget $lastTarget -UndoRequestId $undoRequestId
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($ManualResultPath)) {
                $manualResult = Invoke-MbManualResultRequest -RequestPath $ManualResultPath -Index $index `
                    -EventsDirectory $EventsDirectory -IgnoreTitlePatterns $IgnoreTitlePatterns `
                    -ProcessedRequests $processedResultRequests
                if ($null -ne $manualResult) {
                    if ([bool]$manualResult.saved -and $pendingResultIndex -eq $index) {
                        $pendingResultIndex = 0; $pendingResultWindowHandle = 0L
                    }
                    Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'paused' -Count $index `
                        -Message $(if ([bool]$manualResult.saved) { '直前の手順へ結果画面を追加しました' } else { '結果画面を追加できませんでした' }) `
                        -LastTarget $lastTarget -ResultRequestId ([string]$manualResult.requestId)
                }
            }
            # 終了要求と取消が重なった場合も、取消を状態へ反映してから停止する。
            if ((Test-Path -LiteralPath $StopPath -PathType Leaf) -or $index -ge $MaxEvents -or
                $watch.Elapsed.TotalMinutes -ge $MaxMinutes) { break }
            if (([int]$watch.ElapsedMilliseconds - $lastStatusMs) -ge 400) {
                $lastStatusMs = [int]$watch.ElapsedMilliseconds
                Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'paused' -Count $index `
                    -Message '記録を一時停止しています' -LastTarget $lastTarget
            }
            Start-Sleep -Milliseconds 50
            continue
        } elseif ($paused) {
            # 休止中に押したボタンを再開直後の操作として拾わない。
            if ($mouseHookActive) {
                [MbRecorderNative]::ClearMouseClicks()
                [MbRecorderNative]::ClearKeyboardActivities()
            }
            $leftState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkLeftButton)
            $rightState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkRightButton)
            $leftWasDown = Test-MbAsyncKeyStateDown -State $leftState
            $rightWasDown = Test-MbAsyncKeyStateDown -State $rightState
            $paused = $false
            Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count $index `
                -Message '操作の記録を再開しました' -LastTarget $lastTarget
        }

        if (-not [string]::IsNullOrWhiteSpace($UndoPath) -and (Test-Path -LiteralPath $UndoPath -PathType Leaf)) {
            $undoRequestId = try { [string](Get-Content -Raw -LiteralPath $UndoPath -ErrorAction Stop) } catch { '' }
            Remove-Item -LiteralPath $UndoPath -Force -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace($undoRequestId)) {
                $undoRemoved = $false
                if ($processedUndoRequests.ContainsKey($undoRequestId)) {
                    $undoRemoved = [bool]$processedUndoRequests[$undoRequestId]
                } else {
                    $undo = Remove-MbLastRecordingEvent -EventsPath $EventsPath -EventsDirectory $EventsDirectory `
                        -LedgerPath $LedgerPath -JobId $JobId
                    $undoRemoved = [bool]$undo.removed
                    $processedUndoRequests[$undoRequestId] = $undoRemoved
                    if ($undoRemoved) {
                        $index = [int]$undo.count; $lastTarget = [string]$undo.lastTarget
                        $pendingResultIndex = 0
                        $pendingResultWindowHandle = 0L
                    }
                }
                Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count $index `
                    -Message $(if ($undoRemoved) { '直前の操作を取り消しました' } else { '取り消せる操作がありませんでした' }) `
                    -LastTarget $lastTarget -UndoRequestId $undoRequestId
            }
        }

        # 常駐パネルの「結果画面を追加」は、対象アプリへフォーカスを戻してから
        # ウィンドウハンドル付きで要求される。このパネル自体を結果画像へ混ぜず、
        # 直前の操作へ利用者が見ている確定画面を関連付ける。
        if (-not [string]::IsNullOrWhiteSpace($ManualResultPath)) {
            $manualResult = Invoke-MbManualResultRequest -RequestPath $ManualResultPath -Index $index `
                -EventsDirectory $EventsDirectory -IgnoreTitlePatterns $IgnoreTitlePatterns `
                -ProcessedRequests $processedResultRequests
            if ($null -ne $manualResult) {
                if ([bool]$manualResult.saved -and $pendingResultIndex -eq $index) {
                    $pendingResultIndex = 0; $pendingResultWindowHandle = 0L
                }
                Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count $index `
                    -Message $(if ([bool]$manualResult.saved) { '直前の手順へ結果画面を追加しました' } else { '結果画面を追加できませんでした' }) `
                    -LastTarget $lastTarget -ResultRequestId ([string]$manualResult.requestId)
            }
        }

        # 常駐パネルの未処理要求を先に取り込み、「終了と同時に押した削除・結果追加」を落とさない。
        if ((Test-Path -LiteralPath $StopPath -PathType Leaf) -or $index -ge $MaxEvents -or
            $watch.Elapsed.TotalMinutes -ge $MaxMinutes) { break }

        $hookClick = $null
        $hookTextActivity = $false
        $hookCommitActivity = $false
        $hookKeyboardWindowHandle = 0L
        $hookKeyboardElapsedMs = -1
        $hookKeyboardContinuationQueued = $false
        $skipTypingPoll = $false
        if ($mouseHookActive) {
            # マウスとキーボードを別キューのまま真偽値へ潰すと、重い撮影中に
            # A入力→Bクリック→B入力が溜まった際、2つの入力をBへ誤結合する。
            # 先頭時刻を比較して1件ずつ処理し、記録時のHWNDも保持する。
            $nextMouse = [MbRecorderNative]::PeekMouseClick()
            $nextKeyboard = [MbRecorderNative]::PeekKeyboardActivity()
            $keyboardIsNext = $null -ne $nextKeyboard -and
                ($null -eq $nextMouse -or [long]$nextKeyboard.Timestamp -le [long]$nextMouse.Timestamp)
            $keyboardWindowChanged = $keyboardIsNext -and $typingActive -and $null -ne $typingWindow -and
                [int]$nextKeyboard.Kind -eq 1 -and [long]$nextKeyboard.WindowHandle -ne 0 -and
                [long]$nextKeyboard.WindowHandle -ne [long]$typingWindow.handle
            if ($keyboardWindowChanged) {
                # 新しいウィンドウの入力は次巡回までキューに残し、現在の入力を先に確定する。
                $hookCommitActivity = $true
                $hookKeyboardWindowHandle = [long]$typingWindow.handle
                $hookKeyboardElapsedMs = [int][Math]::Max(0, [Math]::Round(
                    (([long]$nextKeyboard.Timestamp - $recordingStartedTimestamp) * 1000.0) /
                    [double][MbRecorderNative]::TimestampFrequency) - 1)
                $skipTypingPoll = $true
            } elseif ($keyboardIsNext) {
                $keyboardActivity = [MbRecorderNative]::DequeueKeyboardActivity()
                $hookKeyboardWindowHandle = [long]$keyboardActivity.WindowHandle
                $hookKeyboardElapsedMs = [int][Math]::Max(0, [Math]::Round(
                    (([long]$keyboardActivity.Timestamp - $recordingStartedTimestamp) * 1000.0) /
                    [double][MbRecorderNative]::TimestampFrequency))
                if ([int]$keyboardActivity.Kind -eq 1) { $hookTextActivity = $true }
                elseif ([int]$keyboardActivity.Kind -eq 2) { $hookCommitActivity = $true }
                if ($hookTextActivity) {
                    # 画面取得中に複数キーが滞留しても、古いhook時刻だけを見て
                    # 1文字目を即idle確定しない。次の同一窓キー/確定操作が実時間で
                    # 連続していれば、次巡回まで同じ入力として保持する。
                    $hookKeyboardContinuationQueued = Test-MbQueuedKeyboardContinuation `
                        -CurrentActivity $keyboardActivity `
                        -NextActivity ([MbRecorderNative]::PeekKeyboardActivity()) `
                        -NextMouseClick ([MbRecorderNative]::PeekMouseClick()) `
                        -TypingIdleMs $TypingIdleMs `
                        -TimestampFrequency ([long][MbRecorderNative]::TimestampFrequency)
                }
            } else {
                $hookClick = [MbRecorderNative]::DequeueMouseClick()
            }
        }
        $leftState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkLeftButton)
        $rightState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkRightButton)
        $leftDown = Test-MbAsyncKeyStateDown -State $leftState
        $rightDown = Test-MbAsyncKeyStateDown -State $rightState
        if ($mouseHookActive) {
            $leftClicked = $null -ne $hookClick -and [int]$hookClick.Message -eq 0x0201
            $rightClicked = $null -ne $hookClick -and [int]$hookClick.Message -eq 0x0204
        } else {
            $leftClicked = $pendingLeftClick -or (Test-MbAsyncKeyStatePressed -State $leftState) -or ($leftDown -and -not $leftWasDown)
            $rightClicked = $pendingRightClick -or (Test-MbAsyncKeyStatePressed -State $rightState) -or ($rightDown -and -not $rightWasDown)
        }
        $pendingLeftClick = $false
        $pendingRightClick = $false
        $clicked = ($leftClicked -or $rightClicked)
        $clickElapsedMs = if ($null -ne $hookClick) {
            [int][Math]::Max(0, [Math]::Round(
                (([long]$hookClick.Timestamp - $recordingStartedTimestamp) * 1000.0) /
                [double][MbRecorderNative]::TimestampFrequency))
        } else { [int]$watch.ElapsedMilliseconds }
        $leftWasDown = $leftDown
        $rightWasDown = $rightDown

        # Ctrl+L / Ctrl+F / Alt+英字などを「入力」として数えない。修飾キーと
        # 文字キーが巡回の間に離されても下位ビットで同じショートカットと判断する。
        $commandModifierActive = $false
        foreach ($modifierVk in @(0x11, 0x12, 0x5B, 0x5C)) { # Ctrl, Alt, 左右Windows
            $modifierState = [int][MbRecorderNative]::GetAsyncKeyState($modifierVk)
            if ((Test-MbAsyncKeyStateDown -State $modifierState) -or
                (Test-MbAsyncKeyStatePressed -State $modifierState)) {
                $commandModifierActive = $true
            }
        }

        $typingNow = $hookTextActivity
        $typingPressed = $hookTextActivity
        $pastePressed = $false
        foreach ($vk in $typingKeys) {
            if ($keyboardHookActive) { break }
            $keyState = [int][MbRecorderNative]::GetAsyncKeyState($vk)
            $keyPressed = Test-MbAsyncKeyStatePressed -State $keyState
            $keyDown = Test-MbAsyncKeyStateDown -State $keyState
            $keyActive = $keyDown -or $keyPressed

            if ($commandModifierActive -and $keyActive -and -not (Test-MbTextChangingShortcutKey -VirtualKey $vk)) {
                [void]$suppressedTypingKeys.Add($vk)
            }
            if ($suppressedTypingKeys.Contains($vk)) {
                if (-not $keyActive) { [void]$suppressedTypingKeys.Remove($vk) }
                continue
            }

            if ($keyActive) {
                $typingNow = $true
                if ($commandModifierActive -and $vk -eq 0x56) { $pastePressed = $true }
            }
            if ($keyPressed) { $typingPressed = $true }
        }
        if ($skipTypingPoll) { $typingNow = $false; $typingPressed = $false }
        $typingCommitPressed = $hookCommitActivity
        foreach ($vk in $typingCommitKeys) {
            if ($keyboardHookActive) { break }
            $commitState = [int][MbRecorderNative]::GetAsyncKeyState($vk)
            if (Test-MbAsyncKeyStatePressed -State $commitState) { $typingCommitPressed = $true }
        }
        if ($hookCommitActivity -and $typingActive -and $null -ne $typingWindow -and
            $hookKeyboardWindowHandle -ne 0 -and
            $hookKeyboardWindowHandle -ne [long]$typingWindow.handle) {
            $typingCommitPressed = $false
        }
        if ($typingPressed -and $typingActive) {
            # 押下検出と同時に撮ると、アプリが最終文字を描画する直前になる。
            # キー内容は読まず、通常のキー解放後になる120msまで描画を待つ。
            # 押下中に撮るとExcelでは最終文字だけ反映前の画面になる。
            $typingCaptureDueAtMs = [int]$watch.ElapsedMilliseconds + 120
        }
        $typingStartedThisIteration = $false
        if ($typingNow) {
            if (-not $typingActive) {
                $typingActive = $true
                $typingStartedThisIteration = $true
                $typingEvidenceKind = ''
                # クリック後画像の待機中に入力が始まった場合、その入力済み画面を
                # 直前クリックの結果として結び付けない。結果なしの方が因果を捏造しない。
                if ($pendingResultIndex -gt 0) {
                    $pendingResultIndex = 0
                    $pendingResultWindowHandle = 0L
                }
                # キーそのものは読まない。入力中の画面だけを保持し、表示された文字を
                # 残すか隠すかは、取り込み後の画像編集（黒塗り）で利用者が決める。
                try {
                    $candidateWindow = if ($hookTextActivity -and $hookKeyboardWindowHandle -ne 0) {
                        Get-MbWindowInfo -Handle ([IntPtr]::new($hookKeyboardWindowHandle))
                    } else { Get-MbForegroundWindowInfo }
                    if (-not (Test-MbIgnoredWindow -Window $candidateWindow -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                        $typingCapture = Copy-MbScreenBitmap
                        $typingCaptureAtMs = [int]$watch.ElapsedMilliseconds
                        $typingWindow = $candidateWindow
                        # 貼り付けは1キー操作なので「次の文字キー」が来ない。開始時の画像は
                        # 貼り付け反映前になり得るため、初回キーだけでも描画後に1度更新する。
                        if ($typingPressed) {
                            $typingCaptureDueAtMs = [int]$watch.ElapsedMilliseconds + 180
                        }
                        # 入力開始時にも同期UIAを呼ばない。ExcelやChromiumではFocusedElementの
                        # 取得だけで数秒止まり、その間の別セル入力をまとめて失っていた。
                        # 直前クリックの非同期キャッシュが十分なアンカーになり、取れない場合も
                        # 入力済み画面と時刻はCopilotが判別できるため、対象なしのまま残す。
                        $typingField = $null
                        if ($null -ne $lastDomTarget -and
                            [string]$lastDomTarget.controlType -eq 'ControlType.Edit' -and
                            $null -ne $candidateWindow -and [long]$candidateWindow.handle -eq $lastDomWindowHandle) {
                            $typingField = $lastDomTarget
                        } elseif ($null -ne $lastInteractionTarget -and $null -ne $candidateWindow -and
                            [long]$candidateWindow.handle -eq $lastInteractionWindowHandle -and
                            ([int]$watch.ElapsedMilliseconds - $lastInteractionAtMs) -le 5000 -and
                            [string]$lastInteractionTarget.controlType -in @('ControlType.Edit', 'ControlType.DataItem')) {
                            $typingField = $lastInteractionTarget
                        }
                    }
                } catch {
                    if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
                    $typingCapture = $null
                    $typingWindow = $null
                    $typingField = $null
                }
            }
            if ($pastePressed) { $typingEvidenceKind = 'paste' }
            $lastTypingMs = if ($hookTextActivity -and $hookKeyboardElapsedMs -ge 0) {
                [int]$hookKeyboardElapsedMs
            } else { [int]$watch.ElapsedMilliseconds }
        }

        if ($typingActive -and $typingCaptureDueAtMs -gt 0 -and
            [int]$watch.ElapsedMilliseconds -ge $typingCaptureDueAtMs -and -not $clicked -and
            -not $typingCommitPressed -and
            (([int]$watch.ElapsedMilliseconds - $typingCaptureAtMs) -ge $typingEvidenceIntervalMs)) {
                # 最終文字のキー解放と描画を待って1回更新し、その後は凍結する。
                # 連続キャプチャはExcelの操作イベントを取り逃すほど重いため行わず、
                # Enterを押した後の計算結果で確定前の式を上書きしない。
                $replacementCapture = $null
                try {
                    $replacementWindow = Get-MbForegroundWindowInfo
                    $sameTypingWindow = $null -ne $replacementWindow -and $null -ne $typingWindow -and
                        [long]$replacementWindow.handle -eq [long]$typingWindow.handle
                    if ($sameTypingWindow) {
                        $replacementCapture = Copy-MbScreenBitmap
                        if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
                        $typingCapture = $replacementCapture
                        $replacementCapture = $null
                        $typingCaptureAtMs = [int]$watch.ElapsedMilliseconds
                    }
                    $typingCaptureDueAtMs = 0
                } catch {
                    if ($null -ne $replacementCapture) { try { $replacementCapture.bitmap.Dispose() } catch { } }
                }
        }

        # 入力が確定した、途切れた、または次のクリックが来たら、保持していた
        # 確定前の最新画面を1手順にする。Enter/Tab検出後には撮り直さない。
        $typingFinishedByIdle = $typingActive -and -not $typingCommitPressed -and
            -not ($clicked -and -not $typingStartedThisIteration) -and
            (([int]$watch.ElapsedMilliseconds - $lastTypingMs) -ge $TypingIdleMs) -and
            -not $hookKeyboardContinuationQueued
        $typingFinished = $typingActive -and ($typingCommitPressed -or
            ($clicked -and -not $typingStartedThisIteration) -or $typingFinishedByIdle)
        if ($typingFinished) {
            $typingActive = $false
            try {
                # 確定後にセルが計算結果へ変わる前の最終表示を、時系列原本へ1枚だけ残す。
                # 例: =SUM(B2:B3) は確定後に 1550 としか見えず、式を復元できない。
                if ($null -ne $typingCapture -and $null -ne $typingWindow -and
                    -not [string]::IsNullOrWhiteSpace($FramesDirectory) -and
                    -not [string]::IsNullOrWhiteSpace($FramesPath)) {
                    try {
                        $frameIndex++
                        [void](Save-MbRecordingTimelineFrame -Capture $typingCapture -Index $frameIndex `
                            -ElapsedMs $typingCaptureAtMs -FramesDirectory $FramesDirectory `
                            -FramesPath $FramesPath -Window $typingWindow -Role 'input-evidence' `
                            -EvidenceKind $typingEvidenceKind)
                    } catch {
                        if ($frameIndex -gt 0) { $frameIndex-- }
                    }
                }
                # 待機で入力が完了した場合は、確定後の文字が見える最新画面へ更新する。
                # クリックで完了した場合は画面遷移後を撮らないよう、最後の入力時点を使う。
                if (-not $clicked -and $null -ne $typingCapture) {
                    $replacementCapture = $null
                    try {
                        $replacementCapture = Copy-MbScreenBitmap
                        $typingCapture.bitmap.Dispose()
                        $typingCapture = $replacementCapture
                        $replacementCapture = $null
                    } catch {
                        if ($null -ne $replacementCapture) { try { $replacementCapture.bitmap.Dispose() } catch { } }
                    }
                }
                if ($null -ne $typingCapture -and $index -lt $MaxEvents) {
                    $index++
                    # 次のクリックが入力を確定した場合、対象解析や画像保存に時間が
                    # かかっても入力をそのクリックより後へ並べない。後続ボタンのOCR名を
                    # 入力欄へ誤って引き継ぐ原因になるため、実クリック時刻の直前に置く。
                    $inputElapsedMs = if ($hookCommitActivity -and $hookKeyboardElapsedMs -ge 0 -and -not $clicked) {
                        [int]$hookKeyboardElapsedMs
                    } elseif ($typingFinishedByIdle) {
                        # フックキューが滞留して現在時刻が先へ進んでいても、入力を
                        # 後から処理するクリックより未来へ並べない。
                        [int][Math]::Min([int]$watch.ElapsedMilliseconds, $lastTypingMs + $TypingIdleMs)
                    } else {
                        Resolve-MbTypingEventElapsedMs -CurrentElapsedMs ([int]$watch.ElapsedMilliseconds) `
                            -Clicked $clicked -ClickElapsedMs $clickElapsedMs
                    }
                    $record = Save-MbRecordingEvent -Capture $typingCapture -Index $index -ElapsedMs $inputElapsedMs `
                        -Kind 'input' -EventsDirectory $EventsDirectory -EventsPath $EventsPath `
                        -EvidenceDirectory $EvidenceDirectory -LedgerPath $LedgerPath -JobId $JobId -Target $typingField `
                        -Window $typingWindow
                    $lastTarget = [string]$record.targetName
                }
            } catch {
                # 1件取り損ねても記録は続ける。
            } finally {
                if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
            }
            $typingField = $null
            $typingWindow = $null
            $typingCapture = $null
            $typingCaptureAtMs = 0
            $typingCaptureDueAtMs = 0
            $typingEvidenceKind = ''
            if (-not $clicked) {
                # 確定済みセルを次の入力へ引き継ぐと、クリック取り逃し時に1つ前の
                # セルへ赤枠が付く。次の操作が明示クリックでなければ対象不明を優先する。
                $lastInteractionTarget = $null
                $lastInteractionWindowHandle = 0L
                $lastInteractionAtMs = -10000
            }
        }

        # 前のクリック結果を保存する。通常は遷移が落ち着くまで少し待ってから撮る。
        # 利用者が先に次の操作を始めた場合は、その操作の直前に保持した安定画面が
        # 直前操作の結果そのものなので、同じビットマップを結果画像として使う。
        if ($pendingResultIndex -gt 0 -and -not $leftDown -and -not $rightDown -and
            ($clicked -or ([int]$watch.ElapsedMilliseconds -ge $pendingResultDueAtMs))) {
            $resultCapture = $null
            try {
                if ($clicked -and $null -ne $preClickCapture -and $null -ne $preClickWindow) {
                    if ($pendingResultWindowHandle -gt 0 -and
                        [long]$preClickWindow.handle -eq $pendingResultWindowHandle) {
                        [void](Save-MbRecordingResultImage -Capture $preClickCapture -Index $pendingResultIndex `
                            -EventsDirectory $EventsDirectory -Window $preClickWindow)
                    }
                    $pendingResultIndex = 0
                    $pendingResultWindowHandle = 0L
                } elseif (-not $clicked) {
                    $windowBeforeResult = Get-MbForegroundWindowInfo
                    if (-not (Test-MbIgnoredWindow -Window $windowBeforeResult -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                        $samePendingWindow = $pendingResultWindowHandle -gt 0 -and $null -ne $windowBeforeResult -and
                            [long]$windowBeforeResult.handle -eq $pendingResultWindowHandle
                        if ($samePendingWindow) {
                            $resultCapture = Copy-MbScreenBitmap
                            $windowAfterResult = Get-MbForegroundWindowInfo
                            $stableResultWindow = $null -ne $windowAfterResult -and
                                [long]$windowBeforeResult.handle -eq [long]$windowAfterResult.handle -and
                                [string]$windowBeforeResult.title -eq [string]$windowAfterResult.title
                            if ($stableResultWindow) {
                                [void](Save-MbRecordingResultImage -Capture $resultCapture -Index $pendingResultIndex `
                                    -EventsDirectory $EventsDirectory -Window $windowBeforeResult)
                                $pendingResultIndex = 0
                                $pendingResultWindowHandle = 0L
                            } else {
                                # 遷移中なら次の巡回で再試行する。
                                $pendingResultDueAtMs = [int]$watch.ElapsedMilliseconds + 150
                            }
                        } else {
                            # 別アプリへ移った画面を、直前操作の結果として結び付けない。
                            $pendingResultIndex = 0
                            $pendingResultWindowHandle = 0L
                        }
                    } else {
                        # ManualBuilderへ戻った画面を操作結果として混ぜない。
                        $pendingResultIndex = 0
                        $pendingResultWindowHandle = 0L
                    }
                }
            } catch {
                # 結果画像だけ失敗しても、クリック前画像を持つ手順は残す。
                $pendingResultDueAtMs = [int]$watch.ElapsedMilliseconds + 200
            } finally {
                if ($null -ne $resultCapture) { try { $resultCapture.bitmap.Dispose() } catch { } }
            }
        }

        if ($clicked -and $index -lt $MaxEvents) {
            $capture = $null
            try {
                $point = New-Object 'MbRecorderNative+POINT'
                if ($null -ne $hookClick) {
                    $point.X = [int]$hookClick.X
                    $point.Y = [int]$hookClick.Y
                } else {
                    [void][MbRecorderNative]::GetCursorPos([ref]$point)
                }
                $window = if ($null -ne $hookClick -and [long]$hookClick.WindowHandle -ne 0) {
                    Get-MbWindowInfo -Handle ([IntPtr]::new([long]$hookClick.WindowHandle))
                } else { Get-MbForegroundWindowInfo }
                if ($null -eq $window) { $window = Get-MbForegroundWindowInfo }
                # キュー待ちの間に撮った画像をクリック前画像と誤認しない。
                # フックが保持した実際の押下時刻を基準に、新しいバッファは除外する。
                $captureAgeMs = ($clickElapsedMs - $preClickCaptureAtMs)
                $sameBufferedWindow = $null -ne $preClickWindow -and $null -ne $window -and
                    [long]$preClickWindow.handle -eq [long]$window.handle
                $bufferedDomTarget = $null
                $bufferedCachedTarget = $null
                $bufferTitleMatches = $sameBufferedWindow -and
                    [string]$preClickWindow.title -eq [string]$window.title
                $canUseBufferedCapture = $null -ne $preClickCapture -and $sameBufferedWindow -and
                    $captureAgeMs -ge 0 -and $captureAgeMs -le $preClickCaptureMaxAgeMs
                if ($canUseBufferedCapture -and -not $bufferTitleMatches) {
                    # 通常のEdgeや標準ダイアログでも、クリック前UIAが同じウィンドウと
                    # クリック点を示すなら、タイトル変更前の画像を安全に採用できる。
                    if (-not [string]::IsNullOrWhiteSpace($UiaTargetPath)) {
                        $bufferedCachedTarget = Get-MbUiaTargetFromCache -Path $UiaTargetPath `
                            -X ([int]$point.X) -Y ([int]$point.Y) -Window $preClickWindow -MaxAgeMs 700
                    }
                    if (-not [string]::IsNullOrWhiteSpace($DomTargetPath)) {
                        # 旧DOM監視を明示的に使う実験経路ではpointerdownも証拠にできる。
                        for ($domAttempt = 0; $domAttempt -lt 7 -and $null -eq $bufferedDomTarget; $domAttempt++) {
                            if ($domAttempt -gt 0) { Start-Sleep -Milliseconds 20 }
                            $bufferedDomTarget = Get-MbDomTargetFromCache -Path $DomTargetPath `
                                -X ([int]$point.X) -Y ([int]$point.Y) -Window $preClickWindow
                        }
                    }
                    # 証拠がない古い画面は、自動遷移後に押した次の操作へ誤適用しない。
                    $canUseBufferedCapture = $null -ne $bufferedDomTarget -or $null -ne $bufferedCachedTarget
                }
                if ($canUseBufferedCapture) {
                    # pointerdownを検出した時点では遷移済みでも、同じEdgeウィンドウの
                    # 遷移前画像とタイトルを使う。所有権を移し、次のクリックでは再利用しない。
                    $capture = $preClickCapture
                    $window = $preClickWindow
                    $preClickCapture = $null
                    $preClickWindow = $null
                    $preClickCaptureAtMs = -1000
                } else {
                    # 別ウィンドウへ切り替えてすぐ押した場合は古い画像を使わない。
                    if ($null -ne $preClickCapture) { try { $preClickCapture.bitmap.Dispose() } catch { } }
                    $preClickCapture = $null
                    $preClickWindow = $null
                    $preClickCaptureAtMs = -1000
                    $capture = Copy-MbScreenBitmap
                }
                if (-not (Test-MbIgnoredWindow -Window $window -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                    $target = $bufferedDomTarget
                    if (-not [string]::IsNullOrWhiteSpace($DomTargetPath)) {
                        # 画面はすでに押下直前で確保済み。Edgeからpointerdown通知が届くまで
                        # 最大約120msだけ待ち、速いマウス移動直後のクリックもDOMで拾う。
                        for ($domAttempt = 0; $domAttempt -lt 7 -and $null -eq $target; $domAttempt++) {
                            if ($domAttempt -gt 0) { Start-Sleep -Milliseconds 20 }
                            $target = Get-MbDomTargetFromCache -Path $DomTargetPath -X ([int]$point.X) -Y ([int]$point.Y) -Window $window
                        }
                    }
                    $domTarget = $target
                    # DOMが取れていてもUIAを代替候補として残す。以前はDOMが誤っていると
                    # UIAを一度も比較せず、その矩形だけがCopilotへ渡っていた。
                    $cachedTarget = $bufferedCachedTarget
                    if ($null -eq $cachedTarget -and -not [string]::IsNullOrWhiteSpace($UiaTargetPath)) {
                        $cachedTarget = Get-MbUiaTargetFromCache -Path $UiaTargetPath `
                            -X ([int]$point.X) -Y ([int]$point.Y) -Window $window
                    }
                    $domFileInput = $null -ne $target -and
                        $target.PSObject.Properties.Name -contains 'provider' -and [string]$target.provider -eq 'DOM' -and
                        $target.PSObject.Properties.Name -contains 'inputType' -and [string]$target.inputType -eq 'file'
                    if ($null -eq $target -or $domFileInput) {
                        # エクスプローラーや標準ダイアログではクリック後すぐに元要素が消える。
                        # 別プロセスがクリック前に保持した対象を、同じクリック点の新しい記録に限って使う。
                        $useCachedTarget = $null -eq $target -and $null -ne $cachedTarget
                        if ($domFileInput -and $null -ne $cachedTarget) {
                            # type=fileはDOM上ではボタンと未選択表示が1矩形になる。Windowsが
                            # 内側のButtonを公開する場合は、その小さい実コントロールへ赤枠を寄せる。
                            $domArea = [double]$target.width * [double]$target.height
                            $cachedArea = [double]$cachedTarget.width * [double]$cachedTarget.height
                            $useCachedTarget = [string]$cachedTarget.controlType -eq 'ControlType.Button' -or
                                ($domArea -gt 0 -and $cachedArea -lt ($domArea * 0.8))
                        }
                        if ($useCachedTarget) { $target = $cachedTarget }
                        if ($useCachedTarget -and $target.PSObject.Properties.Name -contains 'captureWindow' -and
                            $null -ne $target.captureWindow) {
                            # スクリーンショットもクリック前なので、切り出すウィンドウ範囲とタイトルを揃える。
                            $window = $target.captureWindow
                        }
                    }
                    # クリック後の同期UIA検索は行わない。Chromiumの大きなアクセシビリティ木では
                    # 数秒止まることがあり、その間の入力や次クリックを別画面へ誤結合していた。
                    # 操作前に別プロセスで得たキャッシュが無ければクリック座標を事実として残し、
                    # 対象名と必要場面の精査は、連続フレームを見られるCopilotへ委ねる。
                    $postTarget = $null
                    if ($null -eq $target) {
                        $target = if ($null -ne $cachedTarget) { $cachedTarget } else { $postTarget }
                    }
                    $pointTarget = New-MbClickPointTargetInfo -X ([int]$point.X) -Y ([int]$point.Y) -Window $window
                    $recordingTargetCandidates = @($domTarget, $cachedTarget, $postTarget, $pointTarget) | Where-Object { $null -ne $_ }
                    $target = Select-MbRecordingPrimaryTarget -Target $target -PointTarget $pointTarget -Window $window
                    if ($null -ne $target -and $target.PSObject.Properties.Name -contains 'provider' -and
                        [string]$target.provider -eq 'DOM') {
                        $lastDomTarget = $target
                        $lastDomWindowHandle = [long]$window.handle
                    } else {
                        $lastDomTarget = $null
                        $lastDomWindowHandle = 0L
                    }
                    # 直前画面バッファを持てないほど速い連続操作でも、現在のクリック用に
                    # 確保した操作前画面を、ひとつ前の操作結果として欠落させない。
                    if ($pendingResultIndex -gt 0) {
                        try {
                            if ($pendingResultWindowHandle -gt 0 -and
                                [long]$window.handle -eq $pendingResultWindowHandle) {
                                [void](Save-MbRecordingResultImage -Capture $capture -Index $pendingResultIndex `
                                    -EventsDirectory $EventsDirectory -Window $window)
                            }
                            $pendingResultIndex = 0
                            $pendingResultWindowHandle = 0L
                        } catch { }
                    }
                    $index++
                    $clickKind = if ($rightClicked) { 'right-click' } else { 'click' }
                    $record = Save-MbRecordingEvent -Capture $capture -Index $index -ElapsedMs $clickElapsedMs `
                        -Kind $clickKind -EventsDirectory $EventsDirectory -EventsPath $EventsPath `
                        -EvidenceDirectory $EvidenceDirectory -LedgerPath $LedgerPath -JobId $JobId -Target $target `
                        -TargetCandidates $recordingTargetCandidates -Window $window `
                        -ScreenX ([int]$point.X) -ScreenY ([int]$point.Y)
                    if (-not [string]::IsNullOrWhiteSpace($FramesDirectory) -and
                        -not [string]::IsNullOrWhiteSpace($FramesPath)) {
                        try {
                            $frameIndex++
                            [void](Save-MbRecordingTimelineEventFrame -EventRecord $record -Index $frameIndex `
                                -ElapsedMs $clickElapsedMs -EventsDirectory $EventsDirectory `
                                -FramesDirectory $FramesDirectory -FramesPath $FramesPath -Window $window)
                        } catch {
                            if ($frameIndex -gt 0) { $frameIndex-- }
                        }
                    }
                    $lastTarget = [string]$record.targetName
                    $lastInteractionTarget = $target
                    $lastInteractionWindowHandle = [long]$window.handle
                    $lastInteractionAtMs = $clickElapsedMs
                    # フォーカス用クリックと最初のキーを同じ巡回で検出した場合、入力開始時には
                    # まだ lastInteractionTarget がない。クリックの事実を入力イベントへ引き継ぎ、
                    # 対象名が取れないアプリでもクリックと入力を別手順にしない。
                    if ($typingActive -and $null -eq $typingField -and $null -ne $typingWindow -and
                        [long]$typingWindow.handle -eq [long]$window.handle) {
                        $typingField = $target
                    }
                    $pendingResultIndex = $index
                    $pendingResultWindowHandle = [long]$window.handle
                    $pendingResultDueAtMs = [int]$watch.ElapsedMilliseconds + $ResultCaptureDelayMs
                }
            } catch {
                # 応答しないアプリを押した場合など。記録は続ける。
                try {
                    $safeMessage = [regex]::Replace([string]$_.Exception.Message, '[\r\n]+', ' ')
                    [IO.File]::AppendAllText($diagnosticPath,
                        ([DateTime]::UtcNow.ToString('o') + ' click=' + $safeMessage + [Environment]::NewLine),
                        [Text.UTF8Encoding]::new($false))
                } catch { }
            } finally {
                if ($null -ne $capture) { try { $capture.bitmap.Dispose() } catch { } }
            }
        }

        # クリックを検出してから撮ると、遷移の速いページでは間に合わない。
        # 押下が無い巡回だけで低頻度に更新し、撮影中にタイトル／前面ウィンドウが
        # 変わった不安定なフレームは保持しない。
        if (-not $clicked -and -not $leftDown -and -not $rightDown -and
            (([int]$watch.ElapsedMilliseconds - $preClickCaptureAttemptAtMs) -ge $preClickCaptureIntervalMs)) {
            $replacementCapture = $null
            try {
                $preClickCaptureAttemptAtMs = [int]$watch.ElapsedMilliseconds
                $windowBeforeCapture = Get-MbForegroundWindowInfo
                $replacementCapture = Copy-MbScreenBitmap
                $windowAfterCapture = Get-MbForegroundWindowInfo
                # CopyFromScreen中に短いクリックと遷移が完了しても、その画像を
                # 「クリック前」として採用せず、押下履歴は次の巡回へ渡す。
                $leftAfterCapture = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkLeftButton)
                $rightAfterCapture = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkRightButton)
                $leftPressedDuringCapture = Test-MbAsyncKeyStatePressed -State $leftAfterCapture
                $rightPressedDuringCapture = Test-MbAsyncKeyStatePressed -State $rightAfterCapture
                $leftDownAfterCapture = Test-MbAsyncKeyStateDown -State $leftAfterCapture
                $rightDownAfterCapture = Test-MbAsyncKeyStateDown -State $rightAfterCapture
                if (-not $mouseHookActive) {
                    $pendingLeftClick = $leftPressedDuringCapture -or $leftDownAfterCapture
                    $pendingRightClick = $rightPressedDuringCapture -or $rightDownAfterCapture
                }
                $sameStableWindow = $null -ne $windowBeforeCapture -and $null -ne $windowAfterCapture -and
                    [long]$windowBeforeCapture.handle -eq [long]$windowAfterCapture.handle -and
                    [string]$windowBeforeCapture.title -eq [string]$windowAfterCapture.title
                $ignoredWindow = -not $sameStableWindow -or
                    (Test-MbIgnoredWindow -Window $windowBeforeCapture -IgnoreTitlePatterns $IgnoreTitlePatterns)

                # AI用の時系列原本は、クリック前画像より緩い条件で、新しく撮った画面を直接保存する。
                # CopyFromScreen中に短いクリックが完了していても、ボタンが現在離れていれば
                # 操作後の安定画面として有用。ここをクリック前画像と共用すると、操作中だけ
                # 古い画面が繰り返され、AIが必要な状態変化を選べなくなる。
                if (-not $ignoredWindow -and -not $leftDownAfterCapture -and -not $rightDownAfterCapture -and
                    -not [string]::IsNullOrWhiteSpace($FramesDirectory) -and
                    -not [string]::IsNullOrWhiteSpace($FramesPath) -and
                    (([int]$watch.ElapsedMilliseconds - $lastFrameAtMs) -ge $timelineFrameIntervalMs)) {
                    try {
                        $frameIndex++
                        [void](Save-MbRecordingTimelineFrame -Capture $replacementCapture -Index $frameIndex `
                            -ElapsedMs ([int]$watch.ElapsedMilliseconds) -FramesDirectory $FramesDirectory `
                            -FramesPath $FramesPath -Window $windowAfterCapture)
                        $lastFrameAtMs = [int]$watch.ElapsedMilliseconds
                    } catch {
                        if ($frameIndex -gt 0) { $frameIndex-- }
                        $lastFrameAtMs = [int]$watch.ElapsedMilliseconds
                    }
                }

                # クリック前画像は、撮影中にも押下が無かった厳密な場合だけ更新する。
                if (-not $ignoredWindow -and -not $pendingLeftClick -and -not $pendingRightClick) {
                    if ($null -ne $preClickCapture) { try { $preClickCapture.bitmap.Dispose() } catch { } }
                    $preClickCapture = $replacementCapture
                    $replacementCapture = $null
                    $preClickWindow = $windowBeforeCapture
                    $preClickCaptureAtMs = [int]$watch.ElapsedMilliseconds
                }
            } catch {
                # 直前画面を更新できなくても、クリック時の通常撮影へ戻れる。
            } finally {
                if ($null -ne $replacementCapture) { try { $replacementCapture.bitmap.Dispose() } catch { } }
            }
        }

        if (([int]$watch.ElapsedMilliseconds - $lastStatusMs) -ge 400) {
            $lastStatusMs = [int]$watch.ElapsedMilliseconds
            Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count $index `
                -Message '操作を記録しています' -LastTarget $lastTarget
        }
        Start-Sleep -Milliseconds $PollIntervalMs
    }

    if ($mouseHookActive) {
        $droppedMouseClicks = try { [long][MbRecorderNative]::DroppedMouseClicks } catch { 0L }
        $droppedKeyboardActivities = try { [long][MbRecorderNative]::DroppedKeyboardActivities } catch { 0L }
        if ($droppedMouseClicks -gt $lastDroppedMouseClicks -or
            $droppedKeyboardActivities -gt $lastDroppedKeyboardActivities) {
            $script:MbRecorderCaptureCompleteness = 'known-gaps'
            $script:MbRecorderCaptureWarning = '操作が短時間に集中し、一部を記録できなかった可能性があります。手順の抜けを確認してください。'
            Write-MbRecordingLedgerRecord -LedgerPath $LedgerPath -Record ([ordered]@{
                recordType = 'capture-gap'; formatVersion = 2; sessionId = $JobId
                recordedAt = [DateTime]::UtcNow.ToString('o')
                droppedMouseClicks = [Math]::Max(0L, $droppedMouseClicks - $lastDroppedMouseClicks)
                droppedKeyboardActivities = [Math]::Max(0L, $droppedKeyboardActivities - $lastDroppedKeyboardActivities)
                reason = 'capture-queue-overflow'
            })
        }
        try { [MbRecorderNative]::StopMouseHook() } catch { }
        $mouseHookActive = $false
    }

    # 終了ボタンへ戻る直前まで保持していた安定画面があれば、最後のクリック結果に使う。
    # これにより「詳細を表示」してすぐ記録を止めても、結果だけが欠けにくい。
    if ($pendingResultIndex -gt 0 -and $pendingResultWindowHandle -gt 0 -and
        $null -ne $preClickCapture -and $null -ne $preClickWindow -and
        [long]$preClickWindow.handle -eq $pendingResultWindowHandle -and
        -not (Test-MbIgnoredWindow -Window $preClickWindow -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
        try {
            [void](Save-MbRecordingResultImage -Capture $preClickCapture -Index $pendingResultIndex `
                -EventsDirectory $EventsDirectory -Window $preClickWindow)
            $pendingResultIndex = 0
            $pendingResultWindowHandle = 0L
        } catch { }
    }

    # 停止要求が入力の途中で届いても、保持していた入力画面を最後の1手順として残す。
    if ($typingActive -and $null -ne $typingCapture -and $index -lt $MaxEvents) {
        try {
            $index++
            $record = Save-MbRecordingEvent -Capture $typingCapture -Index $index -ElapsedMs ([int]$watch.ElapsedMilliseconds) `
                -Kind 'input' -EventsDirectory $EventsDirectory -EventsPath $EventsPath `
                -EvidenceDirectory $EvidenceDirectory -LedgerPath $LedgerPath -JobId $JobId -Target $typingField `
                -Window $typingWindow
            $lastTarget = [string]$record.targetName
        } catch {
            # 最後の1件に失敗しても、それまでの記録は利用できる。
        } finally {
            try { $typingCapture.bitmap.Dispose() } catch { }
            $typingCapture = $null
        }
    } elseif ($null -ne $typingCapture) {
        try { $typingCapture.bitmap.Dispose() } catch { }
        $typingCapture = $null
    }
    if ($null -ne $preClickCapture) {
        try { $preClickCapture.bitmap.Dispose() } catch { }
        $preClickCapture = $null
    }

    $reason = 'stopped'
    if ($index -ge $MaxEvents) { $reason = 'limit' }
    elseif ($watch.Elapsed.TotalMinutes -ge $MaxMinutes) { $reason = 'timeout' }
    $message = switch ($reason) {
        'limit'   { "記録の上限（$MaxEvents 件）に達したため終了しました" }
        'timeout' { "記録の上限（$MaxMinutes 分）に達したため終了しました" }
        default   { "$index 件の操作を記録しました" }
    }
    Write-MbRecordingLedgerRecord -LedgerPath $LedgerPath -Record ([ordered]@{
        recordType = 'capture-end'; formatVersion = 2; sessionId = $JobId
        recordedAt = [DateTime]::UtcNow.ToString('o'); operationCount = $index
        reason = $reason; completeness = [string]$script:MbRecorderCaptureCompleteness
        warning = [string]$script:MbRecorderCaptureWarning
    })
    Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'completed' -Count $index -Message $message -LastTarget $lastTarget
    return $index
}

Export-ModuleMember -Function @(
    'Initialize-MbRecorderNative',
    'Set-MbProcessDpiAware',
    'Initialize-MbRecorderUia',
    'Get-MbRecorderCapability',
    'Get-MbUiaTargetAtPoint',
    'Get-MbUiaHoverTargetAtPoint',
    'Get-MbUiaTargetFromCache',
    'Invoke-MbUiaTargetCacheLoop',
    'Get-MbUiaFocusedElement',
    'ConvertFrom-MbDomSnapshotTarget',
    'Get-MbDomTargetFromCache',
    'Get-MbRecorderTargetEvidence',
    'ConvertTo-MbRecordingTargetCandidates',
    'Select-MbUiaTargetInfo',
    'Select-MbUiaNamedTargetInfo',
    'New-MbMsaaElementInfo',
    'New-MbClickPointTargetInfo',
    'Select-MbRecordingPrimaryTarget',
    'ConvertTo-MbNormalizedClickPoint',
    'Save-MbRecordingResultImage',
    'Save-MbRecordingTimelineFrame',
    'Save-MbRecordingTimelineEventFrame',
    'Get-MbForegroundWindowInfo',
    'Get-MbVirtualScreenBounds',
    'Get-MbCaptureRegion',
    'ConvertTo-MbRecorderTargetName',
    'Copy-MbScreenBitmap',
    'Save-MbBitmapRegion',
    'ConvertTo-MbRegionRect',
    'Test-MbUsableElementInfo',
    'Test-MbIgnoredWindow',
    'Get-MbWatchedTypingKeys',
    'Get-MbWatchedCommitKeys',
    'Test-MbAsyncKeyStateDown',
    'Test-MbAsyncKeyStatePressed',
    'Test-MbTextChangingShortcutKey',
    'Resolve-MbTypingEventElapsedMs',
    'Test-MbQueuedKeyboardContinuation',
    'Remove-MbLastRecordingEvent',
    'Invoke-MbRecordingLoop',
    'Write-MbRecordingStatus'
)
