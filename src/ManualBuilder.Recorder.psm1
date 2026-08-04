# 操作を記録して手順にする。
#
# 録画から場面を切り出す方式は、画面の変化から「押された場所」を推定する。
# こちらは推定しない。Windows の UI Automation が、クリックした先のコントロールの
# 名前・種類・矩形をそのまま持っているので、それを読む。赤枠の位置も操作対象の名前も
# 確定値になり、文字認識も要らなくなる。
#
# 実装上の判断:
#
#   低レベルフック（WH_MOUSE_LL）は使わない。
#     フックのコールバックが LowLevelHooksTimeout（既定 300ms）を超えると、Windows は
#     警告なくフックを外す。コールバック内でスクリーンショットを撮れば確実に超える。
#     60Hz で GetAsyncKeyState を見るポーリングなら、メッセージポンプもデリゲートの
#     寿命管理も要らず、人間の操作速度には十分間に合う。
#
#   キーの文字は記録しない。
#     パスワードや個人情報をそのまま拾ってしまうため。「入力があった」ことと
#     「どの入力欄か」（UIAのフォーカス要素の名前）だけを見る。
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
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class MbRecorderNative
{
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

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
            'RangeValuePatternIdentifiers.Pattern'
        )
        try {
            foreach ($pattern in @($Element.GetSupportedPatterns())) {
                if ($actionPatternNames -contains [string]$pattern.ProgrammaticName) {
                    $isActionable = $true
                    break
                }
            }
        } catch { }
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
        [Parameter(Mandatory = $true)][double]$Y
    )
    if (-not (Test-MbUsableElementInfo -Info $Info)) { return $false }
    $right = [double]$Info.left + [double]$Info.width
    $bottom = [double]$Info.top + [double]$Info.height
    return ($X -ge ([double]$Info.left - 1.0) -and $X -le ($right + 1.0) -and
        $Y -ge ([double]$Info.top - 1.0) -and $Y -le ($bottom + 1.0))
}

# 同じ点に重なる候補から、実際に操作できる最小の要素を選ぶ。
# 名前は操作説明の手がかりだが、名前が空でもボタンの矩形は赤枠として有用なので除外しない。
function Select-MbUiaTargetInfo {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Candidates,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y
    )

    $best = $null
    $bestArea = [double]::PositiveInfinity
    foreach ($candidate in @($Candidates)) {
        if (-not (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y)) { continue }
        if (-not (Test-MbRecorderInteractiveElementInfo -Info $candidate)) { continue }

        $area = [double]$candidate.width * [double]$candidate.height
        $preferNamed = ($area -eq $bestArea -and $null -ne $best -and
            [string]::IsNullOrWhiteSpace([string]$best.name) -and
            -not [string]::IsNullOrWhiteSpace([string]$candidate.name))
        if ($area -lt $bestArea -or $preferNamed) {
            $best = $candidate
            $bestArea = $area
        }
    }
    return $best
}

# クリックした点にあるコントロールを返す。
#
# UIAのFromPointは、多くのアプリでは最も内側の要素を返すが、Chromium系ブラウザーでは
# ページ全体のDocument/Paneを返すことがある。その場合は点を含む子を下へ掘り、
# 文字要素が返った場合は親も調べる。最後は操作可能な最小要素だけを採用する。
function Get-MbUiaTargetAtPoint {
    param([Parameter(Mandatory = $true)][int]$X, [Parameter(Mandatory = $true)][int]$Y)

    if (-not (Initialize-MbRecorderUia)) { return $null }
    try {
        $point = New-Object System.Windows.Point -ArgumentList @([double]$X, [double]$Y)
        $element = [System.Windows.Automation.AutomationElement]::FromPoint($point)
        if ($null -eq $element) { return $null }

        $info = Get-MbAutomationElementInfo -Element $element
        if (-not (Test-MbUsableElementInfo -Info $info)) { return $null }

        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $candidates = New-Object System.Collections.ArrayList
        if (Test-MbPointWithinElementInfo -Info $info -X $X -Y $Y) {
            [void]$candidates.Add($info)
        }

        # FromPointが文字などの葉を返したとき、リンクやボタンの親まで調べる。
        $node = $element
        for ($depth = 0; $depth -lt 6; $depth++) {
            $parent = $null
            try { $parent = $walker.GetParent($node) } catch { $parent = $null }
            if ($null -eq $parent) { break }

            $parentInfo = $null
            try { $parentInfo = Get-MbAutomationElementInfo -Element $parent } catch { $parentInfo = $null }
            if (-not (Test-MbUsableElementInfo -Info $parentInfo)) { break }
            if ([string]$parentInfo.controlType -eq 'ControlType.Window') { break }
            if (Test-MbPointWithinElementInfo -Info $parentInfo -X $X -Y $Y) {
                [void]$candidates.Add($parentInfo)
            }
            $node = $parent
        }

        # FromPointがページ全体を返したとき、クリック点を含む最小の子を選びながら下へ掘る。
        $node = $element
        for ($depth = 0; $depth -lt 10; $depth++) {
            $child = $null
            try { $child = $walker.GetFirstChild($node) } catch { $child = $null }
            if ($null -eq $child) { break }

            $nextNode = $null
            $nextArea = [double]::PositiveInfinity
            $checked = 0
            while ($null -ne $child -and $checked -lt 256) {
                $checked++
                $childInfo = $null
                try { $childInfo = Get-MbAutomationElementInfo -Element $child } catch { $childInfo = $null }
                if (Test-MbPointWithinElementInfo -Info $childInfo -X $X -Y $Y) {
                    [void]$candidates.Add($childInfo)
                    $area = [double]$childInfo.width * [double]$childInfo.height
                    if ($area -lt $nextArea) {
                        $nextNode = $child
                        $nextArea = $area
                    }
                }
                try { $child = $walker.GetNextSibling($child) } catch { $child = $null }
            }
            if ($null -eq $nextNode) { break }
            $node = $nextNode
        }

        # DocumentやPaneしか得られなかった場合はnullにする。
        # 不確かな画面全体の枠より、赤枠なしのほうが手順として誤解が少ない。
        return (Select-MbUiaTargetInfo -Candidates @($candidates) -X $X -Y $Y)
    } catch {
        # 応答しないアプリでは例外になる。記録は止めず、赤枠なしの手順として残す。
        return $null
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

        # 入力欄を特定できないときにDocument全体を黒塗りしない。
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

function Get-MbForegroundWindowInfo {
    Initialize-MbRecorderNative
    $handle = [MbRecorderNative]::GetForegroundWindow()
    if ($handle -eq [IntPtr]::Zero) { return $null }
    $rect = [MbRecorderNative]::GetVisualWindowRect($handle)
    return [pscustomobject]@{
        title  = [MbRecorderNative]::GetWindowTitle($handle)
        class  = [MbRecorderNative]::GetWindowClass($handle)
        left   = [int]$rect.Left
        top    = [int]$rect.Top
        width  = [int]($rect.Right - $rect.Left)
        height = [int]($rect.Bottom - $rect.Top)
    }
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
        [long]$Quality = 88,
        [AllowNull()]$RedactTarget = $null
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

        # キー入力の内容は読み取らないだけでなく、入力欄に表示された文字も画像へ残さない。
        # UI Automation の物理座標を、切り出した画像内の座標へ直して入力欄全体を隠す。
        if ($null -ne $RedactTarget) {
            $redactLeft = [Math]::Max(0, [int][Math]::Floor([double]$RedactTarget.left - ($left + [int]$origin.left)))
            $redactTop = [Math]::Max(0, [int][Math]::Floor([double]$RedactTarget.top - ($top + [int]$origin.top)))
            $redactRight = [Math]::Min($width, [int][Math]::Ceiling([double]$RedactTarget.left + [double]$RedactTarget.width - ($left + [int]$origin.left)))
            $redactBottom = [Math]::Min($height, [int][Math]::Ceiling([double]$RedactTarget.top + [double]$RedactTarget.height - ($top + [int]$origin.top)))
            if ($redactRight -gt $redactLeft -and $redactBottom -gt $redactTop) {
                $redactionGraphics = $null
                try {
                    $redactionGraphics = [Drawing.Graphics]::FromImage($cropped)
                    $redactionGraphics.FillRectangle([Drawing.Brushes]::Black, $redactLeft, $redactTop,
                        ($redactRight - $redactLeft), ($redactBottom - $redactTop))
                } finally {
                    if ($null -ne $redactionGraphics) { try { $redactionGraphics.Dispose() } catch { } }
                }
            }
        }

        $output = $cropped
        $longest = [Math]::Max($width, $height)
        if ($MaxEdge -gt 0 -and $longest -gt $MaxEdge) {
            $scale = $MaxEdge / [double]$longest
            $scaled = New-Object Drawing.Bitmap -ArgumentList @($cropped, [int][Math]::Round($width * $scale), [int][Math]::Round($height * $scale))
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
    if (($x2 - $x1) -lt 0.002 -or ($y2 - $y1) -lt 0.002) { return $null }
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

function Test-MbIgnoredWindow {
    param([AllowNull()]$Window, [string[]]$IgnoreTitlePatterns = @())
    if ($null -eq $Window) { return $true }
    if ($script:MbRecorderIgnoredClasses -contains [string]$Window.class) { return $true }
    $title = [string]$Window.title
    foreach ($pattern in $IgnoreTitlePatterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        if ($title -like ('*' + $pattern + '*')) { return $true }
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

function Write-MbRecordingStatus {
    param(
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$JobId,
        [Parameter(Mandatory = $true)][string]$State,
        [int]$Count = 0,
        [string]$Message = '',
        [string]$LastTarget = ''
    )
    $status = [pscustomobject]@{
        jobId = $JobId; state = $State; count = $Count; message = $Message
        lastTarget = $LastTarget; updatedAt = [DateTime]::UtcNow.ToString('o')
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

# 1件ぶんの操作を記録する。押す直前の画面を先に確保してからUIAを引く。
function Save-MbRecordingEvent {
    param(
        [Parameter(Mandatory = $true)]$Capture,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][int]$ElapsedMs,
        [Parameter(Mandatory = $true)][string]$Kind,
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [Parameter(Mandatory = $true)][string]$EventsPath,
        [AllowNull()]$Target,
        [AllowNull()]$Window,
        [AllowNull()]$RedactTarget = $null,
        [int]$MaxEdge = 1600
    )

    $region = Get-MbCaptureRegion -Window $Window -Target $Target
    $fileName = ('event-{0:d3}.jpg' -f $Index)
    $saved = Save-MbBitmapRegion -Capture $Capture -Region $region -Path (Join-Path $EventsDirectory $fileName) `
        -MaxEdge $MaxEdge -RedactTarget $RedactTarget
    # 実際に切り出せた範囲で正規化する。画面の端では要求した範囲より狭くなる。
    $rect = ConvertTo-MbRegionRect -Region $saved -Target $Target

    $windowTitle = ''
    if ($null -ne $Window) { $windowTitle = [string]$Window.title }
    $targetName = ''
    $targetType = ''
    if ($null -ne $Target) {
        $targetName = [string]$Target.name
        $targetType = [string]$Target.controlType
    }
    $record = @{
        index       = $Index
        kind        = $Kind
        timeMs      = $ElapsedMs
        image       = $fileName
        windowTitle = $windowTitle
        targetName  = $targetName
        targetType  = $targetType
        rect        = $rect
    }
    Write-MbRecordingEvent -EventsPath $EventsPath -Record $record
    return $record
}

function Invoke-MbRecordingLoop {
    param(
        [Parameter(Mandatory = $true)][string]$EventsDirectory,
        [Parameter(Mandatory = $true)][string]$EventsPath,
        [Parameter(Mandatory = $true)][string]$StatusPath,
        [Parameter(Mandatory = $true)][string]$StopPath,
        [Parameter(Mandatory = $true)][string]$JobId,
        [string[]]$IgnoreTitlePatterns = @(),
        [int]$PollIntervalMs = 16,
        [int]$TypingIdleMs = 1200,
        [int]$MaxEvents = 300,
        [int]$MaxMinutes = 60
    )

    Initialize-MbRecorderNative
    [void](Set-MbProcessDpiAware)
    [void](Initialize-MbRecorderUia)

    $typingKeys = Get-MbWatchedTypingKeys
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $index = 0
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
    $lastTypingMs = 0

    Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count 0 -Message '操作を記録しています'

    while ($true) {
        if (Test-Path -LiteralPath $StopPath -PathType Leaf) { break }
        if ($index -ge $MaxEvents) { break }
        if ($watch.Elapsed.TotalMinutes -ge $MaxMinutes) { break }

        $leftState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkLeftButton)
        $rightState = [int][MbRecorderNative]::GetAsyncKeyState($script:MbVkRightButton)
        $leftDown = Test-MbAsyncKeyStateDown -State $leftState
        $rightDown = Test-MbAsyncKeyStateDown -State $rightState
        $leftClicked = (Test-MbAsyncKeyStatePressed -State $leftState) -or ($leftDown -and -not $leftWasDown)
        $rightClicked = (Test-MbAsyncKeyStatePressed -State $rightState) -or ($rightDown -and -not $rightWasDown)
        $clicked = ($leftClicked -or $rightClicked)
        $leftWasDown = $leftDown
        $rightWasDown = $rightDown

        $typingNow = $false
        foreach ($vk in $typingKeys) {
            $keyState = [int][MbRecorderNative]::GetAsyncKeyState($vk)
            if ((Test-MbAsyncKeyStateDown -State $keyState) -or (Test-MbAsyncKeyStatePressed -State $keyState)) {
                $typingNow = $true
                break
            }
        }
        if ($typingNow) {
            if (-not $typingActive) {
                $typingActive = $true
                # 入力後に撮ると、文字列そのものがスクリーンショットへ残る。
                # 最初のキーを検出した時点の画面を保持し、保存時には入力欄も黒塗りする。
                try {
                    $candidateWindow = Get-MbForegroundWindowInfo
                    if (-not (Test-MbIgnoredWindow -Window $candidateWindow -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                        $typingCapture = Copy-MbScreenBitmap
                        $typingWindow = $candidateWindow
                        $typingField = Get-MbUiaFocusedElement
                    }
                } catch {
                    if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
                    $typingCapture = $null
                    $typingWindow = $null
                    $typingField = $null
                }
            }
            $lastTypingMs = [int]$watch.ElapsedMilliseconds
        }

        # 入力が途切れたか、次のクリックが来たら、保持していた入力開始時の画面を1手順にする。
        $typingFinished = $typingActive -and ($clicked -or (([int]$watch.ElapsedMilliseconds - $lastTypingMs) -ge $TypingIdleMs))
        if ($typingFinished) {
            $typingActive = $false
            try {
                if ($null -ne $typingCapture -and $index -lt $MaxEvents) {
                    $index++
                    $record = Save-MbRecordingEvent -Capture $typingCapture -Index $index -ElapsedMs ([int]$watch.ElapsedMilliseconds) `
                        -Kind 'input' -EventsDirectory $EventsDirectory -EventsPath $EventsPath -Target $typingField `
                        -Window $typingWindow -RedactTarget $typingField
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
        }

        if ($clicked -and $index -lt $MaxEvents) {
            $capture = $null
            try {
                # 押した瞬間の画面を最優先で確保する。UIAはこのあと。
                $capture = Copy-MbScreenBitmap
                $point = New-Object 'MbRecorderNative+POINT'
                [void][MbRecorderNative]::GetCursorPos([ref]$point)
                $window = Get-MbForegroundWindowInfo
                if (-not (Test-MbIgnoredWindow -Window $window -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                    $target = Get-MbUiaTargetAtPoint -X ([int]$point.X) -Y ([int]$point.Y)
                    $index++
                    $clickKind = if ($rightClicked) { 'right-click' } else { 'click' }
                    $record = Save-MbRecordingEvent -Capture $capture -Index $index -ElapsedMs ([int]$watch.ElapsedMilliseconds) `
                        -Kind $clickKind -EventsDirectory $EventsDirectory -EventsPath $EventsPath -Target $target -Window $window
                    $lastTarget = [string]$record.targetName
                }
            } catch {
                # 応答しないアプリを押した場合など。記録は続ける。
            } finally {
                if ($null -ne $capture) { try { $capture.bitmap.Dispose() } catch { } }
            }
        }

        if (([int]$watch.ElapsedMilliseconds - $lastStatusMs) -ge 400) {
            $lastStatusMs = [int]$watch.ElapsedMilliseconds
            Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'recording' -Count $index `
                -Message '操作を記録しています' -LastTarget $lastTarget
        }
        Start-Sleep -Milliseconds $PollIntervalMs
    }

    # 停止要求が入力の途中で届いても、保持していた入力前の画面を最後の1手順として残す。
    if ($typingActive -and $null -ne $typingCapture -and $index -lt $MaxEvents) {
        try {
            $index++
            $record = Save-MbRecordingEvent -Capture $typingCapture -Index $index -ElapsedMs ([int]$watch.ElapsedMilliseconds) `
                -Kind 'input' -EventsDirectory $EventsDirectory -EventsPath $EventsPath -Target $typingField `
                -Window $typingWindow -RedactTarget $typingField
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

    $reason = 'stopped'
    if ($index -ge $MaxEvents) { $reason = 'limit' }
    elseif ($watch.Elapsed.TotalMinutes -ge $MaxMinutes) { $reason = 'timeout' }
    $message = switch ($reason) {
        'limit'   { "記録の上限（$MaxEvents 件）に達したため終了しました" }
        'timeout' { "記録の上限（$MaxMinutes 分）に達したため終了しました" }
        default   { "$index 件の操作を記録しました" }
    }
    Write-MbRecordingStatus -StatusPath $StatusPath -JobId $JobId -State 'completed' -Count $index -Message $message -LastTarget $lastTarget
    return $index
}

Export-ModuleMember -Function @(
    'Initialize-MbRecorderNative',
    'Set-MbProcessDpiAware',
    'Initialize-MbRecorderUia',
    'Get-MbRecorderCapability',
    'Get-MbUiaTargetAtPoint',
    'Get-MbUiaFocusedElement',
    'Select-MbUiaTargetInfo',
    'Get-MbForegroundWindowInfo',
    'Get-MbVirtualScreenBounds',
    'Get-MbCaptureRegion',
    'Copy-MbScreenBitmap',
    'Save-MbBitmapRegion',
    'ConvertTo-MbRegionRect',
    'Test-MbUsableElementInfo',
    'Test-MbIgnoredWindow',
    'Get-MbWatchedTypingKeys',
    'Test-MbAsyncKeyStateDown',
    'Test-MbAsyncKeyStatePressed',
    'Invoke-MbRecordingLoop',
    'Write-MbRecordingStatus'
)
