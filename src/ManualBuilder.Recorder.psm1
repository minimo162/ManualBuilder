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

    $best = $null
    $bestArea = [double]::PositiveInfinity
    foreach ($candidate in @($Candidates)) {
        if (-not (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y -Tolerance 12.0)) { continue }
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

# 操作パターンを公開しないWebアプリでも、クリック点の下に名前付きText/Group等があれば
# その最小要素を対象として使う。ウィンドウの大半を占めるコンテナは採用しない。
function Select-MbUiaNamedTargetInfo {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Candidates,
        [Parameter(Mandatory = $true)][double]$X,
        [Parameter(Mandatory = $true)][double]$Y,
        [AllowNull()]$Window
    )

    $best = $null
    $bestArea = [double]::PositiveInfinity
    $windowArea = if ($null -ne $Window) { [double]$Window.width * [double]$Window.height } else { 0.0 }
    foreach ($candidate in @($Candidates)) {
        if (-not (Test-MbPointWithinElementInfo -Info $candidate -X $X -Y $Y -Tolerance 12.0)) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$candidate.name)) { continue }
        if ([string]$candidate.controlType -in @('ControlType.Window', 'ControlType.Document')) { continue }
        $area = [double]$candidate.width * [double]$candidate.height
        if ($windowArea -gt 0 -and $area -gt ($windowArea * 0.4)) { continue }
        if ($area -lt $bestArea) { $best = $candidate; $bestArea = $area }
    }
    if ($null -ne $best) {
        $best | Add-Member -NotePropertyName 'isInferred' -NotePropertyValue $true -Force
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
        isActionable = $true; isFallback = $true
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

function Get-MbForegroundWindowInfo {
    Initialize-MbRecorderNative
    $handle = [MbRecorderNative]::GetForegroundWindow()
    if ($handle -eq [IntPtr]::Zero) { return $null }
    $rect = [MbRecorderNative]::GetVisualWindowRect($handle)
    return [pscustomobject]@{
        handle = [long]$handle.ToInt64()
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
        $width = [double]$Snapshot.rect.width * $dpr
        $height = [double]$Snapshot.rect.height * $dpr
        $left = $X + (([double]$Snapshot.rect.left - $clientX) * $dpr)
        $top = $Y + (([double]$Snapshot.rect.top - $clientY) * $dpr)
        foreach ($number in @($clientX, $clientY, $width, $height, $left, $top)) {
            if ([double]::IsNaN([double]$number) -or [double]::IsInfinity([double]$number)) { return $null }
        }
        if ($width -le 1.0 -or $height -le 1.0 -or $width -gt 12000 -or $height -gt 12000) { return $null }
        if ($X -lt ($left - 16.0) -or $X -gt ($left + $width + 16.0) -or
            $Y -lt ($top - 16.0) -or $Y -gt ($top + $height + 16.0)) { return $null }

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
            left = $left; top = $top; width = $width; height = $height
            isActionable = $true; isFallback = $false
            provider = 'DOM'; confidence = 'high'
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
        $cursorDistance = [Math]::Sqrt(
            [Math]::Pow(([double]$cache.cursorX - $X), 2.0) +
            [Math]::Pow(([double]$cache.cursorY - $Y), 2.0)
        )
        if ($cursorDistance -gt 36.0) { return $null }

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
        $titleMatches = [string]::IsNullOrWhiteSpace($pageTitle) -and [string]::IsNullOrWhiteSpace($snapshotTitle)
        if (-not [string]::IsNullOrWhiteSpace($pageTitle) -and
            $windowTitle.IndexOf($pageTitle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $titleMatches = $true }
        if (-not [string]::IsNullOrWhiteSpace($snapshotTitle) -and
            $windowTitle.IndexOf($snapshotTitle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $titleMatches = $true }
        if (-not $titleMatches) {
            return $null
        }
        return (ConvertFrom-MbDomSnapshotTarget -Snapshot $snapshot -X $X -Y $Y)
    } catch { return $null }
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
        index       = $Index
        kind        = $Kind
        timeMs      = $ElapsedMs
        image       = $fileName
        windowTitle = $windowTitle
        targetName  = $targetName
        targetType  = $targetType
        rect        = $rect
    }
    if ($null -ne $Target -and $Target.PSObject.Properties.Name -contains 'provider') {
        $record.targetSource = [string]$Target.provider
    }
    if ($null -ne $Target -and $Target.PSObject.Properties.Name -contains 'confidence') {
        $record.confidence = [string]$Target.confidence
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
        [AllowEmptyString()][string]$DomTargetPath = '',
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
    $typingCaptureAtMs = 0
    $lastTypingMs = 0
    $lastDomTarget = $null
    $lastDomWindowHandle = 0L

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
        $typingPressed = $false
        foreach ($vk in $typingKeys) {
            $keyState = [int][MbRecorderNative]::GetAsyncKeyState($vk)
            $keyPressed = Test-MbAsyncKeyStatePressed -State $keyState
            if ((Test-MbAsyncKeyStateDown -State $keyState) -or $keyPressed) {
                $typingNow = $true
            }
            if ($keyPressed) { $typingPressed = $true }
        }
        if ($typingNow) {
            if (-not $typingActive) {
                $typingActive = $true
                # キーそのものは読まない。入力中の画面だけを保持し、表示された文字を
                # 残すか隠すかは、取り込み後の画像編集（黒塗り）で利用者が決める。
                try {
                    $candidateWindow = Get-MbForegroundWindowInfo
                    if (-not (Test-MbIgnoredWindow -Window $candidateWindow -IgnoreTitlePatterns $IgnoreTitlePatterns)) {
                        $typingCapture = Copy-MbScreenBitmap
                        $typingCaptureAtMs = [int]$watch.ElapsedMilliseconds
                        $typingWindow = $candidateWindow
                        $typingField = Get-MbUiaFocusedElement
                        if ($null -ne $lastDomTarget -and
                            [string]$lastDomTarget.controlType -eq 'ControlType.Edit' -and
                            $null -ne $candidateWindow -and [long]$candidateWindow.handle -eq $lastDomWindowHandle) {
                            $typingField = $lastDomTarget
                        }
                    }
                } catch {
                    if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
                    $typingCapture = $null
                    $typingWindow = $null
                    $typingField = $null
                }
            } elseif ($typingPressed -and (([int]$watch.ElapsedMilliseconds - $typingCaptureAtMs) -ge 120)) {
                # クリックで画面遷移する直前にも、入力済みの表示がなるべく残るよう更新する。
                $replacementCapture = $null
                try {
                    $replacementCapture = Copy-MbScreenBitmap
                    if ($null -ne $typingCapture) { try { $typingCapture.bitmap.Dispose() } catch { } }
                    $typingCapture = $replacementCapture
                    $replacementCapture = $null
                    $typingCaptureAtMs = [int]$watch.ElapsedMilliseconds
                } catch {
                    if ($null -ne $replacementCapture) { try { $replacementCapture.bitmap.Dispose() } catch { } }
                }
            }
            $lastTypingMs = [int]$watch.ElapsedMilliseconds
        }

        # 入力が途切れたか、次のクリックが来たら、保持していた最新の入力画面を1手順にする。
        $typingFinished = $typingActive -and ($clicked -or (([int]$watch.ElapsedMilliseconds - $lastTypingMs) -ge $TypingIdleMs))
        if ($typingFinished) {
            $typingActive = $false
            try {
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
                    $record = Save-MbRecordingEvent -Capture $typingCapture -Index $index -ElapsedMs ([int]$watch.ElapsedMilliseconds) `
                        -Kind 'input' -EventsDirectory $EventsDirectory -EventsPath $EventsPath -Target $typingField `
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
                    $target = $null
                    if (-not [string]::IsNullOrWhiteSpace($DomTargetPath)) {
                        # 画面はすでに押下直前で確保済み。Edgeからpointerdown通知が届くまで
                        # 最大約120msだけ待ち、速いマウス移動直後のクリックもDOMで拾う。
                        for ($domAttempt = 0; $domAttempt -lt 7 -and $null -eq $target; $domAttempt++) {
                            if ($domAttempt -gt 0) { Start-Sleep -Milliseconds 20 }
                            $target = Get-MbDomTargetFromCache -Path $DomTargetPath -X ([int]$point.X) -Y ([int]$point.Y) -Window $window
                        }
                    }
                    if ($null -eq $target) {
                        $target = Get-MbUiaTargetAtPoint -X ([int]$point.X) -Y ([int]$point.Y) -Window $window
                    }
                    if ($null -ne $target -and $target.PSObject.Properties.Name -contains 'provider' -and
                        [string]$target.provider -eq 'DOM') {
                        $lastDomTarget = $target
                        $lastDomWindowHandle = [long]$window.handle
                    } else {
                        $lastDomTarget = $null
                        $lastDomWindowHandle = 0L
                    }
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

    # 停止要求が入力の途中で届いても、保持していた入力画面を最後の1手順として残す。
    if ($typingActive -and $null -ne $typingCapture -and $index -lt $MaxEvents) {
        try {
            $index++
            $record = Save-MbRecordingEvent -Capture $typingCapture -Index $index -ElapsedMs ([int]$watch.ElapsedMilliseconds) `
                -Kind 'input' -EventsDirectory $EventsDirectory -EventsPath $EventsPath -Target $typingField `
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
    'ConvertFrom-MbDomSnapshotTarget',
    'Get-MbDomTargetFromCache',
    'Select-MbUiaTargetInfo',
    'Select-MbUiaNamedTargetInfo',
    'New-MbMsaaElementInfo',
    'New-MbClickPointTargetInfo',
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
    'Test-MbAsyncKeyStateDown',
    'Test-MbAsyncKeyStatePressed',
    'Invoke-MbRecordingLoop',
    'Write-MbRecordingStatus'
)
