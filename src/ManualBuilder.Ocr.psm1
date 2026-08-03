# 画面の文字を読む。
#
# Windows 10/11 に入っている Windows.Media.Ocr を PowerShell 5.1 から呼ぶ。
# 追加インストールも管理者権限も要らず、外部へ送信しないのが採用理由。
#
# 使う目的は2つ:
#   1. 赤枠を実際のボタンの輪郭へ寄せる。録画から求めた変化領域は「変わった場所」で
#      あって「ボタンの形」ではないため、そこにある文字の矩形と合わせると枠が締まる。
#   2. 操作対象の名前と画面の文字をCopilotへ渡す。Copilotが画像から推測するのではなく、
#      読み取り済みの文字を整えるだけで済むようにする。
#
# OCRが使えない環境（言語データが無い、WinRTが古い）では、この機能だけを止めて
# 録画からの手順づくり自体は続けられるように、例外ではなく理由つきの不可を返す。

$script:MbOcrInitialized = $false
$script:MbOcrEngine = $null
$script:MbOcrLanguage = ''
$script:MbOcrReason = ''
$script:MbOcrAsTask = $null

# WinRTの非同期処理をPS 5.1で待つための橋渡し。
# IAsyncOperation<T> を Task<T> へ変換する AsTask の総称メソッドを1度だけ取り出す。
function Get-MbOcrAsTaskMethod {
    if ($null -ne $script:MbOcrAsTask) { return $script:MbOcrAsTask }
    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    } | Select-Object -First 1
    if (-not $method) { throw 'WinRTの非同期処理を待つためのAsTaskが見つかりません。' }
    $script:MbOcrAsTask = $method
    return $method
}

function Wait-MbOcrOperation {
    param(
        [Parameter(Mandatory = $true)]$Operation,
        [Parameter(Mandatory = $true)][Type]$ResultType
    )

    $asTask = (Get-MbOcrAsTaskMethod).MakeGenericMethod($ResultType)
    $task = $asTask.Invoke($null, @($Operation))
    # Wait の戻り値は真偽値。パイプラインへ漏らさないよう明示的に捨てる。
    [void]$task.Wait(-1)
    if ($task.IsFaulted) { throw $task.Exception.GetBaseException() }
    return $task.Result
}

function Initialize-MbOcr {
    param([switch]$Force)

    if ($script:MbOcrInitialized -and -not $Force) { return ($null -ne $script:MbOcrEngine) }
    $script:MbOcrInitialized = $true
    $script:MbOcrEngine = $null
    $script:MbOcrLanguage = ''
    $script:MbOcrReason = ''

    try {
        # WinRT型の読み込み。Windows 10未満やDesktop Runtimeが無い環境ではここで失敗する。
        [void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
        [void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
        [void][Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime]
        [void][Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]
        [void](Get-MbOcrAsTaskMethod)
    } catch {
        $script:MbOcrReason = 'この環境ではWindowsの文字認識を利用できません。'
        return $false
    }

    # 日本語を最優先。無ければ利用者の言語設定、それも駄目なら英語。
    $engine = $null
    try {
        $japanese = New-Object Windows.Globalization.Language 'ja'
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($japanese)
    } catch { $engine = $null }
    if (-not $engine) {
        try { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() } catch { $engine = $null }
    }
    if (-not $engine) {
        try {
            $english = New-Object Windows.Globalization.Language 'en-US'
            $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($english)
        } catch { $engine = $null }
    }
    if (-not $engine) {
        $script:MbOcrReason = '文字認識に使える言語がインストールされていません。'
        return $false
    }

    $script:MbOcrEngine = $engine
    try { $script:MbOcrLanguage = [string]$engine.RecognizerLanguage.LanguageTag } catch { $script:MbOcrLanguage = '' }
    return $true
}

function Get-MbOcrStatus {
    $available = Initialize-MbOcr
    return [pscustomobject]@{
        available = $available
        language  = $script:MbOcrLanguage
        reason    = $script:MbOcrReason
    }
}

# 画像1枚を読み、語とその矩形（0〜1の正規化）を返す。
function Get-MbOcrSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$MaxTextLength = 2000
    )

    $empty = [pscustomobject]@{
        available = $false; reason = ''; language = $script:MbOcrLanguage
        words = @(); lines = @(); text = ''; width = 0; height = 0
    }

    if (-not (Initialize-MbOcr)) {
        $empty.reason = $script:MbOcrReason
        return $empty
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $empty.reason = '画像が見つかりません。'
        return $empty
    }

    $stream = $null
    $bitmap = $null
    $converted = $null
    try {
        $file = Wait-MbOcrOperation -Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) -ResultType ([Windows.Storage.StorageFile])
        $stream = Wait-MbOcrOperation -Operation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) -ResultType ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Wait-MbOcrOperation -Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) -ResultType ([Windows.Graphics.Imaging.BitmapDecoder])

        $width = [int]$decoder.PixelWidth
        $height = [int]$decoder.PixelHeight
        if ($width -lt 1 -or $height -lt 1) {
            $empty.reason = '画像の大きさを取得できませんでした。'
            return $empty
        }
        $limit = 0
        try { $limit = [int]$script:MbOcrEngine.MaxImageDimension } catch { $limit = 0 }
        if ($limit -gt 0 -and ([Math]::Max($width, $height) -gt $limit)) {
            $empty.reason = "画像が大きすぎて文字認識できません（上限 ${limit}px）。"
            $empty.width = $width
            $empty.height = $height
            return $empty
        }

        $bitmap = Wait-MbOcrOperation -Operation ($decoder.GetSoftwareBitmapAsync()) -ResultType ([Windows.Graphics.Imaging.SoftwareBitmap])
        # OCRが受け付ける画素形式へ揃える。JPEGの既定形式のままだと失敗することがある。
        $converted = [Windows.Graphics.Imaging.SoftwareBitmap]::Convert(
            $bitmap,
            [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)

        $result = Wait-MbOcrOperation -Operation ($script:MbOcrEngine.RecognizeAsync($converted)) -ResultType ([Windows.Media.Ocr.OcrResult])

        $words = New-Object System.Collections.ArrayList
        $lines = New-Object System.Collections.ArrayList
        foreach ($line in @($result.Lines)) {
            $lineText = [string]$line.Text
            if (-not [string]::IsNullOrWhiteSpace($lineText)) { [void]$lines.Add($lineText) }
            foreach ($word in @($line.Words)) {
                $rect = $word.BoundingRect
                [void]$words.Add([pscustomobject]@{
                    text = [string]$word.Text
                    x1   = [Math]::Round(([double]$rect.X / $width), 6)
                    y1   = [Math]::Round(([double]$rect.Y / $height), 6)
                    x2   = [Math]::Round((([double]$rect.X + [double]$rect.Width) / $width), 6)
                    y2   = [Math]::Round((([double]$rect.Y + [double]$rect.Height) / $height), 6)
                })
            }
        }

        $text = ($lines -join "`n")
        if ($MaxTextLength -gt 0 -and $text.Length -gt $MaxTextLength) {
            $text = $text.Substring(0, $MaxTextLength)
        }

        return [pscustomobject]@{
            available = $true
            reason    = ''
            language  = $script:MbOcrLanguage
            words     = @($words)
            lines     = @($lines)
            text      = $text
            width     = $width
            height    = $height
        }
    } catch {
        $empty.reason = '文字認識に失敗しました: ' + $_.Exception.Message
        return $empty
    } finally {
        foreach ($disposable in @($converted, $bitmap, $stream)) {
            if ($null -ne $disposable) { try { $disposable.Dispose() } catch { } }
        }
    }
}

function Get-MbRectCenter {
    param([Parameter(Mandatory = $true)]$Rect)
    return [pscustomobject]@{
        x = (([double]$Rect.x1 + [double]$Rect.x2) / 2)
        y = (([double]$Rect.y1 + [double]$Rect.y2) / 2)
    }
}

# 点から矩形までの距離。矩形の中なら0。
function Get-MbPointRectDistance {
    param([double]$X, [double]$Y, [Parameter(Mandatory = $true)]$Rect)
    $dx = [Math]::Max([Math]::Max([double]$Rect.x1 - $X, 0), $X - [double]$Rect.x2)
    $dy = [Math]::Max([Math]::Max([double]$Rect.y1 - $Y, 0), $Y - [double]$Rect.y2)
    return [Math]::Sqrt(($dx * $dx) + ($dy * $dy))
}

# 変化領域を、そこにある文字の矩形と突き合わせて赤枠の位置と操作対象の名前を決める。
#
#   inside  : 変化領域の中に文字がある。変化領域がボタンそのものなので枠はそのまま使い、
#             文字だけを操作対象の名前として取り出す。
#   nearest : 変化領域が縁の一部しか捉えていない場合。近くの文字と合わせて枠を広げる。
#   none    : 近くに文字が無い。アイコンだけのボタンなど。枠はそのままにして名前は空。
function Resolve-MbOperationRect {
    param(
        [Parameter(Mandatory = $true)]$Rect,
        [AllowNull()]$Snapshot,
        [double]$NearestLimit = 0.08
    )

    $clamp = { param([double]$Value) [Math]::Round([Math]::Min(1, [Math]::Max(0, $Value)), 6) }
    $result = [pscustomobject]@{
        rect    = [pscustomobject]@{
            x1 = & $clamp ([double]$Rect.x1); y1 = & $clamp ([double]$Rect.y1)
            x2 = & $clamp ([double]$Rect.x2); y2 = & $clamp ([double]$Rect.y2)
        }
        label   = ''
        matched = 'none'
    }
    if ($null -eq $Snapshot -or -not $Snapshot.available) { return $result }
    $words = @($Snapshot.words)
    if ($words.Count -eq 0) { return $result }

    $inside = @($words | Where-Object {
        $center = Get-MbRectCenter -Rect $_
        $center.x -ge [double]$Rect.x1 -and $center.x -le [double]$Rect.x2 -and
        $center.y -ge [double]$Rect.y1 -and $center.y -le [double]$Rect.y2
    })
    if ($inside.Count -gt 0) {
        # 読み順（上から、同じ行なら左から）に並べて名前にする。
        $ordered = @($inside | Sort-Object @{ Expression = { [Math]::Round([double]$_.y1, 2) } }, @{ Expression = { [double]$_.x1 } })
        $result.label = (($ordered | ForEach-Object { [string]$_.text }) -join ' ').Trim()
        $result.matched = 'inside'
        return $result
    }

    $center = Get-MbRectCenter -Rect $Rect
    $nearest = $null
    $nearestDistance = [double]::MaxValue
    foreach ($word in $words) {
        $distance = Get-MbPointRectDistance -X $center.x -Y $center.y -Rect $word
        if ($distance -lt $nearestDistance) {
            $nearestDistance = $distance
            $nearest = $word
        }
    }
    if ($null -eq $nearest -or $nearestDistance -gt $NearestLimit) { return $result }

    # 変化領域と文字の矩形を合わせた範囲を、文字の高さの4分の1だけ広げる。
    # ボタンは文字より一回り大きいため、この余白があると輪郭に近くなる。
    $pad = ([double]$nearest.y2 - [double]$nearest.y1) / 4
    $result.rect = [pscustomobject]@{
        x1 = & $clamp ([Math]::Min([double]$Rect.x1, [double]$nearest.x1) - $pad)
        y1 = & $clamp ([Math]::Min([double]$Rect.y1, [double]$nearest.y1) - $pad)
        x2 = & $clamp ([Math]::Max([double]$Rect.x2, [double]$nearest.x2) + $pad)
        y2 = & $clamp ([Math]::Max([double]$Rect.y2, [double]$nearest.y2) + $pad)
    }
    $result.label = ([string]$nearest.text).Trim()
    $result.matched = 'nearest'
    return $result
}

Export-ModuleMember -Function @(
    'Initialize-MbOcr',
    'Get-MbOcrStatus',
    'Get-MbOcrSnapshot',
    'Resolve-MbOperationRect'
)
