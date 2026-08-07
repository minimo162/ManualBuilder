# 操作しながら話した内容を文字にする。
#
# Win+H（音声入力）が使っているのと同じ Windows.Media.SpeechRecognition を呼ぶ。
# System.Speech（旧SAPI）より日本語の精度が段違いに高い。
#
# なぜ操作記録モードでだけ使えるのか:
#   Windows.Media.SpeechRecognition には、System.Speech の SetInputToWaveFile に
#   相当するものが無い。入力は既定のマイクに固定されており、録画ファイルから
#   取り出した音声を流し込むことはできない。
#   操作記録モードは実時間で走っているので、その制約に当たらない。
#
# なぜ別プロセスなのか:
#   RecognizeAsync は発話が終わるまで待つ。60Hzでマウスを見ている記録ループと
#   同居させられない。またWinRTのイベント購読はPS 5.1から扱いづらく、C#で受け口を
#   書こうにも Windows.winmd の参照にWindows SDKが要る。認識だけを別プロセスにして
#   結果をファイルへ落とし、取り込み時に時刻で突き合わせるほうが確実に動く。
#
# 送信先について:
#   ディクテーションはMicrosoftのオンライン音声認識を使う。音声は端末の外へ出る。
#   既定では無効にし、記録のたびに利用者が選ぶ。

Set-StrictMode -Version 2.0

$script:MbDictationReady = $false
$script:MbDictationAsTask = $null
$script:MbDictationReason = ''
$script:MbDictationLanguage = ''

# WinRTの非同期処理をPS 5.1で待つための橋渡し。
# Ocrモジュールにも同じものがあるが、モジュール間の依存を作らないためここにも持つ。
function Get-MbDictationAsTaskMethod {
    if ($null -ne $script:MbDictationAsTask) { return $script:MbDictationAsTask }
    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    } | Select-Object -First 1
    if (-not $method) { throw 'WinRTの非同期処理を待つためのAsTaskが見つかりません。' }
    $script:MbDictationAsTask = $method
    return $method
}

function Wait-MbDictationOperation {
    param(
        [Parameter(Mandatory = $true)]$Operation,
        [Parameter(Mandatory = $true)][Type]$ResultType
    )
    $asTask = (Get-MbDictationAsTaskMethod).MakeGenericMethod($ResultType)
    $task = $asTask.Invoke($null, @($Operation))
    # Wait の戻り値は真偽値。パイプラインへ漏らさないよう明示的に捨てる。
    [void]$task.Wait(-1)
    if ($task.IsFaulted) { throw $task.Exception.GetBaseException() }
    return $task.Result
}

function Initialize-MbDictation {
    if ($script:MbDictationReady) { return $true }
    try {
        [void][Windows.Media.SpeechRecognition.SpeechRecognizer, Windows.Foundation, ContentType = WindowsRuntime]
        [void][Windows.Media.SpeechRecognition.SpeechRecognitionTopicConstraint, Windows.Foundation, ContentType = WindowsRuntime]
        [void][Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]
        [void](Get-MbDictationAsTaskMethod)
        $script:MbDictationReady = $true
        return $true
    } catch {
        $script:MbDictationReason = 'この環境ではWindowsの音声入力を利用できません。'
        return $false
    }
}

# 日本語のディクテーションが使えるかを調べる。
# 使えない理由はそのまま画面に出すので、対処が分かる文にする。
function Get-MbDictationCapability {
    if (-not (Initialize-MbDictation)) {
        return [pscustomobject]@{ available = $false; reason = $script:MbDictationReason; language = '' }
    }

    $japanese = $null
    try {
        $supported = [Windows.Media.SpeechRecognition.SpeechRecognizer]::SupportedTopicLanguages
        foreach ($language in @($supported)) {
            if ([string]$language.LanguageTag -like 'ja*') { $japanese = $language; break }
        }
    } catch {
        return [pscustomobject]@{
            available = $false
            reason = 'この端末では音声入力の対応言語を確認できませんでした。'
            language = ''
        }
    }
    if ($null -eq $japanese) {
        # オンライン音声認識が切られていると、対応言語の一覧自体が空になる。
        return [pscustomobject]@{
            available = $false
            reason = '日本語の音声入力が使えません。［設定］→［プライバシーとセキュリティ］→［音声認識］で「オンライン音声認識」をオンにしてください。'
            language = ''
        }
    }
    $script:MbDictationLanguage = [string]$japanese.LanguageTag
    return [pscustomobject]@{ available = $true; reason = ''; language = $script:MbDictationLanguage }
}

function New-MbDictationRecognizer {
    if (-not (Initialize-MbDictation)) { throw $script:MbDictationReason }
    $capability = Get-MbDictationCapability
    if (-not $capability.available) { throw ([string]$capability.reason) }

    $language = New-Object Windows.Globalization.Language -ArgumentList ([string]$capability.language)
    $recognizer = New-Object Windows.Media.SpeechRecognition.SpeechRecognizer -ArgumentList $language

    # 口述筆記の制約。これがオンラインの認識サービスを使う指定になる。
    $constraint = New-Object Windows.Media.SpeechRecognition.SpeechRecognitionTopicConstraint -ArgumentList @(
        [Windows.Media.SpeechRecognition.SpeechRecognitionScenario]::Dictation, 'dictation')
    $recognizer.Constraints.Add($constraint)

    # 黙っている間に長く待たせない。待ち続けると次の発話を取り逃す。
    try {
        $recognizer.Timeouts.InitialSilenceTimeout = [TimeSpan]::FromSeconds(6)
        $recognizer.Timeouts.EndSilenceTimeout = [TimeSpan]::FromSeconds(1)
        $recognizer.Timeouts.BabbleTimeout = [TimeSpan]::FromSeconds(4)
    } catch {
        # 値の範囲は環境によって異なる。既定のままでも動く。
    }

    $compiled = Wait-MbDictationOperation -Operation ($recognizer.CompileConstraintsAsync()) `
        -ResultType ([Windows.Media.SpeechRecognition.SpeechRecognitionCompilationResult])
    $status = [string]$compiled.Status
    if ($status -ne 'Success') {
        try { $recognizer.Dispose() } catch { }
        if ($status -eq 'TopicLanguageNotSupported') {
            throw '日本語の音声入力を準備できませんでした。［設定］→［プライバシーとセキュリティ］→［音声認識］で「オンライン音声認識」をオンにしてください。'
        }
        throw ("音声入力を準備できませんでした（$status）。")
    }
    return $recognizer
}

# 認識した発話を1件ぶんの記録へ直す。
# 受信した時刻ではなく PhraseStartTime を使う。認識の遅れに引きずられないため。
function ConvertTo-MbDictationRecord {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)][DateTime]$StartedAtUtc,
        [Parameter(Mandatory = $true)][int]$ReceivedAtMs
    )

    $text = ''
    try { $text = [string]$Result.Text } catch { $text = '' }
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $durationMs = 0
    try { $durationMs = [int]([TimeSpan]$Result.PhraseDuration).TotalMilliseconds } catch { $durationMs = 0 }

    $startMs = -1
    try {
        $phraseStart = [DateTimeOffset]$Result.PhraseStartTime
        if ($phraseStart.Ticks -gt 0) {
            $startMs = [int]($phraseStart.UtcDateTime - $StartedAtUtc).TotalMilliseconds
        }
    } catch {
        $startMs = -1
    }
    # PhraseStartTime が取れない環境では、受信時刻から発話の長さを引いて代用する。
    if ($startMs -lt 0) { $startMs = [Math]::Max(0, $ReceivedAtMs - $durationMs) }
    if ($durationMs -le 0) { $durationMs = 1000 }

    $confidence = ''
    try { $confidence = [string]$Result.Confidence } catch { $confidence = '' }

    return [pscustomobject]@{
        startMs    = [Math]::Max(0, $startMs)
        endMs      = [Math]::Max(0, $startMs) + $durationMs
        text       = $text.Trim()
        confidence = $confidence
    }
}

function Invoke-MbDictationLoop {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$StopPath,
        [AllowEmptyString()][string]$PausePath = '',
        [Parameter(Mandatory = $true)][long]$StartedAtUtcTicks,
        [int]$MaxPhrases = 2000
    )

    $startedAtUtc = New-Object DateTime -ArgumentList @($StartedAtUtcTicks, [DateTimeKind]::Utc)
    $recognizer = New-MbDictationRecognizer
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $encoding = New-Object Text.UTF8Encoding($false)
    $count = 0

    try {
        while ($count -lt $MaxPhrases) {
            if (Test-Path -LiteralPath $StopPath -PathType Leaf) { break }
            if (-not [string]::IsNullOrWhiteSpace($PausePath) -and
                (Test-Path -LiteralPath $PausePath -PathType Leaf)) {
                Start-Sleep -Milliseconds 200
                continue
            }

            $result = $null
            try {
                $result = Wait-MbDictationOperation -Operation ($recognizer.RecognizeAsync()) `
                    -ResultType ([Windows.Media.SpeechRecognition.SpeechRecognitionResult])
            } catch {
                # 発話が無いまま待ち時間が過ぎた場合もここへ来る。止めずに待ち直す。
                Start-Sleep -Milliseconds 200
                continue
            }
            if ($null -eq $result) { continue }
            # 認識待ちの途中で一時停止された発話は、操作ログと対応しないため保存しない。
            if (-not [string]::IsNullOrWhiteSpace($PausePath) -and
                (Test-Path -LiteralPath $PausePath -PathType Leaf)) { continue }

            # 認識できなかった発話は捨てる。誤った文をCopilotへ渡すより空のほうがよい。
            $status = ''
            try { $status = [string]$result.Status } catch { $status = '' }
            if ($status -ne 'Success') { continue }
            $confidence = ''
            try { $confidence = [string]$result.Confidence } catch { $confidence = '' }
            if ($confidence -eq 'Rejected') { continue }

            $record = ConvertTo-MbDictationRecord -Result $result -StartedAtUtc $startedAtUtc -ReceivedAtMs ([int]$watch.ElapsedMilliseconds)
            if ($null -eq $record) { continue }
            [IO.File]::AppendAllText($OutputPath, (($record | ConvertTo-Json -Compress) + [Environment]::NewLine), $encoding)
            $count++
        }
    } finally {
        try { $recognizer.Dispose() } catch { }
    }
    return $count
}

Export-ModuleMember -Function @(
    'Initialize-MbDictation',
    'Get-MbDictationCapability',
    'New-MbDictationRecognizer',
    'ConvertTo-MbDictationRecord',
    'Invoke-MbDictationLoop'
)
