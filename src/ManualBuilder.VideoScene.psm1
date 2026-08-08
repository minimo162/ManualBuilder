# 録画から選ばれた1コマを、外部AIへ送らず編集可能な手順として取り込む。
#
# 実行時に使われる唯一の実装をここに置く。以前は Start-ManualBuilder.ps1 と
# ManualBuilder.CopilotServer.psm1 に同じ関数が二重にあり、回帰試験は
# 出荷されない側(CopilotServer)を検証していた。

Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Capture.psm1')
Import-Module (Join-Path $PSScriptRoot 'ManualBuilder.Ocr.psm1')

function Import-MbVideoScene {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [int]$TimeMs = 0,
        [AllowEmptyString()][string]$RectJson = '',
        [AllowEmptyString()][string]$CandidatesJson = '',
        [switch]$SkipOcr
    )

    $sha = [Security.Cryptography.SHA256]::Create()
    try { $sceneHash = [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
    $existingImage = @($Project.images | Where-Object { [string]$_.sha256 -eq $sceneHash }) | Select-Object -First 1
    if ($null -ne $existingImage) {
        $targetSheet = @($Project.sheets | Where-Object { [string]$_.id -eq $SheetId }) | Select-Object -First 1
        if ($null -ne $targetSheet) {
            foreach ($existingStep in @($targetSheet.steps | Where-Object { [string]$_.imageId -eq [string]$existingImage.id })) {
                if ($existingStep.PSObject.Properties.Name -contains 'capture' -and $null -ne $existingStep.capture -and
                    [string]$existingStep.capture.kind -eq 'video-scene' -and [int]$existingStep.capture.videoTimeMs -eq $TimeMs) {
                    return [pscustomobject]@{ status = 'duplicate'; stepId = [string]$existingStep.id; clickLabel = ''; ocrAvailable = $false }
                }
            }
        }
    }

    $added = Add-MbImageStep -Project $Project -ProjectPath $ProjectPath -SheetId $SheetId -Bytes $Bytes `
        -Source 'video' -AllowDuplicateStep
    if ([string]$added.Status -ne 'added') {
        return [pscustomobject]@{ status = [string]$added.Status; stepId = ''; clickLabel = ''; ocrAvailable = $false }
    }
    $stepId = [string]$added.Step.id

    $rect = $null
    if (-not [string]::IsNullOrWhiteSpace($RectJson)) {
        try { $rect = $RectJson | ConvertFrom-Json } catch { $rect = $null }
        if (-not (Test-MbNormalizedRect -Rect $rect)) { $rect = $null }
    }
    $candidates = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace($CandidatesJson)) {
        $parsed = $null
        try { $parsed = $CandidatesJson | ConvertFrom-Json } catch { $parsed = $null }
        foreach ($candidate in @(@($parsed) | Select-Object -First 4)) {
            if ($null -eq $candidate -or $candidate.PSObject.Properties.Name -notcontains 'rect' -or
                -not (Test-MbNormalizedRect -Rect $candidate.rect)) { continue }
            [void]$candidates.Add([pscustomobject]@{
                id = 'video-diff-' + ($candidates.Count + 1)
                source = 'video-diff'; confidence = 'low'; label = ''; targetType = ''; rect = $candidate.rect
            })
        }
    }
    if ($candidates.Count -eq 0 -and $null -ne $rect) {
        [void]$candidates.Add([pscustomobject]@{
            id = 'video-diff-1'; source = 'video-diff'; confidence = 'low'; label = ''; targetType = ''; rect = $rect
        })
    }

    $screenText = ''; $ocrAvailable = $false
    $imagePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId ([string]$added.Step.imageId)
    if (-not $SkipOcr -and -not [string]::IsNullOrWhiteSpace($imagePath)) {
        try {
            $snapshot = Get-MbOcrSnapshot -Path $imagePath
            $ocrAvailable = [bool]$snapshot.available
            if ($ocrAvailable) { $screenText = [string]$snapshot.text }
            if ($ocrAvailable) {
                foreach ($candidate in @($candidates)) {
                    $resolved = Resolve-MbOperationRect -Rect $candidate.rect -Snapshot $snapshot
                    $candidate.rect = $resolved.rect
                    $candidate.label = [string]$resolved.label
                }
            }
        } catch { $ocrAvailable = $false }
    }

    $captured = Set-MbStepCapture -Project $Project -StepId $stepId -Kind 'video-scene' -VideoTimeMs $TimeMs `
        -ClickLabel '' -ScreenText $screenText -TargetSource 'video-diff' `
        -TargetConfidence $(if ($candidates.Count -gt 0) { [string]$candidates[0].confidence } else { '' }) `
        -TargetCandidateId '' `
        -TargetCandidatesJson $(if ($candidates.Count -gt 0) { ConvertTo-Json -InputObject @($candidates) -Depth 8 -Compress } else { '' })
    # 動画差分だけでは正しい操作箇所を確定できないため、候補は保存しても赤枠は付けない。
    $captured.capture.targetCandidateId = ''
    [void](Set-MbStepDraft -Project $Project -StepId $stepId -Title '録画の場面を確認' `
        -Description '画面の内容を確認し、必要な操作を説明します。' -Note '')
    [void](Set-MbStepReview -Project $Project -StepId $stepId -Action 'review' `
        -Reason '録画の画面変化から作成した手順です。必要な場面か、文章と操作箇所を確認してください。')
    return [pscustomobject]@{
        status = 'added'; stepId = $stepId; clickLabel = ''; hasRect = $false; ocrAvailable = $ocrAvailable
    }
}

Export-ModuleMember -Function @('Import-MbVideoScene')
