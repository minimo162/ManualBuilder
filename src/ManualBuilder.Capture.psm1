# ManualBuilder image validation and project-local storage.

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Drawing

function Get-MbImageKind {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    if ($Bytes.Length -ge 8 -and
        $Bytes[0] -eq 0x89 -and $Bytes[1] -eq 0x50 -and $Bytes[2] -eq 0x4E -and $Bytes[3] -eq 0x47 -and
        $Bytes[4] -eq 0x0D -and $Bytes[5] -eq 0x0A -and $Bytes[6] -eq 0x1A -and $Bytes[7] -eq 0x0A) {
        return [pscustomobject]@{ Extension = 'png'; MimeType = 'image/png' }
    }
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xD8 -and $Bytes[2] -eq 0xFF) {
        return [pscustomobject]@{ Extension = 'jpg'; MimeType = 'image/jpeg' }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0x42 -and $Bytes[1] -eq 0x4D) {
        return [pscustomobject]@{ Extension = 'bmp'; MimeType = 'image/bmp' }
    }
    return $null
}

function Get-MbByteSha256 {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '')
    } finally {
        $sha.Dispose()
    }
}

function Get-MbImageDimensions {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    $stream = New-Object IO.MemoryStream
    $image = $null
    try {
        $stream.Write($Bytes, 0, $Bytes.Length)
        $stream.Position = 0
        $image = [Drawing.Image]::FromStream($stream, $true, $true)
        $width = [int]$image.Width
        $height = [int]$image.Height
        if ($width -lt 1 -or $height -lt 1 -or $width -gt 12000 -or $height -gt 12000) {
            throw "画像サイズ ${width}x${height}px は範囲外です（各辺12000pxまで）。"
        }
        if (([long]$width * [long]$height) -gt 100000000) {
            throw "画像の総画素数が大きすぎます: ${width}x${height}px"
        }
        return [pscustomobject]@{ Width = $width; Height = $height }
    } catch {
        throw "画像として読み込めません: $($_.Exception.Message)"
    } finally {
        if ($image) { $image.Dispose() }
        $stream.Dispose()
    }
}

function Get-MbImageDirectory {
    param([Parameter(Mandatory = $true)][string]$ProjectPath)
    return Join-Path (Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))) 'images'
}

function Get-MbImageFilePath {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$ImageId
    )

    $image = @($Project.images | Where-Object { $_.id -eq $ImageId }) | Select-Object -First 1
    if (-not $image) { return $null }
    if ([string]$image.fileName -notmatch '^image-[a-f0-9]{32}\.(png|jpg|bmp)$') { return $null }
    return Join-Path (Get-MbImageDirectory -ProjectPath $ProjectPath) ([string]$image.fileName)
}

function Add-MbImageAsset {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [ValidateSet('watcher', 'paste', 'drop', 'file', 'video', 'recorder')][string]$Source = 'file'
    )

    if ($Bytes.Length -lt 1) { throw '画像データが空です。' }
    if ($Bytes.Length -gt (20 * 1024 * 1024)) { throw '画像は20MB以下にしてください。' }
    $kind = Get-MbImageKind -Bytes $Bytes
    if (-not $kind) { throw 'PNG、JPEG、BMP画像に対応しています。' }

    $hash = Get-MbByteSha256 -Bytes $Bytes
    $existing = @($Project.images | Where-Object { $_.sha256 -eq $hash }) | Select-Object -First 1
    if ($existing) {
        return [pscustomobject]@{ Status = 'existing'; Image = $existing }
    }

    $dimensions = Get-MbImageDimensions -Bytes $Bytes
    $imageId = 'image-' + [guid]::NewGuid().ToString('N')
    $fileName = "$imageId.$($kind.Extension)"
    $imageDirectory = Get-MbImageDirectory -ProjectPath $ProjectPath
    if (-not (Test-Path -LiteralPath $imageDirectory)) {
        [void](New-Item -ItemType Directory -Path $imageDirectory -Force)
    }
    $destination = Join-Path $imageDirectory $fileName
    $temporary = Join-Path $imageDirectory ('.image-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllBytes($temporary, $Bytes)
        [IO.File]::Move($temporary, $destination)
    } finally {
        if (Test-Path -LiteralPath $temporary) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    }

    $now = [DateTime]::UtcNow.ToString('o')
    $image = [pscustomobject]@{
        id         = $imageId
        fileName   = $fileName
        sha256     = $hash
        width      = $dimensions.Width
        height     = $dimensions.Height
        byteLength = [long]$Bytes.Length
        mimeType   = $kind.MimeType
        source     = $Source
        createdAt  = $now
    }
    $Project.images = @($Project.images) + @($image)
    return [pscustomobject]@{ Status = 'added'; Image = $image }
}

function Add-MbImageStep {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$SheetId,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [ValidateSet('watcher', 'paste', 'drop', 'file', 'video', 'recorder')][string]$Source = 'file',
        [switch]$AllowDuplicateStep = $false
    )

    $targetSheet = @($Project.sheets | Where-Object { $_.id -eq $SheetId }) | Select-Object -First 1
    if (-not $targetSheet) { throw '対象シートが見つかりません。' }
    if (@($targetSheet.steps).Count -ge 500) { throw '1シートの手順は500件までです。' }
    $asset = Add-MbImageAsset -Project $Project -ProjectPath $ProjectPath -Bytes $Bytes -Source $Source
    if ($asset.Status -eq 'existing' -and -not $AllowDuplicateStep) {
        return [pscustomobject]@{ Status = 'duplicate'; Step = $null; Image = $asset.Image }
    }

    $now = [DateTime]::UtcNow.ToString('o')
    $step = Add-MbStep -Project $Project -SheetId $SheetId
    $step.imageId = [string]$asset.Image.id
    $step.updatedAt = $now
    return [pscustomobject]@{ Status = 'added'; Step = $step; Image = $asset.Image }
}

function Set-MbStepResultImage {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$StepId,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [ValidateSet('paste', 'drop', 'file', 'recorder')][string]$Source = 'file'
    )
    $step = Get-MbStepById -Project $Project -StepId $StepId
    if (-not $step) { throw '対象手順が見つかりません。' }
    if ([string]::IsNullOrWhiteSpace([string]$step.imageId)) { throw '先に1枚目の画像を追加してください。' }
    $asset = Add-MbImageAsset -Project $Project -ProjectPath $ProjectPath -Bytes $Bytes -Source $Source
    if ([string]$step.resultImageId -eq [string]$asset.Image.id) {
        return [pscustomobject]@{ Status = 'duplicate'; Step = $step; Image = $asset.Image; RemovedPath = $null }
    }
    $previousImageId = [string]$step.resultImageId
    $step.resultImageId = [string]$asset.Image.id
    $step.resultAnnotations = @()
    $step.resultCrop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
    if ([string]$step.imageLayout -eq 'before') { $step.imageLayout = 'side-by-side' }
    $step.updatedAt = [DateTime]::UtcNow.ToString('o')
    $removedPath = Remove-MbUnusedImage -Project $Project -ProjectPath $ProjectPath -ImageId $previousImageId
    return [pscustomobject]@{ Status = 'set'; Step = $step; Image = $asset.Image; RemovedPath = $removedPath }
}

function Remove-MbStepResultImage {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$StepId
    )
    $step = Get-MbStepById -Project $Project -StepId $StepId
    if (-not $step) { throw '対象手順が見つかりません。' }
    $removedImageId = [string]$step.resultImageId
    $step.resultImageId = $null
    $step.resultAnnotations = @()
    $step.resultCrop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
    $step.imageLayout = 'before'
    $step.imageOrder = 'before-after'
    $step.updatedAt = [DateTime]::UtcNow.ToString('o')
    $removedPath = Remove-MbUnusedImage -Project $Project -ProjectPath $ProjectPath -ImageId $removedImageId
    return [pscustomobject]@{ Step = $step; RemovedPath = $removedPath }
}

function Set-MbStepImage {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$StepId,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [ValidateSet('paste', 'drop', 'file')][string]$Source = 'file'
    )

    $step = $null
    foreach ($sheet in @($Project.sheets)) {
        $step = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($step) { break }
    }
    if (-not $step) { throw '対象手順が見つかりません。' }

    $asset = Add-MbImageAsset -Project $Project -ProjectPath $ProjectPath -Bytes $Bytes -Source $Source
    if ([string]$step.imageId -eq [string]$asset.Image.id) {
        return [pscustomobject]@{ Status = 'duplicate'; Step = $step; Image = $asset.Image; Previous = $null }
    }

    $previous = [pscustomobject]@{
        imageId     = [string]$step.imageId
        annotations = @($step.annotations)
        crop        = $step.crop
    }
    $step.imageId = [string]$asset.Image.id
    $step.annotations = @()
    $step.crop = [pscustomobject]@{ x = 0.0; y = 0.0; width = 1.0; height = 1.0 }
    $step.updatedAt = [DateTime]::UtcNow.ToString('o')
    return [pscustomobject]@{
        Status   = 'replaced'
        Step     = $step
        Image    = $asset.Image
        Previous = $previous
        Created  = ($asset.Status -eq 'added')
    }
}

function Restore-MbStepImage {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$StepId,
        [Parameter(Mandatory = $true)][object]$Previous
    )

    $step = $null
    foreach ($sheet in @($Project.sheets)) {
        $step = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($step) { break }
    }
    if (-not $step) { throw '対象手順が見つかりません。' }
    $previousImageId = [string]$Previous.imageId
    if ([string]::IsNullOrWhiteSpace($previousImageId) -or -not (@($Project.images | Where-Object { $_.id -eq $previousImageId }) | Select-Object -First 1)) {
        throw '元の画像が見つからないため復元できません。'
    }

    $replacedImageId = [string]$step.imageId
    $step.imageId = $previousImageId
    $step.annotations = @($Previous.annotations)
    $step.crop = $Previous.crop
    $step.updatedAt = [DateTime]::UtcNow.ToString('o')
    return [pscustomobject]@{ Step = $step; ReplacedImageId = $replacedImageId }
}

function Remove-MbUnusedImage {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [AllowEmptyString()][string]$ImageId
    )

    if ([string]::IsNullOrWhiteSpace($ImageId)) { return $null }
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            if ([string]$step.imageId -eq $ImageId) { return $null }
            if ($step.PSObject.Properties.Name -contains 'resultImageId' -and
                [string]$step.resultImageId -eq $ImageId) { return $null }
        }
    }

    $imagePath = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId $ImageId
    $before = @($Project.images).Count
    $Project.images = @($Project.images | Where-Object { $_.id -ne $ImageId })
    if (@($Project.images).Count -eq $before) { return $null }
    return $imagePath
}

function Remove-MbUnreferencedImages {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath
    )

    $referenced = @{}
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            $imageId = [string]$step.imageId
            if (-not [string]::IsNullOrWhiteSpace($imageId)) { $referenced[$imageId] = $true }
            if ($step.PSObject.Properties.Name -contains 'resultImageId') {
                $resultImageId = [string]$step.resultImageId
                if (-not [string]::IsNullOrWhiteSpace($resultImageId)) { $referenced[$resultImageId] = $true }
            }
        }
    }

    $paths = New-Object System.Collections.ArrayList
    foreach ($image in @($Project.images)) {
        $imageId = [string]$image.id
        if ($referenced.ContainsKey($imageId)) { continue }
        $path = Get-MbImageFilePath -Project $Project -ProjectPath $ProjectPath -ImageId $imageId
        if ($path) { [void]$paths.Add($path) }
    }
    $Project.images = @($Project.images | Where-Object { $referenced.ContainsKey([string]$_.id) })
    return @($paths)
}

# --- 動画（Excel出力から再生する） -----------------------------------------
# 動画はブラウザーへ配信しない。手順カードにはコマから作った静止画を出し、
# 動画本体はExcel出力のときだけファイルとして読む。
# これにより単一スレッドのHttpListenerで大きな配信が走らず、Range要求も不要になる。

$script:MbVideoMaxBytes = 30 * 1024 * 1024

function Get-MbVideoKind {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    if ($Bytes.Length -ge 12 -and
        $Bytes[4] -eq 0x66 -and $Bytes[5] -eq 0x74 -and $Bytes[6] -eq 0x79 -and $Bytes[7] -eq 0x70) {
        return [pscustomobject]@{ Extension = 'mp4'; MimeType = 'video/mp4' }
    }
    if ($Bytes.Length -ge 4 -and
        $Bytes[0] -eq 0x1A -and $Bytes[1] -eq 0x45 -and $Bytes[2] -eq 0xDF -and $Bytes[3] -eq 0xA3) {
        return [pscustomobject]@{ Extension = 'webm'; MimeType = 'video/webm' }
    }
    return $null
}

function Get-MbVideoDirectory {
    param([Parameter(Mandatory = $true)][string]$ProjectPath)
    return Join-Path (Split-Path -Parent ([IO.Path]::GetFullPath($ProjectPath))) 'videos'
}

function Get-MbVideoFilePath {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$VideoId
    )

    $video = @($Project.videos | Where-Object { $_.id -eq $VideoId }) | Select-Object -First 1
    if (-not $video) { return $null }
    if ([string]$video.fileName -notmatch '^video-[a-f0-9]{32}\.(mp4|webm)$') { return $null }
    return Join-Path (Get-MbVideoDirectory -ProjectPath $ProjectPath) ([string]$video.fileName)
}

function Add-MbVideoAsset {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [double]$DurationSec = 0
    )

    if ($Bytes.Length -lt 1) { throw '動画データが空です。' }
    if ($Bytes.Length -gt $script:MbVideoMaxBytes) { throw '動画は30MB以下にしてください。短く撮り直すか、解像度を下げてください。' }
    $kind = Get-MbVideoKind -Bytes $Bytes
    if (-not $kind) { throw 'mp4（H.264）またはwebmの動画に対応しています。' }
    if ([double]::IsNaN($DurationSec) -or [double]::IsInfinity($DurationSec) -or $DurationSec -lt 0 -or $DurationSec -gt 3600) {
        throw '動画の長さが範囲外です。'
    }

    $hash = Get-MbByteSha256 -Bytes $Bytes
    $existing = @($Project.videos | Where-Object { $_.sha256 -eq $hash }) | Select-Object -First 1
    if ($existing) {
        return [pscustomobject]@{ Status = 'existing'; Video = $existing }
    }
    if (@($Project.videos).Count -ge 50) { throw '動画は1マニュアル50本までです。' }

    $videoId = 'video-' + [guid]::NewGuid().ToString('N')
    $fileName = "$videoId.$($kind.Extension)"
    $videoDirectory = Get-MbVideoDirectory -ProjectPath $ProjectPath
    if (-not (Test-Path -LiteralPath $videoDirectory)) {
        [void](New-Item -ItemType Directory -Path $videoDirectory -Force)
    }
    $destination = Join-Path $videoDirectory $fileName
    $temporary = Join-Path $videoDirectory ('.video-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllBytes($temporary, $Bytes)
        [IO.File]::Move($temporary, $destination)
    } finally {
        if (Test-Path -LiteralPath $temporary) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    }

    $video = [pscustomobject]@{
        id          = $videoId
        fileName    = $fileName
        sha256      = $hash
        byteLength  = [long]$Bytes.Length
        mimeType    = $kind.MimeType
        durationSec = [Math]::Round([double]$DurationSec, 1)
        createdAt   = [DateTime]::UtcNow.ToString('o')
    }
    $Project.videos = @($Project.videos) + @($video)
    return [pscustomobject]@{ Status = 'added'; Video = $video }
}

function Set-MbStepVideo {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$StepId,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [double]$DurationSec = 0
    )

    $step = $null
    foreach ($sheet in @($Project.sheets)) {
        $step = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($step) { break }
    }
    if (-not $step) { throw '対象手順が見つかりません。' }

    $previousVideoId = [string]$step.videoId
    $asset = Add-MbVideoAsset -Project $Project -ProjectPath $ProjectPath -Bytes $Bytes -DurationSec $DurationSec
    $step.videoId = [string]$asset.Video.id
    $step.updatedAt = [DateTime]::UtcNow.ToString('o')
    $removedPath = $null
    if ($previousVideoId -and $previousVideoId -ne [string]$asset.Video.id) {
        $removedPath = Remove-MbUnusedVideo -Project $Project -ProjectPath $ProjectPath -VideoId $previousVideoId
    }
    return [pscustomobject]@{ Step = $step; Video = $asset.Video; RemovedPath = $removedPath }
}

function Remove-MbStepVideo {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [Parameter(Mandatory = $true)][string]$StepId
    )

    $step = $null
    foreach ($sheet in @($Project.sheets)) {
        $step = @($sheet.steps | Where-Object { $_.id -eq $StepId }) | Select-Object -First 1
        if ($step) { break }
    }
    if (-not $step) { throw '対象手順が見つかりません。' }
    $videoId = [string]$step.videoId
    if (-not $videoId) { return [pscustomobject]@{ Step = $step; RemovedPath = $null } }
    $step.videoId = $null
    $step.updatedAt = [DateTime]::UtcNow.ToString('o')
    return [pscustomobject]@{
        Step        = $step
        RemovedPath = (Remove-MbUnusedVideo -Project $Project -ProjectPath $ProjectPath -VideoId $videoId)
    }
}

function Remove-MbUnusedVideo {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath,
        [AllowEmptyString()][string]$VideoId
    )

    if ([string]::IsNullOrWhiteSpace($VideoId)) { return $null }
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            if ([string]$step.videoId -eq $VideoId) { return $null }
        }
    }

    $videoPath = Get-MbVideoFilePath -Project $Project -ProjectPath $ProjectPath -VideoId $VideoId
    $before = @($Project.videos).Count
    $Project.videos = @($Project.videos | Where-Object { $_.id -ne $VideoId })
    if (@($Project.videos).Count -eq $before) { return $null }
    return $videoPath
}

function Remove-MbUnreferencedVideos {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$ProjectPath
    )

    $referenced = @{}
    foreach ($sheet in @($Project.sheets)) {
        foreach ($step in @($sheet.steps)) {
            $videoId = [string]$step.videoId
            if (-not [string]::IsNullOrWhiteSpace($videoId)) { $referenced[$videoId] = $true }
        }
    }

    $paths = New-Object System.Collections.ArrayList
    foreach ($video in @($Project.videos)) {
        $videoId = [string]$video.id
        if ($referenced.ContainsKey($videoId)) { continue }
        $path = Get-MbVideoFilePath -Project $Project -ProjectPath $ProjectPath -VideoId $videoId
        if ($path) { [void]$paths.Add($path) }
    }
    $Project.videos = @($Project.videos | Where-Object { $referenced.ContainsKey([string]$_.id) })
    return @($paths)
}

function Get-MbVideoTotalBytes {
    param([Parameter(Mandatory = $true)][object]$Project)
    $total = [long]0
    foreach ($video in @($Project.videos)) { $total += [long]$video.byteLength }
    return $total
}

function Resolve-MbScreenshotDirectory {
    $guid = '{B7BEDE81-DF94-4682-A7D8-57A52620B86F}'
    foreach ($key in @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')) {
        try {
            $value = (Get-ItemProperty -Path $key -Name $guid -ErrorAction Stop).$guid
            if ($value) {
                $path = [Environment]::ExpandEnvironmentVariables([string]$value)
                if (Test-Path -LiteralPath $path -PathType Container) { return $path }
            }
        } catch { }
    }

    $pictures = [Environment]::GetFolderPath('MyPictures')
    foreach ($name in @('Screenshots', 'スクリーンショット')) {
        $path = Join-Path $pictures $name
        if (Test-Path -LiteralPath $path -PathType Container) { return $path }
    }
    foreach ($base in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
        if (-not $base) { continue }
        foreach ($name in @('Pictures\Screenshots', 'ピクチャ\スクリーンショット')) {
            $path = Join-Path $base $name
            if (Test-Path -LiteralPath $path -PathType Container) { return $path }
        }
    }
    return $null
}

Export-ModuleMember -Function @(
    'Get-MbImageKind',
    'Get-MbImageFilePath',
    'Add-MbImageAsset',
    'Add-MbImageStep',
    'Set-MbStepResultImage',
    'Remove-MbStepResultImage',
    'Set-MbStepImage',
    'Restore-MbStepImage',
    'Remove-MbUnusedImage',
    'Remove-MbUnreferencedImages',
    'Get-MbVideoKind',
    'Get-MbVideoFilePath',
    'Add-MbVideoAsset',
    'Set-MbStepVideo',
    'Remove-MbStepVideo',
    'Remove-MbUnusedVideo',
    'Remove-MbUnreferencedVideos',
    'Get-MbVideoTotalBytes',
    'Resolve-MbScreenshotDirectory'
)
