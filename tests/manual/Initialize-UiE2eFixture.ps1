[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [ValidateRange(2, 20)][int]$StepCount = 5
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Project.psm1') -Force
Import-Module (Join-Path $repoRoot 'src\ManualBuilder.Capture.psm1') -Force
Add-Type -AssemblyName System.Drawing

function New-MbUiE2ePng {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][Drawing.Color]$Color
    )

    $bitmap = New-Object Drawing.Bitmap 960, 540
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $font = New-Object Drawing.Font 'Yu Gothic UI', 32, ([Drawing.FontStyle]::Bold)
    $stream = New-Object IO.MemoryStream
    try {
        $graphics.Clear($Color)
        $graphics.DrawString($Text, $font, [Drawing.Brushes]::White, 60, 220)
        $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    } finally {
        $stream.Dispose()
        $font.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$fullProjectPath = [IO.Path]::GetFullPath($ProjectPath)
$projectDirectory = Split-Path -Parent $fullProjectPath
[void](New-Item -ItemType Directory -Path $projectDirectory -Force)
$project = Get-MbProject -Path $fullProjectPath
Set-MbProjectTitle -Project $project -Title 'UI E2E Manual'
$sheet = @($project.sheets)[0]

for ($index = 1; $index -le $StepCount; $index++) {
    $before = New-MbUiE2ePng -Text ("STEP {0} - BEFORE" -f $index) `
        -Color ([Drawing.Color]::FromArgb(40, 80, (100 + ($index * 20))))
    $after = New-MbUiE2ePng -Text ("STEP {0} - AFTER" -f $index) `
        -Color ([Drawing.Color]::FromArgb(40, (100 + ($index * 20)), 80))
    $capture = Add-MbImageStep -Project $project -ProjectPath $fullProjectPath -SheetId $sheet.id -Bytes $before -Source file
    Update-MbStep -Project $project -StepId $capture.Step.id -Title ("Step {0}" -f $index) `
        -Description ("Perform operation {0}." -f $index) -Note ''
    [void](Set-MbStepResultImage -Project $project -ProjectPath $fullProjectPath `
        -StepId $capture.Step.id -Bytes $after -Source file)
}

[void](Save-MbProject -Project $project -Path $fullProjectPath)
Write-Output $fullProjectPath
