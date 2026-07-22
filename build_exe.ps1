param(
    [ValidateSet("none", "patch", "minor", "major")]
    [string]$Bump = "none"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot
$DistDir = Join-Path $ProjectRoot "dist\LabelGenerator"
$VersionFile = Join-Path $ProjectRoot "version.py"

function Get-AppVersion {
    if (!(Test-Path -LiteralPath $VersionFile -PathType Leaf)) {
        throw "Version file does not exist: $VersionFile"
    }

    $content = [System.IO.File]::ReadAllText($VersionFile)
    $match = [regex]::Match($content, '(?m)^__version__\s*=\s*"(\d+)\.(\d+)\.(\d+)"\s*$')
    if (!$match.Success) {
        throw 'version.py must contain __version__ = "major.minor.patch"'
    }

    return @(
        [int]$match.Groups[1].Value,
        [int]$match.Groups[2].Value,
        [int]$match.Groups[3].Value
    )
}

function Set-AppVersion {
    param(
        [Parameter(Mandatory = $true)]
        [int[]]$Parts
    )

    $version = $Parts -join "."
    $content = "`"`"`"Single source of truth for the application version.`"`"`"`r`n`r`n__version__ = `"$version`"`r`n"
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($VersionFile, $content, $utf8WithoutBom)
    return $version
}

$VersionParts = Get-AppVersion
if ($Bump -ne "none") {
    switch ($Bump) {
        "patch" { $VersionParts[2] += 1 }
        "minor" { $VersionParts[1] += 1; $VersionParts[2] = 0 }
        "major" { $VersionParts[0] += 1; $VersionParts[1] = 0; $VersionParts[2] = 0 }
    }
    $AppVersion = Set-AppVersion -Parts $VersionParts
    Write-Host "Version bumped ($Bump): v$AppVersion"
} else {
    $AppVersion = $VersionParts -join "."
    Write-Host "Building existing version: v$AppVersion"
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command,
        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $Command $($Arguments -join ' ')"
    }
}

if (Test-Path $DistDir) {
    Remove-Item -LiteralPath $DistDir -Recurse -Force
}

Invoke-Checked python -m PyInstaller --clean .\LabelGenerator.spec
if (!(Test-Path $DistDir)) {
    throw "Build output directory does not exist: $DistDir"
}

$itemsToCopy = @(
    "template_mapping.xlsx",
    "docs"
)

$itemsToCopy += Get-ChildItem -LiteralPath $ProjectRoot -Filter "*.txt" -File | ForEach-Object { $_.Name }
$itemsToCopy += Get-ChildItem -LiteralPath $ProjectRoot -Filter "*.md" -File | ForEach-Object { $_.Name }
$itemsToCopy = $itemsToCopy | Select-Object -Unique

$packageNameFile = Get-ChildItem -LiteralPath $ProjectRoot -Filter "bom*.txt" -File | Select-Object -First 1
if ($packageNameFile) {
    $itemsToCopy += $packageNameFile.Name
} else {
    Write-Warning "Package name whitelist file was not found: bom*.txt"
}

foreach ($item in $itemsToCopy) {
    $source = Join-Path $ProjectRoot $item
    if (!(Test-Path $source)) {
        Write-Warning "Skip missing file or directory: $item"
        continue
    }

    $target = Join-Path $DistDir $item
    $targetParent = Split-Path -Parent $target
    if (!(Test-Path $targetParent)) {
        New-Item -ItemType Directory -Path $targetParent | Out-Null
    }

    if (Test-Path $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    Copy-Item -LiteralPath $source -Destination $target -Recurse
}

$templateDir = Join-Path $DistDir "Templates"
if (Test-Path $templateDir) {
    Remove-Item -LiteralPath $templateDir -Recurse -Force
}
New-Item -ItemType Directory -Path $templateDir | Out-Null

foreach ($dirName in @("runtime_data", "final_labels", "logs")) {
    $dir = Join-Path $DistDir $dirName
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

Write-Host "Build completed: v$AppVersion -> $DistDir"
