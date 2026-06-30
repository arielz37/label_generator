$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

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

Invoke-Checked python -m PyInstaller --clean .\LabelGenerator.spec

$DistDir = Join-Path $ProjectRoot "dist\LabelGenerator"
if (!(Test-Path $DistDir)) {
    throw "Build output directory does not exist: $DistDir"
}

$itemsToCopy = @(
    "template_mapping.xlsx",
    "字段变量说明.md",
    "docs"
)

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

Write-Host "Build completed: $DistDir"
