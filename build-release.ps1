#Requires -Version 5.0
<#
.SYNOPSIS
    CloudAdmin365 Release Builder

.DESCRIPTION
    Produces a single-file, framework-dependent release EXE using the standard
    .NET PublishSingleFile mechanism with Brotli compression.

    Output is one CloudAdmin365.exe bundling all dependencies (~40-60 MB
    compressed vs ~90 MB uncompressed loose files), plus setup.bat and
    setup.ps1 to check the .NET 8 Desktop Runtime on the target machine.

    The EXE extracts itself to a per-user temp cache on first launch
    (transparent to the user). No ps2exe, no Base64 encoding.

    Requirements:
      - .NET SDK 8.0+  (dotnet CLI in PATH)
      - Run from the project root directory

.PARAMETER Runtime
    Target RID. Default: win-x64. Use win-arm64 for ARM devices.

.PARAMETER OutputDir
    Destination folder for the distributable files.
    Default: dist\<Runtime>

.PARAMETER OpenFolder
    Open the output folder in Explorer after a successful build.

.EXAMPLE
    .\build-release.ps1
    .\build-release.ps1 -Runtime win-arm64
    .\build-release.ps1 -OutputDir "C:\drop\CloudAdmin365" -OpenFolder
#>
param(
    [string]$Runtime   = "win-x64",
    [string]$OutputDir = "",
    [switch]$OpenFolder
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Banner {
    param([string]$Text)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Write-Step {
    param([string]$Text)
    Write-Host ""
    Write-Host "[STEP] $Text" -ForegroundColor Yellow
}

function Write-OK {
    param([string]$Text)
    Write-Host "[OK]   $Text" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Text)
    Write-Host "[FAIL] $Text" -ForegroundColor Red
}

# ---------------------------------------------------------------------------
# Pre-flight
# ---------------------------------------------------------------------------

Write-Banner "CloudAdmin365 Release Builder"

# Locate project root (same directory as this script)
$projectRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($projectRoot)) {
    $projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$csprojPath = Join-Path $projectRoot "CloudAdmin365.csproj"
if (-not (Test-Path $csprojPath)) {
    Write-Fail "CloudAdmin365.csproj not found at: $csprojPath"
    Write-Fail "Run this script from the project root directory."
    exit 1
}

try {
    Get-Command dotnet -ErrorAction Stop | Out-Null
}
catch {
    Write-Fail ".NET SDK not found in PATH. Install from https://dotnet.microsoft.com/download"
    exit 1
}

$sdkVersion = (& dotnet --version 2>&1).Trim()
Write-Host "  .NET SDK : $sdkVersion"
Write-Host "  Runtime  : $Runtime"

# Resolve output directory
if ([string]::IsNullOrEmpty($OutputDir)) {
    $OutputDir = Join-Path $projectRoot "dist\$Runtime"
}
Write-Host "  Output   : $OutputDir"

# ---------------------------------------------------------------------------
# Publish — single-file, framework-dependent, compressed
# ---------------------------------------------------------------------------

Write-Step "Publishing single-file release..."

$publishArgs = @(
    "publish",
    $csprojPath,
    "-c", "Release",
    "-f", "net8.0-windows",
    "-r", $Runtime,
    "--no-self-contained",                      # keep framework-dependent (small EXE)
    "-p:PublishSingleFile=true",                # bundle all DLLs into the EXE
    "-p:EnableCompressionInSingleFile=true",    # Brotli compress bundled files (~40-60 MB)
    "-p:DebugType=None",                        # strip .pdb from output
    "-p:DebugSymbols=false",
    "--nologo"
)

& dotnet @publishArgs

if ($LASTEXITCODE -ne 0) {
    Write-Fail "dotnet publish failed (exit $LASTEXITCODE)."
    exit 1
}

Write-OK "Publish succeeded."

# ---------------------------------------------------------------------------
# Locate the published EXE
# ---------------------------------------------------------------------------

$publishedExe = Join-Path $projectRoot "bin\Release\net8.0-windows\$Runtime\publish\CloudAdmin365.exe"

if (-not (Test-Path $publishedExe)) {
    Write-Fail "Expected output not found: $publishedExe"
    Write-Fail "Check publish output above for errors."
    exit 1
}

$exeSizeMB = [math]::Round((Get-Item $publishedExe).Length / 1MB, 1)
Write-OK "CloudAdmin365.exe  ($exeSizeMB MB)"

# ---------------------------------------------------------------------------
# Assemble distribution folder
# ---------------------------------------------------------------------------

Write-Step "Assembling distribution folder: $OutputDir"

if (Test-Path $OutputDir) {
    Remove-Item $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Copy single-file EXE
Copy-Item $publishedExe -Destination $OutputDir

# Copy launcher scripts
foreach ($launcher in @("setup.bat", "setup.ps1")) {
    $src = Join-Path $projectRoot $launcher
    if (Test-Path $src) {
        Copy-Item $src -Destination $OutputDir
        Write-OK "Copied $launcher"
    }
    else {
        Write-Host "[WARN] $launcher not found — skipping." -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# Create ZIP for easy sharing
# ---------------------------------------------------------------------------

Write-Step "Creating distribution ZIP..."

$zipPath = Join-Path (Split-Path $OutputDir -Parent) "CloudAdmin365-$Runtime.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Compress-Archive -Path "$OutputDir\*" -DestinationPath $zipPath -CompressionLevel Optimal
$zipSizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-OK "CloudAdmin365-$Runtime.zip  ($zipSizeMB MB)"

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Banner "Build Complete"
Write-Host ""
Write-Host "  Distribution folder : $OutputDir"
Write-Host "  Distribution ZIP    : $zipPath"
Write-Host ""
Write-Host "  Contents:" -ForegroundColor White
Get-ChildItem $OutputDir | Sort-Object Name | ForEach-Object {
    $sz = [math]::Round($_.Length / 1MB, 1)
    Write-Host "    $($_.Name.PadRight(35)) $sz MB" -ForegroundColor White
}
Write-Host ""
Write-Host "  Target machine requires: .NET 8 Desktop Runtime" -ForegroundColor Yellow
Write-Host "  Download: https://dotnet.microsoft.com/en-us/download/dotnet/8.0" -ForegroundColor Yellow
Write-Host ""

if ($OpenFolder) {
    explorer.exe $OutputDir
}
