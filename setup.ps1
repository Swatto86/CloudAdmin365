#Requires -Version 5.0
<#
.SYNOPSIS
    CloudAdmin365 Setup Script
    
.DESCRIPTION
    Checks for prerequisites (.NET Runtime 8.0) and launches the application.
    If .NET Runtime 8.0 is not installed, offers to download it.
    
.PARAMETER SkipRuntimeCheck
    If specified, skips the .NET Runtime check (not recommended).
    
.NOTES
    Run from the same directory as CloudAdmin365.exe
#>

param(
    [switch]$SkipRuntimeCheck = $false
)

$ErrorActionPreference = 'Stop'

function Write-Header {
    param([string]$Message)
    Write-Host ""
    Write-Host "=" * 70 -ForegroundColor Green
    Write-Host $Message -ForegroundColor Green
    Write-Host "=" * 70 -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "ℹ️  $Message" -ForegroundColor Cyan
}

function Write-Error-Custom {
    param([string]$Message)
    Write-Host "❌ $Message" -ForegroundColor Red
}

function Write-Success {
    param([string]$Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function Test-DotNetRuntime {
    <#
    .SYNOPSIS
        Checks if .NET Runtime 8.0 is installed.
    #>
    try {
        $output = & dotnet --version 2>&1
        $version = $output.Trim()
        
        if ($version -match '^8\.' -or $version -match '^9\.') {
            Write-Success ".NET Runtime detected: $version"
            return $true
        }
        else {
            Write-Error-Custom ".NET Runtime 8.0 not found. Current version: $version"
            return $false
        }
    }
    catch {
        Write-Error-Custom ".NET Runtime is not installed or not in PATH"
        return $false
    }
}

function Invoke-DotNetRuntimeInstall {
    <#
    .SYNOPSIS
        Prompts user to install .NET Runtime 8.0 and opens the download page.
    #>
    Write-Header "Installing .NET Runtime 8.0"
    
    Write-Info "CloudAdmin365 requires .NET Runtime 8.0 or later."
    Write-Info "Opening download page in your default browser..."
    Write-Info ""
    Write-Info "Download options:"
    Write-Info "  • Hosting Bundle (includes both ASP.NET Core and Console Runtime)"
    Write-Info "  • Desktop Runtime (for Windows desktop applications)"
    Write-Info "  • Runtime (console applications)"
    Write-Info ""
    Write-Info "For CloudAdmin365, the 'Desktop Runtime' or 'Hosting Bundle' is recommended."
    
    $url = "https://dotnet.microsoft.com/en-us/download/dotnet/8.0"
    Start-Process $url
    
    Write-Host ""
    Write-Host "After installing .NET Runtime 8.0:" -ForegroundColor Yellow
    Write-Host "  1. Close this window"
    Write-Host "  2. Run this script again"
    Write-Host ""
    
    Read-Host "Press Enter to close this window"
    exit 1
}

function Invoke-Application {
    <#
    .SYNOPSIS
        Launches CloudAdmin365.exe from the current directory.
    #>
    $appPath = Join-Path $PSScriptRoot "CloudAdmin365.exe"
    
    if (-not (Test-Path $appPath)) {
        Write-Error-Custom "CloudAdmin365.exe not found at: $appPath"
        Read-Host "Press Enter to close"
        exit 1
    }
    
    Write-Info "Launching CloudAdmin365..."
    & $appPath
}

# ============================================================================
# Main
# ============================================================================

Write-Header "CloudAdmin365 Setup"

Write-Info "Current directory: $PSScriptRoot"

if (-not $SkipRuntimeCheck) {
    Write-Host ""
    Write-Info "Checking for .NET Runtime 8.0..."
    
    if (-not (Test-DotNetRuntime)) {
        Invoke-DotNetRuntimeInstall
    }
}

Write-Success "All prerequisites satisfied!"
Write-Info "Launching application..."

Invoke-Application
