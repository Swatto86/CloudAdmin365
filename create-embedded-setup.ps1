# Create Self-Contained Setup with Embedded Application Files
# This bundles all app files into a single setup.exe

param(
    [switch]$Install
)

$ErrorActionPreference = 'Stop'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Self-Contained Setup Builder" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Ensure dotnet is available
try {
    Get-Command dotnet -ErrorAction Stop | Out-Null
}
catch {
    Write-Host "[ERROR] dotnet SDK not found in PATH." -ForegroundColor Red
    Write-Host "Install the .NET SDK and retry." -ForegroundColor Yellow
    exit 1
}

# Check for ps2exe module
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "[WARNING] ps2exe module not found" -ForegroundColor Yellow
    
    if ($Install) {
        Write-Host "Installing ps2exe module..." -ForegroundColor Yellow
        Install-Module ps2exe -Scope CurrentUser -Force
        Write-Host "[SUCCESS] ps2exe installed" -ForegroundColor Green
        Write-Host ""
    }
    else {
        Write-Host "Run with -Install flag to install ps2exe automatically" -ForegroundColor Yellow
        $response = Read-Host "Install ps2exe now? (Y/N)"
        if ($response -eq 'Y') {
            Install-Module ps2exe -Scope CurrentUser -Force
            Write-Host "[SUCCESS] ps2exe installed" -ForegroundColor Green
        }
        else {
            exit 1
        }
    }
}

Import-Module ps2exe

# Check if app files exist
$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptRoot)) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$runtimeId = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "win-arm64" } else { "win-x64" }
$publishPath = Join-Path $scriptRoot "bin\Release\net8.0-windows\$runtimeId\publish"

Write-Host "Publishing framework-dependent release for $runtimeId..." -ForegroundColor Yellow
& dotnet publish -c Release -f net8.0-windows -r $runtimeId --no-self-contained
if ($LASTEXITCODE -ne 0) {
    Write-Host "" 
    Write-Host "[ERROR] dotnet publish failed." -ForegroundColor Red
    exit 1
}

Write-Host "Checking publish output..." -ForegroundColor Yellow
$publishFiles = Get-ChildItem -Path $publishPath -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike "*.pdb" }

function Test-ExcludedCulturePath {
    param([string]$RelativePath)

    $parts = $RelativePath -split '[\\/]+'
    foreach ($part in $parts) {
        if ($part -match '^[a-z]{2}(-[A-Z]{2})?$') {
            if ($part -eq 'en' -or $part -eq 'en-US') {
                return $false
            }

            return $true
        }
    }

    return $false
}

$publishFiles = $publishFiles | Where-Object {
    $relativePath = $_.FullName.Substring($publishPath.Length + 1)
    -not (Test-ExcludedCulturePath $relativePath)
}

$requiredFileNames = @(
    "System.Management.Automation.dll",
    "Microsoft.PowerShell.Commands.Diagnostics.dll",
    "Microsoft.Identity.Client.dll"
)

$requiredFiles = @{}
foreach ($name in $requiredFileNames) {
    $match = Get-ChildItem -Path $publishPath -File -Recurse -Filter $name -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $match) {
        Write-Host "" 
        Write-Host "[ERROR] Required file missing: $name" -ForegroundColor Red
        Write-Host "Re-run publish or restore packages and try again." -ForegroundColor Yellow
        exit 1
    }

    $requiredFiles[$name] = $match.FullName
}

$rootAutomationPath = Join-Path $publishPath "System.Management.Automation.dll"
if (-not (Test-Path $rootAutomationPath)) {
    Copy-Item -Path $requiredFiles["System.Management.Automation.dll"] -Destination $rootAutomationPath -Force
    $publishFiles = Get-ChildItem -Path $publishPath -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike "*.pdb" }
}

if (-not $publishFiles -or $publishFiles.Count -eq 0) {
    Write-Host "" 
    Write-Host "[ERROR] No publish output found at: $publishPath" -ForegroundColor Red
    Write-Host "Please re-run this script to publish and package the app." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

$requiredFiles = $publishFiles | Sort-Object FullName | ForEach-Object {
    $_.FullName.Substring($publishPath.Length + 1)
}
foreach ($file in $requiredFiles) {
    $path = Join-Path $publishPath $file
    $size = [math]::Round((Get-Item $path).Length / 1KB, 0)
    Write-Host "  [OK] $file ($size KB)" -ForegroundColor Green
}

Write-Host ""
Write-Host "Embedding application files as Base64..." -ForegroundColor Yellow

# Read and encode files
$embeddedFiles = @{}
foreach ($file in $requiredFiles) {
    $path = Join-Path $publishPath $file
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $base64 = [Convert]::ToBase64String($bytes)
    $embeddedFiles[$file] = $base64
    Write-Host "  Encoded: $file" -ForegroundColor Green
}

Write-Host ""
Write-Host "Creating embedded installer script..." -ForegroundColor Yellow

# Create the embedded installer script
$installerScript = @"
#Requires -Version 5.1
<#
.SYNOPSIS
    CloudAdmin365 Self-Contained Installer
.DESCRIPTION
    Single-file installer with embedded application files.
    No admin rights required.
#>

[CmdletBinding()]
param(
    [switch]`$Uninstall,
    [switch]`$Silent
)

`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'

`$GuiAvailable = `$false
try {
    Add-Type -AssemblyName System.Windows.Forms
    `$GuiAvailable = `$true
}
catch { }

`$HasConsole = `$false
try {
    `$HasConsole = `$null -ne `$Host.UI -and `$null -ne `$Host.UI.RawUI
}
catch { }

`$IsConsoleHost = `$HasConsole -or (`$Host.Name -match 'ConsoleHost')
`$LogPath = Join-Path `$env:TEMP "CloudAdmin365-Setup.log"

# Configuration
`$AppName = "CloudAdmin365"
`$AppVersion = "2.0.0"
`$Publisher = "Steve Watson"
`$InstallPath = Join-Path `$env:LOCALAPPDATA `$AppName
`$StartMenuPath = Join-Path ([Environment]::GetFolderPath('StartMenu')) "Programs\`$AppName"

# Embedded application files (Base64 encoded)
`$EmbeddedFiles = @{
"@

foreach ($file in $requiredFiles) {
    $installerScript += "`n    '$file' = @'"
    $installerScript += "`n" + $embeddedFiles[$file]
    $installerScript += "`n'@"
}

$installerScript += @"

}

function Write-Status {
    param([string]`$Message, [string]`$Type = "Info")
    try {
        `$timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
        Add-Content -Path `$LogPath -Value "[`$timestamp] [`$Type] `$Message" -Encoding UTF8
    }
    catch { }

    if (`$Silent) { return }

    if (`$IsConsoleHost) {
        `$color = switch (`$Type) {
            "Success" { "Green" }
            "Warning" { "Yellow" }
            "Error" { "Red" }
            default { "White" }
        }
        Write-Host "[`$Type] `$Message" -ForegroundColor `$color
        return
    }

    if (`$Type -eq "Error") {
        Show-MessageBox -Message `$Message `
            -Title "CloudAdmin365 Setup" `
            -Buttons ([System.Windows.Forms.MessageBoxButtons]::OK) `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Show-UiPrompt {
    param(
        [string]`$Message,
        [string]`$Title,
        [System.Windows.Forms.MessageBoxButtons]`$Buttons,
        [System.Windows.Forms.MessageBoxIcon]`$Icon
    )

    if (`$IsConsoleHost) { return $null }

    return Show-MessageBox -Message `$Message -Title `$Title -Buttons `$Buttons -Icon `$Icon
}

function Show-MessageBox {
    param(
        [string]`$Message,
        [string]`$Title,
        [System.Windows.Forms.MessageBoxButtons]`$Buttons,
        [System.Windows.Forms.MessageBoxIcon]`$Icon
    )

    if (-not `$GuiAvailable) { return $null }

    try {
        return [System.Windows.Forms.MessageBox]::Show(`$Message, `$Title, `$Buttons, `$Icon)
    }
    catch {
        try {
            return [System.Windows.Forms.MessageBox]::Show(`$Message, `$Title)
        }
        catch {
            return $null
        }
    }
}

function Test-DotNetRuntime {
    try {
        `$dotnetExe = Get-Command dotnet -ErrorAction SilentlyContinue
        if (`$dotnetExe) {
            `$runtimes = & `$dotnetExe.Source --list-runtimes 2>`$null | Where-Object { `$_ -match "Microsoft\.WindowsDesktop\.App 8\.\d+\.\d+" }
            if (`$null -ne `$runtimes) { return `$true }
        }

        `$roots = @()
        if (`$env:ProgramFiles) { `$roots += Join-Path `$env:ProgramFiles "dotnet" }
        if (`${env:ProgramFiles(x86)}) { `$roots += Join-Path `${env:ProgramFiles(x86)} "dotnet" }

        foreach (`$root in `$roots) {
            `$shared = Join-Path `$root "shared\Microsoft.WindowsDesktop.App"
            if (Test-Path `$shared) {
                `$versions = Get-ChildItem -Path `$shared -Directory -ErrorAction SilentlyContinue
                if (`$versions | Where-Object { `$_.Name -like "8.*" }) { return `$true }
            }
        }
    }
    catch { }

    return `$false
}

function Install-DotNetPrompt {
    Write-Status ".NET 8 Desktop Runtime is required but not installed." "Warning"
    Write-Host ""
    Write-Host "Download from: https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not `$Silent) {
        if (`$IsConsoleHost) {
            `$response = Read-Host "Open download page in browser? (Y/N)"
            if (`$response -eq 'Y') {
                Start-Process "https://dotnet.microsoft.com/download/dotnet/8.0"
                Write-Host ""
                Write-Host "After installing .NET 8 Runtime, run this installer again." -ForegroundColor Yellow
                Read-Host "Press Enter to exit"
                exit 1
            }
        }
        else {
            `$result = Show-UiPrompt `
                -Message ".NET 8 Desktop Runtime is required. Open the download page now?" `
                -Title "CloudAdmin365 Setup" `
                -Buttons ([System.Windows.Forms.MessageBoxButtons]::YesNo) `
                -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
            if (`$result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process "https://dotnet.microsoft.com/download/dotnet/8.0"
            }
        }
    }
    
    Write-Status "Installation cannot continue without .NET 8 Runtime." "Error"
    exit 1
}

function New-Shortcut {
    param(
        [string]`$ShortcutPath,
        [string]`$TargetPath,
        [string]`$Description,
        [string]`$WorkingDirectory
    )
    
    `$WshShell = New-Object -ComObject WScript.Shell
    `$Shortcut = `$WshShell.CreateShortcut(`$ShortcutPath)
    `$Shortcut.TargetPath = `$TargetPath
    `$Shortcut.Description = `$Description
    `$Shortcut.WorkingDirectory = `$WorkingDirectory
    `$Shortcut.Save()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$WshShell) | Out-Null
}

function Install-Application {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  `$AppName v`$AppVersion Installer" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not `$Silent -and -not `$IsConsoleHost) {
        Show-MessageBox -Message "CloudAdmin365 setup is running. Please wait." `
            -Title "CloudAdmin365 Setup" `
            -Buttons ([System.Windows.Forms.MessageBoxButtons]::OK) `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    
    # Check .NET Runtime
    Write-Status "Checking for .NET 8 Runtime..."
    if (-not (Test-DotNetRuntime)) {
        Install-DotNetPrompt
    }
    Write-Status ".NET 8 Runtime found" "Success"
    Write-Host ""
    
    # Create installation directory
    Write-Status "Creating installation directory..."
    if (Test-Path `$InstallPath) {
        Write-Status "Updating existing installation at: `$InstallPath" "Warning"
    }
    else {
        Write-Status "Installing to: `$InstallPath"
    }
    
    New-Item -Path `$InstallPath -ItemType Directory -Force | Out-Null
    Write-Host ""
    
    # Extract embedded files
    Write-Status "Extracting application files..."
    `$extractErrors = @()
    foreach (`$fileName in `$EmbeddedFiles.Keys) {
        try {
            `$bytes = [Convert]::FromBase64String(`$EmbeddedFiles[`$fileName])
            `$destPath = Join-Path `$InstallPath `$fileName
            `$destDir = Split-Path -Parent `$destPath
            if (-not [string]::IsNullOrWhiteSpace(`$destDir)) {
                New-Item -Path `$destDir -ItemType Directory -Force | Out-Null
            }
            [System.IO.File]::WriteAllBytes(`$destPath, `$bytes)
            Write-Status "  Extracted: `$fileName" "Success"
        }
        catch {
            Write-Status "  Failed to extract: `$fileName" "Error"
            Write-Host "  Error: `$_" -ForegroundColor Red
            `$extractErrors += `$fileName
        }
    }
    if (`$extractErrors.Count -gt 0) {
        Write-Status "One or more files failed to extract." "Error"
        exit 1
    }
    Write-Host ""

    # Verify required runtime dependencies are present
    `$requiredRuntimeFiles = @(
        "Microsoft.Identity.Client.dll",
        "System.Management.Automation.dll",
        "Microsoft.PowerShell.Commands.Diagnostics.dll"
    )

    `$missingRuntime = @()
    foreach (`$required in `$requiredRuntimeFiles) {
        `$found = Get-ChildItem -Path `$InstallPath -File -Recurse -Filter `$required -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not `$found) {
            `$missingRuntime += `$required
        }
    }

    if (`$missingRuntime.Count -gt 0) {
        Write-Status "Missing required runtime files: `$( `$missingRuntime -join ', ' )" "Error"
        exit 1
    }
    
    # Create Start Menu shortcuts
    Write-Status "Creating Start Menu shortcuts..."
    New-Item -Path `$StartMenuPath -ItemType Directory -Force | Out-Null
    
    `$exePath = Join-Path `$InstallPath "CloudAdmin365.exe"
    `$shortcutPath = Join-Path `$StartMenuPath "`$AppName.lnk"
    
    New-Shortcut -ShortcutPath `$shortcutPath ``
                 -TargetPath `$exePath ``
                 -Description "Exchange Online Audit Tool" ``
                 -WorkingDirectory `$InstallPath
    
    Write-Status "Start Menu shortcut created" "Success"
    Write-Host ""
    
    # Save uninstall information
    `$uninstallInfo = @{
        Version = `$AppVersion
        InstallDate = (Get-Date).ToString('yyyy-MM-dd')
        InstallPath = `$InstallPath
        Publisher = `$Publisher
    }
    `$uninstallInfo | ConvertTo-Json | Out-File (Join-Path `$InstallPath "uninstall.json") -Encoding UTF8
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  Installation Complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Installed to: `$InstallPath" -ForegroundColor White
    Write-Host "Launch from: Start Menu > `$AppName" -ForegroundColor White
    Write-Host ""
    
    if (-not `$Silent) {
        if (`$IsConsoleHost) {
            `$response = Read-Host "Launch `$AppName now? (Y/N)"
            if (`$response -eq 'Y') {
                Start-Process (Join-Path `$InstallPath "CloudAdmin365.exe")
            }
        }
        else {
            `$result = Show-UiPrompt `
                -Message "Launch `$AppName now?" `
                -Title "CloudAdmin365 Setup" `
                -Buttons ([System.Windows.Forms.MessageBoxButtons]::YesNo) `
                -Icon ([System.Windows.Forms.MessageBoxIcon]::Question)
            if (`$result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process (Join-Path `$InstallPath "CloudAdmin365.exe")
            }
        }
    }
}

function Uninstall-Application {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "  `$AppName Uninstaller" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    
    if (-not (Test-Path `$InstallPath)) {
        Write-Status "`$AppName is not installed." "Warning"
        if (-not `$Silent) { Read-Host "Press Enter to exit" }
        exit 0
    }
    
    if (-not `$Silent) {
        if (`$IsConsoleHost) {
            `$response = Read-Host "Are you sure you want to uninstall `$AppName? (Y/N)"
            if (`$response -ne 'Y') {
                Write-Status "Uninstall cancelled." "Warning"
                exit 0
            }
        }
        else {
            `$result = Show-UiPrompt `
                -Message "Uninstall `$AppName?" `
                -Title "CloudAdmin365 Setup" `
                -Buttons ([System.Windows.Forms.MessageBoxButtons]::YesNo) `
                -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
            if (`$result -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Status "Uninstall cancelled." "Warning"
                exit 0
            }
        }
    }
    
    Write-Status "Removing application files..."
    if (Test-Path `$InstallPath) {
        Remove-Item -Path `$InstallPath -Recurse -Force
        Write-Status "Application files removed" "Success"
    }
    
    Write-Status "Removing Start Menu shortcuts..."
    if (Test-Path `$StartMenuPath) {
        Remove-Item -Path `$StartMenuPath -Recurse -Force
        Write-Status "Start Menu shortcuts removed" "Success"
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  Uninstall Complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    
    if (-not `$Silent) { Read-Host "Press Enter to exit" }
}

# Main execution
try {
    if (`$Uninstall) {
        Uninstall-Application
    }
    else {
        Install-Application
    }
}
catch {
    Write-Status "Installation failed: `$_" "Error"
    Write-Status `$_.Exception.Message "Error"
    if (-not `$Silent) { Read-Host "Press Enter to exit" }
    exit 1
}
"@

# Save the embedded installer script
$tempDir = Join-Path $scriptRoot "dist\_temp"
if (-not (Test-Path $tempDir)) {
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
}
$tempScriptPath = Join-Path $tempDir "setup-embedded.ps1"
$installerScript | Out-File $tempScriptPath -Encoding UTF8
Write-Host "[OK] Embedded installer script created" -ForegroundColor Green

Write-Host ""
Write-Host "Compiling to self-contained setup.exe..." -ForegroundColor Yellow

# Create dist folder if it doesn't exist
$distDir = Join-Path $scriptRoot "dist"
if (-not (Test-Path $distDir)) {
    New-Item -Path $distDir -ItemType Directory | Out-Null
}

$outputExe = Join-Path $distDir "CloudAdmin365-Setup.exe"
$iconFile = Join-Path $scriptRoot "CloudAdmin365.ico"

try {
    # Check if icon exists
    if (-not (Test-Path $iconFile)) {
        Write-Host "[WARNING] Icon file not found at: $iconFile" -ForegroundColor Yellow
        Write-Host "  Run generate-icon.ps1 to create the icon" -ForegroundColor Yellow
    }
    
    Invoke-ps2exe `
        -inputFile $tempScriptPath `
        -outputFile $outputExe `
        -iconFile $(if (Test-Path $iconFile) { $iconFile } else { $null }) `
        -noConsole:$false `
        -title "CloudAdmin365 Setup" `
        -description "CloudAdmin365 Self-Contained Installer" `
        -company "Steve Watson" `
        -product "CloudAdmin365" `
        -version "2.0.0.0" `
        -requireAdmin:$false `
        -supportOS `
        -noError
    
    if (Test-Path $outputExe) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host " Self-Contained Setup.exe Created!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        
        $exeSize = [math]::Round((Get-Item $outputExe).Length / 1KB, 0)
        Write-Host "Location: $outputExe" -ForegroundColor White
        Write-Host "Size: $exeSize KB" -ForegroundColor White
        Write-Host ""
        Write-Host "This is a SINGLE FILE installer!" -ForegroundColor Green
        Write-Host "  - Contains all application files embedded inside" -ForegroundColor White
        Write-Host "  - No other files needed" -ForegroundColor White
        Write-Host "  - Just distribute this one .exe file" -ForegroundColor White
        Write-Host ""
        Write-Host "Users simply run: CloudAdmin365-Setup.exe" -ForegroundColor Cyan
        Write-Host ""
    }
    else {
        throw "setup.exe was not created"
    }
}
catch {
    Write-Host ""
    Write-Host "[ERROR] Failed to create setup.exe" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
if (Test-Path $tempDir) {
    try {
        Remove-Item $tempDir -Recurse -Force -ErrorAction Stop
    }
    catch {
        # Temp directory may still be locked; schedule async cleanup
        Start-Job -ScriptBlock {
            Start-Sleep -Seconds 2
            Remove-Item $args[0] -Recurse -Force -ErrorAction SilentlyContinue
        } -ArgumentList $tempDir | Out-Null
        Write-Host "Temp files will be cleaned up automatically..." -ForegroundColor Gray
    }
}

Write-Host "[DONE]" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"
