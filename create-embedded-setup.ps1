<#
.SYNOPSIS
    DEPRECATED — use build-release.ps1 instead.

.DESCRIPTION
    This script previously bundled application files as Base64 strings inside a
    ps2exe-generated EXE. That approach produced bloated ~120 MB installers due
    to Base64 overhead (+33%) on top of the already-large dependency set.

    It has been replaced by build-release.ps1, which uses the standard .NET
    PublishSingleFile + EnableCompressionInSingleFile mechanism. This produces
    a correctly compressed (~40-60 MB) single-file EXE with no external tools.

.EXAMPLE
    .\build-release.ps1
    .\build-release.ps1 -Runtime win-arm64 -OpenFolder
#>

Write-Host ""
Write-Host "[DEPRECATED] create-embedded-setup.ps1 is no longer used." -ForegroundColor Yellow
Write-Host "             Run  .\build-release.ps1  instead." -ForegroundColor Yellow
Write-Host ""

& (Join-Path $PSScriptRoot "build-release.ps1") @args
