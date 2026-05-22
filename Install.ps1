#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs XL Trim Excel add-in for all users on the machine.

.DESCRIPTION
    Copies the add-in files to Program Files and registers it in HKLM
    so it loads for every user. Intended for SCCM / Software Center deployment.

    Prerequisites (must be installed before running this script):
      - .NET Framework 4.7.2+
      - Microsoft Visual Studio 2010 Tools for Office Runtime (VSTO 4.0)
        https://aka.ms/vsto40

.EXAMPLE
    # Run from the folder containing this script + the build output:
    powershell -ExecutionPolicy Bypass -File Install.ps1
#>

$ErrorActionPreference = "Stop"

$INSTALL_DIR  = Join-Path $env:ProgramFiles "XLTrim"
$ADDIN_KEY    = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\XLTrim"
$ADDIN_KEY_WOW= "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Addins\XLTrim"
$MANIFEST     = "$INSTALL_DIR\XLTrim.vsto|vstolocal"

Write-Host "=== XL Trim Installer ===" -ForegroundColor Cyan

# ── Check VSTO Runtime ────────────────────────────────────────────────────────
$vstoInstalled = Get-ItemProperty `
    "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" `
    -ErrorAction SilentlyContinue

if (-not $vstoInstalled) {
    Write-Warning "VSTO Runtime 4.0 not detected. The add-in may not load."
    Write-Warning "Download: https://aka.ms/vsto40"
}

# ── Copy files ────────────────────────────────────────────────────────────────
Write-Host "Installing to: $INSTALL_DIR"
New-Item -ItemType Directory -Force -Path $INSTALL_DIR | Out-Null

# Copy everything from the script's folder (build output lives alongside Install.ps1)
Get-ChildItem -Path $PSScriptRoot -File | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $INSTALL_DIR -Force
}

# ── Register add-in (both 64-bit and 32-bit hives for Excel compatibility) ────
foreach ($key in @($ADDIN_KEY, $ADDIN_KEY_WOW)) {
    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    Set-ItemProperty -Path $key -Name "Description"  -Value "XL Trim - Excel file optimizer"
    Set-ItemProperty -Path $key -Name "FriendlyName" -Value "XL Trim"
    Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $key -Name "Manifest"     -Value $MANIFEST
}

Write-Host "XL Trim installed successfully." -ForegroundColor Green
Write-Host "Restart Excel to see the XL Trim tab in the ribbon."
