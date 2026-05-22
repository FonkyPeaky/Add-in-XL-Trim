#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Uninstalls XL Trim Excel add-in.
#>

$ErrorActionPreference = "Stop"

$INSTALL_DIR  = Join-Path $env:ProgramFiles "XLTrim"
$ADDIN_KEY    = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\XLTrim"
$ADDIN_KEY_WOW= "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Addins\XLTrim"

Write-Host "=== XL Trim Uninstaller ===" -ForegroundColor Cyan

foreach ($key in @($ADDIN_KEY, $ADDIN_KEY_WOW)) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force
        Write-Host "Registry key removed: $key"
    }
}

if (Test-Path $INSTALL_DIR) {
    Remove-Item -Path $INSTALL_DIR -Recurse -Force
    Write-Host "Files removed: $INSTALL_DIR"
}

Write-Host "XL Trim uninstalled successfully." -ForegroundColor Green
