#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Removes the "Compress PDF" context menu entry.
    Does NOT uninstall Ghostscript.
#>

Write-Host "`nRemoving Compress PDF context menu..." -ForegroundColor Cyan

# Remove HKLM registry entry
Remove-Item -Path "HKLM:\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF" -Recurse -Force -ErrorAction SilentlyContinue

# Remove HKCU registry entry (if any)
Remove-Item -Path "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF" -Recurse -Force -ErrorAction SilentlyContinue

# Remove deployed scripts
Remove-Item -Path "C:\ProgramData\CompressPDF" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Done. Compress PDF has been removed." -ForegroundColor Green
Write-Host "(Ghostscript was NOT uninstalled)" -ForegroundColor Yellow

if ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected) { pause }
