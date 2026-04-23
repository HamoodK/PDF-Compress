@echo off
:: Launcher — runs the PowerShell installer as Administrator
:: Double-click this file to install on any PC

echo Requesting administrator privileges...
powershell -Command "Start-Process powershell -ArgumentList '-STA -ExecutionPolicy Bypass -File ""%~dp0Install-CompressPDF.ps1""' -Verb RunAs"
