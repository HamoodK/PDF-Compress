#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs "Compress PDF" right-click context menu using Ghostscript.
.DESCRIPTION
    - Installs Ghostscript (via Chocolatey or direct download)
    - Deploys the compression script to C:\ProgramData\CompressPDF\
    - Adds "Compress PDF" to the right-click menu for all PDF files
    - Restores the classic Windows 11 context menu (optional)
.NOTES
    Run as Administrator. Works on Windows 10 and 11.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Config ────────────────────────────────────────────────────────────────────
$InstallDir   = "C:\ProgramData\CompressPDF"
$ScriptFile   = "$InstallDir\compress_pdf.ps1"
$GsVersion    = "10.07.0"
$GsExe        = "C:\Program Files\gs\gs$GsVersion\bin\gswin64c.exe"
$GsDownload   = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10070/gs10070w64.exe"
$GsInstaller  = "$env:TEMP\gs_installer.exe"

# ── Helpers ───────────────────────────────────────────────────────────────────
function Write-Step($msg) { Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "   OK: $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "   WARN: $msg" -ForegroundColor Yellow }

# ── Step 1: Install Ghostscript ───────────────────────────────────────────────
Write-Step "Checking Ghostscript..."

$gsFound = $null
Get-ChildItem "C:\Program Files\gs" -ErrorAction SilentlyContinue | ForEach-Object {
    $candidate = "$($_.FullName)\bin\gswin64c.exe"
    if (Test-Path $candidate) { $gsFound = $candidate }
}

if ($gsFound) {
    Write-OK "Ghostscript already installed: $gsFound"
    $GsExe = $gsFound
} else {
    Write-Step "Installing Ghostscript..."

    # Try Chocolatey first
    $chocoAvailable = (Get-Command choco -ErrorAction SilentlyContinue) -ne $null
    if ($chocoAvailable) {
        Write-Host "   Using Chocolatey..." -ForegroundColor Gray
        choco install ghostscript -y --no-progress 2>&1 | Out-Null
    } else {
        # Direct download fallback
        Write-Host "   Downloading Ghostscript $GsVersion..." -ForegroundColor Gray
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($GsDownload, $GsInstaller)
        Write-Host "   Running installer (silent)..." -ForegroundColor Gray
        Start-Process -FilePath $GsInstaller -ArgumentList "/S" -Wait
        Remove-Item $GsInstaller -Force -ErrorAction SilentlyContinue
    }

    # Locate the installed exe
    Get-ChildItem "C:\Program Files\gs" -ErrorAction SilentlyContinue | ForEach-Object {
        $candidate = "$($_.FullName)\bin\gswin64c.exe"
        if (Test-Path $candidate) { $gsFound = $candidate }
    }

    if (-not $gsFound) {
        Write-Host "`nERROR: Ghostscript installation failed. Install manually from:" -ForegroundColor Red
        Write-Host "  https://www.ghostscript.com/releases/gsdnld.html" -ForegroundColor Yellow
        pause; exit 1
    }
    $GsExe = $gsFound
    Write-OK "Ghostscript installed: $GsExe"
}

# ── Step 2: Deploy compression script ────────────────────────────────────────
Write-Step "Deploying compression script to $InstallDir..."
New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null

$psScript = @"
# compress_pdf.ps1  (deployed by Install-CompressPDF.ps1)
param([string]`$InputFile)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Find Ghostscript
`$gs = Get-ChildItem "C:\Program Files\gs" -ErrorAction SilentlyContinue |
       ForEach-Object { "`$(`$_.FullName)\bin\gswin64c.exe" } |
       Where-Object { Test-Path `$_ } |
       Select-Object -Last 1

if (-not `$gs) {
    [System.Windows.Forms.MessageBox]::Show("Ghostscript not found.`nPlease re-run the installer.", "Compress PDF", "OK", "Error")
    exit 1
}

`$form = New-Object System.Windows.Forms.Form
`$form.Text = "Compress PDF"
`$form.Size = New-Object System.Drawing.Size(400, 230)
`$form.StartPosition = "CenterScreen"
`$form.FormBorderStyle = "FixedDialog"
`$form.MaximizeBox = `$false
`$form.TopMost = `$true

`$label = New-Object System.Windows.Forms.Label
`$label.Text = "Quality for: " + [System.IO.Path]::GetFileName(`$InputFile)
`$label.Location = New-Object System.Drawing.Point(20, 20)
`$label.Size = New-Object System.Drawing.Size(360, 40)
`$form.Controls.Add(`$label)

`$combo = New-Object System.Windows.Forms.ComboBox
`$combo.Location = New-Object System.Drawing.Point(20, 70)
`$combo.Size = New-Object System.Drawing.Size(360, 30)
`$combo.DropDownStyle = "DropDownList"
`$combo.Items.AddRange(@(
    "Screen / Email  (smallest ~72 DPI)",
    "eBook           (balanced ~150 DPI)",
    "Print Quality   (high ~300 DPI)",
    "Prepress        (maximum quality)"
))
`$combo.SelectedIndex = 1
`$form.Controls.Add(`$combo)

`$btnOK = New-Object System.Windows.Forms.Button
`$btnOK.Text = "Compress"
`$btnOK.Location = New-Object System.Drawing.Point(190, 155)
`$btnOK.Size = New-Object System.Drawing.Size(90, 30)
`$btnOK.DialogResult = "OK"
`$form.Controls.Add(`$btnOK)
`$form.AcceptButton = `$btnOK

`$btnCancel = New-Object System.Windows.Forms.Button
`$btnCancel.Text = "Cancel"
`$btnCancel.Location = New-Object System.Drawing.Point(290, 155)
`$btnCancel.Size = New-Object System.Drawing.Size(80, 30)
`$btnCancel.DialogResult = "Cancel"
`$form.Controls.Add(`$btnCancel)
`$form.CancelButton = `$btnCancel

if (`$form.ShowDialog() -ne "OK") { exit }

`$qualities = @("/screen", "/ebook", "/printer", "/prepress")
`$quality = `$qualities[`$combo.SelectedIndex]

`$output = [System.IO.Path]::Combine(
    [System.IO.Path]::GetDirectoryName(`$InputFile),
    [System.IO.Path]::GetFileNameWithoutExtension(`$InputFile) + "_compressed.pdf"
)

`$proc = Start-Process -FilePath `$gs `
    -ArgumentList "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dPDFSETTINGS=`$quality", `
                  "-dNOPAUSE", "-dQUIET", "-dBATCH", `
                  "-sOutputFile=``"`$output``"", "``"`$InputFile``"" `
    -Wait -PassThru -WindowStyle Hidden

if (`$proc.ExitCode -eq 0) {
    `$inSize  = [math]::Round((Get-Item `$InputFile).Length / 1MB, 1)
    `$outSize = [math]::Round((Get-Item `$output).Length   / 1MB, 2)
    [System.Windows.Forms.MessageBox]::Show(
        "Done!``n``nOriginal:   `$inSize MB``nCompressed: `$outSize MB``n``nSaved as:``n`$output",
        "Compress PDF", "OK", "Information")
} else {
    [System.Windows.Forms.MessageBox]::Show("Compression failed.", "Compress PDF", "OK", "Error")
}
"@

Set-Content -Path $ScriptFile -Value $psScript -Encoding UTF8
Write-OK "Script written to $ScriptFile"

# ── Step 3: Registry (all users via HKLM) ────────────────────────────────────
Write-Step "Adding context menu entry for all users (HKLM)..."

$regBase = "HKLM:\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF"
Remove-Item -Path $regBase -Recurse -Force -ErrorAction SilentlyContinue

New-Item -Path "$regBase\command" -Force | Out-Null
Set-ItemProperty -Path $regBase -Name "(Default)" -Value "Compress PDF"
Set-ItemProperty -Path $regBase -Name "Icon"      -Value "$GsExe,0"

$cmd = "powershell.exe -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File " +
       [char]34 + $ScriptFile + [char]34 +
       " -InputFile " + [char]34 + "%1" + [char]34

Set-ItemProperty -Path "$regBase\command" -Name "(Default)" -Value $cmd
Write-OK "Registry entry added"

# ── Step 4: Restart Explorer ──────────────────────────────────────────────────
Write-Step "Restarting Explorer to apply changes..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Start-Process explorer
Write-OK "Explorer restarted"

# ── Done ──────────────────────────────────────────────────────────────────────
Write-Host "`n================================================" -ForegroundColor Green
Write-Host "  Compress PDF installed successfully!" -ForegroundColor Green
Write-Host "  Right-click any PDF to use it." -ForegroundColor Green
Write-Host "================================================`n" -ForegroundColor Green
pause
