#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs "Compress PDF" right-click context menu using Ghostscript.
.DESCRIPTION
    - Downloads the latest Ghostscript (falls back to a pinned version if the
      release naming convention has changed or the API is unreachable)
    - Deploys the compression script to C:\ProgramData\CompressPDF\
    - Adds "Compress PDF" to the right-click menu for all PDF files
.NOTES
    Run as Administrator. Works on Windows 10 and 11.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Pinned fallback (used when the GitHub API / asset pattern changes) ────────
$FallbackVersion      = "10.07.1"
$FallbackDownload     = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10071/gs10071w64.exe"
$FallbackHashAlg      = "SHA512"
$FallbackExpectedHash = "BAE62C525FFE6D6D8A747DC92256A16C90FA20C59BE5BA98494CD2E408395528542212B94E461ACD68F09DC0FD5B96AB04072BB087554854335A9B2AD28A5EB9"

# ── Config ────────────────────────────────────────────────────────────────────
$InstallDir  = "C:\ProgramData\CompressPDF"
$ScriptFile  = "$InstallDir\compress_pdf.ps1"
$GsInstaller = "$env:TEMP\gs_installer.exe"

# ── Helpers ───────────────────────────────────────────────────────────────────
function Write-Step($msg) { Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "   OK: $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "   WARN: $msg" -ForegroundColor Yellow }

function Pause-IfInteractive {
    if ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected) { pause }
}

function Find-GhostscriptExe {
    Get-ChildItem "C:\Program Files\gs" -ErrorAction SilentlyContinue |
        Sort-Object { try { [version]($_.Name -replace '^gs','') } catch { [version]'0.0' } } -Descending |
        Select-Object -First 1 |
        ForEach-Object { "$($_.FullName)\bin\gswin64c.exe" } |
        Where-Object { Test-Path -LiteralPath $_ }
}

function Get-LatestGhostscript {
    <#
    Queries the GitHub releases API for the newest Ghostscript Windows 64-bit
    installer.  Returns a hashtable with Version/DownloadUrl/FileName/
    HashAlgorithm/ExpectedHash, or $null if anything looks unexpected so the
    caller can fall back to the pinned version.
    #>
    try {
        $apiHeaders = @{ 'User-Agent' = 'CompressPDF-Installer/1.0' }
        $release = Invoke-RestMethod `
            -Uri     'https://api.github.com/repos/ArtifexSoftware/ghostpdl-downloads/releases/latest' `
            -Headers $apiHeaders `
            -TimeoutSec 15 `
            -ErrorAction Stop

        # Must match gs<digits>w64.exe — bail if naming changed
        $exeAsset = $release.assets |
                    Where-Object { $_.name -match '^gs\d+w64\.exe$' } |
                    Select-Object -First 1
        if (-not $exeAsset) {
            Write-Warn "Latest release has no gs*w64.exe asset — using pinned fallback."
            return $null
        }

        # Prefer SHA512SUMS, accept SHA256SUMS
        $sumsAsset = $release.assets | Where-Object { $_.name -eq 'SHA512SUMS' } | Select-Object -First 1
        $hashAlg   = 'SHA512'
        if (-not $sumsAsset) {
            $sumsAsset = $release.assets | Where-Object { $_.name -eq 'SHA256SUMS' } | Select-Object -First 1
            $hashAlg   = 'SHA256'
        }
        if (-not $sumsAsset) {
            Write-Warn "Latest release has no checksum file — using pinned fallback."
            return $null
        }

        # Download the checksum file (.Content may be string or byte[] depending on PS version)
        $sumsResp = Invoke-WebRequest -UseBasicParsing -Uri $sumsAsset.browser_download_url `
                        -TimeoutSec 15 -ErrorAction Stop
        $sumsText = if ($sumsResp.Content -is [byte[]]) {
            [System.Text.Encoding]::UTF8.GetString($sumsResp.Content)
        } else {
            [string]$sumsResp.Content
        }

        # Lines are: "<hash>  <filename>" or "<hash> *<filename>"
        $hashLine = ($sumsText -split "`n") |
                    Where-Object { $_ -match [regex]::Escape($exeAsset.name) } |
                    Select-Object -First 1
        $hash = ($hashLine -split '\s+')[0].Trim()

        $expectedLen = if ($hashAlg -eq 'SHA512') { 128 } else { 64 }
        if ($hash.Length -ne $expectedLen) {
            Write-Warn "Checksum parse failed for $($exeAsset.name) — using pinned fallback."
            return $null
        }

        # Extract version from release name e.g. "Ghostscript/GhostPDL 10.07.1 Release"
        $version = if ($release.name -match '(\d+\.\d+\.\d+)') { $matches[1] } else { $null }
        if (-not $version) {
            Write-Warn "Cannot parse version from release name '$($release.name)' — using pinned fallback."
            return $null
        }

        return @{
            Version       = $version
            DownloadUrl   = $exeAsset.browser_download_url
            FileName      = $exeAsset.name
            HashAlgorithm = $hashAlg
            ExpectedHash  = $hash.ToUpper()
        }
    } catch {
        Write-Warn "GitHub API unreachable ($_) — using pinned fallback."
        return $null
    }
}

# ── Step 1: Install Ghostscript ───────────────────────────────────────────────
Write-Step "Checking Ghostscript..."

$gsFound = Find-GhostscriptExe

if ($gsFound) {
    Write-OK "Ghostscript already installed: $gsFound"
    $GsExe = $gsFound
} else {
    # Check Chocolatey first — skip API call entirely if choco will handle the install
    $chocoAvailable = Get-Command choco -ErrorAction SilentlyContinue
    if ($chocoAvailable) {
        Write-Step "Installing Ghostscript via Chocolatey..."
        choco install ghostscript -y --no-progress 2>&1 | Out-Null
    } else {
        # Resolve the version to download
        Write-Step "Resolving latest Ghostscript version..."
        $latest = Get-LatestGhostscript

        $GsVersion      = if ($latest) { $latest.Version }       else { $FallbackVersion }
        $GsDownload     = if ($latest) { $latest.DownloadUrl }   else { $FallbackDownload }
        $GsHashAlg      = if ($latest) { $latest.HashAlgorithm } else { $FallbackHashAlg }
        $GsExpectedHash = if ($latest) { $latest.ExpectedHash }  else { $FallbackExpectedHash }

        if ($latest) {
            Write-OK "Latest version: $GsVersion ($($latest.FileName))"
        } else {
            Write-Warn "Falling back to pinned version: $GsVersion"
        }

        Write-Step "Downloading Ghostscript $GsVersion..."
        try {
            Invoke-WebRequest -UseBasicParsing -Uri $GsDownload -OutFile $GsInstaller -TimeoutSec 120
        } catch {
            Write-Host "`nERROR: Download failed — $_" -ForegroundColor Red
            Write-Host "  Install manually from: https://www.ghostscript.com/releases/gsdnld.html" -ForegroundColor Yellow
            Pause-IfInteractive; exit 1
        }

        Write-Host "   Verifying installer integrity ($GsHashAlg)..." -ForegroundColor Gray
        $actualHash = (Get-FileHash $GsInstaller -Algorithm $GsHashAlg).Hash
        if ($actualHash -ne $GsExpectedHash) {
            Remove-Item $GsInstaller -Force -ErrorAction SilentlyContinue
            Write-Host "`nERROR: Installer $GsHashAlg hash mismatch — download may be corrupt or tampered." -ForegroundColor Red
            Write-Host "  Expected: $GsExpectedHash" -ForegroundColor Yellow
            Write-Host "  Got:      $actualHash" -ForegroundColor Yellow
            Pause-IfInteractive; exit 1
        }

        Write-Host "   Running installer (silent)..." -ForegroundColor Gray
        Start-Process -FilePath $GsInstaller -ArgumentList "/S" -Wait
        Remove-Item $GsInstaller -Force -ErrorAction SilentlyContinue
    }

    $gsFound = Find-GhostscriptExe

    if (-not $gsFound) {
        Write-Host "`nERROR: Ghostscript installation failed. Install manually from:" -ForegroundColor Red
        Write-Host "  https://www.ghostscript.com/releases/gsdnld.html" -ForegroundColor Yellow
        Pause-IfInteractive; exit 1
    }
    $GsExe = $gsFound
    Write-OK "Ghostscript installed: $GsExe"
}

# ── Step 2: Deploy compression script ────────────────────────────────────────
Write-Step "Deploying compression script to $InstallDir..."
New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null

$sourceScript = Join-Path $PSScriptRoot "compress_pdf.ps1"
if (-not (Test-Path -LiteralPath $sourceScript)) {
    Write-Host "`nERROR: compress_pdf.ps1 not found alongside installer at:" -ForegroundColor Red
    Write-Host "  $sourceScript" -ForegroundColor Yellow
    Pause-IfInteractive; exit 1
}
Copy-Item -LiteralPath $sourceScript -Destination $ScriptFile -Force
Write-OK "Script copied to $ScriptFile"

# ── Step 3: Registry (all users via HKLM) ────────────────────────────────────
Write-Step "Adding context menu entry for all users (HKLM)..."

$regBase = "HKLM:\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF"
Remove-Item -Path $regBase -Recurse -Force -ErrorAction SilentlyContinue

New-Item -Path "$regBase\command" -Force | Out-Null
Set-ItemProperty -Path $regBase -Name "(Default)" -Value "Compress PDF"
Set-ItemProperty -Path $regBase -Name "Icon"      -Value "$GsExe,0"

$cmd = "powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File " +
       [char]34 + $ScriptFile + [char]34 +
       " -InputFile " + [char]34 + "%1" + [char]34

Set-ItemProperty -Path "$regBase\command" -Name "(Default)" -Value $cmd
Write-OK "Registry entry added"

# ── Done ──────────────────────────────────────────────────────────────────────
Write-Host "`n================================================" -ForegroundColor Green
Write-Host "  Compress PDF installed successfully!" -ForegroundColor Green
Write-Host "  Right-click any PDF to use it." -ForegroundColor Green
Write-Host "================================================`n" -ForegroundColor Green
Pause-IfInteractive
