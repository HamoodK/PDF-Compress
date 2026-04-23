# PDF-Compress
Easily Compress PDF files in windows from sub menu.

# Compress PDF — Windows Context Menu Tool

Right-click any PDF on Windows and compress it instantly using Ghostscript.
No third-party apps, no uploads, works fully offline.

## What it does

- Adds a **"Compress PDF"** option to the right-click menu for all PDF files
- Opens a small dialog to choose compression quality
- Saves the compressed file in the same folder as the original (e.g. `report_compressed.pdf`)
- Original file is never modified

## Quality options

| Option | DPI | Best for |
|--------|-----|----------|
| Screen / Email | ~72 DPI | Sending over email or messaging apps |
| eBook | ~150 DPI | General use, good size/quality balance |
| Print Quality | ~300 DPI | Printing documents |
| Prepress | Max | Professional/commercial printing |

## Files

```
Install-CompressPDF.bat     Double-click to install (handles admin elevation)
Install-CompressPDF.ps1     The installer script
Uninstall-CompressPDF.ps1   Removes the context menu entry
```

## Requirements

- Windows 10 or Windows 11
- Internet connection on first install (to download Ghostscript if not present)
- Administrator rights during installation

## Installation

1. Copy `Install-CompressPDF.bat` and `Install-CompressPDF.ps1` to the same folder
2. Double-click `Install-CompressPDF.bat`
3. Click **Yes** on the UAC prompt
4. Wait for the installer to finish — Explorer will restart automatically

The installer will:
- Install Ghostscript automatically (via Chocolatey if available, otherwise direct download)
- Deploy the compression script to `C:\ProgramData\CompressPDF\`
- Register the context menu entry for all users on the PC
- Restore the classic right-click menu on Windows 11

## Usage

1. Right-click any PDF file
2. Click **Compress PDF**
3. Select quality from the dropdown
4. Click **Compress**
5. A popup confirms the result and shows the size reduction

## Uninstall

Right-click `Uninstall-CompressPDF.ps1` and choose **Run with PowerShell**.

This removes the context menu entry and the deployed script.
Ghostscript itself is NOT uninstalled.

## How it works

The installer registers a shell command under:
```
HKLM\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF
```
This makes the entry appear for all users on the machine for all PDF files, regardless of which app is set as the default PDF viewer.

Compression is handled by [Ghostscript](https://www.ghostscript.com/), an open-source PDF engine.
