# PDF-Compress
Easily Compress PDF files in windows from sub menu.

# Compress PDF — Windows Context Menu Tool

Right-click any PDF on Windows and compress it instantly using Ghostscript.
No third-party apps, no uploads, works fully offline.

---

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

---

## Installation

### Option 1 — Online Installer (recommended)

Downloads Ghostscript automatically during installation. Requires internet on first install.

1. Download `CompressPDF-Setup.exe` from the [Releases](../../releases/latest) page
2. Double-click it and follow the wizard
3. Click **Yes** on the UAC prompt

### Option 2 — Offline Installer

Fully self-contained. No internet required. Best for restricted or corporate environments.

1. Download `CompressPDF-Setup-Offline.exe` from the [Releases](../../releases/latest) page
2. Double-click it and follow the wizard
3. Click **Yes** on the UAC prompt

> Both installers handle everything automatically:
> install Ghostscript, deploy the compression script, and register the right-click menu.

---

## Usage

1. Right-click any PDF file
2. Click **Compress PDF**
3. Select quality from the dropdown
4. Click **Compress**
5. A popup confirms the result and shows the size reduction

---

## Uninstall

Open **Add or Remove Programs** → search for **Compress PDF** → Uninstall.

Or run `Uninstall-CompressPDF.ps1` directly with PowerShell as Administrator.

> Ghostscript itself is NOT uninstalled.

---

## Requirements

- Windows 10 or Windows 11
- Administrator rights during installation
- Internet connection only for the **online** installer

---

## Files

```
Install-CompressPDF.ps1          Online installer script
Install-CompressPDF-Offline.ps1  Offline installer script (uses bundled Ghostscript)
Install-CompressPDF.bat          Legacy launcher (runs online installer as admin)
Uninstall-CompressPDF.ps1        Removes the context menu entry
```

---

## How it works

The installer registers a shell command under:
```
HKLM\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF
```
This makes the entry appear for all users on the machine for all PDF files,
regardless of which app is set as the default PDF viewer.

Compression is handled by [Ghostscript](https://www.ghostscript.com/), an open-source PDF engine.
