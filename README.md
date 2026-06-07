# Compress PDF — Windows Context Menu Tool

Right-click any PDF on Windows and compress it instantly using Ghostscript.
No third-party apps, no uploads, works fully offline.

---

## What it does

- Adds a **"Compress PDF"** option to the right-click menu for all PDF files
- Opens a small dialog to choose compression quality
- Saves the compressed file next to the original (e.g. `report_compressed.pdf`)
- If the output name already exists, a counter suffix is added automatically (`report_compressed_1.pdf`)
- Original file is never modified
- Logs each run to `%TEMP%\compress_pdf.log` for troubleshooting

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

Automatically downloads the latest Ghostscript during installation. Requires internet on first install.

1. Download `CompressPDF-Setup.exe` from the [Releases](../../releases/latest) page
2. Double-click it and follow the wizard
3. Click **Yes** on the UAC prompt

If Ghostscript is already installed, the download is skipped entirely.
If the latest version cannot be determined (no internet, or GitHub changes its release format), the installer automatically falls back to a known-good pinned version.

### Option 2 — Offline Installer

Fully self-contained — Ghostscript is bundled inside the installer. No internet required.
Best for restricted or corporate environments.

1. Download `CompressPDF-Setup-Offline.exe` from the [Releases](../../releases/latest) page
2. Double-click it and follow the wizard
3. Click **Yes** on the UAC prompt

> Both installers handle everything automatically:
> install Ghostscript, deploy the compression script, and register the right-click menu for all users.

---

## Usage

1. Right-click any PDF file
2. Click **Compress PDF**
3. Select a quality level from the dropdown
4. Click **Compress**
5. A popup confirms the result and shows the original vs. compressed size

---

## Uninstall

Open **Settings → Apps** → search for **Compress PDF** → Uninstall.

Or run `Uninstall-CompressPDF.ps1` directly with PowerShell as Administrator.

> Ghostscript itself is **not** uninstalled — only the context menu entry and the deployed script are removed.

---

## Requirements

- Windows 10 or Windows 11 (64-bit)
- Administrator rights during installation
- Internet connection only for the **online** installer (when Ghostscript is not already installed)

---

## Files

```
Install-CompressPDF.ps1          Online installer — downloads latest Ghostscript automatically
Install-CompressPDF-Offline.ps1  Offline installer — uses bundled Ghostscript
Install-CompressPDF.bat          Launcher — runs the online installer as Administrator
Uninstall-CompressPDF.ps1        Removes the context menu entry and deployed script
compress_pdf.ps1                 The compression script deployed to C:\ProgramData\CompressPDF\
CompressPDF-Setup.iss            Inno Setup script for the online installer
CompressPDF-Setup-Offline.iss    Inno Setup script for the offline installer
```

---

## Building the installers

Requires [Inno Setup 6](https://jrsoftware.org/isdl.php).

**Online installer:**
```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "CompressPDF-Setup.iss"
# Output: Output\CompressPDF-Setup.exe
```

**Offline installer** (place `gs10071w64.exe` in the same folder first):
```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "CompressPDF-Setup-Offline.iss"
# Output: Output\CompressPDF-Setup-Offline.exe
```

Or open either `.iss` file in the Inno Setup Compiler GUI and press **F9**.

---

## How it works

The installer registers a shell command under:
```
HKLM\Software\Classes\SystemFileAssociations\.pdf\shell\CompressPDF
```
This makes the entry appear for **all users** on the machine for all PDF files,
regardless of which app is set as the default PDF viewer.

Compression is handled by [Ghostscript](https://www.ghostscript.com/), an open-source PDF engine.
The online installer verifies the Ghostscript download against its official SHA-512 checksum
before running it.
