param([string]$InputFile)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$logFile = "$env:TEMP\compress_pdf.log"

function Write-Log($msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    try { "$ts  $msg" | Out-File -FilePath $logFile -Append -Encoding UTF8 } catch {}
}

# Validate input
if (-not $InputFile -or -not (Test-Path -LiteralPath $InputFile)) {
    $err = "File not found or not specified: $InputFile"
    Write-Log $err
    [System.Windows.Forms.MessageBox]::Show($err, "Compress PDF", "OK", "Error")
    exit 1
}

# Find newest Ghostscript
$gsExe = Get-ChildItem "C:\Program Files\gs" -ErrorAction SilentlyContinue |
         Sort-Object { try { [version]($_.Name -replace '^gs','') } catch { [version]'0.0' } } -Descending |
         Select-Object -First 1 |
         ForEach-Object { "$($_.FullName)\bin\gswin64c.exe" } |
         Where-Object { Test-Path -LiteralPath $_ }

if (-not $gsExe) {
    $err = "Ghostscript not found. Please re-run the installer."
    Write-Log $err
    [System.Windows.Forms.MessageBox]::Show($err, "Compress PDF", "OK", "Error")
    exit 1
}

Write-Log "Using Ghostscript: $gsExe"
Write-Log "Input: $InputFile"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Compress PDF"
$form.Size = New-Object System.Drawing.Size(400, 230)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.TopMost = $true

$label = New-Object System.Windows.Forms.Label
$label.Text = "Quality for: " + [System.IO.Path]::GetFileName($InputFile)
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(360, 40)
$form.Controls.Add($label)

$combo = New-Object System.Windows.Forms.ComboBox
$combo.Location = New-Object System.Drawing.Point(20, 70)
$combo.Size = New-Object System.Drawing.Size(360, 30)
$combo.DropDownStyle = "DropDownList"
$combo.Items.AddRange(@(
    "Screen / Email  (smallest ~72 DPI)",
    "eBook           (balanced ~150 DPI)",
    "Print Quality   (high ~300 DPI)",
    "Prepress        (maximum quality)"
))
$combo.SelectedIndex = 1
$form.Controls.Add($combo)

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Text = "Compress"
$btnOK.Location = New-Object System.Drawing.Point(190, 155)
$btnOK.Size = New-Object System.Drawing.Size(90, 30)
$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnOK)
$form.AcceptButton = $btnOK

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(290, 155)
$btnCancel.Size = New-Object System.Drawing.Size(80, 30)
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($btnCancel)
$form.CancelButton = $btnCancel

if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }

$qualities = @("/screen", "/ebook", "/printer", "/prepress")
$quality   = $qualities[$combo.SelectedIndex]

# Build unique output path — never silently overwrite
$dir      = [System.IO.Path]::GetDirectoryName($InputFile)
if ([string]::IsNullOrEmpty($dir)) { $dir = $PWD.Path }
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$output   = [System.IO.Path]::Combine($dir, "${baseName}_compressed.pdf")
$counter  = 1
while (Test-Path -LiteralPath $output) {
    $output = [System.IO.Path]::Combine($dir, "${baseName}_compressed_$counter.pdf")
    $counter++
}

Write-Log "Output: $output  Quality: $quality"

# Single-string argument list avoids double-quoting artifacts with the array form
$gsArgs = "-sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=$quality " +
          "-dNOPAUSE -dQUIET -dBATCH " +
          "`"-sOutputFile=$output`" `"$InputFile`""

try {
    $proc = Start-Process -FilePath $gsExe -ArgumentList $gsArgs -Wait -PassThru -WindowStyle Hidden
} catch {
    $err = "Failed to launch Ghostscript: $_"
    Write-Log $err
    [System.Windows.Forms.MessageBox]::Show($err, "Compress PDF", "OK", "Error")
    exit 1
}

if ($proc.ExitCode -eq 0) {
    $inSize  = [math]::Round((Get-Item -LiteralPath $InputFile).Length / 1MB, 1)
    $outSize = [math]::Round((Get-Item -LiteralPath $output).Length   / 1MB, 2)
    Write-Log "Success. $inSize MB -> $outSize MB"
    [System.Windows.Forms.MessageBox]::Show(
        "Done!`n`nOriginal:   $inSize MB`nCompressed: $outSize MB`n`nSaved as:`n$output",
        "Compress PDF", "OK", "Information")
} else {
    $err = "Ghostscript exited with code $($proc.ExitCode)"
    Write-Log $err
    [System.Windows.Forms.MessageBox]::Show(
        "Compression failed.`nSee log: $logFile", "Compress PDF", "OK", "Error")
}
