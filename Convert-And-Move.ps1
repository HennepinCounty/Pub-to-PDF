<#
.SYNOPSIS
    Recursively finds all .pub files, converts each using Microsoft's
    Convert-PubFileToPDF.ps1 (must be in the same folder as this script),
    then moves each PDF immediately after conversion.
    Uses a GUI folder picker, logs errors, and writes a timestamped run report.

.NOTES
    Updated with assistance from GitHub Copilot.
    Model used in this session: GPT-5.3-Codex.
    Additional edits may have been made by a human author or other tooling.
#>

Add-Type -AssemblyName System.Windows.Forms

# ── Timestamp for this run ────────────────────────────────────────────────────
$RunStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$RunDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

# ── Output files next to this script ─────────────────────────────────────────
$ErrorLog = Join-Path $PSScriptRoot "Convert-Errors-$RunStamp.log"
$ReportLog = Join-Path $PSScriptRoot "Convert-Report-$RunStamp.txt"

# ── Helper: GUI folder picker ─────────────────────────────────────────────────
function Select-Folder($description) {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $description
    $dialog.ShowNewFolderButton = $true
    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.SelectedPath
    }
    else {
        Write-Host 'No folder selected. Exiting.'
        exit 0
    }
}

# ── Helper: Log error ─────────────────────────────────────────────────────────
function Write-Log($message) {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $ErrorLog -Value "[$timestamp] $message"
    Write-Warning $message
}

# ── Helper: Write to report ───────────────────────────────────────────────────
function Write-Report($message) {
    Add-Content -Path $ReportLog -Value $message
}

# ── Auto-detect Microsoft script ──────────────────────────────────────────────
$MicrosoftScript = Join-Path $PSScriptRoot 'Convert-PubFileToPDF.ps1'

if (-not (Test-Path $MicrosoftScript)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Could not find Convert-PubFileToPDF.ps1 in the same folder as this script.`n`n$PSScriptRoot`n`nMake sure both scripts are in the same folder.",
        'Missing Script',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# ── GUI: Pick folders ─────────────────────────────────────────────────────────
Write-Host 'Select the root folder where the .pub files are located...'
$RootFolder = Select-Folder 'Select the root folder where the .pub files are located'
Write-Host "Root folder   : $RootFolder"

Write-Host 'Select the folder where the converted PDFs should go...'
$OutputFolder = Select-Folder 'Select the folder where the converted PDFs should go'
Write-Host "Output folder : $OutputFolder"

# ── Validate root folder ──────────────────────────────────────────────────────
if (-not (Test-Path $RootFolder)) {
    Write-Log "RootFolder not found: $RootFolder"
    exit 1
}

# ── Create output folder if it doesn't exist ─────────────────────────────────
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "Created output folder: $OutputFolder"
}

# ── Build list of all .pub files ─────────────────────────────────────────────
$pubFiles = Get-ChildItem -Path $RootFolder -Recurse -Filter '*.pub' -File

if ($pubFiles.Count -eq 0) {
    Write-Host "No .pub files found under: $RootFolder"
    exit 0
}

Write-Host "`nFound $($pubFiles.Count) .pub file(s).`n"

# ── Write report header ───────────────────────────────────────────────────────
Write-Report '======================================='
Write-Report '  Publisher to PDF Conversion Report'
Write-Report '======================================='
Write-Report "Run date      : $RunDate"
Write-Report "Root folder   : $RootFolder"
Write-Report "Output folder : $OutputFolder"
Write-Report "Files found   : $($pubFiles.Count)"
Write-Report '---------------------------------------'
Write-Report ''
Write-Report 'FILES FOUND:'
Write-Report ''
foreach ($file in $pubFiles) {
    Write-Report "  $($file.FullName)"
}
Write-Report ''
Write-Report '---------------------------------------'
Write-Report ''
Write-Report 'RESULTS:'
Write-Report ''

# ── Single loop: convert then immediately move each file ──────────────────────
Write-Host "--- Starting conversion and move ---`n"

$convertSuccess = 0
$convertFail = 0
$moveSuccess = 0
$moveFail = 0
$skippedCount = 0
$totalFiles = $pubFiles.Count
$currentIndex = 0

foreach ($file in $pubFiles) {
    $currentIndex++
    $percent = [math]::Round(($currentIndex / $totalFiles) * 100)

    Write-Progress `
        -Activity 'Converting Publisher files' `
        -Status "File $currentIndex of $totalFiles : $($file.Name)" `
        -PercentComplete $percent

    # ── Restart Publisher every 100 files to prevent memory buildup ─────────────
    if ($currentIndex -gt 1 -and ($currentIndex - 1) % 100 -eq 0) {
        Write-Host "  [RESTART] Restarting Publisher after 100 files to clear memory...`n"
        Get-Process -Name 'MSPUB' -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3
    }

    Write-Host "Converting: $($file.FullName)"

    # Convert with timeout
    $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, '.pdf')
    $timedOut = $false

    $job = Start-Job -ScriptBlock {
        param($scriptPath, $filePath)
        . $scriptPath
        Convert-PubFileToPdf -Path $filePath
    } -ArgumentList $MicrosoftScript, $file.FullName

    $completed = Wait-Job -Job $job -Timeout 30

    if (-not $completed) {
        # Timed out — stop the job and kill Publisher
        Stop-Job -Job $job
        Remove-Job -Job $job -Force
        Get-Process -Name 'MSPUB' -ErrorAction SilentlyContinue | Stop-Process -Force
        $timedOut = $true

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -Path $ErrorLog -Value ''
        Add-Content -Path $ErrorLog -Value "[$timestamp] SKIPPED (timeout after 30s): $($file.FullName)"
        Write-Report "  [SKIPPED - TIMEOUT] $($file.FullName)"
        Write-Warning "  [SKIPPED] Timed out after 30 seconds: $($file.FullName)`n"
        $convertFail++
        $skippedCount++
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        continue
    }

    Remove-Job -Job $job -Force

    if (Test-Path $pdfPath) {
        Write-Host '  [CONVERTED]'
        $convertSuccess++
    }
    else {
        Write-Log "CONVERSION FAILED - PDF not created for: $($file.FullName)"
        Write-Report "  [CONVERSION FAILED] $($file.FullName)"
        $convertFail++
        continue
    }

    # Move immediately after successful conversion
    $pdfName = [System.IO.Path]::GetFileName($pdfPath)
    $destPath = Join-Path $OutputFolder $pdfName

    $counter = 1
    while (Test-Path $destPath) {
        $nameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
        $destPath = Join-Path $OutputFolder "${nameNoExt}_$counter.pdf"
        $counter++
    }

    Write-Host "  Moving to: $destPath"

    try {
        Move-Item -Path $pdfPath -Destination $destPath
        Write-Host "  [OK]`n"
        Write-Report "  [OK] $($file.FullName)"
        Write-Report "    -> $destPath"
        Write-Report ''
        $moveSuccess++
    }
    catch {
        Write-Log "MOVE ERROR - $($pdfPath): $_"
        Write-Report "  [MOVE ERROR] $($pdfPath): $_"
        $moveFail++
    }
}

# ── Clear progress bar ──────────────────────────────────────────────────────
Write-Progress -Activity 'Converting Publisher files' -Completed

# ── Write report footer ───────────────────────────────────────────────────────
Write-Report '---------------------------------------'
Write-Report 'SUMMARY'
Write-Report '---------------------------------------'
Write-Report "Converted : $convertSuccess file(s)"
Write-Report "Moved     : $moveSuccess file(s)"
Write-Report "Failed    : $($convertFail + $moveFail) file(s)"
Write-Report "Skipped   : $skippedCount file(s) (timed out - check error log)"
Write-Report ''
Write-Report "Note: Check Convert-Errors-$RunStamp.log for SKIPPED files (timeout) and errors."
Write-Report '======================================='

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host '======================================='
Write-Host "Converted : $convertSuccess file(s)"
Write-Host "Moved     : $moveSuccess file(s)"
Write-Host "Failed    : $($convertFail + $moveFail) file(s)"
Write-Host "Skipped   : $skippedCount file(s) (timed out)"
Write-Host "Output    : $OutputFolder"
Write-Host "Report    : $ReportLog"

if ($convertFail -gt 0 -or $moveFail -gt 0 -or $skippedCount -gt 0) {
    Write-Host "`nErrors were logged to: $ErrorLog"
    [System.Windows.Forms.MessageBox]::Show(
        "Completed with errors.`n`nConverted: $convertSuccess`nMoved: $moveSuccess`nFailed: $($convertFail + $moveFail)`nSkipped (timeout): $skippedCount`n`nCheck the error log for skipped files:`n$ErrorLog`n`nSee report:`n$ReportLog",
        'Completed with Errors',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
}
else {
    [System.Windows.Forms.MessageBox]::Show(
        "All done!`n`nConverted: $convertSuccess`nMoved: $moveSuccess`nOutput: $OutputFolder`n`nSee report:`n$ReportLog",
        'Complete',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
