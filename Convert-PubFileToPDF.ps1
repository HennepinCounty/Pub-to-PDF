#Requires -Version 5.1
Set-StrictMode -Version Latest

function Convert-PubFileToPdf {
    <#
    .SYNOPSIS
    Converts Microsoft Publisher .pub files to PDF format.

    .DESCRIPTION
    Converts Microsoft Publisher .pub files to PDF using Microsoft Publisher COM automation.
    Microsoft Publisher must be installed and accessible on the computer running this function.
    Path values may identify a file, directory, or wildcard pattern. Directories and patterns
    are searched for .pub files; use Recurse to include subdirectories.

    Path accepts pipeline input from strings or objects with a FullName property. Each successful
    conversion writes one object to the output pipeline. Skipped, ignored, and failed items do
    not write success objects; they are reported through the warning, error, or verbose streams.
    Use TransactionLogPath to record all attempted outcomes in an append-only CSV file.

    .PARAMETER Path
    One or more file system paths, wildcard patterns, or directories. A file must have a .pub
    extension. Directories and wildcard patterns select .pub files for conversion.
    Accepts pipeline input by value and pipeline property input through the FullName alias.

    .PARAMETER Recurse
    When Path points to a directory or wildcard pattern, search subdirectories for .pub files.

    .PARAMETER TransactionLogPath
    Optional path to a CSV file. When specified, the function appends one record for each
    successful, skipped, ignored, failed, or no-file-found outcome. The parent directory is
    created when it does not exist. Log columns are Timestamp, SourcePath, PdfPath, Status,
    and Message.

    .PARAMETER ErrorOnExistingPdf
    When set, an existing PDF is reported through the error stream. Without this switch,
    existing PDFs are skipped and reported through the verbose stream. Existing PDFs never
    produce success output.

    .EXAMPLE
    Convert-PubFileToPdf -Path "C:\Docs\Input\Newsletter.pub"

    .EXAMPLE
    Get-ChildItem -Path "C:\Docs\Input" -Filter "*.pub" -File -Recurse |
        Convert-PubFileToPdf -TransactionLogPath "C:\Logs\pub-conversion.csv"

    .EXAMPLE
    Convert-PubFileToPdf -Path "C:\Docs\Input" -Recurse -Verbose

    Searches the directory and its subdirectories, converts each .pub file, and displays
    verbose messages for skipped files and the final conversion summary.

    .OUTPUTS
    PSCustomObject
    On success, each object contains SourcePath, PdfPath, and ConvertedAt properties.

    .INPUTS
    System.String
    System.IO.FileInfo
    System.IO.DirectoryInfo
    Pipeline input is accepted by value as a path string or by the FullName property.

    .NOTES
    Updated with assistance from GitHub Copilot.
    Model used in this session: GPT-5.3-Codex.
    Additional edits may have been made by a human author or other tooling.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Path,

        [switch]
        $Recurse,

        [ValidateNotNullOrEmpty()]
        [string]
        $TransactionLogPath,

        [switch]
        $ErrorOnExistingPdf
    )

    begin {
        $publisherApp = $null
        $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $logEnabled = $PSBoundParameters.ContainsKey('TransactionLogPath')
        $successCount = 0
        $failureCount = 0
        $skippedCount = 0

        $writeTransaction = {
            param (
                [string]$SourcePath,
                [string]$PdfPath,
                [string]$Status,
                [string]$Message
            )

            if (-not $logEnabled) {
                return
            }

            $logRecord = [pscustomobject]@{
                Timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                SourcePath = $SourcePath
                PdfPath    = $PdfPath
                Status     = $Status
                Message    = $Message
            }

            $exportParams = @{
                Path              = $TransactionLogPath
                NoTypeInformation = $true
                Append            = $true
            }

            $logRecord | Export-Csv @exportParams
        }

        if ($logEnabled) {
            $logDirectory = Split-Path -Path $TransactionLogPath -Parent
            if ([string]::IsNullOrWhiteSpace($logDirectory)) {
                $resolvedLocation = Get-Location
                $TransactionLogPath = Join-Path -Path $resolvedLocation.Path -ChildPath $TransactionLogPath
                $logDirectory = Split-Path -Path $TransactionLogPath -Parent
            }

            if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
                $directoryParams = @{
                    Path     = $logDirectory
                    ItemType = 'Directory'
                    Force    = $true
                }
                $null = New-Item @directoryParams
            }
        }

        try {
            # Add-Type -AssemblyName Microsoft.Office.Interop.Publisher # This doesn't work in PowerShell 7+; use Add-Type -Path with the DLL path instead.
            Add-Type -Path 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Publisher\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Publisher.dll'
        }
        catch {
            throw 'Unable to load Microsoft.Office.Interop.Publisher. Ensure Microsoft Publisher is installed.'
        }

        $stopwatch.Restart()
        try {
            $publisherApp = New-Object -ComObject Publisher.Application
        }
        catch {
            throw 'Unable to create Publisher.Application COM object. Ensure Microsoft Publisher is installed and accessible.'
        }
        finally {
            $stopwatch.Stop()
            Write-Verbose "Publisher COM object loaded in $($stopwatch.Elapsed.TotalMilliseconds.ToString('N2')) ms"
        }

        if ($null -eq $publisherApp) {
            throw 'Publisher COM object initialization returned null.'
        }
    }

    process {
        foreach ($inputPath in $Path) {
            $candidateFiles = @()

            if (Test-Path -LiteralPath $inputPath -PathType Leaf) {
                $fileItem = Get-Item -LiteralPath $inputPath
                if ($fileItem.Extension -ne '.pub') {
                    $failureCount++
                    & $writeTransaction $fileItem.FullName '' 'IgnoredNonPublisherFile' 'Input file does not have a .pub extension.'
                    Write-Warning "Ignoring non-.pub file: $($fileItem.FullName)"
                    continue
                }

                $candidateFiles = @($fileItem)
            }
            elseif (Test-Path -LiteralPath $inputPath -PathType Container) {
                $searchParams = @{
                    LiteralPath = $inputPath
                    Filter      = '*.pub'
                    File        = $true
                    Recurse     = [bool]$Recurse
                }
                $candidateFiles = @(Get-ChildItem @searchParams)
            }
            else {
                $searchParams = @{
                    Path        = $inputPath
                    Filter      = '*.pub'
                    File        = $true
                    Recurse     = [bool]$Recurse
                    ErrorAction = 'SilentlyContinue'
                }
                $candidateFiles = @(Get-ChildItem @searchParams)
            }

            if ($candidateFiles.Count -eq 0) {
                $failureCount++
                & $writeTransaction $inputPath '' 'NoPublisherFilesFound' 'No .pub files were found for the supplied path.'
                Write-Warning "No .pub files found for path: $inputPath"
                continue
            }

            foreach ($file in $candidateFiles) {
                $sourcePath = $file.FullName
                $pdfPath = [System.IO.Path]::ChangeExtension($sourcePath, '.pdf')

                if (Test-Path -LiteralPath $pdfPath -PathType Leaf) {
                    $skippedCount++
                    & $writeTransaction $sourcePath $pdfPath 'SkippedExistingPdf' 'PDF already exists.'

                    if ($ErrorOnExistingPdf) {
                        $failureCount++
                        Write-Error "PDF already exists: $pdfPath"
                    }
                    else {
                        Write-Verbose "Skipping because PDF already exists: $pdfPath"
                    }

                    continue
                }

                $document = $null
                $stopwatch.Restart()
                try {
                    $document = $publisherApp.Open($sourcePath)
                    if ($null -eq $document) {
                        throw "Publisher returned null document for '$sourcePath'."
                    }

                    $document.ExportAsFixedFormat(
                        [Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF,
                        $pdfPath
                    )

                    if (-not (Test-Path -LiteralPath $pdfPath -PathType Leaf)) {
                        throw "Export completed without creating expected PDF '$pdfPath'."
                    }

                    $successCount++
                    & $writeTransaction $sourcePath $pdfPath 'Success' 'Converted successfully.'

                    [pscustomobject]@{
                        SourcePath  = $sourcePath
                        PdfPath     = $pdfPath
                        ConvertedAt = Get-Date
                    }
                }
                catch {
                    $failureCount++
                    $errorMessage = $_.Exception.Message
                    & $writeTransaction $sourcePath $pdfPath 'Failed' $errorMessage
                    Write-Warning "Conversion failed for '$sourcePath': $errorMessage"
                }
                finally {
                    $stopwatch.Stop()
                    Write-Verbose "Conversion for '$sourcePath' took $($stopwatch.Elapsed.TotalMilliseconds.ToString('N2')) ms"

                    if ($null -ne $document) {
                        try {
                            $document.Close()
                        }
                        catch {
                            Write-Verbose "Document close failed for '$sourcePath': $($_.Exception.Message)"
                        }

                        try {
                            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
                        }
                        catch {
                            Write-Verbose "ReleaseComObject failed for '$sourcePath': $($_.Exception.Message)"
                        }
                    }
                }
            }
        }
    }

    end {
        if ($null -ne $publisherApp) {
            try {
                $publisherApp.Quit()
            }
            catch {
                Write-Verbose "Publisher quit failed: $($_.Exception.Message)"
            }

            try {
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($publisherApp)
            }
            catch {
                Write-Verbose "ReleaseComObject failed for Publisher app: $($_.Exception.Message)"
            }
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        Write-Verbose "Completed conversion run. Success: $successCount, Failed: $failureCount, Skipped: $skippedCount"
    }
}
