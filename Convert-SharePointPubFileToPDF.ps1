#Requires -Version 5.1
Set-StrictMode -Version Latest

function Convert-SharePointPubFileToPdf {
    <#
    .SYNOPSIS
    Downloads a SharePoint .pub file, converts it to PDF, and uploads the PDF.

    .DESCRIPTION
    Uses the active PnP.PowerShell connection to download a Publisher file to a
    temporary local directory, calls Convert-PubFileToPdf, and uploads the PDF
    beside the source file in the same SharePoint folder. The source file is
    never changed. Dot-source Convert-PubFileToPDF.ps1 before calling this
    function, and establish the PnP connection before invocation.

    .PARAMETER FileUrl
    Fully qualified URL of a .pub file in SharePoint.

    .PARAMETER Connection
    Optional PnP connection object. When omitted, the active PnP connection
    is used by the PnP cmdlets.

    .PARAMETER OverwritePdf
    Allows the uploaded PDF to replace an existing PDF with the same name.

    .EXAMPLE
    Connect-PnPOnline -Url 'https://contoso.sharepoint.com/sites/Team' -Interactive
    . .\Convert-PubFileToPDF.ps1
    . .\Convert-SharePointPubFileToPDF.ps1
    Convert-SharePointPubFileToPdf `
        -FileUrl 'https://contoso.sharepoint.com/sites/Team/Shared Documents/Newsletter.pub'

    .NOTES
    Updated with assistance from GitHub Copilot.
    Model used in this session: GPT-5.6-Luna
    Additional edits may have been made by a human author or other tooling.

    .OUTPUTS
    PSCustomObject
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [uri]
        $FileUrl,

        [Parameter()]
        [AllowNull()]
        [object]
        $Connection,

        [switch]
        $OverwritePdf
    )

    if (-not $FileUrl.IsAbsoluteUri -or $FileUrl.Scheme -notin @('http', 'https')) {
        throw "FileUrl must be a fully qualified SharePoint HTTP or HTTPS URL: $FileUrl"
    }

    if ([System.IO.Path]::GetExtension($FileUrl.AbsolutePath) -ne '.pub') {
        throw "SharePoint file URL must identify a .pub file: $FileUrl"
    }

    if ($null -eq (Get-Command -Name Get-PnPFile -ErrorAction SilentlyContinue)) {
        throw 'Get-PnPFile is unavailable. Install/import PnP.PowerShell before calling this function.'
    }

    if ($null -eq (Get-Command -Name Add-PnPFile -ErrorAction SilentlyContinue)) {
        throw 'Add-PnPFile is unavailable. Install/import PnP.PowerShell before calling this function.'
    }

    if ($null -eq (Get-Command -Name Convert-PubFileToPdf -ErrorAction SilentlyContinue)) {
        throw 'Convert-PubFileToPdf is unavailable. Dot-source Convert-PubFileToPDF.ps1 before calling this function.'
    }

    $decodedFileUrl = [System.Uri]::UnescapeDataString($FileUrl.AbsoluteUri)
    $decodedFilePath = [System.Uri]::UnescapeDataString($FileUrl.AbsolutePath)
    $fileName = [System.IO.Path]::GetFileName($decodedFilePath)
    $folderUrl = (Split-Path -Path $decodedFilePath -Parent).Replace('\', '/')
    $pdfFileName = [System.IO.Path]::ChangeExtension($fileName, '.pdf')
    $pdfFileUrl = [System.Uri]::new($FileUrl, $pdfFileName).AbsoluteUri
    $temporaryDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString('N'))
    $localPubPath = Join-Path -Path $temporaryDirectory -ChildPath $fileName
    $localPdfPath = [System.IO.Path]::ChangeExtension($localPubPath, '.pdf')

    $pnpParams = @{}
    if ($PSBoundParameters.ContainsKey('Connection') -and $null -ne $Connection) {
        $pnpParams.Connection = $Connection
    }

    try {
        $null = New-Item -Path $temporaryDirectory -ItemType Directory -Force

        $downloadParams = @{
            Url         = $decodedFilePath
            Path        = $temporaryDirectory
            FileName    = $fileName
            AsFile      = $true
            Force       = $true
            ErrorAction = 'Stop'
        }
        foreach ($key in $pnpParams.Keys) {
            $downloadParams[$key] = $pnpParams[$key]
        }

        Write-Verbose "Downloading '$decodedFileUrl' to '$localPubPath'."
        $null = Get-PnPFile @downloadParams

        if (-not (Test-Path -LiteralPath $localPubPath -PathType Leaf)) {
            throw "PnP download completed without creating '$localPubPath'."
        }

        $conversion = @(Convert-PubFileToPdf -Path $localPubPath -ErrorAction Stop)
        if ($conversion.Count -ne 1 -or -not (Test-Path -LiteralPath $localPdfPath -PathType Leaf)) {
            throw "Publisher conversion did not create the expected PDF '$localPdfPath'."
        }

        if ($PSCmdlet.ShouldProcess($folderUrl, "Upload '$pdfFileName'")) {
            $uploadParams = @{
                Path        = $localPdfPath
                Folder      = $folderUrl
                NewFileName = $pdfFileName
                ErrorAction = 'Stop'
            }
            if ($OverwritePdf) {
                $uploadParams.Overwrite = $true
            }
            foreach ($key in $pnpParams.Keys) {
                $uploadParams[$key] = $pnpParams[$key]
            }

            $uploadedFile = Add-PnPFile @uploadParams

            [pscustomobject]@{
                SourceFileUrl = $decodedFileUrl
                PdfFileUrl    = $pdfFileUrl
                LocalPdfPath  = $localPdfPath
                UploadedAt    = Get-Date
                UploadedFile  = $uploadedFile
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryDirectory -PathType Container) {
            Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
