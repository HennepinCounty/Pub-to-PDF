Set-StrictMode -Version Latest

Describe 'Convert-SharePointPubFileToPdf' {
    BeforeAll {
        $projectRoot = Split-Path -Path $PSScriptRoot -Parent
        . (Join-Path -Path $projectRoot -ChildPath 'Convert-PubFileToPDF.ps1')
        . (Join-Path -Path $projectRoot -ChildPath 'Convert-SharePointPubFileToPDF.ps1')

        function global:Get-PnPFile { }
        function global:Add-PnPFile { }
    }

    AfterAll {
        Remove-Item -Path Function:\Get-PnPFile -ErrorAction SilentlyContinue
        Remove-Item -Path Function:\Add-PnPFile -ErrorAction SilentlyContinue
    }

    It 'downloads, converts, and uploads the PDF beside the source file' {
        $downloadedPath = $null
        $uploadedPath = $null

        Mock -CommandName Get-PnPFile -MockWith {
            param($Url, $Path, $FileName)
            $downloadedPath = Join-Path -Path $Path -ChildPath $FileName
            Set-Content -LiteralPath $downloadedPath -Value 'mock-pub' -Encoding Ascii
        }

        Mock -CommandName Convert-PubFileToPdf -MockWith {
            param($Path)
            $pdfPath = [System.IO.Path]::ChangeExtension($Path, '.pdf')
            Set-Content -LiteralPath $pdfPath -Value 'mock-pdf' -Encoding Ascii
            [pscustomobject]@{ SourcePath = $Path; PdfPath = $pdfPath }
        }

        Mock -CommandName Add-PnPFile -MockWith {
            param($Path, $Folder, $NewFileName)
            $uploadedPath = $Path
            [pscustomobject]@{ Name = $NewFileName }
        }

        $result = Convert-SharePointPubFileToPdf `
            -FileUrl 'https://contoso.sharepoint.com/sites/Team/Shared Documents/Newsletter.pub' `
            -Confirm:$false

        Assert-MockCalled -CommandName Get-PnPFile -Times 1
        Assert-MockCalled -CommandName Add-PnPFile -Times 1
        $result.PdfFileUrl | Should -Be 'https://contoso.sharepoint.com/sites/Team/Shared%20Documents/Newsletter.pdf'
        Test-Path -LiteralPath $result.LocalPdfPath | Should -BeFalse
    }

    It 'does not upload when conversion does not create a PDF' {
        Mock -CommandName Get-PnPFile -MockWith {
            param($Path, $FileName)
            Set-Content -LiteralPath (Join-Path -Path $Path -ChildPath $FileName) -Value 'mock-pub' -Encoding Ascii
        }
        Mock -CommandName Convert-PubFileToPdf -MockWith { }
        Mock -CommandName Add-PnPFile -MockWith { }

        {
            Convert-SharePointPubFileToPdf `
                -FileUrl 'https://contoso.sharepoint.com/sites/Team/Shared Documents/Newsletter.pub'
        } | Should -Throw -ExpectedMessage '*did not create the expected PDF*'

        Assert-MockCalled -CommandName Add-PnPFile -Times 0
    }
}
