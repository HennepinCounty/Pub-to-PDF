Set-StrictMode -Version Latest

function script:New-MockPublisherDocument {
    param (
        [bool]$CreatePdf,
        [bool]$ThrowOnExport
    )

    $document = [pscustomobject]@{}

    $closeScript = {
        $null = $true
    }

    $createPdfLocal = $CreatePdf
    $throwOnExportLocal = $ThrowOnExport
    $exportScript = {
        param(
            $format,
            $pdfPath
        )

        if ($throwOnExportLocal) {
            throw 'Mock export failure.'
        }

        if ($createPdfLocal) {
            Set-Content -Path $pdfPath -Value 'mock-pdf-content' -Encoding Ascii
        }
    }.GetNewClosure()

    $document | Add-Member -MemberType ScriptMethod -Name Close -Value $closeScript
    $document | Add-Member -MemberType ScriptMethod -Name ExportAsFixedFormat -Value $exportScript

    $document
}

function script:New-MockPublisherApplication {
    param (
        $Document
    )

    $app = [pscustomobject]@{}

    $documentLocal = $Document
    $openScript = {
        param($path)
        $documentLocal
    }.GetNewClosure()
    $quitScript = {
        $null = $true
    }

    $app | Add-Member -MemberType ScriptMethod -Name Open -Value $openScript
    $app | Add-Member -MemberType ScriptMethod -Name Quit -Value $quitScript

    $app
}

Describe 'Convert-PubFileToPdf' {
    BeforeAll {
        $projectRoot = Split-Path -Path $PSScriptRoot -Parent
        $scriptUnderTest = Join-Path -Path $projectRoot -ChildPath 'Convert-PubFileToPDF.ps1'
        . $scriptUnderTest

        if ($null -eq ('Microsoft.Office.Interop.Publisher.PbFixedFormatType' -as [type])) {
            $enumDefinition = @'
namespace Microsoft.Office.Interop.Publisher {
    public enum PbFixedFormatType {
        pbFixedFormatTypePDF = 2
    }
}
'@

            Microsoft.PowerShell.Utility\Add-Type -TypeDefinition $enumDefinition
        }
    }

    BeforeEach {
        Mock -CommandName Add-Type -MockWith {
            if (
                $PSBoundParameters.ContainsKey('AssemblyName') -and
                $AssemblyName -eq 'Microsoft.Office.Interop.Publisher' -and
                $null -eq ('Microsoft.Office.Interop.Publisher.PbFixedFormatType' -as [type])
            ) {
                $enumDefinition = @'
namespace Microsoft.Office.Interop.Publisher {
    public enum PbFixedFormatType {
        pbFixedFormatTypePDF = 2
    }
}
'@

                Microsoft.PowerShell.Utility\Add-Type -TypeDefinition $enumDefinition
            }
        }
    }

    It 'is available after dot-sourcing the script' {
        $command = Get-Command -Name Convert-PubFileToPdf -ErrorAction Stop
        $command.CommandType | Should -Be 'Function'
    }

    It 'converts a .pub file from pipeline path input and emits success object' {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath 'sample.pub'
        Set-Content -Path $sourcePath -Value 'dummy-pub-content' -Encoding Ascii

        $document = New-MockPublisherDocument -CreatePdf $true -ThrowOnExport $false
        $publisherApp = New-MockPublisherApplication -Document $document
        $script:publisherAppForMock = $publisherApp

        Mock -CommandName New-Object -ParameterFilter { $ComObject -eq 'Publisher.Application' } -MockWith {
            $script:publisherAppForMock
        }

        $result = $sourcePath | Convert-PubFileToPdf

        @($result).Count | Should -Be 1
        $result.SourcePath | Should -Be $sourcePath
        $result.PdfPath | Should -Be ([System.IO.Path]::ChangeExtension($sourcePath, '.pdf'))
        Test-Path -LiteralPath $result.PdfPath | Should -BeTrue
    }

    It 'skips existing PDF by default, emits no output, and writes skipped status to transaction log' {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath 'already.pub'
        $pdfPath = Join-Path -Path $TestDrive -ChildPath 'already.pdf'
        $logPath = Join-Path -Path $TestDrive -ChildPath 'transactions.csv'

        Set-Content -Path $sourcePath -Value 'dummy-pub-content' -Encoding Ascii
        Set-Content -Path $pdfPath -Value 'existing-pdf-content' -Encoding Ascii

        $document = New-MockPublisherDocument -CreatePdf $true -ThrowOnExport $false
        $publisherApp = New-MockPublisherApplication -Document $document
        $script:publisherAppForMock = $publisherApp

        Mock -CommandName New-Object -ParameterFilter { $ComObject -eq 'Publisher.Application' } -MockWith {
            $script:publisherAppForMock
        }

        $result = Convert-PubFileToPdf -Path $sourcePath -TransactionLogPath $logPath

        @($result).Count | Should -Be 0

        $rows = Import-Csv -Path $logPath
        @($rows).Count | Should -Be 1
        $rows[0].Status | Should -Be 'SkippedExistingPdf'
        $rows[0].SourcePath | Should -Be $sourcePath
    }

    It 'writes an error when existing PDF is found and -ErrorOnExistingPdf is used' {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath 'existing.pub'
        $pdfPath = Join-Path -Path $TestDrive -ChildPath 'existing.pdf'

        Set-Content -Path $sourcePath -Value 'dummy-pub-content' -Encoding Ascii
        Set-Content -Path $pdfPath -Value 'existing-pdf-content' -Encoding Ascii

        $document = New-MockPublisherDocument -CreatePdf $true -ThrowOnExport $false
        $publisherApp = New-MockPublisherApplication -Document $document
        $script:publisherAppForMock = $publisherApp

        Mock -CommandName New-Object -ParameterFilter { $ComObject -eq 'Publisher.Application' } -MockWith {
            $script:publisherAppForMock
        }

        Remove-Variable -Name existingPdfErrors -ErrorAction SilentlyContinue
        $conversionParams = @{
            Path               = $sourcePath
            ErrorOnExistingPdf = $true
            ErrorAction        = 'SilentlyContinue'
            ErrorVariable      = 'existingPdfErrors'
        }

        $result = Convert-PubFileToPdf @conversionParams

        @($result).Count | Should -Be 0
        @($existingPdfErrors).Count | Should -BeGreaterOrEqual 1
        @($existingPdfErrors | Where-Object { $_.Exception.Message -match 'PDF already exists' }).Count | Should -BeGreaterOrEqual 1
    }

    It 'logs failed conversions and emits no output when export does not produce a PDF' {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath 'fails.pub'
        $logPath = Join-Path -Path $TestDrive -ChildPath 'failed-transactions.csv'

        Set-Content -Path $sourcePath -Value 'dummy-pub-content' -Encoding Ascii

        $document = New-MockPublisherDocument -CreatePdf $false -ThrowOnExport $false
        $publisherApp = New-MockPublisherApplication -Document $document
        $script:publisherAppForMock = $publisherApp

        Mock -CommandName New-Object -ParameterFilter { $ComObject -eq 'Publisher.Application' } -MockWith {
            $script:publisherAppForMock
        }

        $result = Convert-PubFileToPdf -Path $sourcePath -TransactionLogPath $logPath

        @($result).Count | Should -Be 0

        $rows = Import-Csv -Path $logPath
        @($rows).Count | Should -Be 1
        $rows[0].Status | Should -Be 'Failed'
        $rows[0].SourcePath | Should -Be $sourcePath
    }

    It 'appends transaction log rows across multiple runs' {
        $firstSourcePath = Join-Path -Path $TestDrive -ChildPath 'first.pub'
        $secondSourcePath = Join-Path -Path $TestDrive -ChildPath 'second.pub'
        $logPath = Join-Path -Path $TestDrive -ChildPath 'append-transactions.csv'

        Set-Content -Path $firstSourcePath -Value 'dummy-pub-content' -Encoding Ascii
        Set-Content -Path $secondSourcePath -Value 'dummy-pub-content' -Encoding Ascii

        $document = New-MockPublisherDocument -CreatePdf $true -ThrowOnExport $false
        $publisherApp = New-MockPublisherApplication -Document $document
        $script:publisherAppForMock = $publisherApp

        Mock -CommandName New-Object -ParameterFilter { $ComObject -eq 'Publisher.Application' } -MockWith {
            $script:publisherAppForMock
        }

        Convert-PubFileToPdf -Path $firstSourcePath -TransactionLogPath $logPath | Out-Null
        Convert-PubFileToPdf -Path $secondSourcePath -TransactionLogPath $logPath | Out-Null

        $rows = Import-Csv -Path $logPath
        @($rows).Count | Should -Be 2
        $rows[0].Status | Should -Be 'Success'
        $rows[1].Status | Should -Be 'Success'
    }

    It 'fails in begin block when Publisher COM object cannot be created' {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath 'any.pub'
        Set-Content -Path $sourcePath -Value 'dummy-pub-content' -Encoding Ascii

        Mock -CommandName New-Object -ParameterFilter { $ComObject -eq 'Publisher.Application' } -MockWith {
            throw 'mocked com create failure'
        }

        $invoke = {
            Convert-PubFileToPdf -Path $sourcePath
        }

        $invoke | Should -Throw -ExpectedMessage 'Unable to create Publisher.Application COM object. Ensure Microsoft Publisher is installed and accessible.'
    }
}
