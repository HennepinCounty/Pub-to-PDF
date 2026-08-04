Set-StrictMode -Version Latest

$script:integrationPubPath = $null
$script:hasPublisher = $false

Describe 'Convert-PubFileToPdf Integration' -Tag 'Integration' {
    BeforeAll {
        $projectRoot = Split-Path -Path $PSScriptRoot -Parent
        $scriptUnderTest = Join-Path -Path $projectRoot -ChildPath 'Convert-PubFileToPDF.ps1'
        . $scriptUnderTest

        $script:integrationPubPath = $env:PUBTOPDF_TEST_PUB_PATH
        $script:hasPublisher = $false

        try {
            $publisherApp = New-Object -ComObject Publisher.Application
            if ($null -ne $publisherApp) {
                $script:hasPublisher = $true
                $publisherApp.Quit()
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($publisherApp)
            }
        }
        catch {
            $script:hasPublisher = $false
        }
    }

    It 'converts a real .pub file when Publisher is installed and test file is provided' -Skip:(
        (-not $script:hasPublisher) -or
        [string]::IsNullOrWhiteSpace($script:integrationPubPath) -or
        (-not (Test-Path -LiteralPath $script:integrationPubPath -PathType Leaf))
    ) {
        $result = Convert-PubFileToPdf -Path $script:integrationPubPath

        @($result).Count | Should -Be 1
        $result[0].SourcePath | Should -Be (Resolve-Path -LiteralPath $script:integrationPubPath).Path
        Test-Path -LiteralPath $result[0].PdfPath | Should -BeTrue
    }
}
