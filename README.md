# Pub-to-PDF

Convert Microsoft Publisher files to PDF using a reusable PowerShell function.

## Core function

The script `Convert-PubFileToPDF.ps1` now defines the advanced function `Convert-PubFileToPdf`.

- Designed for dot-sourcing or module inclusion.
- Accepts path input from the pipeline.
- Writes only successful conversion objects to the output pipeline.
- Can optionally log all transactions (success and non-success) to a CSV log file.

## Usage

### Single file

```powershell
. .\Convert-PubFileToPDF.ps1
Convert-PubFileToPdf -Path "C:\Docs\Input\Newsletter.pub"
```

### Pipeline input

```powershell
. .\Convert-PubFileToPDF.ps1
Get-ChildItem -Path "C:\Docs\Input" -Filter "*.pub" -File -Recurse |
	Convert-PubFileToPdf
```

### Log all transactions

```powershell
. .\Convert-PubFileToPDF.ps1
Get-ChildItem -Path "C:\Docs\Input" -Filter "*.pub" -File |
	Convert-PubFileToPdf -TransactionLogPath "C:\Logs\pub-conversion.csv"
```

The transaction log is appended when the file already exists.

### Treat existing PDFs as errors

```powershell
. .\Convert-PubFileToPDF.ps1
Convert-PubFileToPdf -Path "C:\Docs\Input\Newsletter.pub" -ErrorOnExistingPdf
```

## Testing

This project includes Pester tests in the `tests` directory.

### Test files

- `tests/Convert-PubFileToPdf.Tests.ps1`: unit tests that use mocks and do not require Microsoft Publisher.
- `tests/Convert-PubFileToPdf.Integration.Tests.ps1`: optional real conversion test that requires Microsoft Publisher and a valid `.pub` file.

### Run all tests

```powershell
Invoke-Pester -Path .\tests
```

By default, integration tests are included in discovery but automatically skipped unless prerequisites are provided.

### Run only unit tests

```powershell
Invoke-Pester -Path .\tests\Convert-PubFileToPdf.Tests.ps1
```

### Run integration tests

1. Ensure Microsoft Publisher is installed.
2. Set an environment variable with a real `.pub` file path.

```powershell
$env:PUBTOPDF_TEST_PUB_PATH = "C:\Path\To\Sample.pub"
Invoke-Pester -Path .\tests -Tag Integration
```

### Expected behavior

- Unit tests validate function behavior without COM dependencies.
- Integration test performs a real conversion only when all prerequisites are met.
