@echo off
setlocal

set "PowerShellExe=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "ScriptPath=%~dp0Convert-OneDrivePubFilesToPDF.ps1"

if not exist "%ScriptPath%" exit /b 2

pushd "%~dp0" || exit /b 3
"%PowerShellExe%" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "%ScriptPath%" %*
set "ExitCode=%ERRORLEVEL%"
popd
exit /b %ExitCode%
