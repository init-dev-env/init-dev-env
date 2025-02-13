@:: ==== Extract Zip Archive File of This Repository to C:\TempInstall\InstallScript ====
@:: ---- Batch File Part ----
@CLS
@Echo Off
Set CurrentDirPath=%~dp0
Set CurrentDirPath=%CurrentDirPath:~0,-1%
PushD "%CurrentDirPath%"
Net Session 1> Nul 2> Nul
If ErrorLevel 1 (
    PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -Verb RunAs Cmd.exe -ArgumentList \"/C Call %CurrentDirPath%\%~nx0\""
    GoTo :EOF
)
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command Invoke-Expression ( ( @(\"\") * 15 ) + ( ( Get-Content \"%~dpnx0\" ) ^| Select-Object -Skip 15 ) -join [char]10 ).Replace(\"{CurrentDirPath}\", \"%CurrentDirPath%\")
GoTo :EOF
# ---- PowerShell Script Part ----
using namespace System.Collections.Generic
using namespace System.IO
$ErrorActionPreference = "Stop"
Push-Location "{CurrentDirPath}"
if (Test-Path C:\TempInstall\init-dev-env-main -PathType Container) {
    Remove-Item C:\TempInstall\init-dev-env-main -Recurse -Force
}
Expand-Archive -Path init-dev-env-main.zip -DestinationPath C:\TempInstall -Force
if (-not (Test-Path C:\TempInstall\InstallScript -PathType Container)) {
    Rename-Item C:\TempInstall\init-dev-env-main C:\TempInstall\InstallScript -Force
} else {
    Move-Item C:\TempInstall\init-dev-env-main\*.* C:\TempInstall\InstallScript -Force
    Remove-Item C:\TempInstall\init-dev-env-main -Recurse -Force
}
