@:: ==== Extract Zip Archive File of This Repository to C:\TempInstall\InstallScript ====
@:: ---- Batch File Part ----
@CLS
@Echo Off
Set CurrentDirPath=%~dp0
Set CurrentDirPath=%CurrentDirPath:~0,-1%
PushD "%CurrentDirPath%"
Net Session 1> Nul 2> Nul
If ErrorLevel 1 (
    PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -Verb RunAs -ArgumentList \"%CurrentDirPath%\%~nx0\""
    GoTo :EOF
)
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command Invoke-Expression ( ( @(\"\") * 15 ) + ( ( Get-Content \"%~dpnx0\" ) ^| Select-Object -Skip 15 ) -join [char]10 ).Replace(\"{CurrentDirPath}\", \"%CurrentDirPath%\")
GoTo :EOF
# ---- PowerShell Script Part ----
using namespace System.Collections.Generic
using namespace System.IO
$ErrorActionPreference = "Stop"
Push-Location "{CurrentDirPath}"
$TargetFileName = "init-dev-env.zip"
Expand-Archive -Path $TargetFileName -DestinationPath "C:\TempInstall"
Rename-Item C:\TempInstall\init-dev-env C:\TempInstall\InstallScript
