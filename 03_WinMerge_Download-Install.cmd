@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@rem =-=-= WinMerge‚ðƒ_ƒEƒ“ƒ[ƒhEƒCƒ“ƒXƒg[ƒ‹    =-=-=
@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

@Echo Off
Echo [36m[44m=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=[0m
Echo [36m[44m=-=-=[33m[40m WinMerge‚ðƒ_ƒEƒ“ƒ[ƒhEƒCƒ“ƒXƒg[ƒ‹    [36m[44m=-=-=[0m
Echo [36m[44m=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=[0m

Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1

rem ==== ŠÇ—ŽÒŒ ŒÀƒ`ƒFƒbƒN ====
Net Session 1> Nul 2>&1
If ErrorLevel 1 (
    rem ŠÇ—ŽÒŒ ŒÀ‚Å‚È‚¢ê‡‚ÍŠÇ—ŽÒŒ ŒÀ‚Å‚±‚Ìƒoƒbƒ`ƒtƒ@ƒCƒ‹Ž©g‚ðŽÀs‚µ’¼‚·
    Echo.
    Echo ŠÇ—ŽÒŒ ŒÀ‚ÅŽÀs‚µ’¼‚µ‚Ü‚·B
    Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1
    If Exist "%LocalAppData%\Microsoft\WindowsApps\WT.exe" (
        rem Windows Terminal‚ªƒCƒ“ƒXƒg[ƒ‹‚³‚ê‚Ä‚¢‚éê‡‚ÍAWindows Terminal‚ÅŽ©ƒoƒbƒ`ƒRƒ}ƒ“ƒh‚ðŽÀs
        Echo Windows Terminal‚ð‹N“®‚µ‚Ü‚·B
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        Echo.
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Start-Process -Verb RunAs \"%LocalAppData%\Microsoft\WindowsApps\WT.exe\" \"--profile `\"ƒRƒ}ƒ“ƒh ƒvƒƒ“ƒvƒg`\" Cmd.exe /C Call `\"%~dpnx0`\"\""
    ) Else (
        rem Windows Terminal‚ªƒCƒ“ƒXƒg[ƒ‹‚³‚ê‚Ä‚¢‚È‚¢ê‡‚ÍAƒRƒ}ƒ“ƒhƒvƒƒ“ƒvƒg‚ÅŽ©ƒoƒbƒ`ƒRƒ}ƒ“ƒh‚ðŽÀs
        Echo ƒRƒ}ƒ“ƒhƒvƒƒ“ƒvƒg‚ð‹N“®‚µ‚Ü‚·B
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        Echo.
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Start-Process -Verb RunAs Cmd.exe \"/C Call "%~dpnx0"\""
    )
    Echo.
    If Not ErrorLevel 1 (
        Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    ) Else (
        Echo [31mŠÇ—ŽÒŒ ŒÀ‚Å‚Ì‹N“®‚ªŽ¸”s‚µ‚Ü‚µ‚½B[0m
        Ping -n 1 -l 1 -w 1000 1.2.3.4 1> Nul 2>&1
    )
    Echo [90mƒEƒBƒ“ƒhƒE‚ÍŽ©“®“I‚É•Â‚¶‚ç‚ê‚Ü‚·B[0m
    Echo.
    Ping -n 1 -l 1 -w 3000 1.2.3.4 1> Nul 2>&1
    Exit /B 0
)

rem ==== ƒoƒbƒ`ƒtƒ@ƒCƒ‹Ši”[ƒtƒHƒ‹ƒ_[‚ðƒJƒŒƒ“ƒgƒtƒHƒ‹ƒ_[‚ÉÝ’è ====
Set CurrentDirPath=%~dp0
Set CurrentDirPath=%CurrentDirPath:~0,-1%
PushD "%CurrentDirPath%"

rem ==== ƒoƒbƒ`ƒtƒ@ƒCƒ‹–¼‚ÌŠg’£Žq‚ðps1‚É•Ï‚¦‚½ƒtƒ@ƒCƒ‹–¼‚ÌPowerShellƒXƒNƒŠƒvƒg‚ðŽÀs ====
Set NoWaitForInputEnterKey=true
PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -File "%~n0.ps1" %*
If Not ErrorLevel 1 (
    Echo [32m³íI—¹[0m
) Else (
    Echo [31mˆÙíI—¹[0m
)

rem ==== ƒL[“ü—Í‘Ò‚¿Œã‚ÉI—¹ ====
If Not "%NoPause%" == "true" (
    Echo.
    Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    Echo [90mI—¹‚µ‚Ü‚·B‰½‚©ƒL[‚ð‰Ÿ‚µ‚Ä‚­‚¾‚³‚¢B[0m
    Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    Pause 1> Nul 2>&1
    Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1
)
PopD
Exit /B 0
