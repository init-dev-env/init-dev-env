@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@rem =-=-= WinMergeをダウンロード・インストール    =-=-=
@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

@Echo Off
Echo [36m[44m=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=[0m
Echo [36m[44m=-=-=[33m[40m WinMergeをダウンロード・インストール    [36m[44m=-=-=[0m
Echo [36m[44m=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=[0m

Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1

rem ==== 管理者権限チェック ====
Net Session 1> Nul 2>&1
If ErrorLevel 1 (
    rem 管理者権限でない場合は管理者権限でこのバッチファイル自身を実行し直す
    Echo.
    Echo 管理者権限で実行し直します。
    Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1
    If Exist "%LocalAppData%\Microsoft\WindowsApps\WT.exe" (
        rem Windows Terminalがインストールされている場合は、Windows Terminalで自バッチコマンドを実行
        Echo Windows Terminalを起動します。
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        Echo.
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Start-Process -Verb RunAs \"%LocalAppData%\Microsoft\WindowsApps\WT.exe\" \"--profile `\"コマンド プロンプト`\" Cmd.exe /C Call `\"%~dpnx0`\"\""
    ) Else (
        rem Windows Terminalがインストールされていない場合は、コマンドプロンプトで自バッチコマンドを実行
        Echo コマンドプロンプトを起動します。
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        Echo.
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Start-Process -Verb RunAs Cmd.exe \"/C Call "%~dpnx0"\""
    )
    Echo.
    If Not ErrorLevel 1 (
        Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    ) Else (
        Echo [31m管理者権限での起動が失敗しました。[0m
        Ping -n 1 -l 1 -w 1000 1.2.3.4 1> Nul 2>&1
    )
    Echo [90mウィンドウは自動的に閉じられます。[0m
    Echo.
    Ping -n 1 -l 1 -w 3000 1.2.3.4 1> Nul 2>&1
    Exit /B 0
)

rem ==== バッチファイル格納フォルダーをカレントフォルダーに設定 ====
Set CurrentDirPath=%~dp0
Set CurrentDirPath=%CurrentDirPath:~0,-1%
PushD "%CurrentDirPath%"

rem ==== バッチファイル名の拡張子をps1に変えたファイル名のPowerShellスクリプトを実行 ====
Set NoWaitForInputEnterKey=true
PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -File "%~n0.ps1" %*
If Not ErrorLevel 1 (
    Echo [32m正常終了[0m
) Else (
    Echo [31m異常終了[0m
)

rem ==== キー入力待ち後に終了 ====
If Not "%NoPause%" == "true" (
    Echo.
    Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    Echo [90m終了します。何かキーを押してください。[0m
    Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    Pause 1> Nul 2>&1
    Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1
)
PopD
Exit /B 0
