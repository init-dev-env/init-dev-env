@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@rem =-=-= インストールバッチコマンド、PowerShellスクリプトをテンプレートからコピー   =-=-=
@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@CLS
@Echo Off

rem 同名拡張子違いのPowerShellスクリプトファイルを実行
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dpn0.ps1" %*
Pause
GoTo :EOF
