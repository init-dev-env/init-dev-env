@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@rem =-=-= WinMerge���_�E�����[�h�E�C���X�g�[��    =-=-=
@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

@Echo Off
Echo [36m[44m=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=[0m
Echo [36m[44m=-=-=[33m[40m WinMerge���_�E�����[�h�E�C���X�g�[��    [36m[44m=-=-=[0m
Echo [36m[44m=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=[0m

Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1

rem ==== �Ǘ��Ҍ����`�F�b�N ====
Net Session 1> Nul 2>&1
If ErrorLevel 1 (
    rem �Ǘ��Ҍ����łȂ��ꍇ�͊Ǘ��Ҍ����ł��̃o�b�`�t�@�C�����g�����s������
    Echo.
    Echo �Ǘ��Ҍ����Ŏ��s�������܂��B
    Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1
    If Exist "%LocalAppData%\Microsoft\WindowsApps\WT.exe" (
        rem Windows Terminal���C���X�g�[������Ă���ꍇ�́AWindows Terminal�Ŏ��o�b�`�R�}���h�����s
        Echo Windows Terminal���N�����܂��B
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        Echo.
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Start-Process -Verb RunAs \"%LocalAppData%\Microsoft\WindowsApps\WT.exe\" \"--profile `\"�R�}���h �v�����v�g`\" Cmd.exe /C Call `\"%~dpnx0`\"\""
    ) Else (
        rem Windows Terminal���C���X�g�[������Ă��Ȃ��ꍇ�́A�R�}���h�v�����v�g�Ŏ��o�b�`�R�}���h�����s
        Echo �R�}���h�v�����v�g���N�����܂��B
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        Echo.
        Ping -n 1 -l 1 -w 300 1.2.3.4 1> Nul 2>&1
        PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Start-Process -Verb RunAs Cmd.exe \"/C Call "%~dpnx0"\""
    )
    Echo.
    If Not ErrorLevel 1 (
        Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    ) Else (
        Echo [31m�Ǘ��Ҍ����ł̋N�������s���܂����B[0m
        Ping -n 1 -l 1 -w 1000 1.2.3.4 1> Nul 2>&1
    )
    Echo [90m�E�B���h�E�͎����I�ɕ����܂��B[0m
    Echo.
    Ping -n 1 -l 1 -w 3000 1.2.3.4 1> Nul 2>&1
    Exit /B 0
)

rem ==== �o�b�`�t�@�C���i�[�t�H���_�[���J�����g�t�H���_�[�ɐݒ� ====
Set CurrentDirPath=%~dp0
Set CurrentDirPath=%CurrentDirPath:~0,-1%
PushD "%CurrentDirPath%"

rem ==== �o�b�`�t�@�C�����̊g���q��ps1�ɕς����t�@�C������PowerShell�X�N���v�g�����s ====
Set NoWaitForInputEnterKey=true
PowerShell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -File "%~n0.ps1" %*
If Not ErrorLevel 1 (
    Echo [32m����I��[0m
) Else (
    Echo [31m�ُ�I��[0m
)

rem ==== �L�[���͑҂���ɏI�� ====
If Not "%NoPause%" == "true" (
    Echo.
    Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    Echo [90m�I�����܂��B�����L�[�������Ă��������B[0m
    Ping -n 1 -l 1 -w 100 1.2.3.4 1> Nul 2>&1
    Pause 1> Nul 2>&1
    Ping -n 1 -l 1 -w 1200 1.2.3.4 1> Nul 2>&1
)
PopD
Exit /B 0
