@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@rem =-=-= �C���X�g�[���o�b�`�R�}���h�APowerShell�X�N���v�g���e���v���[�g����R�s�[   =-=-=
@rem =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
@CLS
@Echo Off

rem �����g���q�Ⴂ��PowerShell�X�N���v�g�t�@�C�������s
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dpn0.ps1" %*
Pause
GoTo :EOF
