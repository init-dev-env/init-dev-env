# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
# =-=-= 7-Zip���_�E�����[�h�E�C���X�g�[��   =-=-=
# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

# ==== �����ݒ� ====
# �ȗ�����l�[���X�y�[�X�\�L
using namespace System.IO
using namespace System.Net
using namespace System.Text
using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace System.Security.Principal

# �G���[�������ɃX�N���v�g�������p����������~����
$ErrorActionPreference = "Stop"

# ==== �^�C�g���\�� ====
Start-Sleep -Milliseconds 300
Clear-Host
Write-Host -NoNewLine ([char]0x1B + "[0d")
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-=" -NoNewLine;Write-Host -ForegroundColor Yellow " 7-Zip���_�E�����[�h�E�C���X�g�[��   " -NoNewLine;Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=`r`n"

Start-Sleep -Milliseconds 1250

# ==== �ݒ�ǂݍ��� ====
Set-Location $PSScriptRoot
# ���s���̃X�N���v�g�t�@�C���̃t�@�C�����i�g���q�����j�Ɂu_��Config��.ini�v��t�����t�@�C�����̐ݒ�t�@�C����ǂݍ���
$ConfigFileName = Join-Path $PSScriptRoot ([Path]::GetFileNameWithoutExtension($PSCommandPath) + "_��Config��.ini")
$Config = [Collections.Generic.Dictionary[String, String]]::New()
[File]::ReadAllLines($ConfigFileName)`
| Where-Object { [String]::IsNullOrWhiteSpace($_) -eq $False }`
| Where-Object { -not ($_.StartsWith("#") -or $_.StartsWith("//") -or $_.StartsWith(";")) }`
| ForEach-Object {
    $Key, $Value = $_.Split("=", 2)
    $Config[$Key.Trim()] = $Value.Trim()
}

# ==== �Ǘ��Ҍ����`�F�b�N ====

# ���݂�Windows���[�U�[�̃A�J�E���g�����擾
$CurrentUserWindowsIdentity = [WindowsIdentity]::GetCurrent()
# ���݂�Windows���[�U�[�̌��������擾
$CurrentUserWindowsPrincipal = [WindowsPrincipal]$CurrentUserWindowsIdentity
# ���݂�Windows���[�U�[�̌������ɊǗ��Ҍ������܂܂�Ă��Ȃ�������
if ($CurrentUserWindowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $False) {
    # PowerShell�X�N���v�g�Ăяo�����̃o�b�`�t�@�C��������΃o�b�`�t�@�C���p�X���擾
    $CallerBatchFilePath = Join-Path $PSScriptRoot ([Path]::GetFileNameWithoutExtension($PSCommandPath) + ".cmd")
    if (Test-Path $CallerBatchFilePath -PathType Leaf) {
        # �Ǘ��Ҍ����Ńo�b�`�t�@�C������N��������
        if (Test-Path "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" -PathType Leaf) {
            # Windows Terminal���C���X�g�[������Ă���ꍇ�́AWindows Terminal�Ŏ��o�b�`�R�}���h�����s
            Start-Process -Verb RunAs "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" "--profile `"�R�}���h �v�����v�g`" Cmd.exe /C Call `"$CallerBatchFilePath`""
        } else {
            # Windows Terminal���C���X�g�[������Ă��Ȃ��ꍇ�́A�R�}���h�v�����v�g�Ŏ��o�b�`�R�}���h�����s
            Start-Process -Verb RunAs Cmd.exe "/C Call `"$CallerBatchFilePath`""
        }
        exit 99
    } else {
        # ���݂�PowerShell�X�N���v�g���Ǘ��Ҍ����ŋN��������
        if (Test-Path "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" -PathType Leaf) {
            # Windows Terminal���C���X�g�[������Ă���ꍇ�́AWindows Terminal�Ŏ��o�b�`�R�}���h�����s
            Start-Process -Verb RunAs "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" "--profile `"Windows PowerShell`" PowerShell.exe -File `"$PSCommandPath`" -ExecutionPolicy Bypass -NoLogo -NoProfile"
        } else {
            # Windows Terminal���C���X�g�[������Ă��Ȃ��ꍇ�́A�R�}���h�v�����v�g�Ŏ��o�b�`�R�}���h�����s
            Start-Process -Verb RunAs PowerShell.exe "-File `"$PSCommandPath`" -ExecutionPolicy Bypass -NoLogo -NoProfile"
        }
        exit 98
    }
}

# ==== ���C���������` ====
# ��`�������C���������Ăяo���ӏ��̓X�N���v�g�̖����ɋL�q
function Main {
    # ==== �o�[�W�����`�F�b�N ====
    # �o�[�W�����`�F�b�N
    $VersionCheckResult = Check-VersionIsJustOrNewer

    # ���łɑΏۃo�[�W������������V�����o�[�W�������C���X�g�[������Ă���Ȃ�I������
    if ($VersionCheckResult.IsJustOrNewer) {
        # ==== �Ώۃt�H���_�[���J�� ====
        # �C���X�g�[����̃t�H���_�[�p�X
        $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]

        Write-Log "���: �C���X�g�[���������I�����܂��B"
        Start-Sleep -Milliseconds 1000
        Write-Log "���: �C���X�g�[���ς݂̃t�H���_�[���J���܂��B[$InstallDestinationFolderPath]"
        Start-Sleep -Milliseconds 1000

        Explorer.exe $InstallDestinationFolderPath

        # ==== �C���X�g�[����̐ݒ� ====
        Do-AfterInstall

        End-ThisScript
    }

    # ==== �C���X�g�[�����ނ̃_�E�����[�h ====
    # �C���X�g�[�����ނ̑��݃`�F�b�N
    $InstallArtifactCheckResult = Check-InstallArtifactExists

    # �C���X�g�[�����ނ��Ȃ���΃_�E�����[�h����
    if ($InstallArtifactCheckResult.IsNotExist) {
        $DownloadResult = Download-InstallArtifact

        # �_�E�����[�h���s���̓G���[�I������
        if ($DownloadResult.IsFailed) {
            End-ThisScriptAsFailure
        }
    }

    # ==== �C���X�g�[�� ====
    $InstallResult = Install-Application

    # �C���X�g�[�����s���̓G���[�I������
    if ($InstallResult.IsFailed) {
        End-ThisScriptAsFailure
    }

    # ==== �C���X�g�[����̐ݒ� ====
    Do-AfterInstall

    # ==== �Ώۃt�H���_�[���J�� ====
    # �C���X�g�[����̃t�H���_�[�p�X
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    Explorer.exe $InstallDestinationFolderPath

    End-ThisScript
}

# ==== �C���X�g�[������Ă���o�[�W�������擾���� ====
function Get-InstalledVersion {
    # �C���X�g�[����̃t�H���_�[�p�X
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]

    # �`�F�b�N�Ώۂ̃t�@�C���̃p�X
    $CheckTargetFilePath = $Config["CheckTargetFilePath"]
    $CheckTargetFilePath = $CheckTargetFilePath.Replace("{InstallDestinationFolderPath}", $InstallDestinationFolderPath)

    # �Ώۂ̃t�@�C�������݂��Ȃ��ꍇ�̓o�[�W����0��Ԃ�
    if (-not (Test-Path $CheckTargetFilePath -PathType Leaf)) {
        Write-Log "���: �C���X�g�[���ς݂̃t�@�C���͂���܂���ł����B[$CheckTargetFilePath]"
        Start-Sleep -Milliseconds 800

        return New-Object Version "0.0"
    }

    # ���s�t�@�C���̃o�[�W���������擾����
    $VersionInfo = [Version](Get-ItemProperty $CheckTargetFilePath).VersionInfo.FileVersion

    Write-Log "���: �C���X�g�[���ς݃o�[�W����: $($VersionInfo.ToString()) [$CheckTargetFilePath]"
    Start-Sleep -Milliseconds 800

    return $VersionInfo
}

# ==== �o�[�W�����`�F�b�N ====
function Check-VersionIsJustOrNewer {
    param (
        [Switch]$NoOutput
    )
    # �C���X�g�[���Ώۃo�[�W����
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    Write-Log "���: �C���X�g�[���Ώۃo�[�W����: $TargetVersionString"
    Start-Sleep -Milliseconds 800

    # �C���X�g�[���ς݃o�[�W����
    $InstalledVersion = Get-InstalledVersion

    # CompareTo���\�b�h�ɂ���r
    # 0�܂���-1�Ȃ�΃C���X�g�[������Ă���o�[�W�����͑Ώۃo�[�W�����Ɠ����܂��͂��V����
    $CompareResult = $TargetVersion.CompareTo($InstalledVersion)

    # ���ʂɂ���ă��b�Z�[�W���o��
    if ($CompareResult -eq 0) {
        if (-not $NoOutput) {
            Write-Log "���: �Ώۃo�[�W�����Ɠ����o�[�W���������łɃC���X�g�[������Ă��܂��B($InstalledVersion)"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $True
        }
    } elseif ($CompareResult -lt 0) {
        if (-not $NoOutput) {
            Write-Log "���: �Ώۃo�[�W�������V�����o�[�W���������łɃC���X�g�[������Ă��܂��B(�Ώ�: $TargetVersion�A �C���X�g�[���ς�: $InstalledVersion)"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $True
        }
    } elseif ($InstalledVersion -eq (New-Object Version "0.0")) {
        if (-not $NoOutput) {
            Write-Log "���: �Ώۂ̓C���X�g�[������Ă��܂���B"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $False
        }
    } elseif ($CompareResult -gt 0) {
        if (-not $NoOutput) {
            Write-Log "���: �Ώۃo�[�W�������Â��o�[�W�������C���X�g�[������Ă��܂��B(�Ώ�: $TargetVersion�A �C���X�g�[���ς�: $InstalledVersion)"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $False
        }
    }
    Write-Log "�x��: ���B�s�\�R�[�h�ł��B(�Ώ�: $TargetVersion�A �C���X�g�[���ς�: $InstalledVersion)"
    Start-Sleep -Milliseconds 1500
    exit 1
}

# ==== �C���X�g�[�����ނ̑��݃`�F�b�N ====
function Check-InstallArtifactExists {
    # �C���X�g�[���Ώۃo�[�W����
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    # �C���X�g�[�����ނ̃t�@�C����
    $InstallArtifactFileName = $Config["InstallArtifactFileName"]
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionWithoutDot}", $TargetVersionString.Replace(".", ""))
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{Version}", $TargetVersionString)

    # �C���X�g�[�����ނ̃_�E�����[�h��̃��[�J���f�B�X�N��̃t�@�C���p�X
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # �t�@�C���T�[�o�[��Ɋi�[����ꍇ�̃t�@�C���p�X
    $FileServerDownloadDirPath = Join-Path $Config["FileServerDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # �C���X�g�[�����ނ����[�J���f�B�X�N�ɑ��݂��邩�`�F�b�N
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName
    if (Test-Path $InstallArtifactFilePath -PathType Leaf) {
        return [PSCustomObject]@{
            IsNotExist = $False
            InstallArtifactFilePath = $InstallArtifactFilePath
        }
    }

    # �C���X�g�[�����ނ��t�@�C���T�[�o�[�ɑ��݂��邩�`�F�b�N
    $InstallArtifactFilePath = Join-Path $FileServerDownloadDirPath $InstallArtifactFileName
    if (Test-Path $InstallArtifactFilePath -PathType Leaf) {
        return [PSCustomObject]@{
            IsNotExist = $False
            InstallArtifactFilePath = $InstallArtifactFilePath
        }
    }

    return [PSCustomObject]@{
        IsNotExist = $True
    }
}

# ==== �C���X�g�[�����ނ̃_�E�����[�h ====
function Download-InstallArtifact {
    # �C���X�g�[���Ώۃo�[�W����
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    # �C���X�g�[�����ނ̃t�@�C����
    $InstallArtifactFileName = $Config["InstallArtifactFileName"]
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionWithoutDot}", $TargetVersionString.Replace(".", ""))
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{Version}", $TargetVersionString)

    # �C���X�g�[�����ނ̃_�E�����[�h��̃��[�J���f�B�X�N��̃t�@�C���p�X
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # �C���X�g�[�����ނ̃_�E�����[�h��̃t�@�C���p�X
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName

    # �C���X�g�[�����ނ̃_�E�����[�h��t�H���_�[�����݂��Ȃ���΍쐬����
    if (-not (Test-Path $LocalDownloadDirPath -PathType Container)) {
        New-Item -ItemType Directory -Path $LocalDownloadDirPath | Out-Null
    }

    # �_�E�����[�h��URL
    $InstallArtifactDownloadURL = $Config["InstallArtifactDownloadURL"]
    $InstallArtifactDownloadURL = $InstallArtifactDownloadURL.Replace("{InstallArtifactFileName}", $InstallArtifactFileName)

    # ==== �_�E�����[�h���� ====
    # .NET Framework��System.Net.WebClient�N���X�̃C���X�^���X�𐶐�
    $WebClient = New-Object WebClient

    # ���[�v����
    $InLoopFlag = $True
    $LoopCount = 0
    while ($InLoopFlag) {
        # ���[�v�͊�{�I��1��Ŕ�����z��
        $LoopCount++
        $InLoopFlag = $False
        try {
            # �_�E�����[�h
            Write-Log "���: �C���X�g�[�����ނ̃_�E�����[�h���J�n���܂��B"
            Start-Sleep -Milliseconds 800
            Write-Log "���:  - �_�E�����[�h��URL: $InstallArtifactDownloadURL"
            Start-Sleep -Milliseconds 400
            Write-Log "���:  - �_�E�����[�h��p�X: $InstallArtifactFilePath"
            Start-Sleep -Milliseconds 800
            $WebClient.DownloadFile($InstallArtifactDownloadURL, $InstallArtifactFilePath)

            # �_�E�����[�h�����t�@�C�������݂��邩�`�F�b�N
            if (-not (Test-Path $InstallArtifactFilePath -PathType Leaf)) {
                throw "�x��: �C���X�g�[�����ނ̃_�E�����[�h�Ɏ��s���܂����B"
            }
        }
        catch {
            # BXO����
            if ($Env:ComputerName.StartsWith("B063")) {
                # �v���L�V�F�؏��̓��͂𑣂�
                [PSCredential]$Credential = Get-Credential -Message "�v���L�V�F�؂̃��[�U�[���ƃp�X���[�h����͂��Ă��������B"
                if (-not [String]::IsNullOrWhiteSpace($Credential.UserName)) {
                    [WebRequest]::DefaultWebProxy = [WebRequest]::GetSystemWebProxy()
                    [WebRequest]::DefaultWebProxy.Credentials = $Credential
                } else {
                    Write-Log "���: �v���L�V�F�؂̃��[�U�[���ƃp�X���[�h�̓��͂��L�����Z������܂����B"
                    Start-Sleep -Milliseconds 1250
                }
                # ���[�v���p��������
                if ($LoopCount -lt 2) {
                    $InLoopFlag = $True
                }
            }
        }
    }

    # �_�E�����[�h�����C���X�g�[�����ރt�@�C�������݂��邩�`�F�b�N
    if (Test-Path $InstallArtifactFilePath -PathType Leaf) {
        Write-Log "���: �C���X�g�[�����ނ̃_�E�����[�h���������܂����B"
        Start-Sleep -Milliseconds 1250
        return [PSCustomObject]@{
            IsFailed = $False
        }
    } else {
        Write-Log "�G���[: �C���X�g�[�����ނ̃_�E�����[�h�Ɏ��s���܂����B[$InstallArtifactDownloadURL]"
        Start-Sleep -Milliseconds 1250
        return [PSCustomObject]@{
            IsFailed = $True
        }
    }
}

# ==== �C���X�g�[������ ====
function Install-Application {
    # �C���X�g�[���Ώۃo�[�W����
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    # �C���X�g�[�����ނ̃_�E�����[�h��̃��[�J���f�B�X�N��̃t�@�C���p�X
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # �C���X�g�[�����ނ̃t�@�C����
    $InstallArtifactFileName = $Config["InstallArtifactFileName"]
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionWithoutDot}", $TargetVersionString.Replace(".", ""))
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{Version}", $TargetVersionString)

    # �C���X�g�[�����ނ̃_�E�����[�h��̃t�@�C���p�X
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName

    # �C���X�g�[����t�H���_�[�p�X
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $InstallDestinationFolderRealPath = $InstallDestinationFolderPath
    if ($Config["InstallDestinationFolderTargetIsParentFlag"] -eq "True") {
        $InstallDestinationFolderRealPath = Split-Path $InstallDestinationFolderPath -Parent
    }

    # ==== 7-Zip�̃C���X�g�[�� ====
    # ���s���������w�肵�T�C�����g�C���X�g�[��
    Write-Log "���: �C���X�g�[�����J�n���܂��B"
    Start-Sleep -Milliseconds 800
    Write-Log "���:  - �C���X�g�[�����ރt�@�C��: $InstallArtifactFilePath"
    Start-Sleep -Milliseconds 400
    Write-Log "���:  - �C���X�g�[����t�H���_�[: $InstallDestinationFolderPath"
    Start-Sleep -Milliseconds 800
    Write-Log "���: �T�C�����g�C���X�g�[����..."
    Start-Sleep -Milliseconds 100
    Start-Process $InstallArtifactFilePath "/S /D=`"$InstallDestinationFolderPath`"" -Wait -WindowStyle Hidden
    Start-Sleep -Milliseconds 400
    Write-Log "���: �C���X�g�[�����������܂����B"
    Start-Sleep -Milliseconds 1250

    # �C���X�g�[������Ă��邩�`�F�b�N
    $VersionCheckResult = Check-VersionIsJustOrNewer -NoOutput
    if (-not $VersionCheckResult.IsJustOrNewer) {
        Write-Log "�G���[: �C���X�g�[����̃A�v���P�[�V�����̃t�@�C������o�[�W������񂪎擾�ł��܂���ł����B�C���X�g�[���Ɏ��s�����\��������܂��B[$InstallDestinationFolderPath]"
        Start-Sleep -Milliseconds 1250
        End-ThisScriptAsFailure
    }
}

# ==== �C���X�g�[����̏����ݒ� ====
function Do-AfterInstall {
    # �C���X�g�[����t�H���_�[�p�X
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $InstallDestinationFolderRealPath = $InstallDestinationFolderPath
    if ($Config["InstallDestinationFolderTargetIsParentFlag"] -eq "True") {
        $InstallDestinationFolderRealPath = Split-Path $InstallDestinationFolderPath -Parent
    }

    # �A�v���P�[�V�������w�����W�X�g���L�[�̑��݃`�F�b�N
    if (-not (Test-Path "HKCU:\Software\7-Zip")) {
        Write-Log "�G���[: �C���X�g�[����̃A�v���P�[�V�����̃��W�X�g���L�[�����݂��܂���B�C���X�g�[���Ɏ��s�����\��������܂��B[$("HKCU:\Software\7-Zip".Replace("HKCU:", "HKEY_CURRENT_USER"))]"
        Start-Sleep -Milliseconds 1250
        End-ThisScriptAsFailure
    }

    # Options�T�u�L�[���Ȃ���΍쐬����
    if (-not (Test-Path "HKCU:\Software\7-Zip\Options")) {
        New-Item "HKCU:\Software\7-Zip\Options" -Force | Out-Null
        Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]"
        Start-Sleep -Milliseconds 700
    }

    # Options�T�u�L�[���擾����
    $OptionsKey = Get-Item "HKCU:\Software\7-Zip\Options"

    # ContextMenu�G���g�����Ȃ���΍쐬����
    if (-not $OptionsKey.Property.Contains("ContextMenu")) {
        New-ItemProperty -Path "HKCU:\Software\7-Zip\Options" -Name "ContextMenu" -PropertyType DWord -Value 4967 -Force | Out-Null
        Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[ContextMenu]��l[4967]�ō쐬���܂����B"
        Start-Sleep -Milliseconds 700
    } else {
        # ContextMenu�G���g���̒l���擾
        $ContextMenuValue = Get-ItemPropertyValue $OptionsKey.PSPath -Name "ContextMenu"
        # �擾�����l�Ń}�X�N�����l���擾
        $ContextMenuMaskedValue = $ContextMenuValue -band 4967
        if ($ContextMenuValue -ne $ContextMenuMaskedValue) {
            Write-Log "�x��: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[ContextMenu]�̒l��[$ContextMenuMaskedValue]�ł͂���܂���B(�l:[$ContextMenuValue])"
            Start-Sleep -Milliseconds 1250
            # ContextMenu�G���g���̒l���X�V����
            Set-ItemProperty -Path "HKCU:\Software\7-Zip\Options" -Name "ContextMenu" -Value $ContextMenuMaskedValue
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[ContextMenu]�̒l��[$ContextMenuMaskedValue]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 1250
        }
    }

    # MenuIcons�G���g�����Ȃ���΍쐬����
    if (-not $OptionsKey.Property.Contains("MenuIcons")) {
        New-ItemProperty -Path "HKCU:\Software\7-Zip\Options" -Name "MenuIcons" -PropertyType DWord -Value 1 -Force | Out-Null
        Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[MenuIcons]��l[1]�ō쐬���܂����B"
        Start-Sleep -Milliseconds 700
    } else {
        # MenuIcons�G���g���̒l���擾
        $MenuIconsValue = Get-ItemPropertyValue $OptionsKey.PSPath -Name "MenuIcons"
        if ($MenuIconsValue -ne 1) {
            Write-Log "�x��: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[MenuIcons]�̒l��[1]�ł͂���܂���B(�l:[$MenuIconsValue])"
            Start-Sleep -Milliseconds 1250
            # MenuIcons�G���g���̒l���X�V����
            Set-ItemProperty -Path "HKCU:\Software\7-Zip\Options" -Name "MenuIcons" -Value 1
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Options".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[MenuIcons]�̒l��[1]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 1250
        }
    }

    # Extraction�T�u�L�[���Ȃ���΍쐬����
    if (-not (Test-Path "HKCU:\Software\7-Zip\Extraction")) {
        New-Item "HKCU:\Software\7-Zip\Extraction" -Force | Out-Null
        Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\7-Zip\Extraction".Replace("HKCU:", "HKEY_CURRENT_USER"))]"
        Start-Sleep -Milliseconds 700
    }

    # Extraction�T�u�L�[���擾����
    $ExtractionKey = Get-Item "HKCU:\Software\7-Zip\Extraction"

    # ShowPassword�G���g�����Ȃ���΍쐬����
    if (-not $ExtractionKey.Property.Contains("ShowPassword")) {
        New-ItemProperty -Path "HKCU:\Software\7-Zip\Extraction" -Name "ShowPassword" -PropertyType DWord -Value 1 -Force | Out-Null
        Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Extraction".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[ShowPassword]��l[1]�ō쐬���܂����B"
        Start-Sleep -Milliseconds 700
    } else {
        # ShowPassword�G���g���̒l���擾
        $ShowPasswordValue = Get-ItemPropertyValue $ExtractionKey.PSPath -Name "ShowPassword"
        if ($ShowPasswordValue -ne 1) {
            Write-Log "�x��: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Extraction".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[ShowPassword]�̒l��[1]�ł͂���܂���B(�l:[$ShowPasswordValue])"
            Start-Sleep -Milliseconds 1250
            # ShowPassword�G���g���̒l���X�V����
            Set-ItemProperty -Path "HKCU:\Software\7-Zip\Extraction" -Name "ShowPassword" -Value 1
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\7-Zip\Extraction".Replace("HKCU:", "HKEY_CURRENT_USER"))]�̃G���g��[ShowPassword]�̒l��[1]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 1250
        }
    }

    # �g���q���Ƃ̊֘A�t��
    $ExtensionInfoList = @(
        @{
            Extension = "7z"
            IconIndex = 0
        }
        @{
            Extension = "zip"
            IconIndex = 1
        }
        @{
            Extension = "lzh"
            IconIndex = 6
        }
    )
    foreach ($ExtensionInfo in $ExtensionInfoList) {
        $Extension = $ExtensionInfo.Extension
        $IconIndex = $ExtensionInfo.IconIndex

        # �g���q���Ƃ̃L�[����Q�Ƃ����L�[���Ȃ���΍쐬����
        if (-not (Test-Path "HKCU:\Software\Classes\7-Zip.$Extension")) {
            New-Item "HKCU:\Software\Classes\7-Zip.$Extension" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\Classes\7-Zip.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]"
            Start-Sleep -Milliseconds 700
        }

        # �g���q���Ƃ̃L�[����Q�Ƃ����L�[���擾����
        $CalleeKey = Get-Item "HKCU:\Software\Classes\7-Zip.$Extension"

        # ����̃G���g��(�t�@�C���̎�ނ̖��O)���Ȃ����`�F�b�N���A�Ȃ���Βl��ݒ肷��
        if (-not $CalleeKey.Property.Contains("(default)")) {
            New-ItemProperty -Path "HKCU:\Software\Classes\7-Zip.$Extension" -Name "(default)" -PropertyType String -Value "$Extension �A�[�J�C�u�t�@�C��" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]��l[$Extension �A�[�J�C�u�t�@�C��]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 700
        } else {
            # ����̃G���g���̒l���擾
            $CalleeValue = Get-ItemPropertyValue $CalleeKey.PSPath -Name "(default)"
            if ($CalleeValue -ne "$Extension �A�[�J�C�u�t�@�C��") {
                Write-Log "�x��: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]�̒l��[$Extension �A�[�J�C�u�t�@�C��]�ł͂���܂���B(�l:[$CalleeValue])"
                Start-Sleep -Milliseconds 1250
                # ����̃G���g���̒l���X�V����
                Set-ItemProperty -Path "HKCU:\Software\Classes\7-Zip.$Extension" -Name "(default)" -Value "$Extension �A�[�J�C�u�t�@�C��"
                Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]�̒l��[$Extension �A�[�J�C�u�t�@�C��]�ōX�V���܂����B"
                Start-Sleep -Milliseconds 1250
            }
        }

        # �g���q���Ƃ̃A�C�R�����w�肷��L�[���Ȃ���΍쐬����
        if (-not (Test-Path "HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon")) {
            New-Item "HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]"
            Start-Sleep -Milliseconds 700
        }

        # �g���q���Ƃ̃A�C�R�����w�肷��L�[���擾����
        $DefaultIconKey = Get-Item "HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon"

        # ����̃G���g��(�A�C�R�����\�[�X�擾���t�@�C���p�X�A�A�C�R���C���f�b�N�X)���Ȃ����`�F�b�N���A�Ȃ���Βl��ݒ肷��
        if (-not $DefaultIconKey.Property.Contains("(default)")) {
            New-ItemProperty -Path "HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon" -Name "(default)" -PropertyType String -Value "$InstallDestinationFolderPath\7z.dll,$IconIndex" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]��l[$InstallDestinationFolderPath\7z.dll,$IconIndex]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 700
        } else {
            # ����̃G���g���̒l���擾
            $DefaultIconValue = Get-ItemPropertyValue $DefaultIconKey.PSPath -Name "(default)"
            if ($DefaultIconValue -ne "$InstallDestinationFolderPath\7z.dll,$IconIndex") {
                Write-Log "�x��: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]�̒l��[$InstallDestinationFolderPath\7z.dll,$IconIndex]�ł͂���܂���B(�l:[$DefaultIconValue])"
                Start-Sleep -Milliseconds 1250
                # ����̃G���g���̒l���X�V����
                Set-ItemProperty -Path "HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon" -Name "(default)" -Value "$InstallDestinationFolderPath\7z.dll,$IconIndex"
                Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension\DefaultIcon".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]�̒l��[$InstallDestinationFolderPath\7z.dll,$IconIndex]�ōX�V���܂����B"
                Start-Sleep -Milliseconds 1250
            }
        }

        # Shell�L�[���Ȃ���΍쐬����
        if (-not (Test-Path "HKCU:\Software\Classes\7-Zip.$Extension\Shell")) {
            New-Item "HKCU:\Software\Classes\7-Zip.$Extension\Shell" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\Classes\7-Zip.$Extension\Shell".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]"
            Start-Sleep -Milliseconds 700
        }

        # Shell\Open�L�[���Ȃ���΍쐬����
        if (-not (Test-Path "HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open")) {
            New-Item "HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]"
            Start-Sleep -Milliseconds 700
        }

        # Shell\Open\Command�L�[���Ȃ���΍쐬����
        if (-not (Test-Path "HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open\Command")) {
            New-Item "HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open\Command" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open\Command".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]"
            Start-Sleep -Milliseconds 700
        }

        # Shell\Open\Command�L�[���擾����
        $CommandKey = Get-Item "HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open\Command"

        # ����̃G���g��(���s�R�}���h)���Ȃ����`�F�b�N���A�Ȃ���Βl��ݒ肷��
        if (-not $CommandKey.Property.Contains("(default)")) {
            New-ItemProperty -Path "HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open\Command" -Name "(default)" -PropertyType ExpandString -Value "`"$InstallDestinationFolderPath\7zFM.exe`" `"%1`"" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\7-Zip.$Extension\Shell\Open\Command".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]��l[`"$InstallDestinationFolderPath\7zFM.exe`" `"%1`"]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 700
        }

        # �g���q���Ƃ̃L�[���Ȃ���΍쐬����
        if (-not (Test-Path "HKCU:\Software\Classes\.$Extension")) {
            New-Item "HKCU:\Software\Classes\.$Extension" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[���쐬���܂����B[$("HKCU:\Software\Classes\.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]"
            Start-Sleep -Milliseconds 700
        }

        # �g���q���Ƃ̃L�[���擾����
        $CallerKey = Get-Item "HKCU:\Software\Classes\.$Extension"

        # ����̃G���g��(�t�@�C���̎�ނ̖��O)���Ȃ����`�F�b�N���A�Ȃ���Βl��ݒ肷��
        if (-not $CallerKey.Property.Contains("(default)")) {
            New-ItemProperty -Path "HKCU:\Software\Classes\.$Extension" -Name "(default)" -PropertyType String -Value "7-Zip.$Extension" -Force | Out-Null
            Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]��l[7-Zip.$Extension]�ōX�V���܂����B"
            Start-Sleep -Milliseconds 700
        } else {
            # ����̃G���g���̒l���擾
            $CallerValue = Get-ItemPropertyValue $CallerKey.PSPath -Name "(default)"
            if ($CallerValue -ne "7-Zip.$Extension") {
                Write-Log "�x��: ���W�X�g���L�[[$("HKCU:\Software\Classes\.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]�̒l��[7-Zip.$Extension]�ł͂���܂���B(�l:[$CallerValue])"
                Start-Sleep -Milliseconds 1250
                # ����̃G���g���̒l���X�V����
                Set-ItemProperty -Path "HKCU:\Software\Classes\.$Extension" -Name "(default)" -Value "7-Zip.$Extension"
                Write-Log "���: ���W�X�g���L�[[$("HKCU:\Software\Classes\.$Extension".Replace("HKCU:\Software\Classes", "HKEY_CLASSES_ROOT"))]�̊���̃G���g��[(default)]�̒l��[7-Zip.$Extension]�ōX�V���܂����B"
                Start-Sleep -Milliseconds 1250
            }
        }
    }

    # �p�X��ݒ�
    Add-Path $InstallDestinationFolderPath "7z.exe"

    # ���W�X�g���̕ύX��S�E�B���h�E�ɒʒm����
    Notify-RegistryChanged
}

# ==== ���W�X�g���̒l�̌^(���)�̒萔�l���` ====
Set-Variable -Name REG_NONE                       -Value  0 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_SZ                         -Value  1 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_EXPAND_SZ                  -Value  2 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_BINARY                     -Value  3 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_DWORD                      -Value  4 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_DWORD_LITTLE_ENDIAN        -Value  4 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_DWORD_BIG_ENDIAN           -Value  5 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_LINK                       -Value  6 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_MULTI_SZ                   -Value  7 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_RESOURCE_LIST              -Value  8 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_FULL_RESOURCE_DESCRIPTOR   -Value  9 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_RESOURCE_REQUIREMENTS_LIST -Value 10 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_QWORD                      -Value 11 -Option Constant -Scope Script -ErrorAction SilentlyContinue
Set-Variable -Name REG_QWORD_LITTLE_ENDIAN        -Value 11 -Option Constant -Scope Script -ErrorAction SilentlyContinue

# ==== ���W�X�g���̌^���̕����񂩂�Ή�����l��Ԃ� ====
function Convert-RegistryTypeNameToIntValue {
    param (
        [String]$RegistryTypeName
    )
    # ���W�X�g���^���ƑΉ����鐮���l�̘A�z�z����`
    $registryTypeMap = @{
        "REG_NONE"                       = 0
        "REG_SZ"                         = 1
        "REG_EXPAND_SZ"                  = 2
        "REG_BINARY"                     = 3
        "REG_DWORD"                      = 4
        "REG_DWORD_LITTLE_ENDIAN"        = 4
        "REG_DWORD_BIG_ENDIAN"           = 5
        "REG_LINK"                       = 6
        "REG_MULTI_SZ"                   = 7
        "REG_RESOURCE_LIST"              = 8
        "REG_FULL_RESOURCE_DESCRIPTOR"   = 9
        "REG_RESOURCE_REQUIREMENTS_LIST" = 10
        "REG_QWORD"                      = 11
        "REG_QWORD_LITTLE_ENDIAN"        = 11
    }

    # �L�[�����݂���΂��̒l��Ԃ�
    if ($registryTypeMap.ContainsKey($RegistryTypeName)) {
        return [int]$registryTypeMap[$RegistryTypeName]
    }

    # �L�[�����݂��Ȃ��ꍇ�̃G���[���b�Z�[�W
    Write-Log "�G���[: ���m�̃��W�X�g���^���ł��B[$RegistryTypeName]"
    exit 1
}

# ==== ���W�X�g���̌^���̒l����Ή����镶�����Ԃ� ====
function Convert-RegistryTypeIntValueToName {
    param (
        [int]$RegistryTypeValue
    )
    # ���l�ƌ^���̃}�b�s���O���`
    $typeNameMap = @{
        0  = "REG_NONE"
        1  = "REG_SZ"
        2  = "REG_EXPAND_SZ"
        3  = "REG_BINARY"
        4  = "REG_DWORD"
        5  = "REG_DWORD_BIG_ENDIAN"
        6  = "REG_LINK"
        7  = "REG_MULTI_SZ"
        8  = "REG_RESOURCE_LIST"
        9  = "REG_FULL_RESOURCE_DESCRIPTOR"
        10 = "REG_RESOURCE_REQUIREMENTS_LIST"
        11 = "REG_QWORD"
    }

    # �L�[�����݂���Ό^����Ԃ�
    if ($typeNameMap.ContainsKey($RegistryTypeValue)) {
        return $typeNameMap[$RegistryTypeValue]
    }

    # �L�[�����݂��Ȃ��ꍇ�̃G���[���b�Z�[�W
    Write-Log "�G���[: ���m�̃��W�X�g���^�l�ł��B[$RegistryTypeValue]"
    exit 1
}

# ==== ���W�X�g���̒l�����łȂ��^����⏕�I�Ɏ������邽�߂̃N���X���` ====
class RegistryValue {
    [Object]$Value
    [int]$Type

    # �R���X�g���N�^�[
    RegistryValue([Object]$Value, [int]$Type) {
        $this.Value = $Value
        $this.Type = $Type
    }

    # �R���X�g���N�^�[(�l�̂�)
    RegistryValue([Object]$Value) {
        $this.Value = $Value
        $this.Type = $Script:REG_SZ
    }

    # �l��Ԃ�
    [String] ToString() {
        return $this.Value
    }
}

# ==== ���W�X�g���̃L�[�����݂��邩�Ԃ� ====
function Is-ExistRegistryKey {
    [OutputType([Boolean])]
    param (
        [String]$RegistryKey
    )
    $RegistryKey = $RegistryKey.TrimEnd("\")
    # ���W�X�g���n�C�u����PowerShell�����ɐ��K��
    $RegistryKey = $RegistryKey -replace "^HKEY_CURRENT_USER\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_LOCAL_MACHINE\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_CLASSES_ROOT\\", "HKCR:\"
    $RegistryKey = $RegistryKey -replace "^HKCU\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKLM\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKCR\\", "HKCR:\"
    # �L�[�̑��݃`�F�b�N
    return (Test-Path $RegistryKey)
}

# ==== ���W�X�g���̃L�[�ɃG���g�������݂��邩�Ԃ� ====
function Is-ExistRegistryEntry {
    [OutputType([Boolean])]
    param (
        [String]$RegistryKey,
        [String]$EntryName
    )
    # �L�[�̑��݃`�F�b�N
    if (-not (Test-Path $RegistryKey)) {
        return $False
    }
    $RegistryKey = $RegistryKey.TrimEnd("\")
    # ���W�X�g���n�C�u����PowerShell�����ɐ��K��
    $RegistryKey = $RegistryKey -replace "^HKEY_CURRENT_USER\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_LOCAL_MACHINE\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_CLASSES_ROOT\\", "HKCR:\"
    $RegistryKey = $RegistryKey -replace "^HKCU\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKLM\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKCR\\", "HKCR:\"
    # �G���g���̑��݃`�F�b�N
    $RegistryKeyObject = Get-Item $RegistryKey
    $EntryFoundCount = 0
    $RegistryKeyObject.Property | Where-Object { $_.ToLower() -eq $EntryName.ToLower() } | ForEach-Object { $EntryFoundCount++ }
    return ($EntryFoundCount -eq 1)
}

# ==== ���W�X�g���̒l��ǂݏo�� ====
function Get-RegistryValue {
    [OutputType([RegistryValue])]
    param (
        [String]$RegistryKey,
        [String]$EntryName
    )
    if (-not (Is-ExistRegistryKey $RegistryKey)) {
        return $Null
    }
    if (-not (Is-ExistRegistryEntry $RegistryKey $EntryName)) {
        return $Null
    }
    # ���W�X�g���n�C�u����Reg�R�}���h�����ɐ��K��
    $RegistryKey = $RegistryKey -replace "^HKCU:?\\", "HKEY_CURRENT_USER\"
    $RegistryKey = $RegistryKey -replace "^HKLM:?\\", "HKEY_LOCAL_MACHINE\"
    $RegistryKey = $RegistryKey -replace "^HKCR:?\\", "HKEY_CLASSES_ROOT\"
    try {
        if ($EntryName -eq "(default)") {
            $RegCmdResultLines = (Reg.exe Query $RegistryKey /ve          ).Split("`r`n")
        } else {
            $RegCmdResultLines = (Reg.exe Query $RegistryKey /v $EntryName).Split("`r`n")
        }
    }
    catch {
        return $Null
    }
    $RegCmdResultLines = $RegCmdResultLines | Where-Object { $_.Trim() -ne ""} | Where-Object { -not $_.StartsWith($RegistryKey) }
    if ($RegCmdResultLines -is [String]) {
        $RegCmdResultLine = [String]$RegCmdResultLines
    } else {
        if ($RegCmdResultLines.Count -ne 1) {
            return $Null
        }
        $RegCmdResultLine = $RegCmdResultLines[0]
    }
    $Dummy, $EntryName2, $RegistryType, $RegistryValue = $RegCmdResultLine -split "\s+", 4
    $ReturnValue = New-Object RegistryValue $RegistryValue, (Convert-RegistryTypeNameToIntValue $RegistryType)
    return $ReturnValue
}

# ==== ���W�X�g���̒l��ݒ肷�� ====
function Set-RegistryValue {
    param (
        [String]$RegistryKey,
        [String]$EntryName,
        [String]$EntryValue,
        [int]$EntryType = -1
    )
    # ���W�X�g���n�C�u����Reg�R�}���h�����ɐ��K��
    $RegistryKey = $RegistryKey -replace "^HKCU:?\\", "HKEY_CURRENT_USER\"
    $RegistryKey = $RegistryKey -replace "^HKLM:?\\", "HKEY_LOCAL_MACHINE\"
    $RegistryKey = $RegistryKey -replace "^HKCR:?\\", "HKEY_CLASSES_ROOT\"
    # ���w���$REG_NONE��0����ʂł���悤�����Ŕ���
    if ($EntryType -eq -1) {
        $EntryType = $Script:REG_SZ
    }
    $EntryTypeName = Convert-RegistryTypeIntValueToName $EntryType
    $ErrorOccurred = $False
    if ($EntryName -eq "(default)") {
        Reg.exe Add `"$RegistryKey`" /ve               /t `"$EntryTypeName`" /d `"$EntryValue`" /f | Out-Null
        if (-not $?) {
            $ErrorOccurred = $True
        }
    } else {
        Reg.exe Add `"$RegistryKey`" /v `"$EntryName`" /t `"$EntryTypeName`" /d `"$EntryValue`" /f | Out-Null
        if (-not $?) {
            $ErrorOccurred = $True
        }
    }
    if ($ErrorOccurred) {
        Write-Log "�G���[: ���W�X�g���̒l�̐ݒ�Ɏ��s���܂����B"
        End-ThisScriptAsFailure
    }
}

# ==== ���W�X�g���̒l��ݒ肷�� ====
function Set-RegistryValue2 {
    [OutputType([RegistryValue])]
    param (
        [String]$RegistryKey,
        [String]$EntryName,
        [RegistryValue]$RegistryValue
    )
    Set-RegistryValue $RegistryKey $EntryName $RegistryValue.Value $RegistryValue.Type
}

# ==== ���ϐ�Path��ݒ肷�� ====
function Add-Path {
    param (
        [String]$PathToAdd,
        [String]$ExampleExecFilePath = $Null
    )
    # �����̒l���擾
    [RegistryValue]$PathEnvValue = Get-RegistryValue "HKCU:\Environment" "Path"
    # �Z�~�R�����ŕ���
    if ($PathEnvValue -ne $Null) {
        $PathList = [List[String]]::New()
        foreach ($PathItem in $PathEnvValue.Value.Split(';')) {
            $PathItem = $PathItem.Trim()
            if (-not ([String]::IsNullOrWhiteSpace($PathItem))) {
                $PathList.Add($PathItem)
            }
        }
    } else {
        $PathList = [List[String]]::New()
    }
    # �����̒l�ɐݒ肵�悤�Ƃ����l�����邩�`�F�b�N
    $NeedToAdd = $True
    # �啶���������̍��فA%�݂̗͂L���̍��ق��z�����邽�߂ɏ��������A���ϐ��W�J���s��
    $PathToAddExtractedLower = ([Environment]::ExpandEnvironmentVariables($PathToAdd)).ToLower()
    foreach ($PathItem in $PathList) {
        $PathItemExtractedLower = ([Environment]::ExpandEnvironmentVariables($PathItem)).ToLower()
        # �����p�X�����łɓo�^����Ă���Βǉ����Ȃ�
        if ($PathItemExtractedLower -eq $PathToAddExtractedLower) {
            $NeedToAdd = $False
            break
        }
        # �ǉ����悤�Ƃ����t�H���_�[�p�X�Ƃ͕ʂ̃p�X�ł��������s�t�@�C�������݂���ꍇ�͒ǉ����Ȃ�
        if ($ExampleExecFilePath -ne $Null) {
            $TempPathItemExecFilePath = Join-Path $PathItem $ExampleExecFilePath
            $TempPathItemExecFilePathExtractedLower = ([Environment]::ExpandEnvironmentVariables($TempPathItemExecFilePath)).ToLower()
            if ($PathItemExtractedLower -eq $TempPathItemExecFilePathExtractedLower) {
                Write-Log "���: ���ϐ�Path(���[�U�[�P��)�ɂ��łɃp�X[$PathToAdd]�Ɠ������s�t�@�C�������p�X[$PathItem]���o�^����Ă������ߒǉ����܂���ł����B"
                Start-Sleep -Milliseconds 600
                $NeedToAdd = $False
                break
            }
        }
    }
    # �����̒l�ɐݒ肵�悤�Ƃ����l���Ȃ���Βǉ�
    if ($NeedToAdd) {
        $PathList.Add($PathToAdd)
    }
    # �V����Path��ݒ�
    $NewPathValue = [String]::Join(";", $PathList)
    if ($NewPathValue -eq $PathEnvValue.Value) {
        Write-Log "���: ���ϐ�Path(���[�U�[�P��)�ɂ��łɃp�X[$PathToAdd]���o�^����Ă������ߒǉ����܂���ł����B"
        Start-Sleep -Milliseconds 1250
        return
    }
    Set-RegistryValue "HKCU:\Environment" "Path" $NewPathValue $Script:REG_EXPAND_SZ
    Write-Log "���: ���ϐ�Path(���[�U�[�P��)�Ƀp�X[$PathToAdd]��ǉ����܂����B"
    Start-Sleep -Milliseconds 600
    Notify-RegistryChanged
}

# ==== ���W�X�g���̕ύX��S�E�B���h�E�ɒʒm���� ====
# (���ʂ̂قǂ͕s��)
function Notify-RegistryChanged {
    # ���W�X�g���̕ύX��S�E�B���h�E�ɒʒm����

    $SendMessageTimeoutDefinition = @"
    using System;
    using System.Runtime.InteropServices;

    public static class NativeMethods {
        private const long HWND_BROADCAST = 0xffff;
        private const int WM_SETTINGCHANGE = 0x1a;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            int Msg,
            int wParam,
            int lParam,
            int fuFlags,
            int uTimeout,
            out UIntPtr lpdwResult
        );

        public static UIntPtr NotifyChanged() {
            UIntPtr lpdwResult;
            SendMessageTimeout(
                (IntPtr)HWND_BROADCAST,
                WM_SETTINGCHANGE,
                0,
                1,
                2,
                5000,
                out lpdwResult
            );
            return lpdwResult;
        }
    }
"@
    $NativeMethods = Add-Type -TypeDefinition $SendMessageTimeoutDefinition -Language CSharp -PassThru
    $NativeMethods::NotifyChanged() | Out-Null
}

# ==== ���O���� ====
function Initialize-Log {
    param (
        [String]$LogId
    )
    $Script:LogId = $LogId
    $Script:LogFilePath = "$PSScriptRoot\..\Logs\$LogId`_$((Get-Date).ToString("yyyy-MM-dd_HH-mm-ss"))_Log.txt"
    $LogDirPath = [Path]::GetDirectoryName($Script:LogFilePath)
    if (-not (Test-Path $LogDirPath -PathType Container)) {
        New-Item -ItemType Directory $LogDirPath -Force | Out-Null
    }
    if (-not (Test-Path $Script:LogFilePath -PathType Leaf)) {
        New-Item -ItemType File $LogFilePath -Force | Out-Null
    }
}

# ==== ���O�o�� ====
function Write-Log {
    param (
        [String]$Message,
        [Switch]$NoFileOutput = $False, # �t���O���w�肳�ꂽ�ꍇ�̓��O�t�@�C���ɏo�͂��Ȃ�
        [Switch]$NoConsoleOutput = $False # �t���O���w�肳�ꂽ�ꍇ�̓R���\�[���ɏo�͂��Ȃ�
    )
    # ���O�o�͓���
    $LogDateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.ffffff")
    # ���O�t�@�C���ǋL
    if (-not $NoFileOutput) {
        Add-Content -Path $Script:LogFilePath -Value ("[$LogDateTime] $Message")
    }
    # �R���\�[���o��
    if (-not $NoConsoleOutput) {
        Write-Host -ForegroundColor DarkGray "[$LogDateTime] " -NoNewline
        Write-Host $Message
    }
}

# ==== �I������ ====
function End-ThisScript {
    # ���ϐ�NoWaitForInputEnterKey�ɒl"true"���Z�b�g����Ă���Ƃ��͒l���Z�b�g�����Ăь��̃o�b�`�t�@�C���ɂăL�[���͑҂����s���Ɖ��߂��A���̃X�N���v�g�t�@�C���ł̓L�[���͑҂����Ȃ�
    if ($Env:NoWaitForInputEnterKey -ne "true") {
        # �v�����v�g��\�����ăL�[���͂�҂�
        Write-Log ""
        Start-Sleep -Milliseconds 300
        Write-Host ("�I�����܂��B") -NoNewLine; Write-Host ("Enter") -NoNewLine -ForegroundColor Green; Write-Host ("�L�[�������Ă��������B")
        Write-Log "�I�����܂��B" -NoConsoleOutput
        Start-Sleep -Milliseconds 100
        Read-Host
        Start-Sleep -Milliseconds 700
    }
    exit 0
}

# ==== �G���[�ɂȂ�I������ ====
function End-ThisScriptAsFailure {
    # ���ϐ�NoWaitForInputEnterKey�ɒl"true"���Z�b�g����Ă���Ƃ��͒l���Z�b�g�����Ăь��̃o�b�`�t�@�C���ɂăL�[���͑҂����s���Ɖ��߂��A���̃X�N���v�g�t�@�C���ł̓L�[���͑҂����Ȃ�
    if ($Env:NoWaitForInputEnterKey -ne "true") {
        # �v�����v�g��\�����ăL�[���͂�҂�
        Write-Log ""
        Start-Sleep -Milliseconds 300
        Write-Host ("�I�����܂��B") -NoNewLine; Write-Host ("Enter") -NoNewLine -ForegroundColor Green; Write-Host ("�L�[�������Ă��������B")
        Write-Log "�I�����܂��B" -NoConsoleOutput
        Start-Sleep -Milliseconds 100
        Read-Host
        Start-Sleep -Milliseconds 700
    }
    exit 1
}

Initialize-Log -LogId "7-Zip"

Main
