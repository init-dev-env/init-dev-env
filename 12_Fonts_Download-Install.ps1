# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
# =-=-= Fonts���_�E�����[�h�E�C���X�g�[��   =-=-=
# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

# ==== �����ݒ� ====
# �ȗ�����l�[���X�y�[�X�\�L
using namespace System.IO
using namespace System.Net
using namespace System.Text
using namespace System.Threading
using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace Microsoft.PowerShell.Core
using namespace System.Security.Principal

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSPossibleIncorrectComparisonWithNull', '')]
param()

# �G���[�������ɃX�N���v�g�������p����������~����
$ErrorActionPreference = "Stop"

# ==== �^�C�g���\�� ====
Start-Sleep -Milliseconds 300
Clear-Host
Write-Host -NoNewLine ([char]0x1B + "[0d")
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-=" -NoNewLine;Write-Host -ForegroundColor Yellow " Fonts���_�E�����[�h�E�C���X�g�[��   " -NoNewLine;Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-="
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

# ==== �A�Z���u����ǂݍ��� ====
[Reflection.Assembly]::LoadWithPartialName("PresentationCore") | Out-Null

# ==== ���W���[���ǂݍ��� ====
Set-Location $PSScriptRoot
if (-not (Test-Path .\Lib\Microsoft.PowerShell.ThreadJob.dll -PathType Leaf)) {
    Write-Host "�G���[: ���W���[���t�@�C����������܂���B[Lib\Microsoft.PowerShell.ThreadJob.dll]"
    exit 1
}
Import-Module .\Lib\Microsoft.PowerShell.ThreadJob.dll -Cmdlet "Start-ThreadJob" -Scope Local -Force

# ==== ���C���������` ====
# ��`�������C���������Ăяo���ӏ��̓X�N���v�g�̖����ɋL�q
function Main {
    # ==== �t�H���g���_�E�����[�h ====
    # �t�H���g�A�[�J�C�u�t�@�C���W�J��t�H���_�[
    $LocalDownloadRootFolderPath = $Config["LocalDownloadRootFolderPath"]

    # �t�H���g���_�E�����[�h
    Download-FontArchive $Config["FontDownloadUrlListFileName"] $LocalDownloadRootFolderPath

    # ==== �t�H���g�̃A�[�J�C�u�t�@�C����W�J���� ====
    Extract-FontArchive $LocalDownloadRootFolderPath

    # ==== �t�H���g���C���X�g�[������ ====
    Install-Font $LocalDownloadRootFolderPath

    End-ThisScript
}

# ==== �t�H���g�̃A�[�J�C�u�t�@�C�����_�E�����[�h���� ====
function Download-FontArchive {
    param (
        [String]$FontDownloadUrlListFileName,
        [String]$LocalDownloadRootFolderPath
    )
    # �t�H���_�[���Ȃ���΍쐬����
    if (-not (Test-Path $LocalDownloadRootFolderPath -PathType Container)) {
        New-Item -ItemType Directory -Path $LocalDownloadRootFolderPath -Force | Out-Null
    }

    # �t�H���g�_�E�����[�hURL�̃��X�g
    if (-not (Test-Path $FontDownloadUrlListFileName -PathType Leaf)) {
        Write-Log "�G���[: �t�H���g�_�E�����[�hURL�̃��X�g�t�@�C����������܂���B[$FontDownloadUrlListFileName]"
        End-ThisScriptAsFailure
    }
    $FontDownloadUrlList = [List[String]]::New()
    $FontDownloadListFilePath = Join-Path (Get-Location) $FontDownloadUrlListFileName
    $Lines = [File]::ReadAllLines($FontDownloadListFilePath)
    foreach ($Line in $Lines) {
        if (-not ($Line.StartsWith("http"))) {
            continue
        }
        $FontDownloadUrlList.Add($Line)
    }
    $FontDownloadUrlCount = $FontDownloadUrlList.Count
    if ($FontDownloadUrlCount -eq 0) {
        Write-Log "�G���[: �t�H���g�_�E�����[�hURL�̃��X�g�t�@�C����URL���L�ڂ���Ă��܂���B[$FontDownloadUrlListFileName]"
        End-ThisScriptAsFailure
    }
    Write-Log "���: �t�H���g�̃A�[�J�C�u�t�@�C�����_�E�����[�h���܂��B"

    # PowerShell 5.1 �ŉ\�ȕ��񏈗��Ń_�E�����[�h����
    $ScriptBlock = {
        param (
            [String]$FontDownloadUrl,
            [String]$LocalDownloadRootFolderPath,
            [Int32]$FontDownloadUrlIndex,
            [Int32]$FontDownloadUrlCount,
            [String[]]$WriteLogFunctionDef,
            [String]$LogFilePath,
            [Object]$LockObject
        )
        # �֐���`��ǂݍ���
        $Global:LockObject = $LockObject
        ${Function:Write-Log} = $WriteLogFunctionDef

        # �t�H���g�A�[�J�C�u�t�@�C����
        $FontArchiveFileName = [IO.Path]::GetFileName($FontDownloadUrl.Replace("/", "\"))
        $FontArchiveFileBaseName = [IO.Path]::GetFileNameWithoutExtension($FontArchiveFileName)

        # �_�E�����[�h��t�H���_�[�p�X
        $FontDownloadFolderPath = Join-Path $LocalDownloadRootFolderPath $FontArchiveFileBaseName

        # �t�H���_�[������΃X�L�b�v
        if (Test-Path $FontDownloadFolderPath -PathType Container) {
            Write-Log ("���: - ({0:D2}/{1:D2}) URL: $FontDownloadUrl [�_�E�����[�h�ς�]" -f $FontDownloadUrlIndex, $FontDownloadUrlCount)
            continue
        }
        # �_�E�����[�h��t�H���_�[���쐬����
        New-Item -ItemType Directory -Path $FontDownloadFolderPath -Force | Out-Null

        Write-Log ("���: - ({0:D2}/{1:D2}) URL: $FontDownloadUrl" -f $FontDownloadUrlIndex, $FontDownloadUrlCount)

        # �t�H���g���_�E�����[�h����
        try {
            Start-BitsTransfer -Source $FontDownloadUrl -Destination $FontDownloadFolderPath -Description "�t�H���g�_�E�����[�h:$FontArchiveFileBaseName" -DisplayName "�t�H���g�_�E�����[�h:$FontArchiveFileBaseName" -Priority High -TransferType Download
        } finally {
            Write-Progress -Completed -Activity "�t�H���g�_�E�����[�h:$FontArchiveFileBaseName" -Status "����" -PercentComplete 100
        }
    }
    $BackupProgressPreference = $ProgressPreference
    $ProgressPreference = "SilentlyContinue"
    try {
        $FontDownloadUrlIndex = 0
        $Jobs = [List[Job]]::New()
        foreach ($FontDownloadUrl in $FontDownloadUrlList) {
            $FontDownloadUrlIndex++
            $Job = Start-ThreadJob -ScriptBlock $ScriptBlock -ArgumentList $FontDownloadUrl, $LocalDownloadRootFolderPath, $FontDownloadUrlIndex, $FontDownloadUrlCount, ${Function:Write-Log}.ToString(), $Script:LogFilePath, $Global:LockObject -StreamingHost $Host
            $Jobs.Add($Job)
        }
        $Jobs | Wait-Job -Force | Receive-Job -Wait -Force | Remove-Job -Force | Out-Null
    } finally {
        $ProgressPreference = $BackupProgressPreference
    }
    Write-Log "���: �t�H���g�̃A�[�J�C�u�t�@�C���̃_�E�����[�h���������܂����B"
}

# ==== �t�H���g�̃A�[�J�C�u�t�@�C����W�J���� ====
function Extract-FontArchive {
    param (
        [String]$LocalDownloadRootFolderPath
    )
    # �t�H���g���Ƃ̃t�H���_�[���擾
    $FontArchiveFolderList = Get-ChildItem -Path $LocalDownloadRootFolderPath -Directory
    # �t�H���g�̃A�[�J�C�u�t�@�C�����擾
    $FontArchiveFileList = [List[FileInfo]]::New()
    foreach ($FontArchiveFolder in $FontArchiveFolderList) {
        Get-ChildItem -Path $FontArchiveFolder.FullName -File -Recurse -Include "*.zip", "*.7z" | ForEach-Object { $FontArchiveFileList.Add($_) }
    }
    $FontArchiveFileCount = $FontArchiveFileList.Count
    if ($FontArchiveFileCount -eq 0) {
        Write-Log "�G���[: �t�H���g�̃A�[�J�C�u�t�@�C����������܂���B"
        End-ThisScriptAsFailure
    }
    Write-Log "���: �t�H���g�̃A�[�J�C�u�t�@�C����W�J���܂��B"

    # PowerShell 5.1 �ŉ\�ȕ��񏈗��Ń_�E�����[�h����
    $ScriptBlock = {
        param (
            [System.IO.FileInfo]$FontArchiveFile,
            [Int32]$FontArchiveFileIndex,
            [Int32]$FontArchiveFileCount,
            [String[]]$WriteLogFunctionDef,
            [String]$LogFilePath,
            [Object]$LockObject
        )
        # �֐���`��ǂݍ���
        $Global:LockObject = $LockObject
        ${Function:Write-Log} = $WriteLogFunctionDef

        # �A�[�J�C�u�t�@�C����W�J����
        $FontArchiveFilePath = $FontArchiveFile.FullName
        $ExtractDestinationFolderPath = Join-Path ([System.IO.Path]::GetDirectoryName($FontArchiveFilePath)) ([System.IO.Path]::GetFileNameWithoutExtension($FontArchiveFilePath))
        # ���łɓW�J�ς݂̃t�H���_�[�����݂���ꍇ�̓X�L�b�v
        if (Test-Path $ExtractDestinationFolderPath -PathType Container) {
            Write-Log ("���: - ({0:D2}/{1:D2}) �t�H���g�A�[�J�C�u�t�@�C��: $($FontArchiveFile.Name) [�W�J�ς�]" -f $FontArchiveFileIndex, $FontArchiveFileCount)
            continue
        }
        Write-Log ("���: - ({0:D2}/{1:D2}) �t�H���g�A�[�J�C�u�t�@�C��: $($FontArchiveFile.Name)" -f $FontArchiveFileIndex, $FontArchiveFileCount)
        7z.exe x "$FontArchiveFilePath" -o"$ExtractDestinationFolderPath" -y -r -bso0
    }
    [Int32]$FontArchiveFileIndex = 0
    $Jobs = [List[Job]]::New()
    foreach ($FontArchiveFile in $FontArchiveFileList) {
        $FontArchiveFileIndex++
        $Job = Start-ThreadJob -ScriptBlock $ScriptBlock -ArgumentList $FontArchiveFile, $FontArchiveFileIndex, $FontArchiveFileCount, ${Function:Write-Log}.ToString(), $Script:LogFilePath, $Global:LockObject -StreamingHost $Host
        $Jobs.Add($Job)
    }
    $Jobs | Wait-Job -Force | Receive-Job -Wait -Force | Remove-Job -Force | Out-Null

    Write-Log "���: �t�H���g�̃A�[�J�C�u�t�@�C���̓W�J���������܂����B"
}

# ==== �t�H���g���C���X�g�[������ ====
function Install-Font {
    param (
        [String]$LocalDownloadRootFolderPath
    )
    # �t�H���g�L���b�V���T�[�r�X���ғ����ł���Ύ~�߂Ă���
    $NeedToRestoreFontCacheService = $False

    # �C���X�g�[���ς݃t�H���g�������W
    $InstalledFontInfoMap =  [System.Collections.Concurrent.ConcurrentDictionary[String, PSCustomObject]]::New()
    $InstalledFontNameList = [System.Collections.Concurrent.ConcurrentQueue[String]]::New()
    # ���W�X�g���̃L�[�̍��ڂ̑������J�E���g
    $InstalledFontCount = 0
    foreach ($RegistryHive in @("HKCU", "HKLM")) {
        $RegistryKeyName = "$RegistryHive`:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
        foreach ($Dummy in ([String[]](Get-Item $RegistryKeyName).Property)) {
            $InstalledFontCount++
        }
    }
    # ���߂ă��[�v���A�t�@�C�����ƃ��W�X�g���n�C�u�ƃt�H���g�t�@�C���p�X�̊֌W���擾
    [Int32]$InstalledFontIndex = 0
    foreach ($RegistryHive in @("HKCU", "HKLM")) {
        $RegistryKeyName = "$RegistryHive`:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
        # ���W�X�g���̊Y���L�[�̍��ڂ��ƂɃ��[�v
        # PowerShell 5.1 �ŉ\�ȕ��񏈗��Ŏ擾����
        $ScriptBlock = {
            param (
                [String]$FontTypeFaceName,
                [String]$RegistryHive,
                [String]$RegistryKeyName,
                [System.Collections.Concurrent.ConcurrentDictionary[String, PSCustomObject]]$InstalledFontInfoMap,
                [System.Collections.Concurrent.ConcurrentQueue[String]]$InstalledFontNameList,
                [Int32]$InstalledFontIndex,
                [Int32]$InstalledFontCount,
                [String[]]$WriteLogFunctionDef,
                [String]$LogFilePath,
                [Object]$LockObject
            )
            # �֐���`��ǂݍ���
            $Global:LockObject = $LockObject
            ${Function:Write-Log} = $WriteLogFunctionDef

            $ShellApplication = New-Object -ComObject Shell.Application

            # ���W�X�g���̊Y���L�[�̊Y���̍��ڂ̒l���t�H���g�t�@�C���̃p�X�Ƃ��Ď擾
            $FontFilePath = (Get-ItemProperty -LiteralPath $RegistryKeyName -Name $FontTypeFaceName).$FontTypeFaceName
            $FontFileName = [System.IO.Path]::GetFileName($FontFilePath)
            # �L�[�͑啶���E���������ꎋ
            $FontFileNameKey = $FontFileName.ToLower()
            # .fon�t�@�C���͓ǂݍ��ݎ��ɃG���[�ɂȂ�̂ŃX�L�b�v
            if ([System.IO.Path]::GetExtension($FontFileNameKey) -eq ".fon") {
                Start-Sleep -Milliseconds 20
                return
            }
            # ���łɒǉ��ς݂̃t�@�C���������݂�����x�����o���ăX�L�b�v
            if ($InstalledFontInfoMap.ContainsKey($FontFileNameKey)) {
                Write-Log ("�x��: - ({0:D2}/{1:D2}) �t�H���g�t�@�C�������d�����Ă��܂��B [$FontFileName]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 20
                return
            }
            # �t�H���g�t�@�C���̃t���p�X���擾
            $FontFileFullPath = $FontFilePath
            # �p�X���t�@�C�����݂̂̏ꍇ�A�f�t�H���g�̃t�H���_�[�p�X��t������
            if (-not ($FontFileFullPath.Contains("\"))) {
                # ���W�X�g���n�C�u���ƂɃf�t�H���g�̃t�H���g�t�H���_�[���擾
                if ($RegistryHive -eq "HKCU") {
                    $FontFileFullPath = Join-Path $Env:UserProfile "AppData\Local\Microsoft\Windows\Fonts\$FontFilePath"
                } elseif ($RegistryHive -eq "HKLM") {
                    $FontFileFullPath = Join-Path $Env:WinDir "Fonts\$FontFilePath"
                }
            }
            # �t�H���g�t�@�C�������݂��Ȃ��ꍇ�͌x�����o���ăX�L�b�v
            if (-not (Test-Path $FontFileFullPath -PathType Leaf)) {
                Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C����������܂���B[$FontFileFullPath]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 20
                return
            }
            # �t�H���g�t�@�C���̃��^�����擾
            $GlyphTypeface = $Null
            try {
                $GlyphTypeface = (New-Object -TypeName Windows.Media.GlyphTypeface -ArgumentList $FontFileFullPath)
            } catch {
                Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̓ǂݍ��݂Ɏ��s���܂����B[$FontFileFullPath]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 1000
                return
            }
            <#
            if ($GlyphTypeface.FamilyNames.ContainsKey("ja-JP")) {
                Write-Log "FamilyNames: ja-JP: [$($GlyphTypeface.FamilyNames["ja-JP"])]"
            }
            if ($GlyphTypeface.FamilyNames.ContainsKey("en-US")) {
                Write-Log "FamilyNames: en-US: [$($GlyphTypeface.FamilyNames["en-US"])]"
            }
            foreach ($CultureName in $GlyphTypeface.FamilyNames.Keys) {
                if (($CultureName -ne "ja-JP") -and ($CultureName -ne "en-US")) {
                    Write-Log "FamilyNames: $CultureName`: [$($GlyphTypeface.FamilyNames[$CultureName])]"
                }
            }
            if ($GlyphTypeface.Win32FamilyNames.ContainsKey("ja-JP")) {
                Write-Log "Win32FamilyNames: ja-JP: [$($GlyphTypeface.Win32FamilyNames["ja-JP"])]"
            }
            if ($GlyphTypeface.Win32FamilyNames.ContainsKey("en-US")) {
                Write-Log "Win32FamilyNames: en-US: [$($GlyphTypeface.Win32FamilyNames["en-US"])]"
            }
            foreach ($CultureName in $GlyphTypeface.Win32FamilyNames.Keys) {
                if (($CultureName -ne "ja-JP") -and ($CultureName -ne "en-US")) {
                    Write-Log "Win32FamilyNames: $CultureName`: [$($GlyphTypeface.Win32FamilyNames[$CultureName])]"
                }
            }
            if ($GlyphTypeface.FaceNames.ContainsKey("ja-JP")) {
                Write-Log "FaceNames: ja-JP: [$($GlyphTypeface.FaceNames["ja-JP"])]"
            }
            if ($GlyphTypeface.FaceNames.ContainsKey("en-US")) {
                Write-Log "FaceNames: en-US: [$($GlyphTypeface.FaceNames["en-US"])]"
            }
            foreach ($CultureName in $GlyphTypeface.FaceNames.Keys) {
                if (($CultureName -ne "ja-JP") -and ($CultureName -ne "en-US")) {
                    Write-Log "FaceNames: $CultureName`: [$($GlyphTypeface.FaceNames[$CultureName])]"
                }
            }
            if ($GlyphTypeface.Win32FaceNames.ContainsKey("ja-JP")) {
                Write-Log "Win32FaceNames: ja-JP: [$($GlyphTypeface.Win32FaceNames["ja-JP"])]"
            }
            if ($GlyphTypeface.Win32FaceNames.ContainsKey("en-US")) {
                Write-Log "Win32FaceNames: en-US: [$($GlyphTypeface.Win32FaceNames["en-US"])]"
            }
            foreach ($CultureName in $GlyphTypeface.Win32FaceNames.Keys) {
                if (($CultureName -ne "ja-JP") -and ($CultureName -ne "en-US")) {
                    Write-Log "Win32FaceNames: $CultureName`: [$($GlyphTypeface.Win32FaceNames[$CultureName])]"
                }
            }
            #>

            # �t�H���g�t�@�~���[�����擾
            $ShellFontFolder = $ShellApplication.Namespace([System.IO.Path]::GetDirectoryName($FontFileFullPath))
            $ShellFontFile = $ShellFontFolder.ParseName([System.IO.Path]::GetFileName($FontFileFullPath))
            $FontFamilyName = $ShellFontFolder.GetDetailsOf($ShellFontFile, 21)
            $FontFamilyName = $FontFamilyName.Replace(";", " &")
            if ([String]::IsNullOrEmpty($FontFamilyName)) {
                Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̃t�@�~���[����������܂���B[$FontFileFullPath]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 20
                return
            }
            # �t�H���g�t�@�C���̃��^���̃t�H���g���ƃ��W�X�g���̃L�[�̍��ږ��������u (TrueType)�v�̍��ق݂̂ł��邩�`�F�b�N
            if ($FontTypeFaceName -ne ($FontFamilyName + " (TrueType)")) {
                # Write-Log "�x��: �t�H���g�t�@�C�����ƃ��W�X�g���̃L�[�̍��ږ����قȂ�܂��B[$FontTypeFaceName] vs [$FontFamilyName] (TrueType) :[$FontFileName]"
                # Start-Sleep -Milliseconds 1000
            }

            # �t�H���g�t�@�C���̃��^��񂩂�o�[�W���������擾
            $Version = $Null
            if ($GlyphTypeface.VersionStrings.ContainsKey("ja-JP")) {
                $Version = $GlyphTypeface.VersionStrings["ja-JP"]
            } elseif ($GlyphTypeface.VersionStrings.ContainsKey("en-US")) {
                $Version = $GlyphTypeface.VersionStrings["en-US"]
            } else {
                foreach ($Key in ($GlyphTypeface.VersionStrings.Keys | Sort-Object)) {
                    # �L�[�����\�[�g���A�㏟���Ŏ擾
                    $Version = $GlyphTypeface.VersionStrings[$key]
                }
            }
            if ([String]::IsNullOrEmpty($Version)) {
                try {
                    $Version = $ShellFontFolder.GetDetailsOf($ShellFontFile, 166)
                } catch {
                    # do nothing
                }
            }
            if ([String]::IsNullOrEmpty($Version)) {
                Write-Log ("�G���[: - ({0:D2}/{1:D2}) �o�[�W������񂪌�����܂���B[$FontFamilyName]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 3000
                return
            }
            # �Z�~�R�����ȍ~�̕�������폜
            if ($Version.Contains(";")) {
                $Version = $Version.Substring(0, $Version.IndexOf(";")).Trim()
            }
            # Version��v�����ɕt���Ă���ꍇ�͍폜
            if ($Version -match "(?:Version(?: |\.)|v)(\d+(?:\.\d+)*)") {
                $Version = $Matches[1].Trim()
            }

            # �����i�[
            $InstalledFontInfo = [PSCustomObject]@{
                RegistryHive = $RegistryHive
                FontFilePath = $FontFilePath
                FontFileName = $FontFileName
                FontFamilyName = $FontFamilyName
                RegistryItemName = $FontTypeFaceName
                Version = $Version
            }
            $InstalledFontNameList.Enqueue($FontFileNameKey)
            $TryResult = $InstalledFontInfoMap.TryAdd($FontFileNameKey, $InstalledFontInfo)
            if (-not $TryResult) {
                Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C�������d�����Ă��܂��B [$FontFileName]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 20
                return
            }
            Write-Log ("- ({0:D2}/{1:D2}) FontFileName[$FontFileName], RegistryItemName[$FontTypeFaceName], RegistryHive[$RegistryHive], FontFamilyName[$FontFamilyName], Version[$Version], FontFilePath[$FontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
        }
        $Jobs = [List[Job]]::New()
        foreach ($FontTypeFaceName in ([String[]](Get-Item $RegistryKeyName).Property)) {
            $InstalledFontIndex++
            $Job = Start-ThreadJob -ScriptBlock $ScriptBlock -ArgumentList $FontTypeFaceName, $RegistryHive, $RegistryKeyName, $InstalledFontInfoMap, $InstalledFontNameList, $InstalledFontIndex, $InstalledFontCount, ${Function:Write-Log}.ToString(), $Script:LogFilePath, $Global:LockObject -StreamingHost $Host
            $Jobs.Add($Job)
        }
        $Jobs | Wait-Job -Force | Receive-Job -Wait -Force | Remove-Job -Force | Out-Null
    }

    # �t�H���g�t�@�C�����擾
    $FontFileList = Get-ChildItem -Path $LocalDownloadRootFolderPath -File -Recurse -Include "*.ttf", "*.ttc", "*.otf", "*.otc"
    $FontFileCount = $FontFileList.Count
    if ($FontFileCount -eq 0) {
        Write-Log "�G���[: �C���X�g�[���Ώۂ̃t�H���g�t�@�C����������܂���B"
        End-ThisScriptAsFailure
    }

    # PowerShell 5.1 �ŉ\�ȕ��񏈗��ŃC���X�g�[������
    $ScriptBlock = {
        param(
            [String]$FontFilePath,
            [String]$TempFolderPath,
            [System.Collections.Concurrent.ConcurrentDictionary[String, PSCustomObject]]$InstalledFontInfoMap,
            [System.Collections.Concurrent.ConcurrentQueue[String]]$InstalledFontNameList,
            [Int32]$InstalledFontIndex,
            [Int32]$InstalledFontCount,
            [String[]]$WriteLogFunctionDef,
            [String]$LogFilePath,
            [Object]$LockObject
        )
        # �����_���ŃE�F�C�g
        Start-Sleep -Milliseconds (5 * (Get-Random -Minimum 1 -Maximum 25))

        # �֐���`��ǂݍ���
        $Global:LockObject = $LockObject
        ${Function:Write-Log} = $WriteLogFunctionDef

        # COM�I�u�W�F�N�gShellApplication���擾
        $ShellApplication = New-Object -ComObject Shell.Application

        # �R�s�[�I�v�V�����t���O
        $CopyOptions = 0
        $CopyOptions += 0x00000004 # FOF_SILENT
        $CopyOptions += 0x00000010 # FOF_NOCONFIRMATION
        $CopyOptions += 0x00000200 # FOF_NOCONFIRMMKDIR
        $CopyOptions += 0x00000400 # FOF_NOERRORUI
        $CopyOptions += 0x00000800 # FOF_NOCOPYSECURITYATTRIBS
        $CopyOptions += 0x00400000 # FOFX_KEEPNEWERFILE

        $FontFileName = [System.IO.Path]::GetFileName($FontFilePath)
        # �L�[�͑啶���E���������ꎋ
        $FontFileNameKey = $FontFileName.ToLower()

        # �t�H���g�t�@�C���̃��^�����擾
        $GlyphTypeface = $Null
        try {
            $GlyphTypeface = (New-Object -TypeName Windows.Media.GlyphTypeface -ArgumentList $FontFilePath)
        } catch {
            Write-Log ("�G���[: - ({0:D2}/{1:D2}) �C���X�g�[���Ώۂ̃t�H���g�t�@�C���̓ǂݍ��݂Ɏ��s���܂����B[$FontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
            Start-Sleep -Milliseconds 1000
            return
        }

        # �t�H���g�t�@�~���[�����擾
        $ShellFontFolder = $ShellApplication.Namespace([System.IO.Path]::GetDirectoryName($FontFilePath))
        $ShellFontFile = $ShellFontFolder.ParseName([System.IO.Path]::GetFileName($FontFilePath))
        $FontFamilyName = $ShellFontFolder.GetDetailsOf($ShellFontFile, 21)
        $FontFamilyName = $FontFamilyName.Replace(";", " &")
        if ([String]::IsNullOrEmpty($FontFamilyName)) {
            Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̃t�@�~���[����������܂���B[$FontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
            Start-Sleep -Milliseconds 20
            return
        }

        # �t�H���g�t�@�C���̃��^��񂩂�o�[�W���������擾
        $Version = $Null
        if ($GlyphTypeface.VersionStrings.ContainsKey("ja-JP")) {
            $Version = $GlyphTypeface.VersionStrings["ja-JP"]
        } elseif ($GlyphTypeface.VersionStrings.ContainsKey("en-US")) {
            $Version = $GlyphTypeface.VersionStrings["en-US"]
        } else {
            foreach ($Key in ($GlyphTypeface.VersionStrings.Keys | Sort-Object)) {
                # �L�[�����\�[�g���A�㏟���Ŏ擾
                $Version = $GlyphTypeface.VersionStrings[$key]
            }
        }
        if ([String]::IsNullOrEmpty($Version)) {
            try {
                $Version = $ShellFontFolder.GetDetailsOf($ShellFontFile, 166)
            } catch {
                # do nothing
            }
        }
        if ([String]::IsNullOrEmpty($Version)) {
            Write-Log ("�G���[: - ({0:D2}/{1:D2}) �o�[�W������񂪌�����܂���B[$FontFamilyName]" -f $InstalledFontIndex, $InstalledFontCount)
            Start-Sleep -Milliseconds 3000
            return
        }
        # �Z�~�R�����ȍ~�̕�������폜
        if ($Version.Contains(";")) {
            $Version = $Version.Substring(0, $Version.IndexOf(";")).Trim()
        }
        # Version��v�����ɕt���Ă���ꍇ�͍폜
        if ($Version -match "(?:Version(?: |\.)|v)(\d+(?:\.\d+)*)") {
            $Version = $Matches[1].Trim()
        }

        # ���W�X�g���̃G���g���̍��ږ������肷��
        # ���łɃ��W�X�g���ɓo�^����Ă��邩�`�F�b�N
        if (-not ($InstalledFontInfoMap.ContainsKey($FontFileNameKey))) {
            # ���o�^�̏ꍇ
            $Kind = "Add"
            $RegistryHive = "HKCU"
            $RegistryKeyName = "$RegistryHive`:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
            $RegistryItemName = "$FontFamilyName (TrueType)"
            $InstallDestinationFontFolderPath = Join-Path $Env:UserProfile "AppData\Local\Microsoft\Windows\Fonts"
            $InstallDestinationFontFilePath = Join-Path $InstallDestinationFontFolderPath $FontFileName
        } else {
            # �o�^�ς݂̏ꍇ
            $Kind = "Update"
            $InstalledFontInfo = $InstalledFontInfoMap[$FontFileNameKey]
            $RegistryHive = $InstalledFontInfo.RegistryHive
            $RegistryKeyName = "$RegistryHive`:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
            $RegistryItemName = $InstalledFontInfo.RegistryItemName
            $InstallDestinationFontFilePath = $InstalledFontInfo.FontFilePath
            $InstallDestinationFontFolderPath = Split-Path $InstallDestinationFontFilePath
            # �o�[�W�����`�F�b�N
            $InstalledVersion = $InstalledFontInfo.Version
            # ������Ƃ��ăo�[�W�����̑召���r
            $CompareResult = [String]::Compare($InstalledVersion, $Version, $True)
            if ($CompareResult -eq 0) {
                Write-Log ("���: - ({0:D2}/{1:D2}) ����o�[�W���������łɃC���X�g�[������Ă��܂��B[$FontFamilyName]" -f $InstalledFontIndex, $InstalledFontCount)
                Start-Sleep -Milliseconds 900
                return
            } elseif ($CompareResult -lt 0) {
                Write-Log ("���: - ({0:D2}/{1:D2}) �C���X�g�[���ς݂̌Â��t�H���g��������܂����B�X�V���܂��B �C���X�g�[���ς�[$InstalledVersion]�A�C���X�g�[���Ώ�[$Version]" -f $InstalledFontIndex, $InstalledFontCount)
                # �t�H���g���A���C���X�g�[��
                # �t�H���g�t�@�C�����ړ�
                $InstalledFontFilePath = $InstalledFontInfo.FontFilePath
                $TempFolder = $ShellApplication.NameSpace($TempFolderPath)
                try {
                    $TempFolder.MoveHere($InstalledFontFilePath, $CopyOptions)
                } catch {
                    Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̈ړ������Ɏ��s���܂����B[$InstalledFontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
                    Start-Sleep -Milliseconds 900
                    return
                }
                # �t�@�C�����Ȃ��Ȃ������Ƃ��m�F����
                if (Test-Path $InstalledFontFilePath -PathType Leaf) {
                    Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̈ړ��Ɏ��s���܂����B[$InstalledFontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
                    Start-Sleep -Milliseconds 900
                    return
                }
                # �ړ������t�@�C�����폜����
                $MovedFontFilePath = Join-Path $TempFolderPath $InstalledFontInfo.FontFileName
                if (Test-Path $MovedFontFilePath -PathType Leaf) {
                    Remove-Item -Path $MovedFontFilePath -Force -ErrorAction SilentlyContinue
                } else {
                    Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̍폜�O�̈ړ��Ɏ��s���܂����B[$MovedFontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
                    Start-Sleep -Milliseconds 900
                    return
                }
                if (Test-Path $MovedFontFilePath -PathType Leaf) {
                    Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�t�@�C���̍폜�Ɏ��s���܂����B[$MovedFontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
                    Start-Sleep -Milliseconds 900
                    return
                }
            }
        }
        # �t�H���g�t�@�C�����C���X�g�[������
        Write-Log ("���: - ({0:D2}/{1:D2}) �t�H���g���C���X�g�[�����܂��B[$Kind] [$FontFamilyName] $RegistryHive`:$RegistryItemName" -f $InstalledFontIndex, $InstalledFontCount)
        $InstallDestinationFontFolder = $ShellApplication.NameSpace($InstallDestinationFontFolderPath)
        try {
            $InstallDestinationFontFolder.CopyHere($FontFilePath, $CopyOptions)
        } catch {
            Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�̃C���X�g�[�������Ɏ��s���܂����B �t�H���g�t�@�C��[$FontFilePath] �C���X�g�[����[$InstallDestinationFontFolderPath]" -f $InstalledFontIndex, $InstalledFontCount)
            Start-Sleep -Milliseconds 900
            return
        }
        if (-not (Test-Path $InstallDestinationFontFilePath -PathType Leaf)) {
            Write-Log ("�G���[: - ({0:D2}/{1:D2}) �t�H���g�̃C���X�g�[���Ɏ��s���܂����B �t�H���g�t�@�C��[$FontFilePath] �C���X�g�[����[$InstallDestinationFontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
            Start-Sleep -Milliseconds 900
            return
        }
        # ���W�X�g���ɓo�^����
        New-ItemProperty -LiteralPath $RegistryKeyName -Name $RegistryItemName -Value $InstallDestinationFontFilePath -PropertyType String -Force | Out-Null

        Write-Log ("���: - ({0:D2}/{1:D2}) �t�H���g���C���X�g�[�����܂����B �t�H���g�t�@�C��[$FontFilePath] �C���X�g�[����[$InstallDestinationFontFilePath]" -f $InstalledFontIndex, $InstalledFontCount)
    }
    # �ꎞ�t�H���_�[�̃p�X
    $TempFolderPath = Join-Path $LocalDownloadRootFolderPath "Temp"
    # �ꎞ�t�H���_�[�����݂���ꍇ�͈�x�폜����
    if (Test-Path $TempFolderPath -PathType Container) {
        Remove-Item -Recurse -Force $TempFolderPath
    }
    if (Test-Path $TempFolderPath -PathType Container) {
        Write-Log "�G���[: �ꎞ�t�H���_�[�̍폜�Ɏ��s���܂����B[$TempFolderPath]"
        Start-Sleep -Milliseconds 20
        End-ThisScriptAsFailure
    }
    # �V�K�Ɉꎞ�t�H���_�[���쐬����
    if (-not (Test-Path $TempFolderPath -PathType Container)) {
        New-Item -ItemType Directory -Path $TempFolderPath -Force | Out-Null
    }
    if (-not (Test-Path $TempFolderPath -PathType Container)) {
        Write-Log "�G���[: �ꎞ�t�H���_�[�̍쐬�Ɏ��s���܂����B[$TempFolderPath]"
        Start-Sleep -Milliseconds 20
        End-ThisScriptAsFailure
    }
    # �t�H���g�L���b�V���T�[�r�X�������Ă��邩�i�ĊJ���K�v���j���m�F����
    $NeedToRestoreFontCacheService = $False
    try {
        [System.ServiceProcess.ServiceController]$FontCacheServiceController = Get-Service -Name "FontCache"
        if (($FontCacheServiceController.Status -eq "Running") -and ($FontCacheServiceController.CanStop)) {
            Stop-Service $FontCacheServiceController
            $NeedToRestoreFontCacheService = $True
        }
    } catch {
        Write-Log "�x��: �t�H���g�L���b�V���T�[�r�X�̒�~�Ɏ��s���܂����B"
    }
    try {
        Write-Log "���: �t�H���g���C���X�g�[�����܂��B"
        [Int32]$FontFileIndex = 0
        $Jobs = [List[Job]]::New()
        foreach ($FontFile in $FontFileList) {
            $FontFileIndex++

            Write-Log ("���: - ({0:D2}/{1:D2}) �t�H���g�t�@�C��: $($FontFile.Name)" -f $FontFileIndex, $FontFileCount)
            $FontFileName = $FontFile.Name
            $FontFilePath = $FontFile.FullName
            $Job = Start-ThreadJob -ScriptBlock $ScriptBlock -ArgumentList $FontFilePath, $TempFolderPath, $InstalledFontInfoMap, $InstalledFontNameList, $FontFileIndex, $FontFileCount, ${Function:Write-Log}.ToString(), $Script:LogFilePath, $Global:LockObject -StreamingHost $Host
            $Jobs.Add($Job)
        }
        $Jobs | Wait-Job -Force | Receive-Job -Wait -Force | Remove-Job -Force | Out-Null
        Write-Log "���: �t�H���g�̃C���X�g�[�����������܂����B"
        Start-Sleep -Milliseconds 200
    } finally {
        if ($NeedToRestoreFontCacheService) {
            try {
                $FontCacheServiceController = Get-Service -Name "FontCache"
                Start-Service $FontCacheServiceController
            } catch {
                Write-Log "�x��: �t�H���g�L���b�V���T�[�r�X�̍ĊJ�Ɏ��s���܂����B"
            }
        }
    }
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

    # ==== 7-Zip�ň��k�t�@�C����W�J ====
    $SevenZipExecutablePath = Get-SevenZipExecutablePath
    # 7z�̎��s�t�@�C���̑��݃`�F�b�N
    if ($SevenZipExecutablePath -eq $Null) {
        Write-Log "�G���[: 7-Zip�̎��s�t�@�C����������܂���B"
        End-ThisScriptAsFailure
    }
    Write-Log "���: �C���X�g�[�����J�n���܂��B"
    Start-Sleep -Milliseconds 800
    Write-Log "���:  - �C���X�g�[�����ރt�@�C��: $InstallArtifactFilePath"
    Start-Sleep -Milliseconds 400
    Write-Log "���:  - �C���X�g�[����t�H���_�[: $InstallDestinationFolderPath"
    Start-Sleep -Milliseconds 800
    Write-Log "���: �C���X�g�[�����ޓW�J��..."
    Start-Sleep -Milliseconds 100
    $Exec = Start-Process $SevenZipExecutablePath "x `"$InstallArtifactFilePath`" -o`"$InstallDestinationFolderRealPath`" -y" -Wait -WindowStyle Hidden -PassThru
    if ($Exec.ExitCode -ne 0) {
        Write-Log "�G���[: 7-Zip�ł̓W�J�Ɏ��s���܂����B(�G���[�R�[�h: $($Exec.ExitCode))"
        Start-Sleep -Milliseconds 1250
        End-ThisScriptAsFailure
    }
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

# ==== 7-Zip�̎��s�t�@�C���p�X���擾���� ====
function Get-SevenZipExecutablePath {
    # Lib�t�H���_�[�ɑ��݂���ꍇ�͗D�悷��
    if (Test-Path (Join-Path $PSScriptRoot "Lib\7za.exe") -PathType Leaf) {
        return (Join-Path $PSScriptRoot "Lib\7za.exe")
    }
    # �p�X���ʂ��Ă���ӏ�����T��
    $SevenZipExecutablePath = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($SevenZipExecutablePath -ne $Null) {
        return $SevenZipExecutablePath.Path
    }
    # ����̏ꏊ��T��
    if (-not (Test-Path "C:\Program Files\7-Zip\7z.exe" -PathType Leaf)) {
        return "C:\Program Files\7-Zip\7z.exe"
    }
    return $Null
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
$Global:LockObject = New-Object Object
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
    # ���b�N�擾
    $HasLock = $False
    [System.Threading.Monitor]::Enter($Global:LockObject)
    $HasLock = $True
    try {
        # ���O�t�@�C���ǋL
        if (-not $NoFileOutput) {
            Add-Content -Path $Script:LogFilePath -Value ("[$LogDateTime] $Message")
        }
        # �R���\�[���o��
        if (-not $NoConsoleOutput) {
            Write-Host -ForegroundColor DarkGray "[$LogDateTime] " -NoNewline
            Write-Host $Message
        }
    } finally {
        # ���b�N���
        if ($HasLock) {
            [System.Threading.Monitor]::Exit($Global:LockObject)
        }
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

Initialize-Log -LogId "Fonts"

Main
