@:: ==== 7zr.exe�E7za.exe�_�E�����[�h�X�N���v�g ====
@:: ---- �o�b�`�t�@�C������ ----
@CLS
@Echo Off
rem ���̃o�b�`�t�@�C�����X�N���v�g�t�@�C����ǂݍ��݁A11�s�ڂ܂ł��X�L�b�v���� PowerShell �X�N���v�g�Ƃ��Ď��s
Echo ���: �o�b�`�t�@�C���J�n���܂����B
Echo ���: PowerShell�����s���܂��B
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command Write-Host (Invoke-Expression ( ( @(\"\") * 11 ) + ( ( Get-Content \"%~dpnx0\" ) ^| Select-Object -Skip 11 ) -join \"`n\" ) )
Echo ���: �o�b�`�t�@�C���I�����܂��B
Ping -n 1 -l 1 -w 3000 1.2.3.4 1> Nul 2> Nul
Exit /B 0
# ---- PowerShell�X�N���v�g���� ----
Write-Host "���: PowerShell�X�N���v�g�J�n���܂����B"
# �����Ώۃo�[�W��������
$TargetVersion = New-Object Version 24, 9
Write-Host ("���: �Ώۃo�[�W���� = {0:D2}.{1:D2}" -f $($TargetVersion.Major), $($TargetVersion.Minor))

# �G���[�������Ɏ��s���f����ݒ�
$ErrorActionPreference = "Stop"

# 7za.exe�̑��݃`�F�b�N
$SevenZipAExist = Test-Path 7za.exe
# 7za.exe�̃o�[�W�������������Ă��邩
$SevenZipAVerSatisfiesTarget = $SevenZipAExist -and ([Version]([IO.FileInfo](Get-Item 7za.exe)).VersionInfo.FileVersion -ge $TargetVersion)
if ($SevenZipAExist -and $SevenZipAVerSatisfiesTarget) {
    Write-Host "���: 7za.exe�͑��݂��܂��B �o�[�W���� = $(([Version]([IO.FileInfo](Get-Item 7za.exe)).VersionInfo.FileVersion).ToString())"
} else {
    # 7zr.exe�̑��݃`�F�b�N
    $SevenZipRExist = Test-Path 7zr.exe
    # 7zr.exe�̃o�[�W�������������Ă��邩
    $SevenZipRVerSatisfiesTarget = $SevenZipRExist -and ([Version]([IO.FileInfo](Get-Item 7zr.exe)).VersionInfo.FileVersion -ge $TargetVersion)
    if ($SevenZipRExist -and $SevenZipRVerSatisfiesTarget) {
        Write-Host "���: 7zr.exe�͑��݂��܂��B"
    } else {
        # 7zr.exe�̃_�E�����[�h
        Write-Host "���: 7zr.exe���_�E�����[�h���܂��B"
        $SourceUrl = "https://www.7-zip.org/a/7zr.exe"
        Invoke-WebRequest -Uri $SourceUrl -OutFile .\7zr.exe
    }
    # 7za�̃A�[�J�C�u�̑��݃`�F�b�N
    # 7zXXYY-extra.7z��XX.YY�̃o�[�W�������Ώۃo�[�W�����ȏ�̂��̂�T��
    # �������݂����ꍇ�̓o�[�W�������ő�̂��̂�I������
    $SevenZipAArchiveFile = Get-ChildItem "7z*-extra.7z" | Where-Object {
        if ($_.Name -match "^7z(\d{2})(\d{2,})(.*)?-extra\.7z$") {
            $MajorVersion = [int]$Matches[1]
            $MinorVersion = [int]$Matches[2]
            $Version = New-Object Version $MajorVersion, $MinorVersion
            return $Version -ge $TargetVersion
        } else {
            return $False
        }
    } | Sort-Object -Property Name -Descending | Select-Object -First 1
    $SevenZipAArchiveExist = $SevenZipAArchiveFile -ne $Null
    if ($SevenZipAArchiveExist) {
        Write-Host "���: 7za�̃A�[�J�C�u�͑��݂��܂��B"
        $SevenZipAArchiveFileName = $SevenZipAArchiveFile.Name
    } else {
        # 7zXXYY-extra.7z�̃_�E�����[�h
        $SevenZipAArchiveFileName = "7z{0:D2}{1:D2}-extra.7z" -f $TargetVersion.Major, $TargetVersion.Minor
        Write-Host "���: $SevenZipAArchiveFileName`���_�E�����[�h���܂��B"
        $SourceUrl = "https://github.com/ip7z/7zip/releases/download/{0:D2}.{1:D2}/$SevenZipAArchiveFileName" -f $TargetVersion.Major, $TargetVersion.Minor
        Invoke-WebRequest -Uri $SourceUrl -OutFile $SevenZipAArchiveFileName
    }
    # �A�[�J�C�u��W�J
    Write-Host "���: 7za�̃A�[�J�C�u����7za.exe��W�J���܂��B"
    .\7zr.exe x $SevenZipAArchiveFileName 7za.exe -y -bso0 -bse0 -bsp0
    if (-not (Test-Path 7za.exe)) {
        Write-Host "�G���[: 7za.exe�̎擾�E�W�J�Ɏ��s���܂����B"
        exit 1
    }
    Write-Host "���: 7za.exe�̎擾�E�W�J���������܂����B �o�[�W���� = $(([Version]([IO.FileInfo](Get-Item 7za.exe)).VersionInfo.FileVersion).ToString())"
}
Write-Host "���: PowerShell�X�N���v�g�I�����܂��B"
Start-Sleep -Milliseconds 1500
