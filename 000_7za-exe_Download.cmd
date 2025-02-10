@:: ==== 7zr.exe・7za.exeダウンロードスクリプト ====
@:: ---- バッチファイル部分 ----
@CLS
@Echo Off
rem このバッチファイル兼スクリプトファイルを読み込み、11行目までをスキップして PowerShell スクリプトとして実行
Echo 情報: バッチファイル開始しました。
Echo 情報: PowerShellを実行します。
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command Write-Host (Invoke-Expression ( ( @(\"\") * 11 ) + ( ( Get-Content \"%~dpnx0\" ) ^| Select-Object -Skip 11 ) -join \"`n\" ) )
Echo 情報: バッチファイル終了します。
Ping -n 1 -l 1 -w 3000 1.2.3.4 1> Nul 2> Nul
Exit /B 0
# ---- PowerShellスクリプト部分 ----
Write-Host "情報: PowerShellスクリプト開始しました。"
# ★★対象バージョン★★
$TargetVersion = New-Object Version 24, 9
Write-Host ("情報: 対象バージョン = {0:D2}.{1:D2}" -f $($TargetVersion.Major), $($TargetVersion.Minor))

# エラー発生時に実行中断する設定
$ErrorActionPreference = "Stop"

# 7za.exeの存在チェック
$SevenZipAExist = Test-Path 7za.exe
# 7za.exeのバージョンが満足しているか
$SevenZipAVerSatisfiesTarget = $SevenZipAExist -and ([Version]([IO.FileInfo](Get-Item 7za.exe)).VersionInfo.FileVersion -ge $TargetVersion)
if ($SevenZipAExist -and $SevenZipAVerSatisfiesTarget) {
    Write-Host "情報: 7za.exeは存在します。 バージョン = $(([Version]([IO.FileInfo](Get-Item 7za.exe)).VersionInfo.FileVersion).ToString())"
} else {
    # 7zr.exeの存在チェック
    $SevenZipRExist = Test-Path 7zr.exe
    # 7zr.exeのバージョンが満足しているか
    $SevenZipRVerSatisfiesTarget = $SevenZipRExist -and ([Version]([IO.FileInfo](Get-Item 7zr.exe)).VersionInfo.FileVersion -ge $TargetVersion)
    if ($SevenZipRExist -and $SevenZipRVerSatisfiesTarget) {
        Write-Host "情報: 7zr.exeは存在します。"
    } else {
        # 7zr.exeのダウンロード
        Write-Host "情報: 7zr.exeをダウンロードします。"
        $SourceUrl = "https://www.7-zip.org/a/7zr.exe"
        Invoke-WebRequest -Uri $SourceUrl -OutFile .\7zr.exe
    }
    # 7zaのアーカイブの存在チェック
    # 7zXXYY-extra.7zのXX.YYのバージョンが対象バージョン以上のものを探す
    # 複数存在した場合はバージョンが最大のものを選択する
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
        Write-Host "情報: 7zaのアーカイブは存在します。"
        $SevenZipAArchiveFileName = $SevenZipAArchiveFile.Name
    } else {
        # 7zXXYY-extra.7zのダウンロード
        $SevenZipAArchiveFileName = "7z{0:D2}{1:D2}-extra.7z" -f $TargetVersion.Major, $TargetVersion.Minor
        Write-Host "情報: $SevenZipAArchiveFileName`をダウンロードします。"
        $SourceUrl = "https://github.com/ip7z/7zip/releases/download/{0:D2}.{1:D2}/$SevenZipAArchiveFileName" -f $TargetVersion.Major, $TargetVersion.Minor
        Invoke-WebRequest -Uri $SourceUrl -OutFile $SevenZipAArchiveFileName
    }
    # アーカイブを展開
    Write-Host "情報: 7zaのアーカイブから7za.exeを展開します。"
    .\7zr.exe x $SevenZipAArchiveFileName 7za.exe -y -bso0 -bse0 -bsp0
    if (-not (Test-Path 7za.exe)) {
        Write-Host "エラー: 7za.exeの取得・展開に失敗しました。"
        exit 1
    }
    Write-Host "情報: 7za.exeの取得・展開が完了しました。 バージョン = $(([Version]([IO.FileInfo](Get-Item 7za.exe)).VersionInfo.FileVersion).ToString())"
}
Write-Host "情報: PowerShellスクリプト終了します。"
Start-Sleep -Milliseconds 1500
