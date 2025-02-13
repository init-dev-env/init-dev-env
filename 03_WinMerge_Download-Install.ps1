# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
# =-=-= WinMergeをダウンロード・インストール    =-=-=
# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

# ==== 初期設定 ====
# 省略するネームスペース表記
using namespace System.IO
using namespace System.Net
using namespace System.Text
using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace System.Security.Principal

# エラー発生時にスクリプトを処理継続させず停止する
$ErrorActionPreference = "Stop"

# ==== タイトル表示 ====
Start-Sleep -Milliseconds 300
Clear-Host
Write-Host -NoNewLine ([char]0x1B + "[0d")
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-=" -NoNewLine;Write-Host -ForegroundColor Yellow " WinMergeをダウンロード・インストール    " -NoNewLine;Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=`r`n"

Start-Sleep -Milliseconds 1250

# ==== 設定読み込み ====
Set-Location $PSScriptRoot
# 実行中のスクリプトファイルのファイル名（拡張子除く）に「_≪Config≫.ini」を付けたファイル名の設定ファイルを読み込む
$ConfigFileName = Join-Path $PSScriptRoot ([Path]::GetFileNameWithoutExtension($PSCommandPath) + "_≪Config≫.ini")
$Config = [Collections.Generic.Dictionary[String, String]]::New()
[File]::ReadAllLines($ConfigFileName)`
| Where-Object { [String]::IsNullOrWhiteSpace($_) -eq $False }`
| Where-Object { -not ($_.StartsWith("#") -or $_.StartsWith("//") -or $_.StartsWith(";")) }`
| ForEach-Object {
    $Key, $Value = $_.Split("=", 2)
    $Config[$Key.Trim()] = $Value.Trim()
}

# ==== 管理者権限チェック ====

# 現在のWindowsユーザーのアカウント情報を取得
$CurrentUserWindowsIdentity = [WindowsIdentity]::GetCurrent()
# 現在のWindowsユーザーの権限情報を取得
$CurrentUserWindowsPrincipal = [WindowsPrincipal]$CurrentUserWindowsIdentity
# 現在のWindowsユーザーの権限情報に管理者権限が含まれていないか判定
if ($CurrentUserWindowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $False) {
    # PowerShellスクリプト呼び出し元のバッチファイルがあればバッチファイルパスを取得
    $CallerBatchFilePath = Join-Path $PSScriptRoot ([Path]::GetFileNameWithoutExtension($PSCommandPath) + ".cmd")
    if (Test-Path $CallerBatchFilePath -PathType Leaf) {
        # 管理者権限でバッチファイルから起動し直す
        if (Test-Path "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" -PathType Leaf) {
            # Windows Terminalがインストールされている場合は、Windows Terminalで自バッチコマンドを実行
            Start-Process -Verb RunAs "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" "--profile `"コマンド プロンプト`" Cmd.exe /C Call `"$CallerBatchFilePath`""
        } else {
            # Windows Terminalがインストールされていない場合は、コマンドプロンプトで自バッチコマンドを実行
            Start-Process -Verb RunAs Cmd.exe "/C Call `"$CallerBatchFilePath`""
        }
        exit 99
    } else {
        # 現在のPowerShellスクリプトを管理者権限で起動し直す
        if (Test-Path "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" -PathType Leaf) {
            # Windows Terminalがインストールされている場合は、Windows Terminalで自バッチコマンドを実行
            Start-Process -Verb RunAs "$Env:LocalAppData\Microsoft\WindowsApps\WT.exe" "--profile `"Windows PowerShell`" PowerShell.exe -File `"$PSCommandPath`" -ExecutionPolicy Bypass -NoLogo -NoProfile"
        } else {
            # Windows Terminalがインストールされていない場合は、コマンドプロンプトで自バッチコマンドを実行
            Start-Process -Verb RunAs PowerShell.exe "-File `"$PSCommandPath`" -ExecutionPolicy Bypass -NoLogo -NoProfile"
        }
        exit 98
    }
}

# ==== メイン処理を定義 ====
# 定義したメイン処理を呼び出す箇所はスクリプトの末尾に記述
function Main {
    # ==== バージョンチェック ====
    # バージョンチェック
    $VersionCheckResult = Check-VersionIsJustOrNewer

    # すでに対象バージョンかそれより新しいバージョンがインストールされているなら終了する
    if ($VersionCheckResult.IsJustOrNewer) {
        # ==== 対象フォルダーを開く ====
        # インストール先のフォルダーパス
        $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]

        Write-Log "情報: インストール処理を終了します。"
        Start-Sleep -Milliseconds 1000
        Write-Log "情報: インストール済みのフォルダーを開きます。[$InstallDestinationFolderPath]"
        Start-Sleep -Milliseconds 1000

        Explorer.exe $InstallDestinationFolderPath

        # ==== インストール後の設定 ====
        Do-AfterInstall

        End-ThisScript
    }

    # ==== インストール資材のダウンロード ====
    # インストール資材の存在チェック
    $InstallArtifactCheckResult = Check-InstallArtifactExists

    # インストール資材がなければダウンロードする
    if ($InstallArtifactCheckResult.IsNotExist) {
        $DownloadResult = Download-InstallArtifact

        # ダウンロード失敗時はエラー終了する
        if ($DownloadResult.IsFailed) {
            End-ThisScriptAsFailure
        }
    }

    # ==== インストール ====
    $InstallResult = Install-Application

    # インストール失敗時はエラー終了する
    if ($InstallResult.IsFailed) {
        End-ThisScriptAsFailure
    }

    # ==== インストール後の設定 ====
    Do-AfterInstall

    # ==== 対象フォルダーを開く ====
    # インストール先のフォルダーパス
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    Explorer.exe $InstallDestinationFolderPath

    End-ThisScript
}

# ==== インストールされているバージョンを取得する ====
function Get-InstalledVersion {
    # インストール先のフォルダーパス
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]

    # チェック対象のファイルのパス
    $CheckTargetFilePath = $Config["CheckTargetFilePath"]
    $CheckTargetFilePath = $CheckTargetFilePath.Replace("{InstallDestinationFolderPath}", $InstallDestinationFolderPath)

    # 対象のファイルが存在しない場合はバージョン0を返す
    if (-not (Test-Path $CheckTargetFilePath -PathType Leaf)) {
        Write-Log "情報: インストール済みのファイルはありませんでした。[$CheckTargetFilePath]"
        Start-Sleep -Milliseconds 800

        return New-Object Version "0.0"
    }

    # 実行ファイルのバージョン情報を取得する
    $VersionInfo = [Version](Get-ItemProperty $CheckTargetFilePath).VersionInfo.FileVersion

    Write-Log "情報: インストール済みバージョン: $($VersionInfo.ToString()) [$CheckTargetFilePath]"
    Start-Sleep -Milliseconds 800

    return $VersionInfo
}

# ==== バージョンチェック ====
function Check-VersionIsJustOrNewer {
    param (
        [Switch]$NoOutput
    )
    # インストール対象バージョン
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    Write-Log "情報: インストール対象バージョン: $TargetVersionString"
    Start-Sleep -Milliseconds 800

    # インストール済みバージョン
    $InstalledVersion = Get-InstalledVersion

    # CompareToメソッドによる比較
    # 0または-1ならばインストールされているバージョンは対象バージョンと同じまたはより新しい
    $CompareResult = $TargetVersion.CompareTo($InstalledVersion)

    # 結果によってメッセージを出力
    if ($CompareResult -eq 0) {
        if (-not $NoOutput) {
            Write-Log "情報: 対象バージョンと同じバージョンがすでにインストールされています。($InstalledVersion)"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $True
        }
    } elseif ($CompareResult -lt 0) {
        if (-not $NoOutput) {
            Write-Log "情報: 対象バージョンより新しいバージョンがすでにインストールされています。(対象: $TargetVersion、 インストール済み: $InstalledVersion)"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $True
        }
    } elseif ($InstalledVersion -eq (New-Object Version "0.0")) {
        if (-not $NoOutput) {
            Write-Log "情報: 対象はインストールされていません。"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $False
        }
    } elseif ($CompareResult -gt 0) {
        if (-not $NoOutput) {
            Write-Log "情報: 対象バージョンより古いバージョンがインストールされています。(対象: $TargetVersion、 インストール済み: $InstalledVersion)"
            Start-Sleep -Milliseconds 900
        }
        return [PSCustomObject]@{
            IsJustOrNewer = $False
        }
    }
    Write-Log "警告: 到達不能コードです。(対象: $TargetVersion、 インストール済み: $InstalledVersion)"
    Start-Sleep -Milliseconds 1500
    exit 1
}

# ==== インストール資材の存在チェック ====
function Check-InstallArtifactExists {
    # インストール対象バージョン
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    # インストール資材のファイル名
    $InstallArtifactFileName = $Config["InstallArtifactFileName"]
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionWithoutDot}", $TargetVersionString.Replace(".", ""))
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{Version}", $TargetVersionString)
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionForWinMerge}", ($TargetVersionString -replace "(\d+)\.(\d+)\.(\d+)\.(\d+)", "`$1.`$2.`$3-jp-`$4"))

    # インストール資材のダウンロード後のローカルディスク上のファイルパス
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # ファイルサーバー上に格納する場合のファイルパス
    $FileServerDownloadDirPath = Join-Path $Config["FileServerDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # インストール資材がローカルディスクに存在するかチェック
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName
    if (Test-Path $InstallArtifactFilePath -PathType Leaf) {
        return [PSCustomObject]@{
            IsNotExist = $False
            InstallArtifactFilePath = $InstallArtifactFilePath
        }
    }

    # インストール資材がファイルサーバーに存在するかチェック
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

# ==== インストール資材のダウンロード ====
function Download-InstallArtifact {
    # インストール対象バージョン
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    # インストール資材のファイル名
    $InstallArtifactFileName = $Config["InstallArtifactFileName"]
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionWithoutDot}", $TargetVersionString.Replace(".", ""))
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{Version}", $TargetVersionString)
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionForWinMerge}", ($TargetVersionString -replace "(\d+)\.(\d+)\.(\d+)\.(\d+)", "`$1.`$2.`$3-jp-`$4"))

    # インストール資材のダウンロード後のローカルディスク上のファイルパス
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # インストール資材のダウンロード先のファイルパス
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName

    # インストール資材のダウンロード先フォルダーが存在しなければ作成する
    if (-not (Test-Path $LocalDownloadDirPath -PathType Container)) {
        New-Item -ItemType Directory -Path $LocalDownloadDirPath | Out-Null
    }

    # ダウンロード元URL
    $InstallArtifactDownloadURL = $Config["InstallArtifactDownloadURL"]
    $InstallArtifactDownloadURL = $InstallArtifactDownloadURL.Replace("{VersionForWinMerge}", ($TargetVersionString -replace "(\d+)\.(\d+)\.(\d+)\.(\d+)", "`$1.`$2.`$3%2B-jp-`$4"))
    $InstallArtifactDownloadURL = $InstallArtifactDownloadURL.Replace("{InstallArtifactFileName}", $InstallArtifactFileName)

    # ==== ダウンロード処理 ====
    # .NET FrameworkのSystem.Net.WebClientクラスのインスタンスを生成
    $WebClient = New-Object WebClient

    # ループ制御
    $InLoopFlag = $True
    $LoopCount = 0
    while ($InLoopFlag) {
        # ループは基本的に1回で抜ける想定
        $LoopCount++
        $InLoopFlag = $False
        try {
            # ダウンロード
            Write-Log "情報: インストール資材のダウンロードを開始します。"
            Start-Sleep -Milliseconds 800
            Write-Log "情報:  - ダウンロード元URL: $InstallArtifactDownloadURL"
            Start-Sleep -Milliseconds 400
            Write-Log "情報:  - ダウンロード先パス: $InstallArtifactFilePath"
            Start-Sleep -Milliseconds 800
            $WebClient.DownloadFile($InstallArtifactDownloadURL, $InstallArtifactFilePath)

            # ダウンロードしたファイルが存在するかチェック
            if (-not (Test-Path $InstallArtifactFilePath -PathType Leaf)) {
                throw "警告: インストール資材のダウンロードに失敗しました。"
            }
        }
        catch {
            # プロキシ判定
            $Uri = New-Object Uri $InstallArtifactDownloadURL
            $ProxyUri = [Net.WebRequest]::GetSystemWebProxy().GetProxy($Uri)
            if ($Uri.AbsolutePath -ne $ProxyUri.AbsolutePath) {
                # プロキシ認証情報の入力を促す
                [PSCredential]$Credential = Get-Credential -Message "プロキシ認証のユーザー名とパスワードを入力してください。"
                if (-not [String]::IsNullOrWhiteSpace($Credential.UserName)) {
                    [WebRequest]::DefaultWebProxy = [WebRequest]::GetSystemWebProxy()
                    [WebRequest]::DefaultWebProxy.Credentials = $Credential
                } else {
                    Write-Log "情報: プロキシ認証のユーザー名とパスワードの入力がキャンセルされました。"
                    Start-Sleep -Milliseconds 1250
                }
                # ループを継続させる
                if ($LoopCount -lt 2) {
                    $InLoopFlag = $True
                }
            }
        }
    }

    # ダウンロードしたインストール資材ファイルが存在するかチェック
    if (Test-Path $InstallArtifactFilePath -PathType Leaf) {
        Write-Log "情報: インストール資材のダウンロードが完了しました。"
        Start-Sleep -Milliseconds 1250
        return [PSCustomObject]@{
            IsFailed = $False
        }
    } else {
        Write-Log "エラー: インストール資材のダウンロードに失敗しました。[$InstallArtifactDownloadURL]"
        Start-Sleep -Milliseconds 1250
        return [PSCustomObject]@{
            IsFailed = $True
        }
    }
}

# ==== インストール処理 ====
function Install-Application {
    # インストール対象バージョン
    $TargetVersionString = $Config["TargetVersion"]
    $TargetVersion = New-Object Version $TargetVersionString

    # インストール資材のダウンロード後のローカルディスク上のファイルパス
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # インストール資材のファイル名
    $InstallArtifactFileName = $Config["InstallArtifactFileName"]
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionWithoutDot}", $TargetVersionString.Replace(".", ""))
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{Version}", $TargetVersionString)
    $InstallArtifactFileName = $InstallArtifactFileName.Replace("{VersionForWinMerge}", ($TargetVersionString -replace "(\d+)\.(\d+)\.(\d+)\.(\d+)", "`$1.`$2.`$3-jp-`$4"))

    # インストール資材のダウンロード先のファイルパス
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName

    # インストール先フォルダーパス
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $InstallDestinationFolderRealPath = $InstallDestinationFolderPath
    if ($Config["InstallDestinationFolderTargetIsParentFlag"] -eq "True") {
        $InstallDestinationFolderRealPath = Split-Path $InstallDestinationFolderPath -Parent
    }

    # ==== 7-Zipで圧縮ファイルを展開 ====
    $SevenZipExecutablePath = Get-SevenZipExecutablePath
    # 7zの実行ファイルの存在チェック
    if ($SevenZipExecutablePath -eq $Null) {
        Write-Log "エラー: 7-Zipの実行ファイルが見つかりません。"
        End-ThisScriptAsFailure
    }
    Write-Log "情報: インストールを開始します。"
    Start-Sleep -Milliseconds 800
    Write-Log "情報:  - インストール資材ファイル: $InstallArtifactFilePath"
    Start-Sleep -Milliseconds 400
    Write-Log "情報:  - インストール先フォルダー: $InstallDestinationFolderPath"
    Start-Sleep -Milliseconds 800
    Write-Log "情報: インストール資材展開中..."
    Start-Sleep -Milliseconds 100
    $Exec = Start-Process $SevenZipExecutablePath "x `"$InstallArtifactFilePath`" -o`"$InstallDestinationFolderRealPath`" -y" -Wait -WindowStyle Hidden -PassThru
    if ($Exec.ExitCode -ne 0) {
        Write-Log "エラー: 7-Zipでの展開に失敗しました。(エラーコード: $($Exec.ExitCode))"
        Start-Sleep -Milliseconds 1250
        End-ThisScriptAsFailure
    }
    Start-Sleep -Milliseconds 400
    Write-Log "情報: インストールを完了しました。"
    Start-Sleep -Milliseconds 1250

    # インストールされているかチェック
    $VersionCheckResult = Check-VersionIsJustOrNewer -NoOutput
    if (-not $VersionCheckResult.IsJustOrNewer) {
        Write-Log "エラー: インストール後のアプリケーションのファイルからバージョン情報が取得できませんでした。インストールに失敗した可能性があります。[$InstallDestinationFolderPath]"
        Start-Sleep -Milliseconds 1250
        End-ThisScriptAsFailure
    }
}

# ==== インストール後の初期設定 ====
function Do-AfterInstall {
    # インストール先フォルダーパス
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $InstallDestinationFolderRealPath = $InstallDestinationFolderPath
    if ($Config["InstallDestinationFolderTargetIsParentFlag"] -eq "True") {
        $InstallDestinationFolderRealPath = Split-Path $InstallDestinationFolderPath -Parent
    }

    # 環境変数PATHに追加
    Add-Path $InstallDestinationFolderPath "WinMergeU.exe"

    # コンテキストメニュー登録バッチを呼ぶ
    $RegisterBatchFilePath = Join-Path $InstallDestinationFolderPath "RegisterPerUser.bat"
    Write-Log "情報: コンテキストメニュー登録バッチ(ユーザー単位)を実行します。"
    Start-Sleep -Milliseconds 100
    Start-Process Cmd.exe "/C Call `"$RegisterBatchFilePath`" /s" -WorkingDirectory $InstallDestinationFolderPath -Wait -WindowStyle Hidden
    Write-Log "情報: コンテキストメニュー登録バッチ(ユーザー単位)を実行しました。"
    Start-Sleep -Milliseconds 700

    try {
        Write-Log "情報: Windows11コンテキストメニュー登録処理を実行します。"
        Start-Sleep -Milliseconds 100
        $MsixContextMenuPackageFilePath = Join-Path $InstallDestinationFolderPath "WinMergeContextMenuPackage.msix"
        Add-AppxPackage $MsixContextMenuPackageFilePath -ExternalLocation $InstallDestinationFolderPath
        Write-Log "情報: Windows11コンテキストメニュー登録処理を実行しました。"
    } catch {
        Write-Log "警告: WinMergeのコンテキストメニュー登録に失敗しました。"
    }

    # レジストリに登録する
    $RegistryKey = "HKCU:\Software\Thingamahoochie\WinMerge"
    $EntryName = "Executable"
    $EntryValue = "$InstallDestinationFolderPath\WinMergeU.exe"
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_SZ
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "ContextMenuEnabled"
    $EntryValue = 7
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $RegistryKey = "HKCU:\Software\Thingamahoochie\WinMerge\Locale"
    $EntryName = "LanguageId"
    $EntryValue = 1041
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $RegistryKey = "HKCU:\Software\Thingamahoochie\WinMerge\Backup"
    $EntryName = "EnableFile"
    $EntryValue = 0
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $RegistryKey = "HKCU:\Software\Thingamahoochie\WinMerge\Settings"
    $EntryName = "ColorScheme"
    $EntryValue = "Solarized Light"
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_SZ
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "DiffAlgorithm"
    $EntryValue = 3
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "IgnoreEol"
    $EntryValue = 1
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "CompMethod2"
    $EntryValue = 1
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "WordWrap"
    $EntryValue = 1
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "ViewLineNumbers"
    $EntryValue = 1
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "ViewEOL"
    $EntryValue = 1
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"

    $EntryName = "ViewWhitespace"
    $EntryValue = 1
    Set-RegistryValue $RegistryKey $EntryName $EntryValue $REG_DWORD
    Write-Log "情報: レジストリキー[$($RegistryKey.Replace("HKCU:", "HKEY_CURRENT_USER"))]のエントリ[$EntryName]を値[$EntryValue]で設定しました。"
}

# ==== 7-Zipの実行ファイルパスを取得する ====
function Get-SevenZipExecutablePath {
    # Libフォルダーに存在する場合は優先する
    if (Test-Path (Join-Path $PSScriptRoot "Lib\7za.exe") -PathType Leaf) {
        return (Join-Path $PSScriptRoot "Lib\7za.exe")
    }
    # パスが通っている箇所から探す
    $SevenZipExecutablePath = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($SevenZipExecutablePath -ne $Null) {
        return $SevenZipExecutablePath.Path
    }
    # 既定の場所を探す
    if (-not (Test-Path "C:\Program Files\7-Zip\7z.exe" -PathType Leaf)) {
        return "C:\Program Files\7-Zip\7z.exe"
    }
    return $Null
}

# ==== レジストリの値の型(種類)の定数値を定義 ====
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

# ==== レジストリの型名の文字列から対応する値を返す ====
function Convert-RegistryTypeNameToIntValue {
    param (
        [String]$RegistryTypeName
    )
    # レジストリ型名と対応する整数値の連想配列を定義
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

    # キーが存在すればその値を返す
    if ($registryTypeMap.ContainsKey($RegistryTypeName)) {
        return [int]$registryTypeMap[$RegistryTypeName]
    }

    # キーが存在しない場合のエラーメッセージ
    Write-Log "エラー: 未知のレジストリ型名です。[$RegistryTypeName]"
    exit 1
}

# ==== レジストリの型名の値から対応する文字列を返す ====
function Convert-RegistryTypeIntValueToName {
    param (
        [int]$RegistryTypeValue
    )
    # 数値と型名のマッピングを定義
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

    # キーが存在すれば型名を返す
    if ($typeNameMap.ContainsKey($RegistryTypeValue)) {
        return $typeNameMap[$RegistryTypeValue]
    }

    # キーが存在しない場合のエラーメッセージ
    Write-Log "エラー: 未知のレジストリ型値です。[$RegistryTypeValue]"
    exit 1
}

# ==== レジストリの値だけでなく型情報を補助的に持たせるためのクラスを定義 ====
class RegistryValue {
    [Object]$Value
    [int]$Type

    # コンストラクター
    RegistryValue([Object]$Value, [int]$Type) {
        $this.Value = $Value
        $this.Type = $Type
    }

    # コンストラクター(値のみ)
    RegistryValue([Object]$Value) {
        $this.Value = $Value
        $this.Type = $Script:REG_SZ
    }

    # 値を返す
    [String] ToString() {
        return $this.Value
    }
}

# ==== レジストリのキーが存在するか返す ====
function Is-ExistRegistryKey {
    [OutputType([Boolean])]
    param (
        [String]$RegistryKey
    )
    $RegistryKey = $RegistryKey.TrimEnd("\")
    # レジストリハイブ名をPowerShell向けに正規化
    $RegistryKey = $RegistryKey -replace "^HKEY_CURRENT_USER\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_LOCAL_MACHINE\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_CLASSES_ROOT\\", "HKCR:\"
    $RegistryKey = $RegistryKey -replace "^HKCU\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKLM\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKCR\\", "HKCR:\"
    # キーの存在チェック
    return (Test-Path $RegistryKey)
}

# ==== レジストリのキーにエントリが存在するか返す ====
function Is-ExistRegistryEntry {
    [OutputType([Boolean])]
    param (
        [String]$RegistryKey,
        [String]$EntryName
    )
    # キーの存在チェック
    if (-not (Test-Path $RegistryKey)) {
        return $False
    }
    $RegistryKey = $RegistryKey.TrimEnd("\")
    # レジストリハイブ名をPowerShell向けに正規化
    $RegistryKey = $RegistryKey -replace "^HKEY_CURRENT_USER\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_LOCAL_MACHINE\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKEY_CLASSES_ROOT\\", "HKCR:\"
    $RegistryKey = $RegistryKey -replace "^HKCU\\", "HKCU:\"
    $RegistryKey = $RegistryKey -replace "^HKLM\\", "HKLM:\"
    $RegistryKey = $RegistryKey -replace "^HKCR\\", "HKCR:\"
    # エントリの存在チェック
    $RegistryKeyObject = Get-Item $RegistryKey
    $EntryFoundCount = 0
    $RegistryKeyObject.Property | Where-Object { $_.ToLower() -eq $EntryName.ToLower() } | ForEach-Object { $EntryFoundCount++ }
    return ($EntryFoundCount -eq 1)
}

# ==== レジストリの値を読み出す ====
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
    # レジストリハイブ名をRegコマンド向けに正規化
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

# ==== レジストリの値を設定する ====
function Set-RegistryValue {
    param (
        [String]$RegistryKey,
        [String]$EntryName,
        [String]$EntryValue,
        [int]$EntryType = -1
    )
    # レジストリハイブ名をRegコマンド向けに正規化
    $RegistryKey = $RegistryKey -replace "^HKCU:?\\", "HKEY_CURRENT_USER\"
    $RegistryKey = $RegistryKey -replace "^HKLM:?\\", "HKEY_LOCAL_MACHINE\"
    $RegistryKey = $RegistryKey -replace "^HKCR:?\\", "HKEY_CLASSES_ROOT\"
    # 未指定と$REG_NONEの0が区別できるようここで判定
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
        Write-Log "エラー: レジストリの値の設定に失敗しました。"
        End-ThisScriptAsFailure
    }
}

# ==== レジストリの値を設定する ====
function Set-RegistryValue2 {
    [OutputType([RegistryValue])]
    param (
        [String]$RegistryKey,
        [String]$EntryName,
        [RegistryValue]$RegistryValue
    )
    Set-RegistryValue $RegistryKey $EntryName $RegistryValue.Value $RegistryValue.Type
}

# ==== 環境変数Pathを設定する ====
function Add-Path {
    param (
        [String]$PathToAdd,
        [String]$ExampleExecFilePath = $Null
    )
    # 既存の値を取得
    [RegistryValue]$PathEnvValue = Get-RegistryValue "HKCU:\Environment" "Path"
    # セミコロンで分割
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
    # 既存の値に設定しようとした値があるかチェック
    $NeedToAdd = $True
    # 大文字小文字の差異、%囲みの有無の差異を吸収するために小文字化、環境変数展開を行う
    $PathToAddExtractedLower = ([Environment]::ExpandEnvironmentVariables($PathToAdd)).ToLower()
    foreach ($PathItem in $PathList) {
        $PathItemExtractedLower = ([Environment]::ExpandEnvironmentVariables($PathItem)).ToLower()
        # 同じパスがすでに登録されていれば追加しない
        if ($PathItemExtractedLower -eq $PathToAddExtractedLower) {
            $NeedToAdd = $False
            break
        }
        # 追加しようとしたフォルダーパスとは別のパスでも同じ実行ファイルが存在する場合は追加しない
        if ($ExampleExecFilePath -ne $Null) {
            $TempPathItemExecFilePath = Join-Path $PathItem $ExampleExecFilePath
            $TempPathItemExecFilePathExtractedLower = ([Environment]::ExpandEnvironmentVariables($TempPathItemExecFilePath)).ToLower()
            if ($PathItemExtractedLower -eq $TempPathItemExecFilePathExtractedLower) {
                Write-Log "情報: 環境変数Path(ユーザー単位)にすでにパス[$PathToAdd]と同じ実行ファイルを持つパス[$PathItem]が登録されていたため追加しませんでした。"
                Start-Sleep -Milliseconds 600
                $NeedToAdd = $False
                break
            }
        }
    }
    # 既存の値に設定しようとした値がなければ追加
    if ($NeedToAdd) {
        $PathList.Add($PathToAdd)
    }
    # 新しいPathを設定
    $NewPathValue = [String]::Join(";", $PathList)
    if ($NewPathValue -eq $PathEnvValue.Value) {
        Write-Log "情報: 環境変数Path(ユーザー単位)にすでにパス[$PathToAdd]が登録されていたため追加しませんでした。"
        Start-Sleep -Milliseconds 1250
        return
    }
    Set-RegistryValue "HKCU:\Environment" "Path" $NewPathValue $Script:REG_EXPAND_SZ
    Write-Log "情報: 環境変数Path(ユーザー単位)にパス[$PathToAdd]を追加しました。"
    Start-Sleep -Milliseconds 600
    Notify-RegistryChanged
}

# ==== レジストリの変更を全ウィンドウに通知する ====
# (効果のほどは不明)
function Notify-RegistryChanged {
    # レジストリの変更を全ウィンドウに通知する

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

# ==== ログ準備 ====
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

# ==== ログ出力 ====
function Write-Log {
    param (
        [String]$Message,
        [Switch]$NoFileOutput = $False, # フラグが指定された場合はログファイルに出力しない
        [Switch]$NoConsoleOutput = $False # フラグが指定された場合はコンソールに出力しない
    )
    # ログ出力日時
    $LogDateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.ffffff")
    # ログファイル追記
    if (-not $NoFileOutput) {
        Add-Content -Path $Script:LogFilePath -Value ("[$LogDateTime] $Message")
    }
    # コンソール出力
    if (-not $NoConsoleOutput) {
        Write-Host -ForegroundColor DarkGray "[$LogDateTime] " -NoNewline
        Write-Host $Message
    }
}

# ==== 終了処理 ====
function End-ThisScript {
    # 環境変数NoWaitForInputEnterKeyに値"true"がセットされているときは値をセットした呼び元のバッチファイルにてキー入力待ちを行うと解釈し、このスクリプトファイルではキー入力待ちしない
    if ($Env:NoWaitForInputEnterKey -ne "true") {
        # プロンプトを表示してキー入力を待つ
        Write-Log ""
        Start-Sleep -Milliseconds 300
        Write-Host ("終了します。") -NoNewLine; Write-Host ("Enter") -NoNewLine -ForegroundColor Green; Write-Host ("キーを押してください。")
        Write-Log "終了します。" -NoConsoleOutput
        Start-Sleep -Milliseconds 100
        Read-Host
        Start-Sleep -Milliseconds 700
    }
    exit 0
}

# ==== エラーになる終了処理 ====
function End-ThisScriptAsFailure {
    # 環境変数NoWaitForInputEnterKeyに値"true"がセットされているときは値をセットした呼び元のバッチファイルにてキー入力待ちを行うと解釈し、このスクリプトファイルではキー入力待ちしない
    if ($Env:NoWaitForInputEnterKey -ne "true") {
        # プロンプトを表示してキー入力を待つ
        Write-Log ""
        Start-Sleep -Milliseconds 300
        Write-Host ("終了します。") -NoNewLine; Write-Host ("Enter") -NoNewLine -ForegroundColor Green; Write-Host ("キーを押してください。")
        Write-Log "終了します。" -NoConsoleOutput
        Start-Sleep -Milliseconds 100
        Read-Host
        Start-Sleep -Milliseconds 700
    }
    exit 1
}

Initialize-Log -LogId "WinMerge"

Main
