# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
# =-=-= Cliborをダウンロード・インストール  =-=-=
# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

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
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-=" -NoNewLine;Write-Host -ForegroundColor Yellow " Cliborをダウンロード・インストール  " -NoNewLine;Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=-=-="
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue "=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=`r`n"

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

    # 追加資材のファイル名
    $AdditionalArtifactFileName = $Config["AdditionalArtifactFileName"]

    # インストール資材のダウンロード後のローカルディスク上のファイルパス
    $LocalDownloadDirPath = Join-Path $Config["LocalDownloadRootFolderPath"] $Config["DownloadFolderName"]

    # インストール資材のダウンロード先のファイルパス
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName

    # 追加資材のダウンロード先のファイルパス
    $AdditionalArtifactFilePath = Join-Path $LocalDownloadDirPath $AdditionalArtifactFileName

    # インストール資材のダウンロード先フォルダーが存在しなければ作成する
    if (-not (Test-Path $LocalDownloadDirPath -PathType Container)) {
        New-Item -ItemType Directory -Path $LocalDownloadDirPath | Out-Null
    }

    # ダウンロード元URL
    $InstallArtifactDownloadURL = $Config["InstallArtifactDownloadURL"]
    $InstallArtifactDownloadURL = $InstallArtifactDownloadURL.Replace("{InstallArtifactFileName}", $InstallArtifactFileName)

    # 追加資材のダウンロード元URL
    $AdditionalArtifactDownloadURL = $Config["AdditionalArtifactDownloadURL"]
    $AdditionalArtifactDownloadURL = $AdditionalArtifactDownloadURL.Replace("{AdditionalArtifactFileName}", $AdditionalArtifactFileName)

    # ==== ダウンロード処理 ====
    # .NET FrameworkのSystem.Net.WebClientクラスのインスタンスを生成
    $WebClient = New-Object WebClient

    # ループ制御
    for ($FileIndex = 1; $FileIndex -le 2; $FileIndex++) {
        if ($FileIndex -eq 1) {
            $ArtifactDownloadURL = $InstallArtifactDownloadURL
            $ArtifactDownloadFilePath = $InstallArtifactFilePath
            $ArtifactDisplayName = "インストール資材"
        } else {
            $ArtifactDownloadURL = $AdditionalArtifactDownloadURL
            $ArtifactDownloadFilePath = $AdditionalArtifactFilePath
            $ArtifactDisplayName = "追加資材"
        }
        $AdditionalArtifactDownloadSuccess = $False
        $InLoopFlag = $True
        $LoopCount = 0
        while ($InLoopFlag) {
            # ループは基本的に1回で抜ける想定
            $LoopCount++
            $InLoopFlag = $False
            try {
                # ダウンロード
                Write-Log "情報: ${ArtifactDisplayName}のダウンロードを開始します。"
                Start-Sleep -Milliseconds 800
                Write-Log "情報:  - ダウンロード元URL: $ArtifactDownloadURL"
                Start-Sleep -Milliseconds 400
                Write-Log "情報:  - ダウンロード先パス: $ArtifactDownloadFilePath"
                Start-Sleep -Milliseconds 800
                $WebClient.DownloadFile($ArtifactDownloadURL, $ArtifactDownloadFilePath)

                # ダウンロードしたファイルが存在するかチェック
                if (-not (Test-Path $ArtifactDownloadFilePath -PathType Leaf)) {
                    throw "エラー: ${ArtifactDisplayName}のダウンロードに失敗しました。"
                }
            }
            catch {
                # プロキシ判定
                if ($Env:ComputerName.StartsWith("B063")) {
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
        if (Test-Path $ArtifactDownloadFilePath -PathType Leaf) {
            Write-Log "情報: ${ArtifactDisplayName}のダウンロードが完了しました。"
            Start-Sleep -Milliseconds 1250
            if ($FileIndex -eq 2) {
                return [PSCustomObject]@{
                    IsFailed = $False
                }
            }
        } else {
            Write-Log "エラー: ${ArtifactDisplayName}のダウンロードに失敗しました。[$DownloadURL]"
            Start-Sleep -Milliseconds 1250
            if ($FileIndex -eq 1) {
                return [PSCustomObject]@{
                    IsFailed = $True
                }
            }
            if ($FileIndex -eq 2) {
                return [PSCustomObject]@{
                    IsFailed = $False
                }
            }
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

    # 追加資材のファイル名
    $AdditionalArtifactFileName = $Config["AdditionalArtifactFileName"]

    # インストール資材のダウンロード先のファイルパス
    $InstallArtifactFilePath = Join-Path $LocalDownloadDirPath $InstallArtifactFileName

    # インストール資材のインストール先フォルダーパス
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $InstallDestinationFolderRealPath = $InstallDestinationFolderPath
    if ($Config["InstallDestinationFolderTargetIsParentFlag"] -eq "True") {
        $InstallDestinationFolderRealPath = Split-Path $InstallDestinationFolderPath -Parent
    }

    # 追加資材のダウンロード先のファイルパス
    $AdditionalArtifactFilePath = Join-Path $LocalDownloadDirPath $AdditionalArtifactFileName

    # 追加資材のインストール先フォルダーパス
    $AdditionalInstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $AdditionalInstallDestinationFolderRealPath = $AdditionalInstallDestinationFolderPath

    # インストール先フォルダーパスを先に作成
    if (-not (Test-Path $InstallDestinationFolderPath -PathType Container)) {
        New-Item -ItemType Directory -Path $InstallDestinationFolderPath | Out-Null
    }

    # 実行中のプロセスが存在すれば終了させる
    (Get-Process -Name Clibor -ErrorAction SilentlyContinue) | ForEach-Object { . $_.Path /xt }

    # ==== 7-Zipで圧縮ファイルを展開 ====
    $SevenZipExecutablePath = Get-SevenZipExecutablePath
    # 7zの実行ファイルの存在チェック
    if ($SevenZipExecutablePath -eq $Null) {
        Write-Log "エラー: 7-Zipの実行ファイルが見つかりません。"
        End-ThisScriptAsFailure
    }
    for ($FileIndex = 1; $FileIndex -le 2; $FileIndex++) {
        if ($FileIndex -eq 1) {
            $ArtifactFilePath = $InstallArtifactFilePath
            $DestinationFilePath = $InstallDestinationFolderRealPath
            $ArtifactDisplayName = "インストール資材"
        } else {
            $ArtifactFilePath = $AdditionalArtifactFilePath
            $DestinationFilePath = $AdditionalInstallDestinationFolderRealPath
            $ArtifactDisplayName = "追加資材"
        }
        if (($FileIndex -eq 2) -and (-not (Test-Path $ArtifactFilePath -PathType Leaf))) {
            continue
        }
        Write-Log "情報: インストールを開始します。"
        Start-Sleep -Milliseconds 800
        Write-Log "情報:  - ${ArtifactDisplayName}ファイル: $ArtifactFilePath"
        Start-Sleep -Milliseconds 400
        Write-Log "情報:  - インストール先フォルダー: $InstallDestinationFolderPath"
        Start-Sleep -Milliseconds 800
        Write-Log "情報: ${ArtifactDisplayName}展開中..."
        Start-Sleep -Milliseconds 100
        $Exec = Start-Process $SevenZipExecutablePath "x `"$ArtifactFilePath`" -o`"$DestinationFilePath`" -y" -Wait -WindowStyle Hidden -PassThru
        if ($Exec.ExitCode -ne 0) {
            Write-Log "エラー: 7-Zipでの展開に失敗しました。(エラーコード: $($Exec.ExitCode))"
            Start-Sleep -Milliseconds 1250
            End-ThisScriptAsFailure
        }
        Start-Sleep -Milliseconds 400
        Write-Log "情報: インストールを完了しました。"
        Start-Sleep -Milliseconds 1250
    }

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
    # XML設定ファイルを生成・更新
    Set-DefaultSettingXmlFile

    # 展開されたMigemoライブラリを移動
    $ExtractedMigemoLibraryFilePath = Join-Path (Join-Path $Config["InstallDestinationFolderPath"] "cmigemo-default-win32") "migemo.dll"
    $DestinationMigemoLibraryFilePath = Join-Path $Config["InstallDestinationFolderPath"] "migemo.dll"
    if (Test-Path $ExtractedMigemoLibraryFilePath -PathType Leaf) {
        if (Test-Path $DestinationMigemoLibraryFilePath -PathType Leaf) {
            Remove-Item $DestinationMigemoLibraryFilePath -Force
        }
        Move-Item $ExtractedMigemoLibraryFilePath $DestinationMigemoLibraryFilePath -Force
    }

    # 展開されたMigemo辞書ファイルのフォルダーを移動
    $ExtractedMigemoDictionaryFolderPath = Join-Path (Join-Path $Config["InstallDestinationFolderPath"] "cmigemo-default-win32") "dict"
    $DestinationMigemoDictionaryFolderPath = Join-Path $Config["InstallDestinationFolderPath"] "dict"
    if (Test-Path $ExtractedMigemoDictionaryFolderPath -PathType Container) {
        if (Test-Path $DestinationMigemoDictionaryFolderPath -PathType Container) {
            Remove-Item $DestinationMigemoDictionaryFolderPath -Recurse -Force
        }
        Move-Item $ExtractedMigemoDictionaryFolderPath $DestinationMigemoDictionaryFolderPath -Force
    }

    # 展開されたフォルダーを削除
    $ExtractedFolderPath = Join-Path $Config["InstallDestinationFolderPath"] "cmigemo-default-win32"
    if (Test-Path $ExtractedFolderPath -PathType Container) {
        Remove-Item $ExtractedFolderPath -Recurse -Force
    }

    # ユーザーごとのスタートアップフォルダーにショートカットを作成
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]
    $StartupFolderPath = [Environment]::GetFolderPath("Startup")
    $ShortcutName = "Clibor"
    $ShortcutLinkFilePath = Join-Path $StartupFolderPath "$ShortcutName.lnk"
    # 存在しなければ新規に作成
    if (-not (Test-Path $ShortcutLinkFilePath -PathType Leaf)) {
        $ShortcutTargetPath = Join-Path $InstallDestinationFolderPath "Clibor.exe"
        $ShortcutArguments = ""
        $ShortcutDescription = "Cliborの起動"
        $ShortcutIconPath = Join-Path $InstallDestinationFolderPath "Clibor.exe"
        $ShortcutIconIndex = 0
        # WSH(Windows Script Host)のShellオブジェクトを生成
        $WScriptShell = New-Object -ComObject WScript.Shell
        # ショートカットの作成
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutLinkFilePath)
        $Shortcut.TargetPath = $ShortcutTargetPath
        $Shortcut.Arguments = $ShortcutArguments
        $Shortcut.Description = $ShortcutDescription
        $Shortcut.IconLocation = "$ShortcutIconPath,$ShortcutIconIndex"
        $Shortcut.WorkingDirectory = $InstallDestinationFolderPath
        $Shortcut.Save()
        Write-Log "情報: スタートアップフォルダーにショートカットを作成しました。[$ShortcutLinkFilePath]"
    } else {
        Write-Log "情報: スタートアップフォルダーにショートカットがすでに存在します。[$ShortcutLinkFilePath]"
    }

    # 開始
    Start-Process $ShortcutTargetPath -ArgumentList $ShortcutArguments -WorkingDirectory $InstallDestinationFolderPath
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

# ==== Cliborの設定XMLファイルを生成する ====
function Set-DefaultSettingXmlFile {
    # インストール先のフォルダーパス
    $InstallDestinationFolderPath = $Config["InstallDestinationFolderPath"]

    # Cliborの設定XMLファイルのパス
    $SettingXmlFilePath = Join-Path $InstallDestinationFolderPath "Clibor.xml"

    # バージョン
    $Version = $Config["TargetVersion"]

    # 変更フラグ
    $IsChanged = $False

    # 設定ファイルが存在しない場合は一度起動させて終了させる
    if (-not (Test-Path $SettingXmlFilePath -PathType Leaf)) {
        $CliborExeFilePath = Join-Path $InstallDestinationFolderPath "Clibor.exe"
        if (Test-Path $CliborExeFilePath -PathType Leaf) {
            $LoopLimit = 10
            while (($LoopLimit -gt 0) -and (-not (Test-Path $SettingXmlFilePath -PathType Leaf))) {
                $LoopLimit--
                try {
                    $Process = Start-Process $CliborExeFilePath -PassThru
                    $Process.WaitForInputIdle()
                    Start-Sleep -Milliseconds (10 * (11 - $LoopLimit))
                    Stop-Process -InputObject $Process -Force -ErrorAction SilentlyContinue
                } catch {
                    Start-Sleep -Milliseconds (20 * (11 - $LoopLimit))
                }
            }
        }
    }

    # 起動中のプロセスがあれば終了させる
    $Process = Get-Process -Name "Clibor" -ErrorAction SilentlyContinue
    if ($Process -ne $Null) {
        Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
    }

    # 設定ファイルが既に存在すれば開く
    if (-not (Test-Path $SettingXmlFilePath -PathType Leaf)) {
        # 空のXMLオブジェクトを生成
        $Xml = New-Object System.Xml.XmlDocument

        # XML宣言の追加
        $Declaration = $Xml.CreateXmlDeclaration("1.0", "UTF-16", $Null)
        $Xml.AppendChild($Declaration) | Out-Null

        $IsChanged = $True
    } else {
        # XML宣言の属性がversionとencodingの順でないと読み込みエラーになるため、一度読み込んで文字列をパースさせる
        $XmlContent = Get-Content $SettingXmlFilePath -Encoding Unicode
        $XmlContent = $XmlContent.Replace("<?xml encoding=`"UTF-16`" version=`"1.0`"?>", "<?xml version=`"1.0`" encoding=`"UTF-16`"?>")

        $Xml = [Xml]($XmlContent)
    }

    # もしルート要素が存在しないならば、ルート要素の作成。存在するならば取得する
    if ($Xml.DocumentElement -eq $Null) {
        $Root = $Xml.CreateElement("root")
        $Xml.AppendChild($Root) | Out-Null

        $IsChanged = $True
    } else {
        $Root = $Xml.DocumentElement
    }

    # もしCLIBOR要素が存在しないならば、CLIBOR要素の作成。存在するならば取得する
    if ($Root.CLIBOR -eq $Null) {
        $Clibor = $Xml.CreateElement("CLIBOR")
        $Root.AppendChild($Clibor) | Out-Null

        $IsChanged = $True
    } else {
        $Clibor = $Root.CLIBOR
    }

    # 子要素の作成と追加
    $Elements = @{
        CLIBOR_VER                = "Clibor ver${Version}"
        # 基本動作
        AUTOPASTEWAIT             = "100"
        # 基本動作/メイン画面
        AP_ABVALUE                = "242"
        FORMWIDTH                 = "640"
        YOHAKU                    = "2"
        MAINLINECNT               = "128"
        MAINTABHIDE               = "-1"
        MAINPAGEHIDE              = "-1"
        MAINSHADOW                = "-1"
        WAKULINE                  = "1"
        # 基本動作/フォント
        MAINFONTNAME              = "InGen UI N"
        MAINFONTSIZE              = "8"
        TTIPFONTNAME              = "InGen UI N"
        TTIPFONTSIZE              = "8"
        MAINSEARCHFONTNAME        = "PlemolJP"
        MAINSEARCHFONTSIZE        = "8"
        MAINTABFONTNAME           = "Yu Gothic UI"
        MAINTABFONTSIZE           = "8"
        MAINTABNOFONTNAME         = "Yu Gothic UI"
        MAINTABNOFONTSIZE         = "8"
        MAINPGFONTNAME            = "Yu Gothic UI"
        MAINPGFONTSIZE            = "8"
        TEIKEIEDITFONTNAME        = "InGen UI N"
        TEIKEIEDITFONTSIZE        = "8"
        MAINMENUFONTNAME          = "InGen UI N"
        MAINMENUFONTSIZE          = "8"
        # クリップボード
        HIST_SIZE                 = "1024"
        CLIPPRICNT                = "30"
        # クリップボード/保存
        AUTOCLIPSAVETIME          = "600"
        BACKUP_CLIP               = "-1"
        BACKUP_CLIP_CNT           = "2"
        # クリップボード/更新
        CLIPDELAY                 = "40"
        CLIPIGNORE                = "300"
        # 定型文
        BACKUP_TEIKEI             = "-1"
        BACKUP_TEIKEI_CNT         = "2"
        # 画面表示制御
        MAINFOCUS                 = "-1"
        SEARCH_ONOFF              = "3"
        # 配色
        BACK_COLOR                = '$00FFF9F2'
        SLTBACK_COLOR             = '$00A85400'
        BACK_COLORF               = '$00313131'
        SLTBACK_COLORF            = "clWhite"
        LINE_COLOR                = '$00793D00'
        LINEHYOJI_COLOR           = '$00AA5500'
        BACK_COLOR2               = '$00FFF4E8'
        BACK_COLORF2              = '$00383838'
        # ホットキー
        HOTKEYFLG                 = "0"
        HOTKEYSYUSYOKU            = "Ctrl+Shift"
        HOTKEYKEY                 = "Q"
        # バージョン
        AUTO_CHECK_UPDATE         = "0"
    }
    $SubElements = @{
        # 検索
        MIGEMO                    = "-1"
        MIGEMOPRIORITY            = "3"
        MIGEMOPRIORITY_T          = "3"
        # 検索/検索ボックス
        MAINSEARCHDETAIL          = "3"
        MAINSEARCHDETAIL_T        = "3"
    }
    if (Test-Path (Join-Path $InstallDestinationFolderPath "migemo.dll") -PathType Leaf) {
        # SubElementsの内容をElementsに追加
        $Elements += $SubElements
    }
    foreach ($ElementName in $Elements.Keys) {
        $Value = $Null
        if ($Clibor.$ElementName -eq $Null) {
            $Element = $Xml.CreateElement($ElementName)
            $Clibor.AppendChild($Element) | Out-Null
        } else {
            $Element = $Clibor[$ElementName]
            if ($Element -eq $Null) {
                Write-Host "警告: $ElementName の要素が取得できませんでした。"
            }
            $Value = $Element.InnerText
        }
        if ($Value -ne $Elements[$ElementName]) {
            $NewValue = $Elements[$ElementName]
            if (($NewValue -eq "InGen UI N") -or ($NewValue -eq "PlemolJP")) {
                $NewValue = "Yu Gothic UI"
            }
            try {
                $Element.InnerText = $NewValue
            } catch {
                Write-Host "警告: $ElementName の値の設定に失敗しました。"
            }
            $IsChanged = $True
        }
    }

    # XMLをファイルに保存
    if ($IsChanged) {
        $Xml.Save($SettingXmlFilePath)
        Write-Log "情報: Cliborの設定XMLファイルを更新しました。[$SettingXmlFilePath]"
    }
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

Initialize-Log -LogId "Clibor"

Main
