# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
# =-=-= フォントリスト作成
# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

# ==== 初期設定 ====
# 省略するネームスペース表記
using namespace System.IO
using namespace System.Net
using namespace System.Text
using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace System.Security.Principal

$FontMetaInfoList = @(
    @{
        Kind = "GitHub"
        Repositories = @(
            @{
                RepositoryName = "yuru7"
                FontNames = @(
                    "moralerspace"
                    "PlemolJP"
                    "udev-gothic"
                    "HackGen"
                    "NOTONOTO"
                    "Explex"
                    "juisee"
                    "bizin-gothic"
                    "pending-mono"
                    "Firge"
                    "BIZTER"
                    "InGenUI"
                )
            }
        )
    }
)

$FontUrlList = [List[String]]::New()
foreach ($FontMetaInfo in $FontMetaInfoList) {
    $Kind = $FontMetaInfo.Kind
    $Repositories = $FontMetaInfo.Repositories
    foreach ($Repository in $Repositories) {
        $RepositoryName = $Repository.RepositoryName
        $FontNames = $Repository.FontNames
        if ($Kind -eq "GitHub") {
            $BaseUrl = "https://github.com/$RepositoryName"
            foreach ($FontName in $FontNames) {
                $FontReleaseUrl = "$BaseUrl/$FontName/releases/latest"
                Write-Host "- FontReleaseUrl: $FontReleaseUrl"
                $HtmlFileName = "FontPage_$FontName.html"
                if (-not (Test-Path $HtmlFileName)) {
                    Invoke-WebRequest -Uri $FontReleaseUrl -OutFile $HtmlFileName
                    Start-Sleep -Milliseconds 5000
                }
                $Html = Get-Content -Path $HtmlFileName -Encoding UTF8
                $Lines = $Html.Split("`n")
                $Include = $Null
                foreach ($Line in $Lines) {
                    if (-not ($Line.Contains("<include-fragment loading=`"lazy`" src=`""))) {
                        continue
                    }
                    $Line = $Line.Substring($Line.IndexOf("<include-fragment loading=`"lazy`" src=`"") + "<include-fragment loading=`"lazy`" src=`"".Length)
                    $Line = $Line.Substring(0, $Line.IndexOf("`""))
                    $Include = $Line
                    break
                }
                if ($Include -eq $Null) {
                    Write-Host "  - Include is not found."
                    continue
                }
                Write-Host "  - Include: $Include"
                $HtmlFileName = "FontPage_$FontName`_Include.html"
                if (-not (Test-Path $HtmlFileName)) {
                    Invoke-WebRequest -Uri $Include -OutFile $HtmlFileName
                    Start-Sleep -Milliseconds 5000
                }
                $Html = Get-Content -Path $HtmlFileName -Encoding UTF8
                $Lines = $Html.Split("`n")
                $DownloadPathList = [List[String]]::New()
                foreach ($Line in $Lines) {
                    if (-not $Line.Contains("<a")) {
                        continue
                    }
                    Write-Host "  - Line: $Line"
                    if ($Line -match "<a.*href=`"([^`"]+)`".*>") {
                        $DownloadPath = $Matches[1]
                        if ((-not ($DownloadPath.Contains("/refs/tags/"))) -and ((-not ($DownloadPath.EndsWith(".7z"))) -or (-not ($DownloadPath.EndsWith(".zip"))))) {
                            $DownloadPathList.Add($DownloadPath)
                        }
                    }
                }
                foreach ($DownloadPath in $DownloadPathList) {
                    $Url = "https://github.com$DownloadPath"
                    Write-Host "    - URL: $Url"
                    $FontUrlList.Add($Url)
                }
            }
        }
    }
}
$Content = [String]::Join("`n", $FontUrlList)
[File]::WriteAllText("12_Fonts__UrlList.txt", $Content)
