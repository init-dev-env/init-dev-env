# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=
# =-=-= インストールバッチコマンド、PowerShellスクリプトをテンプレートからコピー   =-=-=
# =+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=-=+=

# エラー発生時にスクリプトを処理継続させず停止する
$ErrorActionPreference = "Stop"

# ==== メイン処理 ====
function Main {

    $InfoList = @(
        @{
            "TitleTail" = "ダウンロード・インストール"
            "FileNameTail" = "Download-Install"
            "TemplateNumber" = "1"
        }
        # @{
        #     "TitleTail" = "ダウンロード"
        #     "FileNameTail" = "Download"
        #     "TemplateNumber" = "2"
        # }
        # @{
        #     "TitleTail" = "インストール"
        #     "FileNameTail" = "Install"
        #     "TemplateNumber" = "3"
        # }
    )

    $FileTypeList = @(
        @{
            "Kind" = "バッチコマンド"
            "Extension" = "cmd"
            "Comment" = "@rem"
            "TemplateFileName" = "888_Template-BatchCommand.cmd"
        }
        @{
            "Kind" = "PowerShellスクリプト"
            "Extension" = "ps1"
            "Comment" = "#"
            "TemplateFileName" = "888_Template-PowerShellScript.ps1"
        }
        @{
            "Kind" = "設定ファイル"
            "Extension" = "ini"
            "Comment" = "#"
            "TemplateFileName" = "888_Template-Config.ini"
        }
    )

    $Result = $Null

    while ($Result -ne "Y") {

        $SoftwareName = Read-Host -Prompt "ソフトウェア名？"

        $SequenceNumber = Read-Host -Prompt "連番？"

        Write-Host ""
        Write-Host "下記でよろしいですか？"
        Write-Host $SoftwareName
        Write-Host $SequenceNumber

        $Result = Read-Host "[Y/N]?"
        Write-Host ""
    }

    if ($SoftwareName -eq "") {
        Write-Host "ソフトウェア名が入力されていません。"
        return
    }

    if ($SequenceNumber -eq "") {
        Write-Host "連番が入力されていません。"
        return
    }

    foreach ($Info in $InfoList) {
        $TitleTail = $Info.TitleTail
        $FileNameTail = $Info.FileNameTail
        $TemplateNumber = $Info.TemplateNumber

        $Title = $SoftwareName + "を" + $TitleTail

        $TitleLineInfo = Make-Title $Title

        Write-Host $TitleLineInfo.Line
        Write-Host $TitleLineInfo.Title
        Write-Host $TitleLineInfo.Line

        foreach ($FileType in $FileTypeList) {
            $Kind = $FileType.Kind
            $Extension = $FileType.Extension
            $Comment = $FileType.Comment
            $TemplateFileName = $FileType.TemplateFileName

            if ($Kind -eq "設定ファイル") {
                $Postfix = "_≪Config≫"
            } else {
                $Postfix = ""
            }
            $FileName = $SequenceNumber + "_" + $SoftwareName + "_" + $FileNameTail + $Postfix + "." + $Extension

            Make-File $FileName $TitleLineInfo $Comment $TemplateFileName $TemplateNumber
        }
    }
}

function Make-Title {
    [OutputType([Object])]
    [CmdletBinding()]
    param (
        [String] $TitleName
    )

    $Line0 = "=+="
    $Line1 = "-=+="
    $TitleHead = "=-=-="
    $TitleFoot = "=-=-="
    $TitleName = " " + $TitleName + " "

    $Title = $TitleHead + $TitleName + $TitleFoot
    $TitleLength = [Text.Encoding]::Default.GetBytes($Title).Length

    $Line0Length = [Text.Encoding]::Default.GetBytes($Line0).Length
    $Line1Length = [Text.Encoding]::Default.GetBytes($Line1).Length

    $Line1Count = [Math]::Ceiling(($TitleLength - $Line0Length) / $Line1Length)
    $LineString = $Line0 + $Line1 * $Line1Count
    $LineLength = $LineString.Length

    $TitleAfterSpacing = ""
    while ($LineLength -gt $TitleLength) {
        $TitleAfterSpacing += " "
        $Title = $TitleHead + $TitleName + $TitleAfterSpacing + $TitleFoot
        $TitleLength = [Text.Encoding]::Default.GetBytes($Title).Length
    }

    return @{
        "Title" = $Title
        "TitleHead" = $TitleHead
        "TitleFoot" = $TitleFoot
        "TitleName" = $TitleName
        "TitleAfterSpacing" = $TitleAfterSpacing
        "Line" = $LineString
    }
}

function Make-File {
    [CmdletBinding()]
    param (
        [String] $FileName,
        [Object] $TitleLineInfo,
        [String] $Comment,
        [String] $TemplateFileName,
        [String] $TemplateNumber
    )

    $TitleLineHead = $Comment + " "

    $Content = Make-Content $TitleLineInfo $TitleLineHead $TemplateFileName $TemplateNumber

    if (Test-Path $FileName -PathType Leaf) {
        Write-Host "ファイルが存在します。[$TemplateFileName]"
        return
    }
    [IO.File]::WriteAllText($FileName, $Content, [Text.Encoding]::Default)
}

function Make-Content {
    [CmdletBinding()]
    param (
        [Object] $TitleLineInfo,
        [String] $TitleLineHead,
        [String] $TemplateFileName,
        [String] $TemplateNumber
    )

    $Title = $TitleLineInfo.Title
    $TitleHead = $TitleLineInfo.TitleHead
    $TitleFoot = $TitleLineInfo.TitleFoot
    $TitleName = $TitleLineInfo.TitleName
    $TitleAfterSpacing = $TitleLineInfo.TitleAfterSpacing
    $Line = $TitleLineInfo.Line

    $TemplateFileName = $TemplateFileName -replace "{}", $TemplateNumber
    $TemplateContent = [IO.File]::ReadAllText($TemplateFileName, [Text.Encoding]::Default)
    $TitlePart = ""
    $TitlePart += $TitleLineHead + $Line + "`r`n"
    $TitlePart += $TitleLineHead + $Title + "`r`n"
    $TitlePart += $TitleLineHead + $Line + "`r`n"
    $TitlePart += "`r`n"
    $ShowTitle = ""
    if ($TitleLineHead -eq "@rem ") {
        $BorderStart = "[36m[44m"
        $BorderEnd = "[0m"
        $CenterStart = "[33m[40m"
        $BorderEnd = "[0m"
        $ShowTitle += "Echo $BorderStart" + $Line + "$BorderEnd`r`n"
        $ShowTitle += "Echo $BorderStart" + $TitleHead + $CenterStart + $TitleName + $TitleAfterSpacing + $BorderStart + $TitleFoot + "$BorderEnd`r`n"
        $ShowTitle += "Echo $BorderStart" + $Line + "$BorderEnd"
    } else {
        $ShowTitle += "Write-Host -NoNewLine ([char]0x1B + `"[0d`")`r`n"
        $ShowTitle += "Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue `"" + $Line + "`"`r`n"
        $ShowTitle += "Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue `"" + $TitleHead + "`" -NoNewLine;"
        $ShowTitle += "Write-Host -ForegroundColor Yellow `"" + $TitleName + $TitleAfterSpacing + "`" -NoNewLine;"
        $ShowTitle += "Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue `"" + $TitleFoot + "`"`r`n"
        $ShowTitle += "Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue `"" + $Line + "``r``n`""
    }

    $TemplateContent = $TitlePart + $TemplateContent.Replace("{ShowTitle}", $ShowTitle)

    return $TemplateContent
}

Main
