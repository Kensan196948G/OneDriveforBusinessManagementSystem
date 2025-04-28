# OneDrive for Business 運用ツール - ITSM準拠
# GetAllUserDriveRoots.ps1 - ユーザーごとのOneDriveルート情報取得スクリプト

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$ConfigPath = "$PSScriptRoot\..\config.json"
)

$executionTime = Get-Date

# Main.ps1からログ関数をインポート
. "$PSScriptRoot\..\Main.ps1"

# HTMLレポート生成関数
function Generate-HtmlReport {
    param(
        [Parameter(Mandatory=$true)]
        [array]$UserData,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputDir
    )
    
    try {
        $templatePath = Join-Path $PSScriptRoot "..\WebUI_Template\GetAllUserDriveRoots.html"
        if (!(Test-Path $templatePath)) {
            throw "HTMLテンプレートが見つかりません: $templatePath"
        }
        
        $htmlContent = Get-Content $templatePath -Raw
        
        # データをJSON形式に変換
        $jsonData = $UserData | ConvertTo-Json -Depth 10
        
        # テンプレートにデータを埋め込む
        $htmlContent = $htmlContent -replace '// DATA_PLACEHOLDER', "const reportData = $jsonData;"
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $outputPath = Join-Path $OutputDir "GetAllUserDriveRoots_$timestamp.html"
        $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8
        
        return $outputPath
    } catch {
        Write-ErrorLog $_ "HTMLレポート生成中にエラーが発生しました"
        throw
    }
}

# Microsoft Graph接続
try {
    # config.jsonから認証情報を取得
    $config = Get-Content -Path $ConfigPath | ConvertFrom-Json
    
    # 非対話型認証で接続
    Write-Log "Microsoft Graphに接続中 (ClientSecret認証)..." "INFO"
    
    $params = @{
        TenantId = $config.TenantId
        ClientId = $config.ClientId
        ClientSecret = $config.ClientSecret
    }
    
    Connect-MgGraph @params -ErrorAction Stop
    Write-Log "Microsoft Graphに正常に接続されました (アプリケーション認証)" "SUCCESS"
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続中にエラーが発生しました"
    exit 1
}

# データ収集
$userList = @()
try {
    $allUsers = Get-MgUser -All -Property DisplayName,Mail,UserPrincipalName,Id -ConsistencyLevel eventual -CountVariable totalCount
    $totalUsers = $allUsers.Count
    $processedUsers = 0

    foreach ($user in $allUsers) {
        $processedUsers++
        $percentComplete = [math]::Round(($processedUsers / $totalUsers) * 100, 2)
        Write-Progress -Activity "OneDriveルート情報を取得中" -Status "$processedUsers / $totalUsers ユーザー処理中 ($percentComplete%)" -PercentComplete $percentComplete
        
        try {
            $drive = Get-MgUserDriveRoot -UserId $user.UserPrincipalName -ErrorAction Stop
            $driveInfo = [PSCustomObject]@{
                "ユーザー名" = $user.DisplayName
                "メールアドレス" = $user.Mail
                "OneDrive URL" = $drive.WebUrl
                "作成日時" = $drive.CreatedDateTime
                "最終更新日" = $drive.LastModifiedDateTime
                "所有者" = $drive.Owner.User.DisplayName
            }
            $userList += $driveInfo
        } catch {
            Write-Log "$($user.UserPrincipalName) のOneDriveルート情報取得に失敗: $_" "WARNING"
        }
    }
} catch {
    Write-ErrorLog $_ "OneDriveルート情報の取得中にエラーが発生しました"
    exit
}

# HTMLレポート生成
try {
    $htmlReport = Generate-HtmlReport -UserData $userList -OutputDir $OutputDir
    Write-Log "HTMLレポートを生成しました: $htmlReport" "SUCCESS"
} catch {
    Write-ErrorLog $_ "HTMLレポートの生成中にエラーが発生しました"
}

# CSV出力
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$csvFile = "AllUserDriveRoots.$timestamp.csv"
$csvPath = Join-Path -Path $OutputDir -ChildPath $csvFile

try {
    $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSVファイルが生成されました: $csvPath" "SUCCESS"
} catch {
    Write-ErrorLog $_ "CSVファイルの生成中にエラーが発生しました"
}

Write-Log "すべてのOneDriveルート情報収集が完了しました" "SUCCESS"