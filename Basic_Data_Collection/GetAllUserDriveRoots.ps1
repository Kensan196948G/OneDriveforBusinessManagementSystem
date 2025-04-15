# 全ユーザーのOneDriveルート情報を取得しCSV出力するスクリプト
# config.jsonの認証情報を利用（非対話型認証）

param(
    [string]$ConfigPath = "../config.json",
    [string]$OutputPath = "Output/AllUserDriveRoots.csv"
)

# Microsoft Graph PowerShell SDKバージョン確認
Import-Module Microsoft.Graph
$mgVersion = (Get-Module Microsoft.Graph -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
Write-Host "Microsoft.Graph PowerShell SDK Version: $mgVersion"

# config.json読み込み
$config = Get-Content -Path $ConfigPath | ConvertFrom-Json

$tenantId = $config.TenantId
$clientId = $config.ClientId
$clientSecret = $config.ClientSecret
$scopes = $config.DefaultScopes -join " "

# Microsoft Graph PowerShell SDKの非対話型認証
Write-Host "認証処理開始..."
$secureSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential('app', $secureSecret)
Connect-MgGraph -ClientId $clientId -TenantId $tenantId -ClientSecretCredential $credential -NoWelcome
Write-Host "認証完了"

# 全ユーザー取得（動作確認のため最初の10件のみ取得）
Write-Host "ユーザー一覧取得開始..."
$users = Get-MgUser -Top 10
Write-Host ("取得ユーザー数: " + $users.Count)

# 結果格納用
$results = @()

$counter = 1
foreach ($user in $users) {
    Write-Host ("[" + $counter + "/" + $users.Count + "] " + $user.UserPrincipalName + " のOneDriveルート取得中...")
    try {
        $driveRoot = Get-MgUserDriveRoot -UserId $user.Id
        $results += [PSCustomObject]@{
            UserId = $user.Id
            UserPrincipalName = $user.UserPrincipalName
            DriveId = $driveRoot.Id
            WebUrl = $driveRoot.WebUrl
        }
    } catch {
        $results += [PSCustomObject]@{
            UserId = $user.Id
            UserPrincipalName = $user.UserPrincipalName
            DriveId = ""
            WebUrl = ""
            Error = $_.Exception.Message
        }
    }
    $counter++
}
Write-Host "全ユーザー分のOneDriveルート取得処理完了"

# 出力先ディレクトリ作成（なければ）
$dir = Split-Path $OutputPath -Parent
if (!(Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir | Out-Null
}

# CSV出力
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "全ユーザーのOneDriveルート情報を $OutputPath に出力しました。"
