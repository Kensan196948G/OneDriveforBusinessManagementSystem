<#
.SYNOPSIS
    OneDriveストレージクォータ取得スクリプト
.DESCRIPTION
    Microsoft Graph APIを使用してOneDriveのストレージ使用状況を取得します
#>

# 設定ファイル読み込み
$configPath = Join-Path $PSScriptRoot "..\config.json"
$config = Get-Content $configPath -Encoding UTF8 | ConvertFrom-Json

# MSAL.PSモジュールチェック
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Install-Module MSAL.PS -Force -Scope CurrentUser
}

# アクセストークン取得
$tokenParams = @{
    TenantId     = $config.tenantId
    ClientId     = $config.clientId
    ClientSecret = $config.clientSecret
    Scopes       = $config.scopes
}

try {
    $authResult = Get-MsalToken @tokenParams
    $accessToken = $authResult.AccessToken

    $headers = @{
        Authorization = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }

    # 全ユーザー取得
    $usersUrl = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,mail"
    $users = (Invoke-RestMethod -Uri $usersUrl -Headers $headers -Method Get).value

    # 各ユーザーのOneDriveクォータ取得
    $result = @()
    foreach ($user in $users) {
        $driveUrl = "https://graph.microsoft.com/v1.0/users/$($user.id)/drive"
        $drive = Invoke-RestMethod -Uri $driveUrl -Headers $headers -Method Get

        if ($drive.quota) {
            $usedGB = [math]::Round($drive.quota.used / 1GB, 2)
            $totalGB = [math]::Round($drive.quota.total / 1GB, 2)
            $usagePercentage = [math]::Round(($drive.quota.used / $drive.quota.total) * 100, 2)

            $result += [PSCustomObject]@{
                displayName = $user.displayName
                mail = $user.mail
                usedGB = $usedGB
                totalGB = $totalGB
                usagePercentage = $usagePercentage
            }
        }
    }

    # JSON出力
    $result | ConvertTo-Json -Depth 5
} catch {
    Write-Error "エラーが発生しました: $_"
    exit 1
}