<#
.SYNOPSIS
    OneDriveユーザー情報取得スクリプト
.DESCRIPTION
    Microsoft Graph APIを使用してユーザー情報を取得します
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

    # ユーザー情報取得
    $headers = @{
        Authorization = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }

    # 検索パラメータがある場合はフィルタリング
    $search = $args[0]
    $filter = if ($search) {
        "`$filter=startswith(displayName,'$search') or startswith(mail,'$search')"
    } else {
        ""
    }

    $apiUrl = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,mail,accountEnabled,lastLoginDateTime,assignedLicenses&$filter"
    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get

    # 結果を整形
    $result = $response.value | ForEach-Object {
        [PSCustomObject]@{
            displayName       = $_.displayName
            mail             = $_.mail
            accountEnabled   = if ($_.accountEnabled) { "有効" } else { "無効" }
            lastLoginDateTime = if ($_.lastLoginDateTime) { 
                [datetime]::Parse($_.lastLoginDateTime).ToString("yyyy/MM/dd HH:mm") 
            } else { "未ログイン" }
            licenseType      = if ($_.assignedLicenses) { "ライセンスあり" } else { "ライセンスなし" }
        }
    }

    # JSON出力
    $result | ConvertTo-Json -Depth 5
} catch {
    Write-Error "エラーが発生しました: $_"
    exit 1
}