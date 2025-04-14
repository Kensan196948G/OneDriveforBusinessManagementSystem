# OneDrive for Business 運用ツール - ITSM準拠
# GetAllBasicData.ps1 - すべての基本データ収集スクリプト（カラム指定版）

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$LogDir = "$(Get-Location)\Log"
)

$executionTime = Get-Date
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $LogDir -ChildPath "GetAllBasicData.$timestamp.log"
$errorLogPath = Join-Path -Path $LogDir -ChildPath "GetAllBasicData.Error.$timestamp.log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
    Add-Content -Path $logFilePath -Value $logMessage -Encoding UTF8
}

function Write-ErrorLog {
    param (
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorMessage = "[$timestamp] [ERROR] $Message"
    $errorDetails = @"
例外タイプ: $($ErrorRecord.Exception.GetType().FullName)
例外メッセージ: $($ErrorRecord.Exception.Message)
位置: $($ErrorRecord.InvocationInfo.PositionMessage)
スタックトレース:
$($ErrorRecord.ScriptStackTrace)

"@
    Write-Host $errorMessage -ForegroundColor Red
    Add-Content -Path $logFilePath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorDetails -Encoding UTF8
}

Write-Log "すべての基本データ収集を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"
Write-Log "ログディレクトリ: $LogDir" "INFO"

try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Log "Microsoft Graphに接続されていません。Main.ps1から実行してください。" "ERROR"
        exit
    }
    $isAdmin = $true
    $executorName = "アプリケーション認証"
    $userType = "アプリケーション"
    Write-Log "Microsoft Graphアプリケーション権限で実行中" "INFO"
    Write-Log "実行モード: 管理者モード（アプリケーション権限）" "INFO"
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続確認中にエラーが発生しました"
    exit
}

$userList = @()

try {
    Write-Log "すべてのユーザー情報とOneDriveクォータ情報を取得しています..." "INFO"
    $allUsers = Get-MgUser -All -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType,UserPrincipalName,Id
    $allUsers = $allUsers | Where-Object { -not [string]::IsNullOrEmpty($_.UserPrincipalName) }
    $allUsers = $allUsers | Where-Object { -not [string]::IsNullOrEmpty($_.Id) }

    $adminRoles = Get-MgDirectoryRole
    $adminUserIds = @()
    foreach ($role in $adminRoles) {
        try {
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction Stop
            $adminUserIds += $roleMembers | ForEach-Object { $_.Id.ToString() }
        } catch {}
    }
    $adminUserIds = $adminUserIds | Select-Object -Unique

    foreach ($user in $allUsers) {
        if ($user.UserType -eq "Guest") {
            $userTypeDetail = "ゲスト"
        } elseif ($null -ne $user.Id -and $adminUserIds -contains $user.Id.ToString()) {
            $userTypeDetail = "管理者"
        } else {
            $userTypeDetail = "一般ユーザー"
        }

        $onedriveSupport = "未対応"
        try {
            $drive = Get-MgUserDrive -UserId $user.UserPrincipalName -ErrorAction Stop
            $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
            $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
            $remainingGB = [math]::Round(($drive.Quota.Remaining) / 1GB, 2)
            $usagePercent = [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 2)
            $status = if ($usagePercent -ge 90) { "危険" } elseif ($usagePercent -ge 70) { "警告" } else { "正常" }
            $onedriveSupport = "対応"
        } catch {
            $totalGB = "取得不可"
            $usedGB = "取得不可"
            $remainingGB = "取得不可"
            $usagePercent = "取得不可"
            $status = "不明"
        }

        $userList += [PSCustomObject]@{
            "ユーザー名"       = $user.DisplayName
            "メールアドレス"   = $user.Mail
            "ログインユーザー名" = if($user.onPremisesSamAccountName){$user.onPremisesSamAccountName}else{"同期なし"}
            "ユーザー種別"     = $userTypeDetail
            "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
            "最終同期日時"     = if($user.onPremisesLastSyncDateTime){$user.onPremisesLastSyncDateTime}else{"同期情報なし"}
            "OneDrive対応"     = $onedriveSupport
            "総容量(GB)"       = $totalGB
            "使用容量(GB)"     = $usedGB
            "残り容量(GB)"     = $remainingGB
            "使用率(%)"        = $usagePercent
            "状態"             = $status
        }
    }

    Write-Log "ユーザー情報とOneDriveクォータ情報の取得が完了しました。取得件数: $($userList.Count)" "SUCCESS"
} catch {
    Write-ErrorLog $_ "ユーザー情報とOneDriveクォータ情報の取得中にエラーが発生しました"
}

$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$csvFile = "AllBasicData.$timestamp.csv"
$htmlFile = "AllBasicData.$timestamp.html"
$jsFile = "AllBasicData.$timestamp.js"
$csvPath = Join-Path -Path $OutputDir -ChildPath $csvFile
$htmlPath = Join-Path -Path $OutputDir -ChildPath $htmlFile
$jsPath = Join-Path -Path $OutputDir -ChildPath $jsFile

try {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    } else {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        $content = [System.IO.File]::ReadAllText($csvPath)
        [System.IO.File]::WriteAllText($csvPath, $content, [System.Text.Encoding]::UTF8)
    }
    Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"
} catch {
    Write-ErrorLog $_ "CSVファイルの作成中にエラーが発生しました"
}

# WebUI付きHTML出力（GenerateHtmlFromCsv.ps1を呼び出し）
try {
    $generateHtmlScript = Join-Path -Path $PSScriptRoot -ChildPath "GenerateHtmlFromCsv.ps1"
    if (Test-Path $generateHtmlScript) {
        & $generateHtmlScript -CsvPath $csvPath -OutputHtmlPath $htmlPath
        Write-Log "WebUI付きHTMLファイルを作成しました: $htmlPath" "SUCCESS"
    } else {
        Write-Log "GenerateHtmlFromCsv.ps1 が見つかりません: $generateHtmlScript" "ERROR"
    }
} catch {
    Write-ErrorLog $_ "WebUI付きHTMLファイルの作成中にエラーが発生しました"
}

Write-Log "すべての基本データ収集が完了しました" "SUCCESS"
