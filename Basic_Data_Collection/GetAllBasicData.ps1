# OneDrive for Business 運用ツール - ITSM準拠
# GetAllBasicData.ps1 - すべての基本データ収集スクリプト（カラム指定版）

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$ConfigPath = "$PSScriptRoot\..\config.json"
)

$executionTime = Get-Date

# SafeExitModuleをインポート (ルートディレクトリ直下にあるためパス修正)
$modulePath = "$PSScriptRoot\..\SafeExitModule.psm1"
if (-not (Test-Path $modulePath)) {
    Write-Error "SafeExitModuleが見つかりません: $modulePath"
    exit 1
}
Import-Module $modulePath -Force

# 出力ディレクトリを明示的に設定（ユーザー指定を優先）
$BaseDir = $PSScriptRoot
# 正しいパス構造に修正
# 出力ディレクトリ構造を修正（重複フォルダ作成防止）
$outputRootDir = $OutputDir
if ($outputRootDir -like "*Basic_Data_Collection*") {
    $dataCollectionDir = $outputRootDir
} else {
    $dataCollectionDir = Join-Path -Path $outputRootDir -ChildPath "Basic_Data_Collection.$($executionTime.ToString('yyyyMMdd'))"
}
$reportDir = Join-Path -Path $dataCollectionDir -ChildPath "GetAllBasicData"

# ディレクトリ作成（自動生成フォルダを含む）
try {
    # Basic_Data_Collection.日付フォルダが存在しない場合は作成
    if (-not (Test-Path -Path $dataCollectionDir)) {
        New-Item -Path $dataCollectionDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "Basic_Data_Collectionフォルダを作成しました: $dataCollectionDir" "INFO"
    }

    # GetOneDriveQuotaフォルダが存在しない場合は作成
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -Path $reportDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "GetOneDriveQuotaフォルダを作成しました: $reportDir" "INFO"
    }
    
    # GetAllBasicDataサブフォルダが存在しない場合は作成
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -Path $reportDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "レポート出力ディレクトリを作成しました: $reportDir" "INFO"
    }
} catch {
    Write-Log "ディレクトリ作成エラー: $_" "ERROR"
    throw
}

Write-Log "すべての基本データ収集を開始します" "INFO"
Write-Log "ベースディレクトリ: $BaseDir" "INFO"
Write-Log "出力ルートディレクトリ: $outputRootDir" "INFO"
Write-Log "レポートディレクトリ: $reportDir" "INFO"
Write-Host "出力先ディレクトリ: $outputRootDir" -ForegroundColor Cyan

# config.jsonから認証情報を読み込む
try {
    if (!(Test-Path $ConfigPath)) {
        throw "config.jsonが見つかりません: $ConfigPath"
    }
    $config = Get-Content $ConfigPath | ConvertFrom-Json
    
    # Microsoft Graphモジュールの確認とインポート
    if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
        Install-Module Microsoft.Graph -Force -Scope CurrentUser
    }
    Import-Module Microsoft.Graph

    # 非対話型認証で接続
    $params = @{
        TenantId = $config.TenantId
        ClientId = $config.ClientId
        ClientSecret = $config.ClientSecret
    }

    Connect-MgGraph @params -ErrorAction Stop
    Write-Log "Microsoft Graphに接続しました (Tenant: $($config.TenantId))" "INFO"
    Write-Log "実行モード: 管理者モード（アプリケーション権限）" "INFO"
} catch {
    Write-ErrorLog $_ "Microsoft Graphへの接続に失敗しました"
    exit 1
}

$userList = @()
$startTime = Get-Date
$loadingDisplayed = $false

try {
    Write-Log "すべてのユーザー情報とOneDriveクォータ情報を取得しています..." "INFO"
    
    # ローディング表示を1.5秒以上保証
    $loadingJob = Start-Job -ScriptBlock {
        Start-Sleep -Seconds 1
    }

    # ユーザー情報取得 (リトライ機能付き)
    $retryCount = 0
    $maxRetries = 3
    $allUsers = $null
    
    do {
        try {
            $allUsers = Get-MgUser -All -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType,UserPrincipalName,Id -ErrorAction Stop
            $allUsers = $allUsers | Where-Object { -not [string]::IsNullOrEmpty($_.UserPrincipalName) }
            $allUsers = $allUsers | Where-Object { -not [string]::IsNullOrEmpty($_.Id) }
            break
        } catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                throw
            }
            Write-Log "ユーザー情報取得に失敗しました。リトライします ($retryCount/$maxRetries)..." "WARNING"
            Start-Sleep -Seconds (5 * $retryCount)
        }
    } while ($retryCount -lt $maxRetries)

    # 管理者ロール情報取得
    $adminUserIds = @()
    $adminRoles = Get-MgDirectoryRole
    foreach ($role in $adminRoles) {
        try {
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction Stop
            $adminUserIds += $roleMembers | ForEach-Object { $_.Id.ToString() }
        } catch {
            Write-Log "管理者ロールメンバーの取得に失敗しました: $($role.DisplayName)" "WARNING"
        }
    }
    $adminUserIds = $adminUserIds | Select-Object -Unique

    # 各ユーザーの情報を取得
    $totalUsers = $allUsers.Count
    $processed = 0
    
    foreach ($user in $allUsers) {
        $processed++
        $progress = [math]::Round(($processed / $totalUsers) * 100, 2)
        Write-Log "処理中: $progress% ($processed/$totalUsers) - $($user.UserPrincipalName)" "DEBUG"

        if ($user.UserType -eq "Guest") {
            $userTypeDetail = "ゲスト"
        } elseif ($null -ne $user.Id -and $adminUserIds -contains $user.Id.ToString()) {
            $userTypeDetail = "管理者"
        } else {
            $userTypeDetail = "一般ユーザー"
        }

        $onedriveSupport = "未対応"
        $retryCount = 0
        
        do {
            try {
                $drive = Get-MgUserDrive -UserId $user.UserPrincipalName -ErrorAction Stop
                $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
                $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
                $remainingGB = [math]::Round(($drive.Quota.Remaining) / 1GB, 2)
                $usagePercent = [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 2)
                $status = if ($usagePercent -ge 90) { "危険" } elseif ($usagePercent -ge 70) { "警告" } else { "正常" }
                $onedriveSupport = "対応"
                break
            } catch {
                $retryCount++
                if ($retryCount -ge $maxRetries) {
                    $totalGB = "取得不可"
                    $usedGB = "取得不可"
                    $remainingGB = "取得不可"
                    $usagePercent = "取得不可"
                    $status = "不明"
                    Write-Log "OneDrive情報の取得に失敗しました: $($user.UserPrincipalName)" "WARNING"
                    break
                }
                Start-Sleep -Seconds (2 * $retryCount)
            }
        } while ($retryCount -lt $maxRetries)

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

    # ローディング表示を最低1.5秒保証
    $elapsed = (Get-Date) - $startTime
    if ($elapsed.TotalSeconds -lt 1.5) {
        $remaining = (1.5 - $elapsed.TotalSeconds) * 1000
        Start-Sleep -Milliseconds $remaining
    }
    $loadingJob | Stop-Job -PassThru | Remove-Job

    Write-Log "ユーザー情報とOneDriveクォータ情報の取得が完了しました。取得件数: $($userList.Count)" "SUCCESS"
    
    # HTMLレポート生成
    try {
        $htmlReport = Generate-HtmlReport -UserData $userList -OutputDir $OutputDir
        Write-Log "HTMLレポートを生成しました: $htmlReport" "SUCCESS"
    } catch {
        Write-ErrorLog $_ "HTMLレポートの生成中にエラーが発生しました"
    }
    
} catch {
    if ($loadingJob) { $loadingJob | Stop-Job -PassThru | Remove-Job }
    Write-ErrorLog $_ "ユーザー情報とOneDriveクォータ情報の取得中にエラーが発生しました"
    throw
}

# HTMLレポート生成関数
function Generate-HtmlReport {
    param(
        [Parameter(Mandatory=$true)]
        [array]$UserData,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputDir
    )
    
    try {
        $templatePath = Join-Path $PSScriptRoot "..\WebUI_Template\GetAllBasicData.html"
        if (!(Test-Path $templatePath)) {
            throw "HTMLテンプレートが見つかりません: $templatePath"
        }
        
        $htmlContent = Get-Content $templatePath -Raw
        
        # データをJSON形式に変換
        $jsonData = $UserData | ConvertTo-Json -Depth 10
        
        # テンプレートにデータを埋め込む
        $htmlContent = $htmlContent -replace '// DATA_PLACEHOLDER', "const reportData = $jsonData;"
        # 出力ディレクトリはMain.ps1で作成済み
        
        
        $timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $outputPath = Join-Path $reportDir "AllBasicData.$timestamp.html"
        $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8
        
        return $outputPath
    } catch {
        Write-ErrorLog $_ "HTMLレポート生成中にエラーが発生しました"
        throw
    }
}

$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$csvFile = "AllBasicData.$timestamp.csv"
$csvPath = Join-Path -Path $reportDir -ChildPath $csvFile

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

Write-Log "すべての基本データ収集が完了しました" "SUCCESS"
