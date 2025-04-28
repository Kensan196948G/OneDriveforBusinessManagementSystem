# OneDrive for Business 運用ツール - ITSM準拠
# SyncErrorCheck.ps1 - 同期エラー確認スクリプト

param (
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "$(Get-Location)",
    [string]$ConfigPath = "$PSScriptRoot\..\config.json",
    [int]$RetryCount = 3,
    [switch]$AutoRemediate = $false,
    [switch]$GenerateHtmlReport = $false,
    [switch]$IncludePerformanceData = $false,
    [string]$NotificationMethod = "None",
    [switch]$ForceCacheClear = $false
)

# 列名リストを明示的に定義
$columnNames = @("ユーザー名","メールアドレス","アカウント状態","OneDrive対応","エラー種別","ファイル名","ファイルパス",
                "最終更新日時","サイズ","エラー詳細","推奨対応","修復ステータス","エラーコード","影響度","修復アクション")

# ログ関数定義
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    Write-Output "[$timestamp] [$Level] $Message"
}

# 詳細なエラー分類関数
function Get-ErrorClassification {
    param(
        [string]$ErrorMessage,
        [string]$FileName
    )
    
    $errorType = "同期エラー"
    $severity = "低"
    $errorCode = "SYNC-000"
    $recommendation = "OneDriveクライアントを再起動し、ネットワーク接続を確認してください。"
    
    # エラーメッセージとファイル名に基づく詳細分類
    if ($ErrorMessage -match "conflict") {
        $errorType = "ファイル競合"
        $severity = "中"
        $errorCode = "SYNC-101"
        $recommendation = "競合ファイルを確認し、必要に応じて手動で解決してください。"
    } 
    elseif ($ErrorMessage -match "quota") {
        $errorType = "容量不足"
        $severity = "高"
        $errorCode = "SYNC-201"
        $recommendation = "OneDriveストレージ容量を増やすか、不要なファイルを削除してください。"
    } 
    elseif ($ErrorMessage -match "permission") {
        $errorType = "権限エラー"
        $severity = "高"
        $errorCode = "SYNC-301"
        $recommendation = "ファイルの権限設定を確認してください。管理者に連絡が必要な場合があります。"
    }
    elseif ($ErrorMessage -match "timeout") {
        $errorType = "タイムアウト"
        $severity = "中"
        $errorCode = "SYNC-401"
        $recommendation = "ネットワーク接続を確認し、再試行してください。"
    }
    elseif ($ErrorMessage -match "virus") {
        $errorType = "ウイルス検出"
        $severity = "緊急"
        $errorCode = "SYNC-501"
        $recommendation = "セキュリティチームに即時連絡してください。"
    }
    elseif ($FileName -match "\.(exe|msi|bat|ps1|vbs)$") {
        $errorType = "実行ファイルブロック"
        $severity = "高"
        $errorCode = "SYNC-601"
        $recommendation = "セキュリティポリシーによりブロックされています。別の方法で共有してください。"
    }
    elseif ($ErrorMessage -match "path too long") {
        $errorType = "パス長制限"
        $severity = "中"
        $errorCode = "SYNC-701"
        $recommendation = "ファイルパスを短くするか、階層を浅くしてください。"
    }
    elseif ($ErrorMessage -match "invalid character") {
        $errorType = "不正文字"
        $severity = "中"
        $errorCode = "SYNC-801"
        $recommendation = "ファイル名から不正な文字を削除してください。"
    }
    elseif ($ErrorMessage -match "locked") {
        $errorType = "ファイルロック"
        $severity = "中"
        $errorCode = "SYNC-901"
        $recommendation = "ファイルがロックされています。使用中のアプリケーションを閉じてください。"
    }
    
    return @{
        ErrorType = $errorType
        Severity = $severity
        ErrorCode = $errorCode
        Recommendation = $recommendation
    }
}

# 自動修復関数
function Invoke-ErrorRemediation {
    param(
        [string]$ErrorCode,
        [string]$FilePath,
        [string]$UserPrincipalName,
        [string]$FileName
    )
    
    $status = "未実施"
    $action = ""
    
    try {
        switch ($ErrorCode) {
            "SYNC-101" {
                # 競合ファイルの自動解決
                $conflictFile = "$FilePath.conflict"
                if (Test-Path $conflictFile) {
                    Remove-Item $conflictFile -Force
                    $status = "自動修復済"
                    $action = "競合ファイルを削除"
                }
            }
            "SYNC-201" {
                # 容量不足通知
                $mailParams = @{
                    To = $UserPrincipalName
                    Subject = "OneDrive容量不足警告"
                    Body = "ストレージ容量が不足しています。不要なファイルを削除するか、管理者に連絡してください。"
                    From = "noreply@yourdomain.com"
                    SmtpServer = "smtp.yourdomain.com"
                }
                Send-MailMessage @mailParams
                $status = "通知送信済"
                $action = "容量不足警告を送信"
            }
            "SYNC-301" {
                # 権限リセット
                $item = Get-MgDriveItem -DriveId "me" -DriveItemId (Split-Path $FilePath -Leaf)
                if ($item) {
                    Set-MgDriveItemPermission -DriveId "me" -DriveItemId $item.Id -Roles @("read")
                    $status = "自動修復済"
                    $action = "権限を閲覧のみに変更"
                }
            }
            "SYNC-401" {
                # キャッシュクリア
                if ($ForceCacheClear) {
                    $cachePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\cache"
                    if (Test-Path $cachePath) {
                        Remove-Item "$cachePath\*" -Recurse -Force
                        $status = "自動修復済"
                        $action = "キャッシュをクリア"
                    }
                }
            }
            default {
                $status = "手動対応必要"
                $action = "自動修復不可"
            }
        }
    } catch {
        $status = "修復失敗"
        $action = "修復中にエラー: $_"
    }
    
    return @{
        Status = $status
        Action = $action
    }
}

# 通知送信関数
function Send-ErrorNotification {
    param(
        [object]$ErrorDetail,
        [string]$Method
    )
    
    try {
        switch ($Method) {
            "Email" {
                $mailParams = @{
                    To = $ErrorDetail.メールアドレス
                    Subject = "OneDrive同期エラー通知 [$($ErrorDetail.エラーコード)]"
                    Body = @"
以下の同期エラーが検出されました:

- ユーザー: $($ErrorDetail.ユーザー名)
- ファイル名: $($ErrorDetail.ファイル名)
- エラー種別: $($ErrorDetail.エラー種別)
- 影響度: $($ErrorDetail.影響度)
- 推奨対応: $($ErrorDetail.推奨対応)

詳細は管理者にお問い合わせください。
"@
                    From = "noreply@yourdomain.com"
                    SmtpServer = "smtp.yourdomain.com"
                }
                Send-MailMessage @mailParams
            }
            "Teams" {
                $teamsMessage = @{
                    title = "OneDrive同期エラー通知 [$($ErrorDetail.エラーコード)]"
                    text = @"
**ユーザー**: $($ErrorDetail.ユーザー名)
**ファイル**: $($ErrorDetail.ファイル名)
**エラー**: $($ErrorDetail.エラー種別) ($($ErrorDetail.影響度))
**推奨対応**: $($ErrorDetail.推奨対応)
"@
                }
                # Teams Webhook送信ロジック
                Invoke-RestMethod -Uri $config.TeamsWebhookUrl -Method Post -Body ($teamsMessage | ConvertTo-Json) -ContentType "application/json"
            }
            default {
                Write-Log "通知方法が指定されていません: $Method" "INFO"
            }
        }
        
        Write-Log "通知を送信しました: $Method" "INFO"
        return $true
    } catch {
        Write-Log "通知の送信に失敗しました: $_" "ERROR"
        return $false
    }
}

# HTMLレポート生成関数
function Generate-HtmlReport {
    param(
        [array]$Errors,
        [string]$OutputPath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>同期エラーレポート</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .emergency { background-color: #ffcccc; }
        .high { background-color: #ffe6cc; }
        .medium { background-color: #ffffcc; }
        .low { background-color: #e6ffe6; }
        .summary { margin-top: 30px; }
        .chart-container { width: 100%; height: 400px; margin-top: 30px; }
        .details { display: none; margin-top: 10px; }
        .error-details { padding: 10px; }
        .inner-table { width: 100%; border-collapse: collapse; }
        .inner-table td { padding: 5px; border: none; vertical-align: top; }
        .inner-table td:first-child {
            font-weight: bold;
            width: 120px;
            color: #555;
        }
        .inner-table td:first-child { font-weight: bold; width: 120px; }
    </style>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
    <h1>同期エラーレポート - $(Get-Date -Format 'yyyy/MM/dd')</h1>
    <p>生成日時: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')</p>
    <p>検出された同期エラー数: $($Errors.Count)</p>
    
    <table>
        <tr>
            <th>ユーザー名</th>
            <th>エラー種別</th>
            <th>ファイル名</th>
            <th>影響度</th>
            <th>修復ステータス</th>
            <th>アクション</th>
        </tr>
"@

    # エラー種別ごとの集計
    $errorStats = @{}
    $severityStats = @{}
    $remediationStats = @{ Success = 0; Failed = 0; Pending = 0 }
    
    foreach ($error in $Errors) {
        $rowClass = switch ($error.影響度) {
            "緊急" { "emergency" }
            "高" { "high" }
            "中" { "medium" }
            default { "low" }
        }
        
        # 統計情報を更新
        if (-not $errorStats.ContainsKey($error.エラー種別)) {
            $errorStats[$error.エラー種別] = 0
        }
        $errorStats[$error.エラー種別]++
        
        if (-not $severityStats.ContainsKey($error.影響度)) {
            $severityStats[$error.影響度] = 0
        }
        $severityStats[$error.影響度]++
        
        switch ($error.修復ステータス) {
            { $_ -match "修復済|通知送信済" } { $remediationStats.Success++ }
            "修復失敗" { $remediationStats.Failed++ }
            default { $remediationStats.Pending++ }
        }
        
        $html += @"
        <tr class="$rowClass" onclick="toggleDetails('details_$($error.エラーコード)')">
            <td>$($error.ユーザー名)</td>
            <td>$($error.エラー種別)</td>
            <td>$($error.ファイル名)</td>
            <td>$($error.影響度)</td>
            <td>$($error.修復ステータス)</td>
            <td>$($error.修復アクション)</td>
        </tr>
        <tr id="details_$($error.エラーコード)" class="details">
            <td colspan="6">
                <div class="error-details">
                    <strong>詳細情報:</strong><br>
                    <table class="inner-table">
                        <tr><td>ファイルパス:</td><td>$($error.ファイルパス)</td></tr>
                        <tr><td>最終更新日時:</td><td>$($error.最終更新日時)</td></tr>
                        <tr><td>サイズ:</td><td>$([math]::Round($error.サイズ, 2)) KB</td></tr>
                        <tr><td>エラーコード:</td><td>$($error.エラーコード)</td></tr>
                        <tr><td>エラー詳細:</td><td>$($error.エラー詳細)</td></tr>
                        <tr><td>推奨対応:</td><td>$($error.推奨対応)</td></tr>
                    </table>
                </div>
            </td>
        </tr>
"@
    }

    $html += @"
    </table>
    
    <div class="summary">
        <h2>エラー統計</h2>
        <div class="chart-container">
            <canvas id="errorTypeChart"></canvas>
        </div>
        <div class="chart-container">
            <canvas id="severityChart"></canvas>
        </div>
        <div class="chart-container">
            <canvas id="remediationChart"></canvas>
        </div>
    </div>
    
    <script>
        function toggleDetails(id) {
            const element = document.getElementById(id);
            element.style.display = element.style.display === 'none' ? 'table-row' : 'none';
        }
        
        // エラー種別グラフ
        const errorTypeCtx = document.getElementById('errorTypeChart').getContext('2d');
        new Chart(errorTypeCtx, {
            type: 'pie',
            data: {
                labels: [$(($errorStats.Keys | ForEach-Object { "'$_'" }) -join ',')],
                datasets: [{
                    data: [$(($errorStats.Values -join ','))],
                    backgroundColor: [
                        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', '#FF9F40',
                        '#8AC24A', '#FF5722', '#607D8B', '#9C27B0'
                    ]
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'エラー種別分布'
                    }
                }
            }
        });
        
        // 影響度グラフ
        const severityCtx = document.getElementById('severityChart').getContext('2d');
        new Chart(severityCtx, {
            type: 'bar',
            data: {
                labels: [$(($severityStats.Keys | ForEach-Object { "'$_'" }) -join ',')],
                datasets: [{
                    label: '影響度別エラー数',
                    data: [$(($severityStats.Values -join ','))],
                    backgroundColor: [
                        '#FF6384', '#FF9F40', '#FFCE56', '#4BC0C0'
                    ]
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: {
                        beginAtZero: true
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: '影響度別エラー数'
                    }
                }
            }
        });
        
        // 修復状況グラフ
        const remediationCtx = document.getElementById('remediationChart').getContext('2d');
        new Chart(remediationCtx, {
            type: 'doughnut',
            data: {
                labels: ['修復済', '修復失敗', '未対応'],
                datasets: [{
                    data: [$($remediationStats.Success), $($remediationStats.Failed), $($remediationStats.Pending)],
                    backgroundColor: [
                        '#4BC0C0', '#FF6384', '#FFCE56'
                    ]
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: '修復状況'
                    }
                }
            }
        });
    </script>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Log "HTMLレポートを作成しました: $OutputPath" "SUCCESS"
    } catch {
        Write-Log "HTMLレポートの作成中にエラーが発生しました: $_" "ERROR"
    }
}

# 実行開始時刻を記録
$executionTime = Get-Date
Write-Log "同期エラー確認を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"

# パフォーマンスデータ収集
$performanceData = @{
    StartTime = $executionTime
    UserCount = 0
    ErrorCount = 0
    ErrorTypes = @{}
    RemediationStats = @{
        Success = 0
        Failed = 0
        Pending = 0
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
    
    # グローバル管理者かどうかを判定
    $context = Get-MgContext
    $isAdmin = ($context.Scopes -contains "Directory.ReadWrite.All")
    
    Write-Log "Microsoft Graphに正常に接続されました (アプリケーション認証)" "SUCCESS"
    Write-Log "実行モード: $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})" "INFO"
} catch {
    Write-Log "Microsoft Graphの接続中にエラーが発生しました: $_" "ERROR"
    exit 1
}

# 出力用のエラーリスト
$errorList = @()

try {
    if ($isAdmin) {
        # グローバル管理者の場合、すべてのユーザーの同期エラーを確認
        Write-Log "すべてのユーザーの同期エラーを確認しています..." "INFO"
        
        try {
            # 同期エラーを取得 (リトライ付き)
            $syncErrors = Get-MgReportOneDriveUsageAccountDetail -Period "D30" -ErrorAction Stop | 
                Where-Object { $_.SyncErrorCount -gt 0 }
            
            $performanceData.UserCount = $syncErrors.Count
            
            foreach ($errorDetail in $syncErrors) {
                try {
                    $user = Get-MgUser -UserId $errorDetail.UserPrincipalName -ErrorAction SilentlyContinue
                    
                    # 詳細なエラー分類
                    $classification = Get-ErrorClassification -ErrorMessage $errorDetail.ErrorMessage -FileName $errorDetail.FileName
                    
                    # 自動修復
                    $remediation = @{ Status = "未実施"; Action = "" }
                    if ($AutoRemediate) {
                        $remediation = Invoke-ErrorRemediation -ErrorCode $classification.ErrorCode `
                            -FilePath $errorDetail.FilePath -UserPrincipalName $errorDetail.UserPrincipalName -FileName $errorDetail.FileName
                        
                        # 修復統計更新
                        switch ($remediation.Status) {
                            { $_ -match "修復済|通知送信済" } { $performanceData.RemediationStats.Success++ }
                            "修復失敗" { $performanceData.RemediationStats.Failed++ }
                            default { $performanceData.RemediationStats.Pending++ }
                        }
                    }
                    
                    # 通知送信
                    if ($NotificationMethod -ne "None") {
                        Send-ErrorNotification -ErrorDetail @{
                            ユーザー名 = $user.DisplayName
                            メールアドレス = $user.Mail
                            エラー種別 = $classification.ErrorType
                            ファイル名 = $errorDetail.FileName
                            影響度 = $classification.Severity
                            推奨対応 = $classification.Recommendation
                            エラーコード = $classification.ErrorCode
                        } -Method $NotificationMethod
                    }
                    
                    $errorList += [PSCustomObject]@{
                        "ユーザー名"       = $user.DisplayName
                        "メールアドレス"   = $user.Mail
                        "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
                        "OneDrive対応"     = "対応"
                        "エラー種別"       = $classification.ErrorType
                        "ファイル名"       = $errorDetail.FileName
                        "ファイルパス"     = $errorDetail.FilePath
                        "最終更新日時"     = $errorDetail.LastActivityDate
                        "サイズ"           = [math]::Round($errorDetail.FileSize / 1KB, 2)
                        "エラー詳細"       = $errorDetail.ErrorMessage
                        "推奨対応"         = $classification.Recommendation
                        "修復ステータス"   = $remediation.Status
                        "エラーコード"     = $classification.ErrorCode
                        "影響度"           = $classification.Severity
                        "修復アクション"   = $remediation.Action
                    }
                    
                    # パフォーマンスデータ更新
                    $performanceData.ErrorCount++
                    if (-not $performanceData.ErrorTypes.ContainsKey($classification.ErrorType)) {
                        $performanceData.ErrorTypes[$classification.ErrorType] = 0
                    }
                    $performanceData.ErrorTypes[$classification.ErrorType]++
                    
                    Write-Log "同期エラーを検出: $($errorDetail.FileName) (ユーザー: $($user.DisplayName), 種別: $($classification.ErrorType), 修復: $($remediation.Status))" "WARNING"
                } catch {
                    Write-Log "同期エラー詳細の取得に失敗: $($errorDetail.UserPrincipalName) - $_" "WARNING"
                }
            }
            
            Write-Log "同期エラーの確認が完了しました。検出されたエラー: $($errorList.Count)" "SUCCESS"
        } catch {
            Write-Log "同期エラーの取得中にエラーが発生しました: $_" "ERROR"
        }
    } else {
        # 一般ユーザーの場合、自分の同期エラーのみ確認
        Write-Log "自分自身の同期エラーを確認しています..." "INFO"
        
        try {
            $currentUser = Get-MgContext | Select-Object -ExpandProperty Account
            $myErrors = Get-MgReportOneDriveUsageAccountDetail -Period "D30" -UserId $currentUser -ErrorAction Stop | 
                Where-Object { $_.SyncErrorCount -gt 0 }
            
            $performanceData.UserCount = 1
            
            foreach ($errorDetail in $myErrors) {
                # 詳細なエラー分類
                $classification = Get-ErrorClassification -ErrorMessage $errorDetail.ErrorMessage -FileName $errorDetail.FileName
                
                # 自動修復
                $remediation = @{ Status = "未実施"; Action = "" }
                if ($AutoRemediate) {
                    $remediation = Invoke-ErrorRemediation -ErrorCode $classification.ErrorCode `
                        -FilePath $errorDetail.FilePath -UserPrincipalName $currentUser -FileName $errorDetail.FileName
                    
                    # 修復統計更新
                    switch ($remediation.Status) {
                        { $_ -match "修復済|通知送信済" } { $performanceData.RemediationStats.Success++ }
                        "修復失敗" { $performanceData.RemediationStats.Failed++ }
                        default { $performanceData.RemediationStats.Pending++ }
                    }
                }
                
                $errorList += [PSCustomObject]@{
                    "ユーザー名"       = $currentUser
                    "メールアドレス"   = $currentUser
                    "アカウント状態"   = "有効"
                    "OneDrive対応"     = "対応"
                    "エラー種別"       = $classification.ErrorType
                    "ファイル名"       = $errorDetail.FileName
                    "ファイルパス"     = $errorDetail.FilePath
                    "最終更新日時"     = $errorDetail.LastActivityDate
                    "サイズ"           = [math]::Round($errorDetail.FileSize / 1KB, 2)
                    "エラー詳細"       = $errorDetail.ErrorMessage
                    "推奨対応"         = $classification.Recommendation
                    "修復ステータス"   = $remediation.Status
                    "エラーコード"     = $classification.ErrorCode
                    "影響度"           = $classification.Severity
                    "修復アクション"   = $remediation.Action
                }
                
                # パフォーマンスデータ更新
                $performanceData.ErrorCount++
                if (-not $performanceData.ErrorTypes.ContainsKey($classification.ErrorType)) {
                    $performanceData.ErrorTypes[$classification.ErrorType] = 0
                }
                $performanceData.ErrorTypes[$classification.ErrorType]++
                
                Write-Log "同期エラーを検出: $($errorDetail.FileName) (種別: $($classification.ErrorType), 修復: $($remediation.Status))" "WARNING"
            }
            
            Write-Log "同期エラーの確認が完了しました。検出されたエラー: $($errorList.Count)" "SUCCESS"
        } catch {
            Write-Log "自分の同期エラー取得中にエラーが発生しました: $_" "ERROR"
        }
    }
} catch {
    Write-Log "同期エラーの確認中にエラーが発生しました: $_" "ERROR"
}

# エラーが検出されなかった場合の処理
if ($errorList.Count -eq 0) {
    $errorList += [PSCustomObject]@{
        "ユーザー名"       = "N/A"
        "メールアドレス"   = "N/A"
        "アカウント状態"   = "N/A"
        "OneDrive対応"     = "N/A"
        "エラー種別"       = "情報"
        "ファイル名"       = "N/A"
        "ファイルパス"     = "N/A"
        "最終更新日時"     = Get-Date
        "サイズ"           = "N/A"
        "エラー詳細"       = "同期エラーは検出されませんでした。"
        "推奨対応"         = "対応は必要ありません。"
        "修復ステータス"   = "N/A"
        "エラーコード"     = "N/A"
        "影響度"           = "N/A"
        "修復アクション"   = "N/A"
    }
    
    Write-Log "同期エラーは検出されませんでした。" "SUCCESS"
}

# CSV出力
$csvPath = Join-Path -Path $OutputDir -ChildPath "SyncErrorReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
try {
    $errorList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"
} catch {
    Write-Log "CSVファイルの作成中にエラーが発生しました: $_" "ERROR"
}

# HTMLレポート生成
if ($GenerateHtmlReport) {
    $htmlPath = Join-Path -Path $OutputDir -ChildPath "SyncErrorReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    Generate-HtmlReport -Errors $errorList -OutputPath $htmlPath
}

# パフォーマンスデータ出力
if ($IncludePerformanceData) {
    $performanceData.EndTime = Get-Date
    $performanceData.Duration = ($performanceData.EndTime - $performanceData.StartTime).TotalSeconds
    
    $perfPath = Join-Path -Path $OutputDir -ChildPath "SyncErrorPerformance_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    try {
        $performanceData | ConvertTo-Json -Depth 3 | Out-File -FilePath $perfPath -Encoding UTF8
        Write-Log "パフォーマンスデータを出力しました: $perfPath" "INFO"
    } catch {
        Write-Log "パフォーマンスデータの出力中にエラーが発生しました: $_" "WARNING"
    }
}

# Microsoft Graphから切断
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Log "同期エラー確認を終了します" "INFO"
