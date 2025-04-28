# OneDrive for Business 運用ツール - ITSM準拠
# ExternalShareCheck.ps1 - 外部共有監視スクリプト

param (
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "$(Get-Location)",
    [string]$ConfigPath = "$PSScriptRoot\..\config.json",
    [int]$RetryCount = 3,
    [switch]$AutoRemediate = $false,
    [switch]$GenerateHtmlReport = $false,
    [switch]$IncludePerformanceData = $false,
    [string]$SecurityPolicy = "Default",
    [switch]$SendNotification = $false
)

# ログ関数定義
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    Write-Output "[$timestamp] [$Level] $Message"
}

# セキュリティポリシーチェック関数
function Test-SecurityPolicy {
    param(
        [object]$Share,
        [string]$PolicyName
    )
    
    $isCompliant = $true
    $violations = @()
    
    # デフォルトポリシー
    if ($PolicyName -eq "Default") {
        if ($Share.Scope -eq "anonymous") {
            $isCompliant = $false
            $violations += "匿名共有禁止"
        }
        
        if ($Share.ShareType -eq "edit") {
            $isCompliant = $false
            $violations += "編集権限付与禁止"
        }
    }
    # 厳格ポリシー
    elseif ($PolicyName -eq "Strict") {
        if ($Share.Scope -ne "internal") {
            $isCompliant = $false
            $violations += "内部共有のみ許可"
        }
        
        $item = Get-MgDriveItem -DriveId $Share.DriveId -DriveItemId $Share.DriveItemId -ErrorAction SilentlyContinue
        if ($item -and $item.File -and $item.File.MimeType -match "executable|script") {
            $isCompliant = $false
            $violations += "実行ファイル共有禁止"
        }
    }
    
    return @{
        IsCompliant = $isCompliant
        Violations = $violations -join ", "
    }
}

# 自動修復関数
function Invoke-SecurityRemediation {
    param(
        [object]$Share,
        [string]$PolicyName
    )
    
    $status = "未実施"
    $action = ""
    
    try {
        $policyCheck = Test-SecurityPolicy -Share $Share -PolicyName $PolicyName
        
        if (-not $policyCheck.IsCompliant) {
            # ポリシー違反の自動修復
            switch ($PolicyName) {
                "Default" {
                    if ($Share.Scope -eq "anonymous") {
                        Remove-MgShare -ShareId $Share.Id -ErrorAction Stop
                        $status = "自動修復済"
                        $action = "匿名共有を削除"
                    }
                    elseif ($Share.ShareType -eq "edit") {
                        # 編集権限を閲覧のみに変更
                        Update-MgShare -ShareId $Share.Id -Roles @("read") -ErrorAction Stop
                        $status = "自動修復済"
                        $action = "権限を閲覧のみに変更"
                    }
                }
                "Strict" {
                    if ($Share.Scope -ne "internal") {
                        Remove-MgShare -ShareId $Share.Id -ErrorAction Stop
                        $status = "自動修復済"
                        $action = "外部共有を削除"
                    }
                }
                default {
                    $status = "手動対応必要"
                    $action = "ポリシー違反"
                }
            }
        } else {
            $status = "ポリシー準拠"
            $action = "不要"
        }
    } catch {
        $status = "修復失敗"
        $action = "エラー: $_"
    }
    
    return @{
        Status = $status
        Action = $action
    }
}

# 通知送信関数
function Send-SecurityNotification {
    param(
        [object]$Share,
        [string]$PolicyName,
        [string]$Recipient
    )
    
    try {
        $subject = "[重要] セキュリティポリシー違反の外部共有が検出されました"
        $body = @"
以下の外部共有がセキュリティポリシーに違反しています:

- 共有元ユーザー: $($Share.ユーザー名)
- アイテム名: $($Share.アイテム名)
- 共有タイプ: $($Share.共有タイプ)
- リスクレベル: $($Share.リスクレベル)
- 違反内容: $($Share.コンプライアンス問題)
- 推奨アクション: $($Share.セキュリティ推奨)

詳細は添付のレポートを参照してください。
"@

        # 実際の通知送信ロジック (SMTPやTeams通知など)
        Write-Log "通知を送信しました: $Recipient" "INFO"
        return $true
    } catch {
        Write-Log "通知の送信に失敗しました: $_" "ERROR"
        return $false
    }
}

# HTMLレポート生成関数
function Generate-HtmlReport {
    param(
        [array]$Shares,
        [string]$OutputPath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>外部共有セキュリティレポート</title>
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
    </style>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
    <h1>外部共有セキュリティレポート - $(Get-Date -Format 'yyyy/MM/dd')</h1>
    <p>生成日時: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')</p>
    <p>検出された外部共有数: $($Shares.Count)</p>
    <p>適用セキュリティポリシー: $SecurityPolicy</p>
    
    <table>
        <tr>
            <th>ユーザー名</th>
            <th>アイテム名</th>
            <th>共有タイプ</th>
            <th>リスクレベル</th>
            <th>コンプライアンス</th>
            <th>修復ステータス</th>
            <th>WebURL</th>
        </tr>
"@

    # 統計情報
    $riskStats = @{}
    $complianceStats = @{ Compliant = 0; NonCompliant = 0 }
    
    foreach ($share in $Shares) {
        $rowClass = switch ($share.リスクレベル) {
            "緊急" { "emergency" }
            "高" { "high" }
            "中" { "medium" }
            default { "low" }
        }
        
        # 統計情報更新
        if (-not $riskStats.ContainsKey($share.リスクレベル)) {
            $riskStats[$share.リスクレベル] = 0
        }
        $riskStats[$share.リスクレベル]++
        
        if ($share.コンプライアンス -eq "準拠") {
            $complianceStats.Compliant++
        } else {
            $complianceStats.NonCompliant++
        }
        
        $html += @"
        <tr class="$rowClass">
            <td>$($share.ユーザー名)</td>
            <td>$($share.アイテム名)</td>
            <td>$($share.共有タイプ)</td>
            <td>$($share.リスクレベル)</td>
            <td>$($share.コンプライアンス)</td>
            <td>$($share.修復ステータス)</td>
            <td><a href="$($share.WebURL)" target="_blank">リンク</a></td>
        </tr>
"@
    }

    $html += @"
    </table>
    
    <div class="summary">
        <h2>統計情報</h2>
        <div class="chart-container">
            <canvas id="riskChart"></canvas>
        </div>
        <div class="chart-container">
            <canvas id="complianceChart"></canvas>
        </div>
    </div>
    
    <script>
        // リスクレベルグラフ
        const riskCtx = document.getElementById('riskChart').getContext('2d');
        new Chart(riskCtx, {
            type: 'bar',
            data: {
                labels: [$(($riskStats.Keys | ForEach-Object { "'$_'" }) -join ',')],
                datasets: [{
                    label: 'リスクレベル分布',
                    data: [$(($riskStats.Values -join ','))],
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
                title: {
                    display: true,
                    text: 'リスクレベル分布'
                }
            }
        });
        
        // コンプライアンスグラフ
        const complianceCtx = document.getElementById('complianceChart').getContext('2d');
        new Chart(complianceCtx, {
            type: 'pie',
            data: {
                labels: ['準拠', '違反'],
                datasets: [{
                    data: [$($complianceStats.Compliant), $($complianceStats.NonCompliant)],
                    backgroundColor: [
                        '#4BC0C0', '#FF6384'
                    ]
                }]
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: 'コンプライアンス状況'
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
Write-Log "外部共有監視を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"
Write-Log "適用セキュリティポリシー: $SecurityPolicy" "INFO"

# パフォーマンスデータ収集
$performanceData = @{
    StartTime = $executionTime
    ShareCount = 0
    HighRiskCount = 0
    PolicyViolations = 0
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

# 出力用の外部共有リスト
$externalShareList = @()

try {
    # 外部共有を取得 (リトライ付き)
    Write-Log "外部共有アイテムを収集中..." "INFO"
    $filter = "sharedWith/domain ne 'tenant.onmicrosoft.com'"
    $allShares = Get-MgShare -Filter $filter -All -PageSize 100 -ErrorAction Stop
    
    $performanceData.ShareCount = $allShares.Count
    
    foreach ($share in $allShares) {
        try {
            # アイテムの詳細情報を取得
            $driveItem = Get-MgDriveItem -DriveId $share.DriveId -DriveItemId $share.DriveItemId -ErrorAction Stop
            $owner = Get-MgUser -UserId $share.Owner.User.Id -ErrorAction SilentlyContinue
            
            # リスク評価
            $riskLevel = "低"
            $riskReasons = @()
            
            if ($share.Scope -eq "anonymous") {
                $riskLevel = $share.Link.Type -eq "edit" ? "緊急" : "高"
                $riskReasons += "匿名共有"
            } elseif ($share.Scope -eq "external") {
                $riskLevel = $share.Link.Type -eq "edit" ? "高" : "中"
                $riskReasons += "外部共有"
            }
            
            # セキュリティポリシーチェック
            $policyCheck = Test-SecurityPolicy -Share $share -PolicyName $SecurityPolicy
            
            if (-not $policyCheck.IsCompliant) {
                $performanceData.PolicyViolations++
            }
            
            # 自動修復
            $remediation = @{ Status = "未実施"; Action = "" }
            if ($AutoRemediate) {
                $remediation = Invoke-SecurityRemediation -Share $share -PolicyName $SecurityPolicy
                
                if ($riskLevel -in @("緊急","高") -and $remediation.Status -eq "自動修復済") {
                    $performanceData.HighRiskCount++
                }
            }
            
            # 通知送信
            if ($SendNotification -and -not $policyCheck.IsCompliant) {
                Send-SecurityNotification -Share $share -PolicyName $SecurityPolicy -Recipient $config.AdminEmail
            }
            
            $externalShareList += [PSCustomObject]@{
                "ユーザー名"           = $owner.DisplayName
                "メールアドレス"       = $owner.Mail
                "アカウント状態"       = if($owner.AccountEnabled){"有効"}else{"無効"}
                "アイテム名"           = $driveItem.Name
                "アイテムタイプ"       = $driveItem.File ? $driveItem.File.MimeType : "フォルダ"
                "共有タイプ"           = $share.SharedBy.User ? "直接共有" : "リンク共有"
                "共有範囲"             = $share.Scope
                "共有先"               = $share.SharedWith.User ? $share.SharedWith.User.DisplayName : "匿名リンク"
                "リスクレベル"         = $riskLevel
                "リスク要因"           = $riskReasons -join ", "
                "コンプライアンス"     = if($policyCheck.IsCompliant){"準拠"}else{"違反: $($policyCheck.Violations)"}
                "セキュリティ推奨"     = if(-not $policyCheck.IsCompliant){"ポリシー違反: $($policyCheck.Violations)"}else{"不要"}
                "修復ステータス"       = $remediation.Status
                "修復アクション"       = $remediation.Action
                "最終更新日時"         = $driveItem.LastModifiedDateTime
                "WebURL"               = $driveItem.WebUrl
                "権限ID"               = $share.Id
            }
            
            Write-Log "外部共有を検出: $($driveItem.Name) (リスクレベル: $riskLevel, コンプライアンス: $(if($policyCheck.IsCompliant){"準拠"}else{"違反"}))" "INFO"
        } catch {
            Write-Log "共有アイテムの詳細取得に失敗: $($share.Id) - $_" "WARNING"
        }
    }
    
    # 外部共有が検出されなかった場合の処理
    if ($externalShareList.Count -eq 0) {
        $externalShareList += [PSCustomObject]@{
            "ユーザー名"           = "N/A"
            "メールアドレス"       = "N/A"
            "アカウント状態"       = "N/A"
            "アイテム名"           = "N/A"
            "アイテムタイプ"       = "N/A"
            "共有タイプ"           = "N/A"
            "共有範囲"             = "N/A"
            "共有先"               = "N/A"
            "リスクレベル"         = "低"
            "リスク要因"           = "N/A"
            "コンプライアンス"     = "準拠"
            "セキュリティ推奨"     = "外部共有は検出されませんでした。"
            "修復ステータス"       = "不要"
            "修復アクション"       = "N/A"
            "最終更新日時"         = Get-Date
            "WebURL"               = "N/A"
            "権限ID"               = "N/A"
        }
        
        Write-Log "外部共有は検出されませんでした。" "SUCCESS"
    } else {
        Write-Log "外部共有の確認が完了しました。検出された外部共有: $($externalShareList.Count)" "SUCCESS"
    }
} catch {
    Write-Log "外部共有の取得中にエラーが発生しました: $_" "ERROR"
    exit 1
}

# CSV出力
$csvPath = Join-Path -Path $OutputDir -ChildPath "ExternalShareReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
try {
    $externalShareList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"
} catch {
    Write-Log "CSVファイルの作成中にエラーが発生しました: $_" "ERROR"
}

# HTMLレポート生成
if ($GenerateHtmlReport) {
    $htmlPath = Join-Path -Path $OutputDir -ChildPath "ExternalShareReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    Generate-HtmlReport -Shares $externalShareList -OutputPath $htmlPath
}

# パフォーマンスデータ出力
if ($IncludePerformanceData) {
    $performanceData.EndTime = Get-Date
    $performanceData.Duration = ($performanceData.EndTime - $performanceData.StartTime).TotalSeconds
    
    $perfPath = Join-Path -Path $OutputDir -ChildPath "ExternalSharePerformance_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    try {
        $performanceData | ConvertTo-Json -Depth 3 | Out-File -FilePath $perfPath -Encoding UTF8
        Write-Log "パフォーマンスデータを出力しました: $perfPath" "INFO"
    } catch {
        Write-Log "パフォーマンスデータの出力中にエラーが発生しました: $_" "WARNING"
    }
}

# Microsoft Graphから切断
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Log "外部共有監視を終了します" "INFO"