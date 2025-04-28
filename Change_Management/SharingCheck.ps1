# OneDrive for Business 運用ツール - ITSM準拠
# SharingCheck.ps1 - 共有設定確認スクリプト

param (
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "$(Get-Location)",
    [string]$ConfigPath = "$PSScriptRoot\..\config.json",
    [int]$RetryCount = 3,
    [switch]$AutoRemediate = $false,
    [switch]$GenerateHtmlReport = $false,
    [switch]$IncludePerformanceData = $false,
    [string]$CompliancePolicy = "Default"
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

# コンプライアンスポリシーチェック関数
function Test-Compliance {
    param(
        [object]$Share,
        [string]$PolicyName
    )
    
    $isCompliant = $true
    $violations = @()
    
    # デフォルトポリシー
    if ($PolicyName -eq "Default") {
        if ($Share.ShareType -eq "anonymous") {
            $isCompliant = $false
            $violations += "匿名共有禁止"
        }
        
        if ($Share.ShareType -eq "external" -and $Share.Scope -eq "organization") {
            $isCompliant = $false
            $violations += "組織外共有禁止"
        }
    }
    # 厳格ポリシー
    elseif ($PolicyName -eq "Strict") {
        if ($Share.ShareType -ne "internal") {
            $isCompliant = $false
            $violations += "内部共有のみ許可"
        }
        
        $item = Get-MgDriveItem -DriveId $Share.DriveId -DriveItemId $Share.ItemId -ErrorAction SilentlyContinue
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
function Invoke-Remediation {
    param(
        [object]$Share,
        [string]$PolicyName
    )
    
    $status = "未実施"
    $action = ""
    
    try {
        $compliance = Test-Compliance -Share $Share -PolicyName $PolicyName
        
        if (-not $compliance.IsCompliant) {
            # ポリシー違反の自動修復
            switch ($PolicyName) {
                "Default" {
                    if ($Share.ShareType -eq "anonymous") {
                        Remove-MgShare -ShareId $Share.Id -ErrorAction Stop
                        $status = "自動修復済"
                        $action = "匿名共有を削除"
                    }
                }
                "Strict" {
                    if ($Share.ShareType -ne "internal") {
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
            $status = "コンプライアンス準拠"
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
    <title>共有設定レポート</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .high { background-color: #ffcccc; }
        .medium { background-color: #ffe6cc; }
        .low { background-color: #e6ffe6; }
        .summary { margin-top: 30px; }
        .chart-container { width: 100%; height: 400px; margin-top: 30px; }
    </style>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
    <h1>共有設定レポート - $(Get-Date -Format 'yyyy/MM/dd')</h1>
    <p>生成日時: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')</p>
    <p>検出された共有設定数: $($Shares.Count)</p>
    
    <table>
        <tr>
            <th>共有元ユーザー</th>
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
            <td>$($share.共有元ユーザー)</td>
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
                        '#FF6384', '#FF9F40', '#4BC0C0'
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
Write-Log "共有設定確認を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"
Write-Log "適用ポリシー: $CompliancePolicy" "INFO"

# パフォーマンスデータ収集
$performanceData = @{
    StartTime = $executionTime
    ShareCount = 0
    HighRiskCount = 0
    ComplianceViolations = 0
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

# 出力用の共有設定リスト
$sharingList = @()

try {
    # 共有設定を取得 (リトライ付き)
    Write-Log "共有設定を確認しています..." "INFO"
    $allShares = Get-MgShare -All -ErrorAction Stop
    
    $performanceData.ShareCount = $allShares.Count
    
    foreach ($share in $allShares) {
        try {
            # 共有アイテムの詳細を取得
            $item = Get-MgDriveItem -DriveId $share.DriveId -DriveItemId $share.ItemId -ErrorAction SilentlyContinue
            
            # 共有元ユーザー情報を取得
            $owner = Get-MgUser -UserId $share.Owner.User.Id -ErrorAction SilentlyContinue
            
            # リスク評価
            $riskAssessment = @{
                Level = "低"
                Reasons = ""
            }
            
            if ($share.ShareType -eq "anonymous") {
                $riskAssessment.Level = "高"
                $riskAssessment.Reasons = "匿名共有"
            } elseif ($share.ShareType -eq "external") {
                $riskAssessment.Level = "中"
                $riskAssessment.Reasons = "外部共有"
            }
            
            # コンプライアンスチェック
            $compliance = Test-Compliance -Share $share -PolicyName $CompliancePolicy
            
            if (-not $compliance.IsCompliant) {
                $performanceData.ComplianceViolations++
            }
            
            # 自動修復
            $remediation = @{ Status = "未実施"; Action = "" }
            if ($AutoRemediate) {
                $remediation = Invoke-Remediation -Share $share -PolicyName $CompliancePolicy
                
                if ($riskAssessment.Level -eq "高" -and $remediation.Status -eq "自動修復済") {
                    $performanceData.HighRiskCount++
                }
            }
            
            $sharingList += [PSCustomObject]@{
                "共有元ユーザー" = $owner.DisplayName
                "メールアドレス" = $owner.Mail
                "アイテム名"     = $item.Name
                "アイテムタイプ" = if($item.Folder){"フォルダ"}else{"ファイル"}
                "共有タイプ"     = $share.ShareType
                "共有範囲"       = $share.Scope
                "共有相手"       = if($share.SharedWith) {$share.SharedWith.User.DisplayName} else {"N/A"}
                "リスクレベル"   = $riskAssessment.Level
                "リスク要因"     = $riskAssessment.Reasons
                "コンプライアンス" = if($compliance.IsCompliant){"準拠"}else{"違反: $($compliance.Violations)"}
                "推奨アクション" = if(-not $compliance.IsCompliant){"ポリシー違反: $($compliance.Violations)"}else{"不要"}
                "修復ステータス" = $remediation.Status
                "修復アクション" = $remediation.Action
                "共有日時"       = $share.SharedDateTime
                "WebURL"        = $item.WebUrl
                "権限ID"        = $share.Id
            }
            
            Write-Log "共有設定を検出: $($item.Name) (リスクレベル: $($riskAssessment.Level), コンプライアンス: $(if($compliance.IsCompliant){"準拠"}else{"違反"}))" "INFO"
        } catch {
            Write-Log "共有設定の詳細取得中にエラーが発生しました: $_" "WARNING"
        }
    }
    
    Write-Log "共有設定の確認が完了しました。検出された共有設定: $($sharingList.Count)" "SUCCESS"
} catch {
    Write-Log "共有設定の取得中にエラーが発生しました: $_" "ERROR"
    exit 1
}

# CSV出力
$csvPath = Join-Path -Path $OutputDir -ChildPath "SharingReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
try {
    $sharingList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"
} catch {
    Write-Log "CSVファイルの作成中にエラーが発生しました: $_" "ERROR"
}

# HTMLレポート生成
if ($GenerateHtmlReport) {
    $htmlPath = Join-Path -Path $OutputDir -ChildPath "SharingReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    Generate-HtmlReport -Shares $sharingList -OutputPath $htmlPath
}

# パフォーマンスデータ出力
if ($IncludePerformanceData) {
    $performanceData.EndTime = Get-Date
    $performanceData.Duration = ($performanceData.EndTime - $performanceData.StartTime).TotalSeconds
    
    $perfPath = Join-Path -Path $OutputDir -ChildPath "SharingPerformance_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    try {
        $performanceData | ConvertTo-Json -Depth 3 | Out-File -FilePath $perfPath -Encoding UTF8
        Write-Log "パフォーマンスデータを出力しました: $perfPath" "INFO"
    } catch {
        Write-Log "パフォーマンスデータの出力中にエラーが発生しました: $_" "WARNING"
    }
}

# Microsoft Graphから切断
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Log "共有設定確認を終了します" "INFO"
