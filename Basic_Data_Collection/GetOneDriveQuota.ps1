# OneDrive for Business 運用ツール - ITSM準拠
# GetOneDriveQuota.ps1 - ストレージクォータ取得スクリプト

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$LogDir = "$(Get-Location)\Log"
)

# ログ設定
$executionTime = Get-Date
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $LogDir -ChildPath "GetOneDriveQuota.$timestamp.log"
$errorLogPath = Join-Path -Path $LogDir -ChildPath "GetOneDriveQuota.Error.$timestamp.log"

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

# Microsoft Graph接続
try {
    # config.jsonから認証情報を取得
    $config = Get-Content -Path "config.json" | ConvertFrom-Json
    
    # 非対話型認証で接続
    Write-Log "Microsoft Graphに接続中 (ClientSecret認証)..." "INFO"
    try {
        # モジュールバージョンを確認
        $mgModule = Get-Module Microsoft.Graph -ListAvailable | Select-Object -First 1
        Write-Log "Microsoft.Graphモジュールバージョン: $($mgModule.Version)" "DEBUG"
        
        # PSCredentialオブジェクトを作成
        $secString = ConvertTo-SecureString $config.ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($config.ClientId, $secString)
        
        # 認証実行 (最新の認証方法)
        $body = @{
            client_id = $config.ClientId
            client_secret = $config.ClientSecret
            scope = "https://graph.microsoft.com/.default"
            grant_type = "client_credentials"
        }

        $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token" -Body $body -ErrorAction Stop
        $accessToken = ConvertTo-SecureString -String $tokenResponse.access_token -AsPlainText -Force

        Connect-MgGraph -AccessToken $accessToken -NoWelcome -ErrorAction Stop
        
        # スコープを確認
        $context = Get-MgContext
        if (-not $context.Scopes) {
            $context.Scopes = $config.DefaultScopes
        }
        
        Write-Log "Microsoft Graphに正常に接続されました" "SUCCESS"
    } catch {
        Write-Log "Microsoft Graph接続エラー: $($_.Exception.Message)" "ERROR"
        Write-Log "エラーの詳細:`n$($_.Exception|Format-List -Force|Out-String)" "DEBUG"
        exit 1
    }
    
    $context = Get-MgContext
    if (-not $context) {
        Write-Log "Microsoft Graphへの接続に失敗しました。" "ERROR"
        exit
    }
    
    Write-Log "Microsoft Graphに正常に接続されました (アプリケーション認証)" "SUCCESS"
} catch {
    Write-Log "Microsoft Graphの接続中にエラーが発生しました: $_" "ERROR"
    exit
}

# データ収集
$userList = @()
try {
    $allUsers = Get-MgUser -All -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType,UserPrincipalName,Id -ConsistencyLevel eventual -CountVariable totalCount
    $totalUsers = $allUsers.Count
    $processedUsers = 0

    foreach ($user in $allUsers) {
        $processedUsers++
        $percentComplete = [math]::Round(($processedUsers / $totalUsers) * 100, 2)
        Write-Progress -Activity "OneDriveクォータ情報を取得中" -Status "$processedUsers / $totalUsers ユーザー処理中 ($percentComplete%)" -PercentComplete $percentComplete
        
        try {
            $drive = Get-MgUserDrive -UserId $user.UserPrincipalName -ErrorAction Stop
            $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
            $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
            $remainingGB = [math]::Round(($drive.Quota.Remaining) / 1GB, 2)
            $usagePercent = [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 2)
            
            $status = if ($usagePercent -ge 90) { "危険" } elseif ($usagePercent -ge 70) { "警告" } else { "正常" }
            $onedriveStatus = "対応"
        } catch {
            $totalGB = $usedGB = $remainingGB = $usagePercent = "取得不可"
            $status = "不明"
            $onedriveStatus = "未対応"
        }

        $userList += [PSCustomObject]@{
            "ユーザー名" = $user.DisplayName
            "メールアドレス" = $user.Mail
            "総容量(GB)" = $totalGB
            "使用容量(GB)" = $usedGB
            "残り容量(GB)" = $remainingGB
            "使用率(%)" = $usagePercent
            "状態" = $status
        }
    }
} catch {
    Write-Log "OneDriveクォータ情報の取得中にエラーが発生しました: $_" "ERROR"
    exit
}

# レポート出力先設定
$dateFolder = "OneDriveManagement." + (Get-Date -Format "yyyyMMdd")
$reportFolder = Join-Path $dateFolder "Report"

# フォルダが存在しない場合は作成
if (-not (Test-Path $reportFolder)) {
    New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
}

# HTML生成
$htmlFile = "OneDriveQuota.$timestamp.html"
$htmlPath = Join-Path -Path $reportFolder -ChildPath $htmlFile

# CSV出力先をReportフォルダに修正
$csvFile = "OneDriveQuota.$timestamp.csv"
$csvPath = Join-Path -Path $reportFolder -ChildPath $csvFile

$html = @"
<!DOCTYPE html>
<html>
<head>
<title>ストレージクォータレポート</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<style>
body {
    padding: 20px;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}
.table thead th {
    position: sticky;
    top: 0;
    background-color: #f8f9fa;
    z-index: 10;
}
:root {
  --user-color: #64b5f6;
  --email-color: #ffd54f;
  --total-color: #4dd0e1;
  --used-color: #ff8a65;
  --remaining-color: #81c784;
  --usage-color: #ba68c8;
  --status-color: #ffb74d;
}

.column-user {
  background: linear-gradient(135deg, var(--user-color), #42a5f5);
  color: white;
}
.column-email {
  background: linear-gradient(135deg, var(--email-color), #ffca28);
  color: #333;
}
.column-total {
  background: linear-gradient(135deg, var(--total-color), #26c6da);
  color: white;
}
.column-used {
  background: linear-gradient(135deg, var(--used-color), #ff7043);
  color: white;
}
.column-remaining {
  background: linear-gradient(135deg, var(--remaining-color), #43a047);
  color: white;
}
.column-usage {
  background: linear-gradient(135deg, var(--usage-color), #9c27b0);
  color: white;
}
.column-status {
  background: linear-gradient(135deg, var(--status-color), #ff8a65);
  color: white;
}

.status-warning {
  background: linear-gradient(135deg, #ffd740, #ffa000);
  color: #333;
  font-weight: bold;
}
.status-danger {
  background: linear-gradient(135deg, #ff5252, #c62828);
  color: white;
  font-weight: bold;
}
.status-normal {
  background: linear-gradient(135deg, #69f0ae, #00c853);
  color: white;
  font-weight: bold;
}

.table-hover tbody tr:hover {
  transform: scale(1.02);
  box-shadow: 0 4px 12px rgba(0,0,0,0.15);
  transition: all 0.3s ease;
}
.filter-row select {
    width: 100%;
    padding: 8px;
    background-color: white;
    border: 2px solid #dee2e6;
    border-radius: 4px;
    transition: all 0.3s ease;
}
.filter-container {
    background-color: #f5f5f5;
    padding: 10px;
    border-radius: 5px;
    margin-bottom: 15px;
}
.progress {
    height: 20px;
    border-radius: 10px;
}
.progress-bar {
    transition: width 0.6s ease;
}
</style>
</head>
<body>
<div class="container">
    <h1 class="text-center mb-4">
      <i class="fas fa-database me-3 text-primary"></i>
      ストレージクォータレポート
      <i class="fas fa-hdd ms-3 text-primary"></i>
    </h1>
    
    <div class="filter-container">
        <div class="row mb-3">
            <div class="col-md-4">
                <div class="input-group">
                    <span class="input-group-text"><i class="fas fa-search"></i></span>
                    <input type="text" class="form-control" id="globalSearch" placeholder="ユーザー名またはメールアドレスで検索...">
                </div>
            </div>
            <div class="col-md-4">
                <select class="form-select" id="statusFilter">
                    <option value="">すべてのステータス</option>
                    <option value="warning">警告のみ</option>
                    <option value="danger">危険のみ</option>
                </select>
            </div>
            <div class="col-md-4 text-end">
                <span id="displayCount" class="me-3 fw-bold"></span>
                <button class="btn btn-primary me-2" onclick="exportToCsv()">
                    <i class="fas fa-file-export me-1"></i>CSVエクスポート
                </button>
                <button class="btn btn-secondary" onclick="window.print()">
                    <i class="fas fa-print me-1"></i>印刷
                </button>
            </div>
        </div>
    </div>

    <div class="table-responsive">
        <table class="table table-hover" id="quotaTable">
            <thead>
                <tr>
                    <th class="column-user">
                      <i class="fas fa-user-circle me-2"></i>
                      <span class="badge bg-white text-dark">ユーザー名</span>
                    </th>
                    <th class="column-email">
                      <i class="fas fa-envelope me-2"></i>
                      <span class="badge bg-white text-dark">メールアドレス</span>
                    </th>
                    <th class="column-total">
                      <i class="fas fa-database me-2"></i>
                      <span class="badge bg-white text-dark">総容量(GB)</span>
                    </th>
                    <th class="column-used">
                      <i class="fas fa-hdd me-2"></i>
                      <span class="badge bg-white text-dark">使用容量(GB)</span>
                    </th>
                    <th class="column-remaining">
                      <i class="fas fa-sd-card me-2"></i>
                      <span class="badge bg-white text-dark">残り容量(GB)</span>
                    </th>
                    <th class="column-usage">
                      <i class="fas fa-percentage me-2"></i>
                      <span class="badge bg-white text-dark">使用率(%)</span>
                    </th>
                    <th class="column-status">
                      <i class="fas fa-exclamation-circle me-2"></i>
                      <span class="badge bg-white text-dark">状態</span>
                    </th>
                </tr>
                <tr class="filter-row">
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                </tr>
            </thead>
            <tbody>
"@

# データ行を追加
foreach ($user in $userList) {
    $statusClass = switch ($user.状態) {
        "危険" { "status-danger" }
        "警告" { "status-warning" }
        default { "status-normal" }
    }
    
    $progressBarClass = switch ($user.状態) {
        "危険" { "bg-danger" }
        "警告" { "bg-warning" }
        default { "bg-success" }
    }
    
    $html += @"
                <tr class="$statusClass">
                    <td>$($user.'ユーザー名')</td>
                    <td>$($user.'メールアドレス')</td>
                    <td>$($user.'総容量(GB)')</td>
                    <td>$($user.'使用容量(GB)')</td>
                    <td>$($user.'残り容量(GB)')</td>
                    <td>
                        <div class="progress">
                            <div class="progress-bar $progressBarClass" role="progressbar" style="width: $($user.'使用率(%)')%"></div>
                        </div>
                    </td>
                    <td>$($user.'状態')</td>
                </tr>
"@
}

# HTMLフッター
$html += @"
            </tbody>
        </table>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
// データを明確に定義
const quotaData = $($userList | ConvertTo-Json -Depth 10);
const table = document.getElementById('quotaTable');

// テーブルを更新する関数
function updateTable(data) {
    const tbody = table.querySelector('tbody');
    tbody.innerHTML = '';
    
    data.forEach(user => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${user.ユーザー名 || ''}</td>
            <td>${user.メールアドレス || ''}</td>
            <td>${user.総容量GB || ''}</td>
            <td>${user.使用容量GB || ''}</td>
            <td>${user.残り容量GB || ''}</td>
            <td>
                <div class="progress">
                    <div class="progress-bar"
                         style="width: ${user.使用率 || 0}%"
                         aria-valuenow="${user.使用率 || 0}"
                         aria-valuemin="0"
                         aria-valuemax="100">
                        ${user.使用率 || 0}%
                    </div>
                </div>
            </td>
            <td>${user.状態 || ''}</td>
        `;
        tbody.appendChild(row);
    });
    
    document.getElementById('displayCount').textContent = `表示中: ${data.length}件 / 全${quotaData.length}件`;
}

function updateDisplayCount() {
    const visibleCount = document.querySelectorAll('#quotaTable tbody tr:not([style*="display: none"])').length;
    document.getElementById('displayCount').textContent = `表示中: ${visibleCount}件 / 全${quotaData.length}件`;
}

// カラムフィルター初期化関数
function initFilters() {
    const selects = document.querySelectorAll('#quotaTable thead tr.filter-row select');
    
    selects.forEach((select, index) => {
        select.innerHTML = '<option value="">すべて表示</option>';
        const uniqueValues = new Set();
        
        quotaData.forEach(user => {
            const value = Object.values(user)[index];
            if (value) uniqueValues.add(value.toString());
        });
        
        Array.from(uniqueValues).sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        select.addEventListener('change', filterTable);
    });
}

// 初期化関数
function initializeTable() {
    // ステータスフィルターの初期化
    const statusFilter = document.getElementById('statusFilter');
    statusFilter.innerHTML = `
        <option value="">すべてのステータス</option>
        <option value="正常">正常</option>
        <option value="警告">警告</option>
        <option value="危険">危険</option>
    `;
    
    // イベントリスナーの設定
    document.getElementById('globalSearch').addEventListener('input', filterTable);
    statusFilter.addEventListener('change', filterTable);
    
    // 初期フィルタリング
    initFilters();
    filterTable();
}

// フィルタリング関数
function filterTable() {
    const searchTerm = document.getElementById('globalSearch').value.toLowerCase();
    const statusFilter = document.getElementById('statusFilter').value;
    const columnFilters = Array.from(document.querySelectorAll('#quotaTable thead tr.filter-row select'))
        .map(select => select.value);
    
    let filteredData = quotaData;
    
    // ステータスでフィルタ
    if (statusFilter) {
        filteredData = filteredData.filter(user =>
            user.状態 === statusFilter
        );
    }
    
    // 検索でフィルタ
    if (searchTerm) {
        filteredData = filteredData.filter(user =>
            (user.ユーザー名 && user.ユーザー名.toLowerCase().includes(searchTerm)) ||
            (user.メールアドレス && user.メールアドレス.toLowerCase().includes(searchTerm))
        );
    }
    
    // カラムフィルタ
    filteredData = filteredData.filter(user => {
        return columnFilters.every((filter, index) => {
            if (!filter) return true;
            const value = Object.values(user)[index];
            return value && value.toString() === filter;
        });
    });
        );
    }
    
    updateTable(filteredData);
}

// 初期化処理
document.addEventListener('DOMContentLoaded', () => {
    // ステータスフィルターの初期化
    const statusFilter = document.getElementById('statusFilter');
    statusFilter.innerHTML = `
        <option value="">すべてのステータス</option>
        <option value="正常">正常</option>
        <option value="警告">警告</option>
        <option value="危険">危険</option>
    `;
    
    // イベントリスナーの設定
    document.getElementById('globalSearch').addEventListener('input', filterTable);
    statusFilter.addEventListener('change', filterTable);
    
    // 初期表示
    updateTable(quotaData);
});
                (statusFilter === 'danger' && status === '危険') ||
                (statusFilter === 'normal' && status === '正常');
            
            // カラムフィルタ
            let matchesColumns = true;
            cells.forEach((cell, index) => {
                if (columnFilters[index] && cell.textContent.trim() !== columnFilters[index]) {
                    matchesColumns = false;
                }
            });
            
            if (matchesSearch && matchesStatus && matchesColumns) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        });
        
        updateDisplayCount();
    } catch (error) {
        console.error('フィルタリングエラー:', error);
    }
}

// ページ読み込み時に初期化
document.addEventListener('DOMContentLoaded', initializeTable);

function exportToCsv() {
    const headers = ['ユーザー名', 'メールアドレス', '総容量(GB)', '使用容量(GB)', '残り容量(GB)', '使用率(%)', '状態'];
    const csvContent = [
        headers.join(','),
        ...quotaData.map(item =>
            headers.map(header =>
                `"${(item[header] || '').toString().replace(/"/g, '""')}"`
            ).join(',')
        )
    ].join('\n');
    
    const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `OneDriveQuotaReport_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 初期化関数
function initializeTable() {
    console.log("Initializing table filters...");
    
    // ステータスフィルターの初期化
    const statusFilter = document.getElementById('statusFilter');
    if (statusFilter) {
        statusFilter.innerHTML = `
            <option value="">すべてのステータス</option>
            <option value="normal">正常</option>
            <option value="warning">警告</option>
            <option value="danger">危険</option>
        `;
        statusFilter.addEventListener('change', filterTable);
    }

    // グローバル検索の初期化
    const globalSearch = document.getElementById('globalSearch');
    if (globalSearch) {
        globalSearch.addEventListener('keyup', filterTable);
    }

    // カラムフィルターの初期化
    initFilters();
    filterTable();
    
    console.log("Table initialization completed");
}

// DOM完全読み込み後に初期化
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeTable);
} else {
    initializeTable();
}
</script>
</body>
</html>
"@

# HTMLファイルを保存
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Log "HTMLレポートが生成されました: $htmlPath" "SUCCESS"

# CSV生成
$userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSVファイルが生成されました: $csvPath" "SUCCESS"
