# OneDrive for Business 運用ツール - ITSM準拠
# GetOneDriveQuota.ps1 - ストレージクォータ取得スクリプト

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
$reportDir = Join-Path -Path $dataCollectionDir -ChildPath "GetOneDriveQuota"

# ディレクトリ作成（自動生成フォルダを含む）
try {
    # 親フォルダが存在しない場合は作成
    if (-not (Test-Path -Path $dataCollectionDir)) {
        New-Item -Path $dataCollectionDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "Basic_Data_Collectionフォルダを作成しました: $dataCollectionDir" "INFO"
    }

    # レポート出力フォルダが存在しない場合は作成
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -Path $reportDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "GetOneDriveQuotaフォルダを作成しました: $reportDir" "INFO"
    }
} catch {
    Write-Log "ディレクトリ作成エラー: $_" "ERROR"
    throw
}

Write-Log "OneDriveクォータ情報取得を開始します" "INFO"
Write-Log "ベースディレクトリ: $BaseDir" "INFO"
Write-Log "出力ルートディレクトリ: $outputRootDir" "INFO"
Write-Log "レポートディレクトリ: $reportDir" "INFO"
Write-Host "出力先ディレクトリ: $outputRootDir" -ForegroundColor Cyan

# HTMLレポート生成関数
function Generate-HtmlReport {
    param(
        [Parameter(Mandatory=$true)]
        [array]$UserData,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputDir
    )
    
    try {
        $templatePath = Join-Path $PSScriptRoot "..\WebUI_Template\GetOneDriveQuota.html"
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
        $outputPath = Join-Path $reportDir "OneDriveQuota.$timestamp.html"
        $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8
        
        return $outputPath
    } catch {
        Write-ErrorLog $_ "HTMLレポート生成中にエラーが発生しました"
        throw
    }
}

# Microsoft Graph接続 (v2.27.0対応)
try {
    # config.jsonから認証情報を取得
    $config = Get-Content -Path $ConfigPath | ConvertFrom-Json
    
    # 非対話型認証で接続
    Write-Log "Microsoft Graphに接続中 (ClientSecret認証)..." "INFO"
    
    # 認証パラメータを最新仕様に合わせる
    $secureSecret = ConvertTo-SecureString $config.ClientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($config.ClientId, $secureSecret)
    
    # 既存の接続を解除
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    
    # 新しい接続を確立
    Connect-MgGraph -TenantId $config.TenantId -ClientSecretCredential $credential -ErrorAction Stop
    Write-Log "Microsoft Graphに正常に接続されました (アプリケーション認証)" "SUCCESS"
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続中にエラーが発生しました"
    Write-Log "エラーの詳細: $($_.Exception.Message)" "ERROR"
    Write-Log "スタックトレース: $($_.ScriptStackTrace)" "DEBUG"
    exit 1
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

# タイムスタンプをスクリプト開始時に生成
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# HTMLレポート生成
try {
    # 出力ファイルパスを先に生成
    $htmlFile = "OneDriveQuota.$timestamp.html"
    $htmlPath = Join-Path -Path $reportDir -ChildPath $htmlFile
    
    # ディレクトリ存在確認
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -Path $reportDir -ItemType Directory -Force | Out-Null
        Write-Log "レポートディレクトリを作成しました: $reportDir" "INFO"
    }

    # HTMLレポートの基本構造を生成
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
<title>OneDriveストレージクォータレポート</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<style>
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
.container { background-color: white; padding: 20px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
.table-container { overflow-x: auto; }
.header { background-color: #0078d4; color: white; padding: 15px; margin-bottom: 20px; border-radius: 5px; display: flex; align-items: center; }
.header-icon { font-size: 24px; margin-right: 10px; }
h1 { margin: 0; font-size: 24px; }
.info-section { background-color: #f0f0f0; padding: 10px; margin-bottom: 20px; border-radius: 5px; font-size: 14px; }
.info-label { font-weight: bold; margin-right: 5px; }
.toolbar { margin-bottom: 20px; display: flex; gap: 10px; align-items: center; }
#searchInput { padding: 8px; border: 1px solid #ddd; border-radius: 4px; flex-grow: 1; }
button { padding: 8px 12px; background-color: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; }
button:hover { background-color: #106ebe; }
button:disabled { background-color: #cccccc; cursor: not-allowed; }
table { width: 100%; table-layout: fixed; border-collapse: collapse; margin-bottom: 20px; }
th, td { padding: 12px 10px; text-align: left; border-bottom: 1px solid #ddd; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
th select { width: 100%; font-size: 14px; font-weight: normal; margin: 0; padding: 4px; }
.filter-row select { width: 100%; font-size: 14px; font-weight: normal; margin: 0; padding: 4px; }
#pagination { display: flex; justify-content: center; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 20px; }
.page-info { margin: 0 10px; }
.status-normal { background-color: #d4edda; }
.status-warning { background-color: #fff3cd; }
.status-danger { background-color: #f8d7da; }
@media print { .toolbar, button, #pagination, .filter-row { display: none; } .container { box-shadow: none; padding: 0; } }
</style>
</head>
<body>
<div class="container">
<div class="header"><div class="header-icon"><i class="fas fa-database"></i></div><h1>OneDriveストレージクォータレポート</h1></div>
<div class="info-section">
<p><span class="info-label">実行日時:</span> $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")</p>
<p><span class="info-label">出力フォルダ:</span> $reportDir</p>
</div>
<div class="toolbar">
<input type="text" id="searchInput" placeholder="検索...">
<button id="exportBtn">CSVエクスポート</button>
<button id="printBtn">印刷</button>
</div>
<div id="pagination"></div>
<table id="quotaTable">
<thead>
<tr>
<th>ユーザー名</th>
<th>メールアドレス</th>
<th>総容量(GB)</th>
<th>使用容量(GB)</th>
<th>残り容量(GB)</th>
<th>使用率(%)</th>
<th>状態</th>
</tr>
<tr class="filter-row">
<th><select class="column-filter"><option value="">すべて</option></select></th>
<th><select class="column-filter"><option value="">すべて</option></select></th>
<th><select class="column-filter"><option value="">すべて</option></select></th>
<th><select class="column-filter"><option value="">すべて</option></select></th>
<th><select class="column-filter"><option value="">すべて</option></select></th>
<th><select class="column-filter"><option value="">すべて</option></select></th>
<th><select class="column-filter"><option value="">すべて</option></select></th>
</tr>
</thead>
<tbody>
$($userList | ForEach-Object {
$statusClass = switch ($_."状態") {
"正常" { "status-normal" }
"警告" { "status-warning" }
"危険" { "status-danger" }
default { "" }
}
@"
<tr class="$statusClass">
<td>$($_."ユーザー名")</td>
<td>$($_."メールアドレス")</td>
<td>$($_."総容量(GB)")</td>
<td>$($_."使用容量(GB)")</td>
<td>$($_."残り容量(GB)")</td>
<td>$($_."使用率(%)")</td>
<td>$($_."状態")</td>
</tr>
"@
})
</tbody>
</table>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
// ユーザー情報レポートと同じJavaScript機能を実装
let currentPage = 1;
let rowsPerPage = 10;
let filteredRows = [];

function searchTable() {
    const input = document.getElementById('searchInput').value.toLowerCase();
    const table = document.getElementById('quotaTable');
    const rows = table.querySelectorAll('tbody tr');
    filteredRows = [];
    rows.forEach(row => {
        let found = false;
        row.querySelectorAll('td').forEach(cell => {
            if (cell.textContent.toLowerCase().includes(input)) found = true;
        });
        if (found) filteredRows.push(row);
    });
    currentPage = 1;
    updatePagination();
}

function updatePagination() {
    const startIndex = (currentPage - 1) * rowsPerPage;
    const endIndex = Math.min(startIndex + rowsPerPage, filteredRows.length);
    document.querySelectorAll('#quotaTable tbody tr').forEach(tr => tr.style.display = 'none');
    filteredRows.slice(startIndex, endIndex).forEach(tr => tr.style.display = '');
    updatePaginationControls();
}

function updatePaginationControls() {
    const paginationDiv = document.getElementById('pagination');
    const totalPages = Math.ceil(filteredRows.length / rowsPerPage);
    paginationDiv.innerHTML = \`
        <button \${currentPage === 1 ? 'disabled' : ''} onclick="changePage(\${currentPage - 1})">前へ</button>
        <span class="page-info">\${currentPage} / \${totalPages || 1}ページ</span>
        <button \${currentPage === totalPages || totalPages === 0 ? 'disabled' : ''} onclick="changePage(\${currentPage + 1})">次へ</button>
        <select onchange="changeRowsPerPage(this)">
            <option value="10">10件</option>
            <option value="20">20件</option>
            <option value="50">50件</option>
            <option value="100">100件</option>
        </select>
        <span>全 \${filteredRows.length} 件</span>
    \`;
}

function changePage(page) {
    currentPage = page;
    updatePagination();
}

function changeRowsPerPage(select) {
    rowsPerPage = parseInt(select.value);
    currentPage = 1;
    updatePagination();
}

function exportCSV() {
    const csv = [];
    const rows = document.querySelectorAll("#quotaTable tr");
    rows.forEach(row => {
        const rowData = [];
        row.querySelectorAll("td, th").forEach(cell => rowData.push(\`"\${cell.textContent.replace(/"/g, '""')}"\`));
        csv.push(rowData.join(","));
    });
    const blob = new Blob(["\uFEFF" + csv.join("\n")], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "OneDriveQuota_Export.csv";
    link.click();
}

function printTable() {
    window.print();
}

window.onload = function() {
    document.querySelectorAll('#quotaTable tbody tr').forEach(tr => filteredRows.push(tr));
    document.getElementById('searchInput').addEventListener('keyup', searchTable);
    document.getElementById('exportBtn').addEventListener('click', exportCSV);
    document.getElementById('printBtn').addEventListener('click', printTable);
    updatePagination();
};
</script>
</body>
</html>
"@

    # 出力ファイルパス
    $htmlFile = "OneDriveQuota.$timestamp.html"
    $htmlPath = Join-Path -Path $reportDir -ChildPath $htmlFile
    
    # HTMLファイルを出力
    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Log "HTMLレポートを生成しました: $htmlPath" "SUCCESS"
    
    return $htmlPath
} catch {
    Write-ErrorLog $_ "HTMLレポートの生成中にエラーが発生しました"
    throw
}

# CSV出力
$csvFile = "OneDriveQuota.$timestamp.csv"
$csvPath = Join-Path -Path $reportDir -ChildPath $csvFile

# CSV生成
$userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSVファイルが生成されました: $csvPath" "SUCCESS"
