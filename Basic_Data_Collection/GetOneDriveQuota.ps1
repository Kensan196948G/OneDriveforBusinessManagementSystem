# OneDrive for Business 運用ツール - ITSM準拠
# GetOneDriveQuota.ps1 - ストレージクォータ取得スクリプト

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$LogDir = "$(Get-Location)\Log"
)

# 実行開始時刻を記録
$executionTime = Get-Date

# ログファイルのパスを設定
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $LogDir -ChildPath "GetOneDriveQuota.$timestamp.log"
$errorLogPath = Join-Path -Path $LogDir -ChildPath "GetOneDriveQuota.Error.$timestamp.log"

# ログ関数
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # コンソールに出力
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
    
    # ログファイルに出力
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
    
    # コンソールに出力
    Write-Host $errorMessage -ForegroundColor Red
    
    # ログファイルに出力
    Add-Content -Path $logFilePath -Value $errorMessage -Encoding UTF8
    
    # エラーログに詳細を出力
    Add-Content -Path $errorLogPath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorDetails -Encoding UTF8
}

# ストレージクォータ取得開始
Write-Log "ストレージクォータ取得を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"
Write-Log "ログディレクトリ: $LogDir" "INFO"

# Microsoft Graphの接続確認
try {
    # TEMPフォルダのトークンファイルを読み込み
    $tempDir = Join-Path -Path $PSScriptRoot -ChildPath "..\TEMP"
    $tokenFile = Join-Path -Path $tempDir -ChildPath "graph_token.txt"
    if (Test-Path $tokenFile) {
        $global:AccessToken = Get-Content -Path $tokenFile -Raw
    }
    if (-not $global:AccessToken) {
        Write-Log "Microsoft Graphに接続されていません。Main.ps1から実行してください。" "ERROR"
        exit
    }
    
    # ログインユーザーのUPN（メールアドレス）を自動取得
    $context = Get-MgContext
    $UserUPN = if ($context) { $context.Account } else { "" }

    # UPNが空なら -Me で取得
    if (-not $UserUPN) {
        try {
            $meUser = Get-MgUser -Me -Property UserPrincipalName,DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType
            $UserUPN = $meUser.UserPrincipalName
            $currentUser = $meUser
            Write-Log "Get-MgUser -Me でユーザー情報を取得しました: $($UserUPN)" "INFO"
        } catch {
            Write-Log "Graph接続済みですが、ユーザー情報は取得できません。" "WARNING"
        }
    }

    if ($UserUPN) {
        if (-not $currentUser) {
            # ログイン済ユーザー情報を取得
            $currentUser = Get-MgUser -UserId $UserUPN -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType
        }
        
        # グローバル管理者かどうかを判定
        $isAdmin = ($context.Scopes -contains "Directory.ReadWrite.All")
        
        Write-Log "ログインユーザー: $($currentUser.DisplayName) ($UserUPN)" "INFO"
        Write-Log "実行モード: $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})" "INFO"
    } else {
        Write-Log "Graph接続済みですが、ユーザー情報は取得できません。" "WARNING"
    }
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続確認中にエラーが発生しました"
    exit
}

# 出力用のユーザーリスト
$userList = @()

# 管理者ロールID辞書
$adminRoleIds = @{
    "GlobalAdministrator" = "62e90394-69f5-4237-9190-012177145e10"
    "UserAccountAdministrator" = "fe930be7-5e62-47db-91af-98c3a49a38b1"
    "ExchangeAdministrator" = "29232cdf-9323-42fd-ade2-1d097af3e4de"
    "SharePointAdministrator" = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
    "TeamsAdministrator" = "69091246-20e8-4a56-aa4d-066075b2a7a8"
    "SecurityAdministrator" = "194ae4cb-b126-40b2-bd5b-6091b380977d"
}

try {
    # 常に管理者モードで全ユーザーのOneDrive情報を取得
    Write-Log "すべてのユーザーのOneDriveクォータ情報を取得しています..." "INFO"
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
            
            # 使用率に基づいて状態を設定
            $status = if ($usagePercent -ge 90) { "危険" } elseif ($usagePercent -ge 70) { "警告" } else { "正常" }
            $onedriveStatus = "対応"
        } catch {
            $totalGB = "取得不可"
            $usedGB = "取得不可"
            $remainingGB = "取得不可"
            $usagePercent = "取得不可"
            $status = "不明"
            $onedriveStatus = "未対応"
            
            Write-Log "ユーザー $($user.UserPrincipalName) のOneDriveクォータ情報を取得できませんでした: $_" "WARNING"
        }
        
        # 管理者判定 (REST API版)
        $isAdminUser = $false
        try {
            $headers = @{Authorization = "Bearer $global:AccessToken"}
            $memberOfUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/memberOf"
            $memberOfResponse = Invoke-RestMethod -Headers $headers -Uri $memberOfUrl -Method Get
            $memberOf = $memberOfResponse.value
            foreach ($role in $memberOf) {
                if ($role.'@odata.type' -eq "#microsoft.graph.directoryRole") {
                    $roleDetailUrl = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)"
                    $roleDetail = Invoke-RestMethod -Headers $headers -Uri $roleDetailUrl -Method Get
                    if ($adminRoleIds.Values -contains $roleDetail.roleTemplateId) {
                        $isAdminUser = $true
                        break
                    }
                }
            }
        } catch {
            # 無視
        }

        $userTypeValue = if ($isAdminUser) { "Administrator" } elseif ($user.UserType -eq "Guest") { "Guest" } else { "Member" }

        $userList += [PSCustomObject]@{
            "ユーザー名"       = $user.DisplayName
            "メールアドレス"   = $user.Mail
            "ログインユーザー名" = if($user.onPremisesSamAccountName){$user.onPremisesSamAccountName}else{"同期なし"}
            "ユーザー種別"   = $userTypeValue
            "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
            "OneDrive対応"    = $onedriveStatus
            "総容量(GB)"   = $totalGB
            "使用容量(GB)"   = $usedGB
            "残り容量(GB)"   = $remainingGB
            "使用率(%)"     = $usagePercent
            "状態"          = $status
        }
    }
    
    Write-Progress -Activity "OneDriveクォータ情報を取得中" -Completed
    Write-Log "OneDriveクォータ情報の取得が完了しました。取得件数: $($userList.Count)" "SUCCESS"
} catch {
    Write-ErrorLog $_ "OneDriveクォータ情報の取得中にエラーが発生しました"
}

# タイムスタンプ
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# 出力ファイル名の設定
$csvFile = "OneDriveQuota.$timestamp.csv"
$htmlFile = "OneDriveQuota.$timestamp.html"
$jsFile = "OneDriveQuota.$timestamp.js"

# 出力パスの設定
$csvPath = Join-Path -Path $OutputDir -ChildPath $csvFile
$htmlPath = Join-Path -Path $OutputDir -ChildPath $htmlFile
$jsPath = Join-Path -Path $OutputDir -ChildPath $jsFile

# CSV出力（文字化け対策済み）
try {
    # PowerShell Core (バージョン 6.0以上)の場合
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    }
    # PowerShell 5.1以下の場合
    else {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        # BOMを追加して文字化け対策
        $content = [System.IO.File]::ReadAllText($csvPath)
        [System.IO.File]::WriteAllText($csvPath, $content, [System.Text.Encoding]::UTF8)
    }
    Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"
    
    # CSVファイルをExcelで開き、列幅の調整とフィルターの適用を行う
    try {
        Write-Log "Excelでファイルを開いて列幅の調整とフィルターの適用を行います..." "INFO"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Open($csvPath)
        $worksheet = $workbook.Worksheets.Item(1)
        
        # 列幅の自動調整
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        
        # フィルターの適用
        $usedRange.AutoFilter() | Out-Null
        
        # ウィンドウを最前面に表示
        $excel.ActiveWindow.WindowState = -4143 # xlMaximized
        
        # 変更を保存
        $workbook.Save()
        
        Write-Log "Excelでの処理が完了しました。" "SUCCESS"
    }
    catch {
        Write-Log "Excelでの処理中にエラーが発生しました: $_" "WARNING"
        Write-Log "CSVファイルは正常に作成されましたが、Excel処理はスキップされました。" "WARNING"
    }
} catch {
    Write-ErrorLog $_ "CSVファイルの作成中にエラーが発生しました"
}

# JavaScript ファイルの生成
$jsContent = @"
// OneDriveQuota データ操作用 JavaScript

// グローバル変数
let currentPage = 1;
let rowsPerPage = 10;
let filteredRows = [];

// テーブルを検索する関数
function searchTable() {
    var input = document.getElementById('searchInput').value.toLowerCase();
    var table = document.getElementById('quotaTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    filteredRows = [];
    for (var i = 0; i < rows.length; i++) {
        var found = false;
        var cells = rows[i].getElementsByTagName('td');
        var rowData = {};
        for (var j = 0; j < cells.length; j++) {
            var cellText = cells[j].textContent || cells[j].innerText;
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[j].textContent;
            rowData[headerText] = cellText;
            if (cellText.toLowerCase().indexOf(input) > -1) {
                found = true;
            }
        }
        if (found) {
            filteredRows.push({row: rows[i], data: rowData});
        }
    }
    showSearchSuggestions(input);
    currentPage = 1;
    updatePagination();
}

// 検索候補を表示
function showSearchSuggestions(input) {
    var suggestionsDiv = document.getElementById('searchSuggestions');
    suggestionsDiv.innerHTML = '';
    if (input.length < 1) {
        suggestionsDiv.style.display = 'none';
        return;
    }
    var matches = new Set();
    filteredRows.forEach(item => {
        Object.values(item.data).forEach(value => {
            if (value.toLowerCase().indexOf(input.toLowerCase()) > -1) {
                matches.add(value);
            }
        });
    });
    var count = 0;
    matches.forEach(match => {
        if (count < 5) {
            var div = document.createElement('div');
            div.className = 'suggestion-item';
            div.textContent = match;
            div.onclick = function() {
                document.getElementById('searchInput').value = match;
                searchTable();
                suggestionsDiv.style.display = 'none';
            };
            suggestionsDiv.appendChild(div);
            count++;
        }
    });
    if (count > 0) {
        suggestionsDiv.style.display = 'block';
    } else {
        var noResults = document.createElement('div');
        noResults.className = 'suggestion-item no-results';
        noResults.textContent = '検索結果がありません';
        suggestionsDiv.appendChild(noResults);
        suggestionsDiv.style.display = 'block';
    }
}

// 列フィルターを作成
function createColumnFilters() {
    var table = document.getElementById('quotaTable');
    var headers = table.getElementsByTagName('thead')[0].getElementsByTagName('th');
    var filterRow = document.createElement('tr');
    filterRow.className = 'filter-row';
    for (var i = 0; i < headers.length; i++) {
        var cell = document.createElement('th');
        var select = document.createElement('select');
        select.className = 'column-filter';
        select.setAttribute('data-column', i);
        var defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'すべて';
        select.appendChild(defaultOption);
        var uniqueValues = new Set();
        var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
        for (var j = 0; j < rows.length; j++) {
            var cellValue = rows[j].getElementsByTagName('td')[i].textContent;
            uniqueValues.add(cellValue);
        }
        Array.from(uniqueValues).sort().forEach(value => {
            var option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        select.addEventListener('change', applyColumnFilters);
        cell.appendChild(select);
        filterRow.appendChild(cell);
    }
    table.getElementsByTagName('thead')[0].appendChild(filterRow);
}

// 列フィルターを適用
function applyColumnFilters() {
    var table = document.getElementById('quotaTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var filters = table.querySelectorAll('.column-filter');
    filteredRows = [];
    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var cells = row.getElementsByTagName('td');
        var rowData = {};
        var includeRow = true;
        for (var j = 0; j < filters.length; j++) {
            var filter = filters[j];
            var columnIndex = parseInt(filter.getAttribute('data-column'));
            var filterValue = filter.value;
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[columnIndex].textContent;
            var cellValue = cells[columnIndex].textContent;
            rowData[headerText] = cellValue;
            if (filterValue && cellValue !== filterValue) {
                includeRow = false;
                break;
            }
        }
        if (includeRow) {
            filteredRows.push({row: row, data: rowData});
        }
    }
    var searchInput = document.getElementById('searchInput').value.toLowerCase();
    if (searchInput) {
        filteredRows = filteredRows.filter(item => {
            return Object.values(item.data).some(value => value.toLowerCase().indexOf(searchInput) > -1);
        });
    }
    currentPage = 1;
    updatePagination();
}

// ページング更新
function updatePagination() {
    var table = document.getElementById('quotaTable');
    var tbody = table.getElementsByTagName('tbody')[0];
    var rows = tbody.getElementsByTagName('tr');
    for (var i = 0; i < rows.length; i++) {
        rows[i].style.display = 'none';
    }
    var startIndex = (currentPage - 1) * rowsPerPage;
    var endIndex = Math.min(startIndex + rowsPerPage, filteredRows.length);
    for (var i = startIndex; i < endIndex; i++) {
        filteredRows[i].row.style.display = '';
    }
    updatePaginationControls();
}

// ページネーションコントロール更新
function updatePaginationControls() {
    var paginationDiv = document.getElementById('pagination');
    paginationDiv.innerHTML = '';
    var totalPages = Math.ceil(filteredRows.length / rowsPerPage);
    var prevButton = document.createElement('button');
    prevButton.innerHTML = '<span class="button-icon">◀</span>前へ';
    prevButton.disabled = currentPage === 1;
    prevButton.addEventListener('click', function() {
        if (currentPage > 1) {
            currentPage--;
            updatePagination();
        }
    });
    paginationDiv.appendChild(prevButton);
    var pageInfo = document.createElement('span');
    pageInfo.className = 'page-info';
    pageInfo.textContent = currentPage + ' / ' + (totalPages || 1) + ' ページ';
    paginationDiv.appendChild(pageInfo);
    var nextButton = document.createElement('button');
    nextButton.innerHTML = '次へ<span class="button-icon">▶</span>';
    nextButton.disabled = currentPage === totalPages || totalPages === 0;
    nextButton.addEventListener('click', function() {
        if (currentPage < totalPages) {
            currentPage++;
            updatePagination();
        }
    });
    paginationDiv.appendChild(nextButton);
    var rowsPerPageDiv = document.createElement('div');
    rowsPerPageDiv.className = 'rows-per-page';
    var rowsPerPageLabel = document.createElement('span');
    rowsPerPageLabel.textContent = '表示件数: ';
    rowsPerPageDiv.appendChild(rowsPerPageLabel);
    var rowsPerPageSelect = document.createElement('select');
    [10, 20, 50, 100].forEach(function(value) {
        var option = document.createElement('option');
        option.value = value;
        option.textContent = value + '件';
        if (value === rowsPerPage) {
            option.selected = true;
        }
        rowsPerPageSelect.appendChild(option);
    });
    rowsPerPageSelect.addEventListener('change', function() {
        rowsPerPage = parseInt(this.value);
        currentPage = 1;
        updatePagination();
    });
    rowsPerPageDiv.appendChild(rowsPerPageSelect);
    paginationDiv.appendChild(rowsPerPageDiv);
    var totalItems = document.createElement('span');
    totalItems.className = 'total-items';
    totalItems.textContent = '全 ' + filteredRows.length + ' 件';
    paginationDiv.appendChild(totalItems);
}

// 初期化
window.onload = function() {
    createColumnFilters();
    var table = document.getElementById('quotaTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    filteredRows = [];
    for (var i = 0; i < rows.length; i++) {
        var cells = rows[i].getElementsByTagName('td');
        var rowData = {};
        for (var j = 0; j < cells.length; j++) {
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[j].textContent;
            rowData[headerText] = cells[j].textContent;
        }
        filteredRows.push({row: rows[i], data: rowData});
    }
    updatePagination();
    document.getElementById('searchInput').addEventListener('keyup', searchTable);
};

// CSVとしてエクスポートする関数 (文字化け対策済み)
function exportTableToCSV() {
    var table = document.getElementById('quotaTable');
    var headerRow = table.getElementsByTagName('thead')[0].getElementsByTagName('tr')[0]; // ヘッダー行（1行目）のみ
    var bodyRows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var csv = [];
    
    // ヘッダー行を処理
    var headerCols = headerRow.getElementsByTagName('th');
    var headerData = [];
    for (var i = 0; i < headerCols.length; i++) {
        var data = headerCols[i].innerText.replace(/(\r\n|\n|\r)/gm, ' ').replace(/"/g, '""');
        headerData.push('"' + data + '"');
    }
    csv.push(headerData.join(','));
    
    // データ行を処理（フィルター行は除外）
    for (var i = 0; i < bodyRows.length; i++) {
        var row = [], cols = bodyRows[i].getElementsByTagName('td');
        for (var j = 0; j < cols.length; j++) {
            // セル内のテキストから改行や引用符を適切に処理
            var data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, ' ').replace(/"/g, '""');
            row.push('"' + data.trim() + '"');
        }
        csv.push(row.join(','));
    }
    
    // CSVファイルのダウンロード（UTF-8 BOM付きで文字化け対策）
    var csvContent = '\uFEFF' + csv.join('\n'); // BOMを追加
    var csvFile = new Blob([csvContent], {type: 'text/csv;charset=utf-8'});
    var downloadLink = document.createElement('a');
    downloadLink.download = 'OneDriveQuota_Export.csv';
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = 'none';
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

// 印刷機能
function printTable() {
    window.print();
}

// 表の行に色を付ける
function colorizeRows() {
    var table = document.getElementById('quotaTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    
    for (var i = 0; i < rows.length; i++) {
        var usageCell = rows[i].querySelector('td:nth-child(9)'); // 使用率のセル
        var statusCell = rows[i].querySelector('td:nth-child(10)'); // 状態のセル
        
        if (usageCell && statusCell) {
            var usage = parseFloat(usageCell.textContent);
            var status = statusCell.textContent;
            
            if (!isNaN(usage)) {
                if (usage >= 90 || status === '危険') {
                    rows[i].classList.add('danger');
                } else if (usage >= 70 || status === '警告') {
                    rows[i].classList.add('warning');
                } else {
                    rows[i].classList.add('normal');
                }
            }
        }
        
        // アカウント状態によっても色分け
        var accountStatus = rows[i].querySelector('td:nth-child(5)').textContent; // アカウント状態のセル
        if (accountStatus === '無効') {
            rows[i].classList.add('disabled');
        }
    }
}

// ページロード時に実行
window.onload = function() {
    colorizeRows();
    createColumnFilters();
    
    // 検索イベントリスナーを設定
    document.getElementById('searchInput').addEventListener('keyup', function(e) {
        // リアルタイムで検索を実行（インクリメンタル検索）
        searchTable();
    });
    document.getElementById('searchInput').addEventListener('blur', hideSearchSuggestions);
    
    // エクスポートボタンにイベントリスナーを設定
    document.getElementById('exportBtn').addEventListener('click', exportTableToCSV);
    
    // 印刷ボタンにイベントリスナーを設定
    document.getElementById('printBtn').addEventListener('click', printTable);
    
    // 初期ページングの設定
    var table = document.getElementById('quotaTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    for (var i = 0; i < rows.length; i++) {
        var cells = rows[i].getElementsByTagName('td');
        var rowData = {};
        
        for (var j = 0; j < cells.length; j++) {
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[j].textContent;
            rowData[headerText] = cells[j].textContent;
        }
        
        filteredRows.push({row: rows[i], data: rowData});
    }
    
    updatePagination();
};
"@

# JavaScript ファイルを出力
$jsContent | Out-File -FilePath $jsPath -Encoding UTF8
Write-Log "JavaScriptファイルを作成しました: $jsPath" "SUCCESS"

# 実行日時とユーザー情報を取得
$executionDateFormatted = $executionTime.ToString("yyyy/MM/dd HH:mm:ss")
$executorName = $currentUser.DisplayName
$userType = if($currentUser.UserType){$currentUser.UserType}else{"未定義"}

# HTML ファイルの生成
$htmlContent = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>OneDriveクォータレポート</title>
    <script src="$jsFile"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .header {
            background-color: #0078d4;
            color: white;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 5px;
            display: flex;
            align-items: center;
        }
        .header-icon {
            font-size: 24px;
            margin-right: 10px;
        }
        h1 {
            margin: 0;
            font-size: 24px;
        }
        .info-section {
            background-color: #f0f0f0;
            padding: 10px;
            margin-bottom: 20px;
            border-radius: 5px;
            font-size: 14px;
        }
        .info-label {
            font-weight: bold;
            margin-right: 5px;
        }
        .toolbar {
            margin-bottom: 20px;
            display: flex;
            gap: 10px;
            align-items: center;
            position: relative;
        }
        #searchInput {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            flex-grow: 1;
        }
        #searchSuggestions {
            position: absolute;
            top: 100%;
            left: 0;
            width: 100%;
            max-height: 200px;
            overflow-y: auto;
            background-color: white;
            border: 1px solid #ddd;
            border-radius: 0 0 4px 4px;
            z-index: 1000;
            display: none;
        }
        .suggestion-item {
            padding: 8px;
            border-bottom: 1px solid #eee;
            cursor: pointer;
        }
        .suggestion-item:hover {
            background-color: #f0f0f0;
        }
        .suggestion-item.no-results {
            color: #999;
            font-style: italic;
            cursor: default;
        }
        button {
            padding: 8px 12px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            display: flex;
            align-items: center;
        }
        button:hover {
            background-color: #106ebe;
        }
        button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }
        .button-icon {
            margin-right: 5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        .filter-row th {
            padding: 5px;
        }
        .column-filter {
            width: 100%;
            padding: 5px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        tr.danger {
            background-color: #ffebee;
        }
        tr.warning {
            background-color: #fff8e1;
        }
        tr.normal {
            background-color: #f1f8e9;
        }
        tr.disabled {
            color: #999;
            font-style: italic;
        }
        .status-icon {
            margin-right: 5px;
        }
        #pagination {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }
        .page-info {
            margin: 0 10px;
        }
        .rows-per-page {
            margin-left: 20px;
            display: flex;
            align-items: center;
        }
        .total-items {
            margin-left: 15px;
        }
        
        @media print {
            .toolbar, button, #pagination, .filter-row {
                display: none;
            }
            body {
                background-color: white;
                margin: 0;
            }
            .container {
                box-shadow: none;
                padding: 0;
            }
            .header {
                background-color: black !important;
                color: white !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            th {
                background-color: #f2f2f2 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            tr.danger {
                background-color: #ffebee !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            tr.warning {
                background-color: #fff8e1 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            tr.normal {
                background-color: #f1f8e9 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-icon">💾</div>
            <h1>OneDriveクォータレポート</h1>
        </div>
        
        <div class="info-section">
            <p><span class="info-label">実行日時:</span> $executionDateFormatted</p>
            <p><span class="info-label">実行者:</span> $executorName</p>
            <p><span class="info-label">実行者の種別:</span> $userType</p>
            <p><span class="info-label">実行モード:</span> $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})</p>
            <p><span class="info-label">出力フォルダ:</span> $OutputDir</p>
        </div>
        
        <div class="toolbar">
            <input type="text" id="searchInput" placeholder="検索...">
            <div id="searchSuggestions"></div>
            <button id="exportBtn"><span class="button-icon">📥</span>CSVエクスポート</button>
            <button id="printBtn"><span class="button-icon">🖨️</span>印刷</button>
        </div>
        
        <div id="pagination"></div>

        <table id="quotaTable">
            <thead>
                <tr class="filter-row">
                    <th><select class="column-filter"><option value="">ユーザー名</option></select></th>
                    <th><select class="column-filter"><option value="">メールアドレス</option></select></th>
                    <th><select class="column-filter"><option value="">ログインユーザー名</option></select></th>
                    <th><select class="column-filter"><option value="">ユーザー種別</option></select></th>
                    <th><select class="column-filter"><option value="">アカウント状態</option></select></th>
                    <th><select class="column-filter"><option value="">OneDrive対応</option></select></th>
                    <th><select class="column-filter"><option value="">総容量(GB)</option></select></th>
                    <th><select class="column-filter"><option value="">使用容量(GB)</option></select></th>
                    <th><select class="column-filter"><option value="">残り容量(GB)</option></select></th>
                    <th><select class="column-filter"><option value="">使用率(%)</option></select></th>
                    <th><select class="column-filter"><option value="">状態</option></select></th>
                </tr>
            </thead>
            <tbody>
"@

# HTML テーブル本体の作成
foreach ($user in $userList) {
    # アカウント状態に応じたアイコンを設定
    $statusIcon = if ($user.'アカウント状態' -eq "有効") { "✅" } else { "❌" }
    
    # 状態に応じたアイコンを設定
    $quotaStatusIcon = switch ($user.'状態') {
        "危険" { "🔴" }
        "警告" { "🟡" }
        "正常" { "🟢" }
        default { "❓" }
    }
    
    # 行を追加
    $htmlContent += @"
                <tr>
                    <td>$($user.'ユーザー名')</td>
                    <td>$($user.'メールアドレス')</td>
                    <td>$($user.'ログインユーザー名')</td>
                    <td>$($user.'ユーザー種別')</td>
                    <td><span class="status-icon">$statusIcon</span>$($user.'アカウント状態')</td>
                    <td>$($user.'総容量(GB)')</td>
                    <td>$($user.'使用容量(GB)')</td>
                    <td>$($user.'残り容量(GB)')</td>
                    <td>$($user.'使用率(%)')</td>
                    <td><span class="status-icon">$quotaStatusIcon</span>$($user.'状態')</td>
                </tr>
"@
}

# HTML 終了部分
$htmlContent += @"
            </tbody>
        </table>
        
        <div class="info-section">
            <p><span class="info-label">色の凡例:</span></p>
            <p>🟢 緑色の行: 使用率が70%未満のユーザー</p>
            <p>🟡 黄色の行: 使用率が70%以上90%未満のユーザー</p>
            <p>🔴 赤色の行: 使用率が90%以上のユーザー</p>
            <p>⚪ グレーの行: 無効なアカウント</p>
        </div>
    </div>
</body>
</html>
"@

# HTML ファイルを出力
$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Log "HTMLファイルを作成しました: $htmlPath" "SUCCESS"

# ログファイルに出力
$userList | Format-Table -AutoSize | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
Write-Log "ストレージクォータ取得が完了しました" "SUCCESS"

# 出力ディレクトリを開く
try {
    Start-Process -FilePath "explorer.exe" -ArgumentList $OutputDir
    Write-Log "出力ディレクトリを開きました: $OutputDir" "SUCCESS"
} catch {
    Write-Log "出力ディレクトリを開けませんでした: $_" "WARNING"
}
