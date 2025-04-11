# OneDrive for Business 運用ツール - ITSM準拠
# GenerateReport.ps1 - 総合レポート生成スクリプト

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$LogDir = "$(Get-Location)\Log"
)

# 実行開始時刻を記録
$executionTime = Get-Date

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
    $logFilePath = Join-Path -Path $LogDir -ChildPath "GenerateReport.log"
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
    
    # エラーログに詳細を出力
    $errorLogPath = Join-Path -Path $LogDir -ChildPath "GenerateReport.Error.log"
    Add-Content -Path $errorLogPath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorDetails -Encoding UTF8
}

# 総合レポート生成開始
Write-Log "総合レポート生成を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"

# 親ディレクトリを取得（日付ベースのフォルダ）
$parentDir = Split-Path -Parent $OutputDir
Write-Log "親ディレクトリ: $parentDir" "INFO"

# カテゴリフォルダを検索
$categoryFolders = @(
    "Basic_Data_Collection",
    "Incident_Management",
    "Change_Management",
    "Security_Management"
)

# データ収集用のハッシュテーブル
$reportData = @{
    "UserInfo" = @()
    "OneDriveQuota" = @()
    "SyncErrors" = @()
    "SharingSettings" = @()
    "ExternalSharing" = @()
}

# 各カテゴリフォルダからCSVファイルを読み込む
foreach ($category in $categoryFolders) {
    $categoryPattern = "$category.*"
    $categoryFolderPath = Get-ChildItem -Path $parentDir -Directory | Where-Object { $_.Name -like $categoryPattern } | Select-Object -First 1
    
    if ($categoryFolderPath) {
        Write-Log "カテゴリフォルダを検出: $($categoryFolderPath.FullName)" "INFO"
        
        # CSVファイルを検索
        $csvFiles = Get-ChildItem -Path $categoryFolderPath.FullName -Filter "*.csv" -File
        
        foreach ($csvFile in $csvFiles) {
            Write-Log "CSVファイルを処理: $($csvFile.Name)" "INFO"
            
            try {
                # CSVファイルを読み込む
                $csvData = Import-Csv -Path $csvFile.FullName -Encoding UTF8
                
                # ファイル名に基づいてデータを分類
                if ($csvFile.Name -like "*UserInfo*") {
                    $reportData["UserInfo"] += $csvData
                    Write-Log "ユーザー情報データを追加: $($csvData.Count) 件" "SUCCESS"
                }
                elseif ($csvFile.Name -like "*OneDriveQuota*" -or $csvFile.Name -like "*OneDriveCheck*") {
                    $reportData["OneDriveQuota"] += $csvData
                    Write-Log "OneDriveクォータデータを追加: $($csvData.Count) 件" "SUCCESS"
                }
                elseif ($csvFile.Name -like "*SyncError*") {
                    $reportData["SyncErrors"] += $csvData
                    Write-Log "同期エラーデータを追加: $($csvData.Count) 件" "SUCCESS"
                }
                elseif ($csvFile.Name -like "*Sharing*") {
                    $reportData["SharingSettings"] += $csvData
                    Write-Log "共有設定データを追加: $($csvData.Count) 件" "SUCCESS"
                }
                elseif ($csvFile.Name -like "*External*") {
                    $reportData["ExternalSharing"] += $csvData
                    Write-Log "外部共有データを追加: $($csvData.Count) 件" "SUCCESS"
                }
            }
            catch {
                Write-ErrorLog $_ "CSVファイルの読み込み中にエラーが発生しました: $($csvFile.FullName)"
            }
        }
    }
    else {
        Write-Log "カテゴリフォルダが見つかりません: $category" "WARNING"
    }
}

# タイムスタンプ
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# 出力ファイル名の設定
$htmlFile = "OneDriveReport.$timestamp.html"
$jsFile = "OneDriveReport.$timestamp.js"

# 出力パスの設定
$htmlPath = Join-Path -Path $OutputDir -ChildPath $htmlFile
$jsPath = Join-Path -Path $OutputDir -ChildPath $jsFile

# JavaScript ファイルの生成
$jsContent = @"
// OneDrive 総合レポート用 JavaScript

// グローバル変数
let currentPage = 1;
let rowsPerPage = 10; // デフォルトの1ページあたりの行数
let filteredRows = []; // フィルタリングされた行を保持する配列
let currentTab = 'userInfo'; // 現在表示中のタブ

// タブを切り替える関数
function switchTab(tabName) {
    // すべてのタブコンテンツを非表示
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.style.display = 'none';
    });
    
    // すべてのタブボタンから選択状態を解除
    document.querySelectorAll('.tab-button').forEach(button => {
        button.classList.remove('active');
    });
    
    // 選択されたタブを表示
    document.getElementById(tabName + 'Tab').style.display = 'block';
    document.getElementById(tabName + 'Button').classList.add('active');
    
    // 現在のタブを更新
    currentTab = tabName;
    
    // テーブルの初期化
    initializeTable(tabName);
}

// テーブルを初期化する関数
function initializeTable(tabName) {
    // テーブルIDを取得
    const tableId = tabName + 'Table';
    const table = document.getElementById(tableId);
    
    if (!table) return;
    
    // フィルター行を作成
    createColumnFilters(tableId);
    
    // 行に色を付ける
    colorizeRows(tableId);
    
    // 行データを収集
    collectRowData(tableId);
    
    // ページングを更新
    updatePagination();
}

// テーブルを検索する関数（インクリメンタル検索対応）
function searchTable() {
    var input = document.getElementById('searchInput').value.toLowerCase();
    var tableId = currentTab + 'Table';
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    filteredRows = [];
    
    for (var i = 0; i < rows.length; i++) {
        var found = false;
        var cells = rows[i].getElementsByTagName('td');
        var rowData = {};
        
        for (var j = 0; j < cells.length; j++) {
            var cellText = cells[j].textContent || cells[j].innerText;
            // 列のヘッダー名を取得
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
    
    // 検索候補の表示
    showSearchSuggestions(input);
    
    // 検索結果が空の場合は検索候補を非表示
    if (filteredRows.length === 0 && input.length > 0) {
        document.getElementById('searchSuggestions').innerHTML = '<div class="suggestion-item">検索結果がありません</div>';
        document.getElementById('searchSuggestions').style.display = 'block';
    }
    
    // ページングの更新
    currentPage = 1;
    updatePagination();
}

// 検索候補を表示する関数
function showSearchSuggestions(input) {
    var suggestionsDiv = document.getElementById('searchSuggestions');
    suggestionsDiv.innerHTML = '';
    
    if (input.length < 1) {
        suggestionsDiv.style.display = 'none';
        return;
    }
    
    // 一致する値を収集（重複なし）
    var matches = new Set();
    filteredRows.forEach(item => {
        Object.values(item.data).forEach(value => {
            if (value.toLowerCase().indexOf(input.toLowerCase()) > -1) {
                matches.add(value);
            }
        });
    });
    
    // 最大5件まで表示
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
    } else if (input.length > 0) {
        // 検索結果がない場合のメッセージ
        var noResults = document.createElement('div');
        noResults.className = 'suggestion-item no-results';
        noResults.textContent = '検索結果がありません';
        suggestionsDiv.appendChild(noResults);
        suggestionsDiv.style.display = 'block';
    } else {
        suggestionsDiv.style.display = 'none';
    }
}

// 列フィルターを作成する関数
function createColumnFilters(tableId) {
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
    // 既存のフィルター行を削除
    var existingFilterRow = table.querySelector('.filter-row');
    if (existingFilterRow) {
        existingFilterRow.remove();
    }
    
    var headers = table.getElementsByTagName('thead')[0].getElementsByTagName('th');
    var filterRow = document.createElement('tr');
    filterRow.className = 'filter-row';
    
    for (var i = 0; i < headers.length; i++) {
        var cell = document.createElement('th');
        var select = document.createElement('select');
        select.className = 'column-filter';
        select.setAttribute('data-column', i);
        select.setAttribute('data-table', tableId);
        
        // デフォルトのオプション
        var defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'すべて';
        select.appendChild(defaultOption);
        
        // 列の一意の値を取得
        var uniqueValues = new Set();
        var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
        for (var j = 0; j < rows.length; j++) {
            var cellValue = rows[j].getElementsByTagName('td')[i].textContent;
            uniqueValues.add(cellValue);
        }
        
        // 一意の値をソートしてオプションとして追加
        Array.from(uniqueValues).sort().forEach(value => {
            var option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        // 変更イベントリスナーを追加
        select.addEventListener('change', applyColumnFilters);
        
        cell.appendChild(select);
        filterRow.appendChild(cell);
    }
    
    // フィルター行をテーブルヘッダーに追加
    table.getElementsByTagName('thead')[0].appendChild(filterRow);
}

// 列フィルターを適用する関数
function applyColumnFilters() {
    var tableId = this.getAttribute('data-table');
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var filters = table.querySelectorAll('.column-filter');
    filteredRows = [];
    
    // 各行に対してフィルターを適用
    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var cells = row.getElementsByTagName('td');
        var rowData = {};
        var includeRow = true;
        
        // 各フィルターをチェック
        for (var j = 0; j < filters.length; j++) {
            var filter = filters[j];
            var columnIndex = parseInt(filter.getAttribute('data-column'));
            var filterValue = filter.value;
            
            // 列のヘッダー名を取得
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[columnIndex].textContent;
            var cellValue = cells[columnIndex].textContent;
            rowData[headerText] = cellValue;
            
            // フィルター値が設定されていて、セルの値と一致しない場合は行を除外
            if (filterValue && cellValue !== filterValue) {
                includeRow = false;
                break;
            }
        }
        
        if (includeRow) {
            filteredRows.push({row: row, data: rowData});
        }
    }
    
    // 検索フィールドの値も考慮
    var searchInput = document.getElementById('searchInput').value.toLowerCase();
    if (searchInput) {
        filteredRows = filteredRows.filter(item => {
            return Object.values(item.data).some(value => 
                value.toLowerCase().indexOf(searchInput) > -1
            );
        });
    }
    
    // ページングの更新
    currentPage = 1;
    updatePagination();
}

// 行データを収集する関数
function collectRowData(tableId) {
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
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
}

// ページングを更新する関数
function updatePagination() {
    var tableId = currentTab + 'Table';
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
    var tbody = table.getElementsByTagName('tbody')[0];
    var rows = tbody.getElementsByTagName('tr');
    
    // すべての行を非表示にする
    for (var i = 0; i < rows.length; i++) {
        rows[i].style.display = 'none';
    }
    
    // フィルタリングされた行のみを表示
    var startIndex = (currentPage - 1) * rowsPerPage;
    var endIndex = Math.min(startIndex + rowsPerPage, filteredRows.length);
    
    for (var i = startIndex; i < endIndex; i++) {
        filteredRows[i].row.style.display = '';
    }
    
    // ページネーションコントロールを更新
    updatePaginationControls();
}

// ページネーションコントロールを更新する関数
function updatePaginationControls() {
    var paginationDiv = document.getElementById('pagination');
    paginationDiv.innerHTML = '';
    
    var totalPages = Math.ceil(filteredRows.length / rowsPerPage);
    
    // 「前へ」ボタン
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
    
    // ページ番号
    var pageInfo = document.createElement('span');
    pageInfo.className = 'page-info';
    pageInfo.textContent = currentPage + ' / ' + (totalPages || 1) + ' ページ';
    paginationDiv.appendChild(pageInfo);
    
    // 「次へ」ボタン
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
    
    // 1ページあたりの行数を選択
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
    
    // 総件数表示
    var totalItems = document.createElement('span');
    totalItems.className = 'total-items';
    totalItems.textContent = '全 ' + filteredRows.length + ' 件';
    paginationDiv.appendChild(totalItems);
}

// 検索入力フィールドからフォーカスが外れたときに検索候補を非表示にする
function hideSearchSuggestions() {
    // 少し遅延させて、候補をクリックする時間を確保
    setTimeout(function() {
        document.getElementById('searchSuggestions').style.display = 'none';
    }, 200);
}

// CSVとしてエクスポートする関数
function exportTableToCSV() {
    var tableId = currentTab + 'Table';
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
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
    for (var i = 0; i < filteredRows.length; i++) {
        var row = [], cells = filteredRows[i].row.getElementsByTagName('td');
        for (var j = 0; j < cells.length; j++) {
            // セル内のテキストから改行や引用符を適切に処理
            var data = cells[j].innerText.replace(/(\r\n|\n|\r)/gm, ' ').replace(/"/g, '""');
            row.push('"' + data.trim() + '"'); // 余分な空白を削除
        }
        csv.push(row.join(','));
    }
    
    // CSVファイルのダウンロード（UTF-8 BOM付きで文字化け対策）
    var csvContent = '\uFEFF' + csv.join('\n'); // BOMを追加
    var csvFile = new Blob([csvContent], {type: 'text/csv;charset=utf-8'});
    var downloadLink = document.createElement('a');
    downloadLink.download = 'OneDriveReport_' + currentTab + '_Export.csv';
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
function colorizeRows(tableId) {
    var table = document.getElementById(tableId);
    
    if (!table) return;
    
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    
    // テーブルに応じた色分けルールを適用
    if (tableId === 'oneDriveQuotaTable') {
        // OneDriveクォータテーブルの色分け
        for (var i = 0; i < rows.length; i++) {
            var usageCell = rows[i].querySelector('td:nth-child(10)'); // 使用率のセル
            if (usageCell) {
                var usage = parseFloat(usageCell.textContent);
                if (!isNaN(usage)) {
                    if (usage >= 90) {
                        rows[i].classList.add('danger');
                    } else if (usage >= 70) {
                        rows[i].classList.add('warning');
                    } else {
                        rows[i].classList.add('normal');
                    }
                }
            }
            
            // アカウント状態によっても色分け
            var accountStatus = rows[i].querySelector('td:nth-child(5)'); // アカウント状態のセル
            if (accountStatus && accountStatus.textContent === '無効') {
                rows[i].classList.add('disabled');
            }
        }
    } else if (tableId === 'syncErrorsTable') {
        // 同期エラーテーブルの色分け
        for (var i = 0; i < rows.length; i++) {
            var errorTypeCell = rows[i].querySelector('td:nth-child(4)'); // エラー種別のセル
            if (errorTypeCell) {
                var errorType = errorTypeCell.textContent;
                if (errorType.includes('エラー')) {
                    rows[i].classList.add('danger');
                } else if (errorType.includes('警告')) {
                    rows[i].classList.add('warning');
                } else {
                    rows[i].classList.add('info');
                }
            }
        }
    } else if (tableId === 'sharingSettingsTable' || tableId === 'externalSharingTable') {
        // 共有設定テーブルと外部共有テーブルの色分け
        for (var i = 0; i < rows.length; i++) {
            var riskCell = rows[i].querySelector('td:nth-child(9)'); // リスクレベルのセル
            if (riskCell) {
                var risk = riskCell.textContent;
                if (risk.includes('高')) {
                    rows[i].classList.add('danger');
                } else if (risk.includes('中')) {
                    rows[i].classList.add('warning');
                } else {
                    rows[i].classList.add('normal');
                }
            }
        }
    } else {
        // ユーザー情報テーブルなど、その他のテーブルの色分け
        for (var i = 0; i < rows.length; i++) {
            var accountStatus = rows[i].querySelector('td:nth-child(5)'); // アカウント状態のセル
            if (accountStatus) {
                if (accountStatus.textContent === '無効') {
                    rows[i].classList.add('disabled');
                } else {
                    rows[i].classList.add('normal');
                }
            }
        }
    }
}

// ページロード時に実行
window.onload = function() {
    // 初期タブを表示
    switchTab('userInfo');
    
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
};
"@

# JavaScript ファイルを出力
$jsContent | Out-File -FilePath $jsPath -Encoding UTF8
Write-Log "JavaScriptファイルを作成しました: $jsPath" "SUCCESS"

# 実行日時とユーザー情報を取得
$executionDateFormatted = $executionTime.ToString("yyyy/MM/dd HH:mm:ss")

# HTML ファイルの生成
$htmlContent = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>OneDrive 総合レポート</title>
    <script src="$jsFile"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        /* 横スクロール防止＆自動調整 */
        .container {
            overflow-x: auto;
        }
        table {
            width: 100%;
            max-width: 100%;
            table-layout: auto; /* 内容に応じて自動調整 */
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        th, td {
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        th {
            white-space: nowrap;
        }
        /* URLやメール列は強制折り返し＋最大幅制限 */
        td:nth-child(2), td:nth-child(11), td:nth-child(12),
        th:nth-child(2), th:nth-child(11), th:nth-child(12) {
            max-width: 250px;
            word-break: break-all;
            overflow-wrap: break-word;
        }
        /* ユーザー名列の幅を狭く固定 */
        th:nth-child(1), td:nth-child(1) {
            width: 120px;
            max-width: 150px;
            word-break: break-all;
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
        .tabs {
            display: flex;
            margin-bottom: 20px;
            border-bottom: 1px solid #ddd;
        }
        .tab-button {
            padding: 10px 15px;
            background-color: #f0f0f0;
            border: none;
            border-radius: 4px 4px 0 0;
            cursor: pointer;
            margin-right: 5px;
        }
        .tab-button:hover {
            background-color: #e0e0e0;
        }
        .tab-button.active {
            background-color: #0078d4;
            color: white;
        }
        .tab-content {
            display: none;
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
        tr.info {
            background-color: #e3f2fd;
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
            .toolbar, button, #pagination, .filter-row, .tabs {
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
            tr.info {
                background-color: #e3f2fd !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            .tab-content {
                display: block !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-icon">📊</div>
            <h1>OneDrive 総合レポート</h1>
        </div>
        
        <div class="info-section">
            <p><span class="info-label">実行日時:</span> $executionDateFormatted</p>
            <p><span class="info-label">出力フォルダ:</span> $OutputDir</p>
        </div>
        
        <div class="toolbar">
            <input type="text" id="searchInput" placeholder="検索...">
            <div id="searchSuggestions"></div>
            <button id="exportBtn"><span class="button-icon">📥</span>CSVエクスポート</button>
            <button id="printBtn"><span class="button-icon">🖨️</span>印刷</button>
        </div>
        
        <div class="tabs">
            <button id="userInfoButton" class="tab-button" onclick="switchTab('userInfo')">ユーザー情報</button>
            <button id="oneDriveQuotaButton" class="tab-button" onclick="switchTab('oneDriveQuota')">OneDriveクォータ</button>
            <button id="syncErrorsButton" class="tab-button" onclick="switchTab('syncErrors')">同期エラー</button>
            <button id="sharingSettingsButton" class="tab-button" onclick="switchTab('sharingSettings')">共有設定</button>
            <button id="externalSharingButton" class="tab-button" onclick="switchTab('externalSharing')">外部共有</button>
        </div>
        
        <div id="pagination"></div>
        
        <!-- ユーザー情報タブ -->
        <div id="userInfoTab" class="tab-content">
            <table id="userInfoTable">
                <thead>
                    <tr>
                        <th>表示名</th>
                        <th>メール</th>
                        <th>ログインユーザー名</th>
                        <th>ユーザー種別</th>
                        <th>アカウント状態</th>
                        <th>最終同期日時</th>
                    </tr>
                </thead>
                <tbody>
"@

# ユーザー情報テーブルの作成
if ($reportData["UserInfo"].Count -gt 0) {
    foreach ($user in $reportData["UserInfo"]) {
        # アカウント状態に応じたアイコンを設定
        $statusIcon = if ($user.'アカウント状態' -eq "有効") { "✅" } else { "❌" }
        
        # 行を追加
        $htmlContent += @"
                    <tr>
                        <td>$($user.'ユーザー名')</td>
                        <td>$($user.'メールアドレス')</td>
                        <td>$($user.'ログインユーザー名')</td>
                        <td>$($user.'ユーザー種別')</td>
                        <td><span class="status-icon">$statusIcon</span>$($user.'アカウント状態')</td>
                        <td>$($user.'最終同期日時')</td>
                    </tr>
"@
    }
} else {
    $htmlContent += @"
                    <tr>
                        <td colspan="6" style="text-align: center;">データがありません</td>
                    </tr>
"@
}

$htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <!-- OneDriveクォータタブ -->
        <div id="oneDriveQuotaTab" class="tab-content">
            <table id="oneDriveQuotaTable">
                <thead>
                    <tr>
                        <th>表示名</th>
                        <th>メール</th>
                        <th>ログインユーザー名</th>
                        <th>ユーザー種別</th>
                        <th>アカウント状態</th>
                        <th>総容量(GB)</th>
                        <th>使用容量(GB)</th>
                        <th>残り容量(GB)</th>
                        <th>使用率(%)</th>
                    </tr>
                </thead>
                <tbody>
"@

# OneDriveクォータテーブルの作成
if ($reportData["OneDriveQuota"].Count -gt 0) {
    foreach ($quota in $reportData["OneDriveQuota"]) {
        # アカウント状態に応じたアイコンを設定
        $statusIcon = if ($quota.'アカウント状態' -eq "有効") { "✅" } else { "❌" }
        
        # 行を追加
        $htmlContent += @"
                    <tr>
                        <td>$($quota.'ユーザー名')</td>
                        <td>$($quota.'メールアドレス')</td>
                        <td>$($quota.'ログインユーザー名')</td>
                        <td>$($quota.'ユーザー種別')</td>
                        <td><span class="status-icon">$statusIcon</span>$($quota.'アカウント状態')</td>
                        <td>$($quota.'総容量(GB)')</td>
                        <td>$($quota.'使用容量(GB)')</td>
                        <td>$($quota.'残り容量(GB)')</td>
                        <td>$($quota.'使用率(%)')</td>
                    </tr>
"@
    }
} else {
    $htmlContent += @"
                    <tr>
                        <td colspan="9" style="text-align: center;">データがありません</td>
                    </tr>
"@
}

$htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <!-- 同期エラータブ -->
        <div id="syncErrorsTab" class="tab-content">
            <table id="syncErrorsTable">
                <thead>
                    <tr>
                        <th>表示名</th>
                        <th>メール</th>
                        <th>アカウント状態</th>
                        <th>エラー種別</th>
                        <th>ファイル名</th>
                        <th>ファイルパス</th>
                        <th>最終更新日時</th>
                        <th>サイズ(KB)</th>
                        <th>エラー詳細</th>
                        <th>推奨対応</th>
                    </tr>
                </thead>
                <tbody>
"@

# 同期エラーテーブルの作成
if ($reportData["SyncErrors"].Count -gt 0) {
    foreach ($error in $reportData["SyncErrors"]) {
        # 行を追加
        $htmlContent += @"
                    <tr>
                        <td>$($error.'ユーザー名')</td>
                        <td>$($error.'メールアドレス')</td>
                        <td>$($error.'アカウント状態')</td>
                        <td>$($error.'エラー種別')</td>
                        <td>$($error.'ファイル名')</td>
                        <td>$($error.'ファイルパス')</td>
                        <td>$($error.'最終更新日時')</td>
                        <td>$($error.'サイズ(KB)')</td>
                        <td>$($error.'エラー詳細')</td>
                        <td>$($error.'推奨対応')</td>
                    </tr>
"@
    }
} else {
    $htmlContent += @"
                    <tr>
                        <td colspan="10" style="text-align: center;">データがありません</td>
                    </tr>
"@
}

$htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <!-- 共有設定タブ -->
        <div id="sharingSettingsTab" class="tab-content">
            <table id="sharingSettingsTable">
                <thead>
                    <tr>
                        <th>表示名</th>
                        <th>メール</th>
                        <th>アカウント状態</th>
                        <th>共有方向</th>
                        <th>アイテム名</th>
                        <th>アイテムタイプ</th>
                        <th>共有タイプ</th>
                        <th>共有範囲</th>
                        <th>リスクレベル</th>
                        <th>最終更新日時</th>
                        <th>WebURL</th>
                    </tr>
                </thead>
                <tbody>
"@

# 共有設定テーブルの作成
if ($reportData["SharingSettings"].Count -gt 0) {
    foreach ($sharing in $reportData["SharingSettings"]) {
        # 行を追加
        $htmlContent += @"
                    <tr>
                        <td>$($sharing.'ユーザー名')</td>
                        <td>$($sharing.'メールアドレス')</td>
                        <td>$($sharing.'アカウント状態')</td>
                        <td>$($sharing.'共有方向')</td>
                        <td>$($sharing.'アイテム名')</td>
                        <td>$($sharing.'アイテムタイプ')</td>
                        <td>$($sharing.'共有タイプ')</td>
                        <td>$($sharing.'共有範囲')</td>
                        <td>$($sharing.'リスクレベル')</td>
                        <td>$($sharing.'最終更新日時')</td>
                        <td>$($sharing.'WebURL')</td>
                    </tr>
"@
    }
} else {
    $htmlContent += @"
                    <tr>
                        <td colspan="11" style="text-align: center;">データがありません</td>
                    </tr>
"@
}

$htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <!-- 外部共有タブ -->
        <div id="externalSharingTab" class="tab-content">
            <table id="externalSharingTable">
                <thead>
                    <tr>
                        <th>表示名</th>
                        <th>メール</th>
                        <th>アカウント状態</th>
                        <th>アイテム名</th>
                        <th>アイテムタイプ</th>
                        <th>共有タイプ</th>
                        <th>共有範囲</th>
                        <th>共有先</th>
                        <th>リスクレベル</th>
                        <th>最終更新日時</th>
                        <th>WebURL</th>
                        <th>セキュリティ推奨</th>
                    </tr>
                </thead>
                <tbody>
"@

# 外部共有テーブルの作成
if ($reportData["ExternalSharing"].Count -gt 0) {
    foreach ($external in $reportData["ExternalSharing"]) {
        # 行を追加
        $htmlContent += @"
                    <tr>
                        <td>$($external.'ユーザー名')</td>
                        <td>$($external.'メールアドレス')</td>
                        <td>$($external.'アカウント状態')</td>
                        <td>$($external.'アイテム名')</td>
                        <td>$($external.'アイテムタイプ')</td>
                        <td>$($external.'共有タイプ')</td>
                        <td>$($external.'共有範囲')</td>
                        <td>$($external.'共有先')</td>
                        <td>$($external.'リスクレベル')</td>
                        <td>$($external.'最終更新日時')</td>
                        <td>$($external.'WebURL')</td>
                        <td>$($external.'セキュリティ推奨')</td>
                    </tr>
"@
    }
} else {
    $htmlContent += @"
                    <tr>
                        <td colspan="12" style="text-align: center;">データがありません</td>
                    </tr>
"@
}

$htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <div class="info-section">
            <p><span class="info-label">色の凡例:</span></p>
            <p>🟢 緑色の行: 正常なユーザー、使用率が70%未満のユーザー、低リスクの共有</p>
            <p>🟡 黄色の行: 使用率が70%以上90%未満のユーザー、中リスクの共有、警告</p>
            <p>🔴 赤色の行: 使用率が90%以上のユーザー、高リスクの共有、エラー</p>
            <p>🔵 青色の行: 情報メッセージ</p>
            <p>⚪ グレーの行: 無効なアカウント</p>
        </div>
    </div>
</body>
</html>
"@

# HTML ファイルを出力
$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Log "HTMLファイルを作成しました: $htmlPath" "SUCCESS"

# 出力ディレクトリを開く
try {
    Start-Process -FilePath "explorer.exe" -ArgumentList $OutputDir
    Write-Log "出力ディレクトリを開きました: $OutputDir" "SUCCESS"
} catch {
    Write-ErrorLog $_ "出力ディレクトリを開けませんでした: $OutputDir"
}

Write-Log "総合レポート生成が完了しました" "SUCCESS"
