param(
    [string]$CsvPath,
    [string]$OutputHtmlPath = "$(Get-Location)\OneDriveQuota_FromCsv.html"
)

if (-not (Test-Path $CsvPath)) {
    Write-Host "CSVファイルが見つかりません: $CsvPath" -ForegroundColor Red
    exit 1
}

$data = Import-Csv -Path $CsvPath -Encoding UTF8
$json = $data | ConvertTo-Json -Compress

# HTMLテンプレート（完全版）
$html = @"
<!DOCTYPE html>
<html>
<head>
<title>OneDriveレポート</title>
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
.column-user { background-color: #e3f2fd; }
.column-email { background-color: #fff8e1; }
.column-login { background-color: #f3e5f5; }
.column-type { background-color: #e8f5e9; }
.column-status { background-color: #fff3e0; }
.column-onedrive { background-color: #f1f8e9; }
.column-total { background-color: #e0f7fa; }
.column-used { background-color: #fce4ec; }
.column-remaining { background-color: #e8eaf6; }
.column-usage { background-color: #f1f8e9; }
.column-state { background-color: #ffebee; }
.status-danger { background-color: #ffcdd2; }
.status-warning { background-color: #ffecb3; }
.status-normal { background-color: #c8e6c9; }
.filter-row select {
    width: 100%;
    padding: 5px;
    background-color: white;
}
.filter-container {
    background-color: #f5f5f5;
    padding: 10px;
    border-radius: 5px;
    margin-bottom: 15px;
}
.count-info {
    font-size: 0.9rem;
    color: #666;
    margin-left: 10px;
}
.badge-count {
    font-size: 0.8rem;
    margin-left: 5px;
}
</style>
</head>
<body>
<div class="container">
    <h1 class="text-center mb-4">OneDrive クォータレポート</h1>
    
    <div class="filter-container">
        <div class="row mb-3">
            <div class="col-md-4">
                <div class="input-group">
                    <span class="input-group-text"><i class="fas fa-search"></i></span>
                    <input type="text" class="form-control" id="globalSearch" placeholder="全体検索..." onkeyup="filterTable()">
                </div>
            </div>
            <div class="col-md-3">
                <div class="input-group">
                    <span class="input-group-text"><i class="fas fa-list-ol"></i></span>
                    <select class="form-select" id="pageSize" onchange="updatePageSize()">
                        <option value="10">10件表示</option>
                        <option value="20">20件表示</option>
                        <option value="50">50件表示</option>
                        <option value="100">100件表示</option>
                        <option value="0">全件表示</option>
                    </select>
                </div>
            </div>
            <div class="col-md-5 text-end">
                <button class="btn btn-outline-secondary me-2" onclick="exportToCsv()">
                    <i class="fas fa-file-csv me-1"></i>CSVエクスポート
                </button>
                <button class="btn btn-outline-primary" onclick="window.print()">
                    <i class="fas fa-print me-1"></i>印刷
                </button>
            </div>
        </div>
        <div id="filterInfo" class="count-info"></div>
    </div>

    <div class="table-responsive">
        <table class="table table-hover" id="quotaTable">
            <thead>
                <tr>
                    <th class="column-user"><i class="fas fa-user me-1"></i>ユーザー名</th>
                    <th class="column-email"><i class="fas fa-envelope me-1"></i>メールアドレス</th>
                    <th class="column-login"><i class="fas fa-sign-in-alt me-1"></i>ログインユーザー名</th>
                    <th class="column-type"><i class="fas fa-users me-1"></i>ユーザー種別</th>
                    <th class="column-status"><i class="fas fa-power-off me-1"></i>アカウント状態</th>
                    <th class="column-onedrive"><i class="fas fa-cloud me-1"></i>OneDrive対応</th>
                    <th class="column-total"><i class="fas fa-database me-1"></i>総容量(GB)</th>
                    <th class="column-used"><i class="fas fa-hdd me-1"></i>使用容量(GB)</th>
                    <th class="column-remaining"><i class="fas fa-sd-card me-1"></i>残り容量(GB)</th>
                    <th class="column-usage"><i class="fas fa-percentage me-1"></i>使用率(%)</th>
                    <th class="column-state"><i class="fas fa-exclamation-circle me-1"></i>状態</th>
                </tr>
                <tr class="filter-row">
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                    <th><select class="form-select" onchange="filterTable()"></select></th>
                </tr>
            </thead>
            <tbody id="tableBody"></tbody>
        </table>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
window.quotaData = $json;

const displayColumns = [
    'ユーザー名',
    'メールアドレス',
    'ログインユーザー名',
    'ユーザー種別',
    'アカウント状態',
    'OneDrive対応',
    '総容量(GB)',
    '使用容量(GB)',
    '残り容量(GB)',
    '使用率(%)',
    '状態'
];

function prepareData(data) {
    return data.map(item => {
        const filteredItem = {};
        displayColumns.forEach(col => {
            filteredItem[col] = item[col];
        });
        return filteredItem;
    });
}

let currentPage = 1;
let pageSize = 10;
let filteredData = [];

function updatePageSize() {
    pageSize = parseInt(document.getElementById('pageSize').value);
    currentPage = 1;
    filterTable();
}

function filterTable() {
    try {
        console.log('フィルタ処理開始');
        const filters = [];
        const filterInputs = document.querySelectorAll('.filter-row select');
        const globalSearch = document.getElementById('globalSearch').value.toLowerCase();
        
        // 各カラムのフィルタ値を取得
        filterInputs.forEach((select, index) => {
            filters[index] = select.value.toLowerCase();
            console.log(`カラム${index} フィルタ値: ${filters[index]}`);
        });

        // データフィルタリング
        const allData = prepareData(window.quotaData);
        console.log('全データ件数:', allData.length);
        
        filteredData = allData.filter(item => {
            // グローバル検索
            if (globalSearch) {
                const found = displayColumns.some(col => 
                    String(item[col]).toLowerCase().includes(globalSearch)
                );
                if (!found) return false;
            }

            // カラムごとのフィルタ
            return displayColumns.every((col, colIndex) => {
                if (!filters[colIndex]) return true;
                return String(item[col]).toLowerCase().includes(filters[colIndex]);
            });
        });
        
        console.log('フィルタ後データ件数:', filteredData.length);

        renderPaginatedTable();
    } catch (error) {
        console.error('フィルタ処理エラー:', error);
    }
}

function renderPaginatedTable() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';
    
    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = pageSize > 0 ? startIdx + pageSize : filteredData.length;
    const paginatedData = pageSize > 0 ? 
        filteredData.slice(startIdx, endIdx) : filteredData;
    
    paginatedData.forEach(item => {
        const row = document.createElement('tr');
        
        if (item['状態'] === '危険') {
            row.classList.add('status-danger');
        } else if (item['状態'] === '警告') {
            row.classList.add('status-warning');
        } else {
            row.classList.add('status-normal');
        }
        
        displayColumns.forEach(col => {
            const td = document.createElement('td');
            td.textContent = item[col];
            row.appendChild(td);
        });
        
        tableBody.appendChild(row);
    });
    
    updatePaginationInfo();
}

function updatePaginationInfo() {
    try {
        const totalItems = filteredData.length;
        const totalPages = pageSize > 0 ? Math.ceil(totalItems / pageSize) : 1;
        const showingFrom = Math.min((currentPage-1)*pageSize+1, totalItems);
        const showingTo = Math.min(currentPage*pageSize, totalItems);
        
        console.log(`ページ情報更新: 全${totalItems}件, ページ${currentPage}/${totalPages}, 表示${showingFrom}-${showingTo}`);
        
        const filterInfo = document.getElementById('filterInfo');
        if (filterInfo) {
            filterInfo.innerHTML = @"
                <span class="badge bg-primary badge-count">全 $(totalItems) 件</span>
                <span class="badge bg-success badge-count">表示中 $(showingFrom)-$(showingTo) 件</span>
                $($(totalPages > 1 ? '<span class="badge bg-info badge-count">ページ $(currentPage)/$(totalPages)</span>' : ''))
            "@;
        }
    } catch (error) {
        console.error('ページ情報更新エラー:', error);
    }
    
    const activeFilters = Array.from(document.querySelectorAll('.filter-row select, .filter-row input'))
        .filter(el => el.value).length;
    
    if (activeFilters > 0) {
            filterInfo.innerHTML += `
                <span class="badge bg-warning text-dark badge-count">
                    <i class="fas fa-filter me-1"></i>${activeFilters} フィルタ適用中
                </span>
            `;
    }
}

function exportToCsv() {
    const headers = displayColumns.join(',') + '\n';
    const csvContent = headers + filteredData.map(item => 
        displayColumns.map(col => `"${item[col]}"`).join(',')
    ).join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'OneDriveReport.csv');
    link.click();
}

window.onload = function() {
    try {
        // 各カラムにプルダウンメニューを設定
        document.querySelectorAll('.filter-row th').forEach((th, index) => {
            if (index >= displayColumns.length) return;
            
            const select = document.createElement('select');
            select.className = 'form-select';
            select.onchange = filterTable;
            
            // デフォルトオプション
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = 'すべて';
            select.appendChild(defaultOption);
            
            // ユニークな値を取得してオプション追加
            const uniqueValues = [...new Set(
                window.quotaData.map(item => String(item[displayColumns[index]]))
            )].filter(Boolean).sort();
            
            uniqueValues.forEach(value => {
                const option = document.createElement('option');
                option.value = value;
                option.textContent = value;
                select.appendChild(option);
            });
            
            th.innerHTML = '';
            th.appendChild(select);
        });
        
        // 初期データ表示
        console.log('初期データ読み込み開始');
        try {
            filteredData = prepareData(window.quotaData);
            console.log('データ準備完了:', filteredData.length, '件');
            renderPaginatedTable();
            console.log('初期表示完了');
        } catch (error) {
            console.error('データ表示エラー:', error);
        }
        
        // ツールチップ初期化
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(tooltipTriggerEl => new bootstrap.Tooltip(tooltipTriggerEl));
        
    } catch (error) {
        console.error('初期化エラー:', error);
    }
};
</script>
</body>
</html>
"@
$html | Out-File -FilePath $OutputHtmlPath -Encoding UTF8
Write-Host "HTMLファイルが生成されました: $OutputHtmlPath"
