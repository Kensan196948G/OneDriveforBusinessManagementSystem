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
