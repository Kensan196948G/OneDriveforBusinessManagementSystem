/**
 * OneDrive管理ツール共通JavaScriptライブラリ
 * 非対話型認証対応版
 */

// リスクレベルスタイル定義
const riskStyles = {
    '高': 'risk-high',
    '中': 'risk-medium',
    '低': 'risk-low'
};

// リスクレベル説明
const riskDescriptions = {
    '高': '外部共有 + 編集権限',
    '中': '外部共有 or 社内広範囲共有',
    '低': '社内限定の限定共有'
};

// config.jsonから認証情報を取得
async function fetchConfig() {
    try {
        // 相対パスで直接インポート
        const configPath = './config.json';
        const response = await fetch(configPath, {
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`設定ファイル読み込み失敗: ${response.status}`);
        }
        
        const config = await response.json();
        
        // 必須フィールドの検証
        if (!config.tenantId || !config.clientId || !config.clientSecret) {
            throw new Error('設定ファイルに必要な認証情報が不足しています');
        }
        
        return config;
    } catch (error) {
        console.error('認証情報取得エラー:', error);
        showError(
            '設定エラー',
            `設定ファイルの読み込みに失敗しました:<br>
            <strong>${error.message}</strong><br><br>
            設定ファイルのパスと内容を確認してください`
        );
        throw error;
    }
}

// エラーメッセージ表示関数
function showError(title, message) {
    const errorContainer = document.getElementById('errorContainer');
    if (errorContainer) {
        errorContainer.querySelector('.alert-heading').textContent = title;
        errorContainer.querySelector('p').innerHTML = message;
        errorContainer.style.display = 'block';
    }
}

// MSALを使用した非対話型認証
async function getAuthToken() {
    const config = await fetchConfig();
    const response = await fetch('/api/auth', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            tenantId: config.tenantId,
            clientId: config.clientId,
            clientSecret: config.clientSecret,
            scopes: config.scopes
        })
    });
    
    if (!response.ok) throw new Error(`認証失敗: ${response.status}`);
    return await response.json();
}

// API呼び出し共通関数
async function callApi(endpoint, method = 'GET', body = null) {
    const { token } = await getAuthToken();
    
    const options = {
        method,
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        }
    };
    
    // SharingCheck.html用ページネーション関数
    window.changePage = function(newPage) {
        const itemsPerPage = 10;
        const resultCount = document.getElementById('resultCount');
        const data = window.currentData || [];
        const totalItems = data.length;
        const totalPages = Math.ceil(totalItems / itemsPerPage);
    
        // ページ範囲チェック
        if (newPage < 1) newPage = 1;
        if (newPage > totalPages) newPage = totalPages;
    
        window.currentPage = newPage;
        
        // 表示件数更新
        const startItem = (newPage - 1) * itemsPerPage + 1;
        const endItem = Math.min(newPage * itemsPerPage, totalItems);
        resultCount.textContent = `${totalItems}件中 ${startItem}～${endItem}件を表示`;
        
        // テーブル更新 (各ページで実装が必要)
        if (typeof window.updateTable === 'function') {
            window.updateTable();
        }
    };
    
    if (body) options.body = JSON.stringify(body);
    
    const response = await fetch(endpoint, options);
    if (!response.ok) throw new Error(`APIエラー: ${response.status}`);
    return await response.json();
}

// ローディング表示制御
function showLoading(show = true) {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = show ? 'flex' : 'none';
    }
}

// エラーメッセージ表示
function showError(message, details = '') {
    const errorContainer = document.getElementById('errorContainer');
    if (errorContainer) {
        errorContainer.innerHTML = `
            <h4><i class="fas fa-exclamation-triangle me-2"></i>${message}</h4>
            <p>${details}</p>
            <button class="btn btn-danger mt-2" onclick="location.reload()">
                <i class="fas fa-sync-alt me-1"></i>再試行
            </button>
        `;
        errorContainer.style.display = 'block';
    }
}

// テーブルデータ表示共通関数 (リスクレベル対応版)
function renderTable(data, tableId, columns, riskColumn = '') {
    const tbody = document.getElementById(tableId);
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach(item => {
        const row = document.createElement('tr');
        if (riskColumn && item[riskColumn]) {
            row.className = riskStyles[item[riskColumn]] || '';
        }
        
        columns.forEach(col => {
            const cell = document.createElement('td');
            cell.textContent = item[col] || '';
            row.appendChild(cell);
        });
        
        tbody.appendChild(row);
    });
}

// 検索結果表示更新
function updateResultCount(data, filteredData, currentPage, itemsPerPage) {
    const resultCount = document.getElementById('resultCount');
    if (resultCount) {
        const start = (currentPage - 1) * itemsPerPage + 1;
        const end = Math.min(currentPage * itemsPerPage, filteredData.length);
        resultCount.textContent = `${filteredData.length}件中 ${start}～${end}件を表示`;
    }
}

// ページネーション設定
function setupPagination(data, itemsPerPage, renderCallback) {
    let currentPage = 1;
    const pageCount = Math.ceil(data.length / itemsPerPage);
    const pagination = document.getElementById('pagination');
    if (!pagination) return;
    
    // 検索結果表示を更新
    updateResultCount(data, data, currentPage, itemsPerPage);
    
    function updatePagination() {
        pagination.innerHTML = '';
        
        // 前へボタン
        const prevItem = document.createElement('li');
        prevItem.className = `page-item ${currentPage === 1 ? 'disabled' : ''}`;
        prevItem.innerHTML = `<a class="page-link" href="#">前へ</a>`;
        prevItem.addEventListener('click', () => {
            if (currentPage > 1) changePage(currentPage - 1);
        });
        pagination.appendChild(prevItem);

        // ページ番号
        for (let i = 1; i <= pageCount; i++) {
            const pageItem = document.createElement('li');
            pageItem.className = `page-item ${i === currentPage ? 'active' : ''}`;
            pageItem.innerHTML = `<a class="page-link" href="#">${i}</a>`;
            pageItem.addEventListener('click', () => changePage(i));
            pagination.appendChild(pageItem);
        }

        // 次へボタン
        const nextItem = document.createElement('li');
        nextItem.className = `page-item ${currentPage === pageCount ? 'disabled' : ''}`;
        nextItem.innerHTML = `<a class="page-link" href="#">次へ</a>`;
        nextItem.addEventListener('click', () => {
            if (currentPage < pageCount) changePage(currentPage + 1);
        });
        pagination.appendChild(nextItem);
    }
    
    function changePage(page) {
        currentPage = page;
        const start = (page - 1) * itemsPerPage;
        const end = start + itemsPerPage;
        renderCallback(data.slice(start, end));
        updateResultCount(data, data, currentPage, itemsPerPage);
        updatePagination();
    }
    
    updatePagination();
    changePage(1);
}

// CSVエクスポート共通関数
function exportToCsv(data, filename, columns) {
    const headers = columns.join(',');
    const rows = data.map(item => 
        columns.map(col => 
            `"${String(item[col] || '').replace(/"/g, '""')}"`
        ).join(',')
    );
    
    const csvContent = [headers, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
}

// フィルタリング処理
function applyFilters(data, columns) {
    const filters = {};
    document.querySelectorAll('.filter-row select').forEach((select, index) => {
        if (select.value) filters[columns[index]] = select.value;
    });
    
    const globalSearch = document.getElementById('globalSearch')?.value.toLowerCase();
    
    return data.filter(item => {
        // グローバル検索
        if (globalSearch) {
            const found = columns.some(col =>
                String(item[col]).toLowerCase().includes(globalSearch)
            );
            if (!found) return false;
        }
        
        // カラムフィルター
        return Object.entries(filters).every(([key, value]) =>
            String(item[key]) === value
        );
    });
}

// フィルターメニュー初期化
function initFilters(data, columns) {
    columns.forEach((col, index) => {
        const select = document.querySelectorAll('.filter-row select')[index];
        if (!select) return;
        
        select.innerHTML = '';
        
        // デフォルトオプション
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = `すべての${col}`;
        select.appendChild(defaultOption);
        
        // ユニークな値を取得してオプション追加
        const uniqueValues = [...new Set(data.map(item => item[col]))];
        uniqueValues.forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
    });
}

// ツールチップ初期化
function initTooltips() {
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(tooltipTriggerEl => {
        return new bootstrap.Tooltip(tooltipTriggerEl, {
            html: true
        });
    });
}

// 簡易ページネーション制御 (SharingCheck.html用)
window.setupSimplePagination = function(totalItems, itemsPerPage = 10) {
    let currentPage = 1;
    const totalPages = Math.ceil(totalItems / itemsPerPage);

    function updatePagination() {
        const startItem = (currentPage - 1) * itemsPerPage + 1;
        const endItem = Math.min(currentPage * itemsPerPage, totalItems);
        
        // 表示件数更新
        document.getElementById('paginationInfo').textContent =
            `${totalItems}件中 ${startItem}～${endItem}件を表示`;
        
        // ボタン状態更新
        document.getElementById('prevPage').disabled = currentPage <= 1;
        document.getElementById('nextPage').disabled = currentPage >= totalPages;
    }

    window.prevPage = function() {
        if (currentPage > 1) {
            currentPage--;
            updatePagination();
            updateTable(); // 各ページで実装が必要
        }
    };

    window.nextPage = function() {
        if (currentPage < totalPages) {
            currentPage++;
            updatePagination();
            updateTable(); // 各ページで実装が必要
        }
    };

    updatePagination();
}

// ブラウザで使用するためmodule.exportsは不要