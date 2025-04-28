/**
 * OneDrive管理ツール共通JavaScriptライブラリ
 * 非対話型認証対応版 + PSスクリプト連携機能
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

// 設定ファイルキャッシュ
let configCache = null;

// config.jsonから認証情報を取得（キャッシュ付き）
async function fetchConfig() {
    if (configCache) return configCache;
    
    try {
        const response = await fetch('./config.json', {
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
        const requiredFields = ['tenantId', 'clientId', 'clientSecret'];
        const missingFields = requiredFields.filter(field => !config[field]);
        
        if (missingFields.length > 0) {
            throw new Error(`設定ファイルに必要な認証情報が不足しています: ${missingFields.join(', ')}`);
        }
        
        configCache = config;
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

// データ取得方法の切り替え（API or PSスクリプト出力）
async function fetchData(source, endpoint, dataFile) {
    try {
        showLoading(true);
        
        if (source === 'api') {
            // API経由でデータ取得
            return await callApi(endpoint);
        } else {
            // PSスクリプトの出力を直接取得
            const response = await fetch(dataFile);
            if (!response.ok) throw new Error(`データファイル読み込み失敗: ${response.status}`);
            return await response.json();
        }
    } catch (error) {
        console.error('データ取得エラー:', error);
        showError('データ取得エラー', error.message);
        throw error;
    } finally {
        showLoading(false);
    }
}

// 認証トークン取得（リトライ付き）
async function getAuthToken(maxRetries = 3) {
    let lastError;
    
    for (let i = 0; i < maxRetries; i++) {
        try {
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
                    scopes: config.scopes || ['https://graph.microsoft.com/.default']
                })
            });
            
            if (!response.ok) throw new Error(`認証失敗: ${response.status}`);
            return await response.json();
        } catch (error) {
            lastError = error;
            if (i < maxRetries - 1) {
                await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
            }
        }
    }
    
    throw lastError;
}

// API呼び出し共通関数（リトライ付き）
async function callApi(endpoint, method = 'GET', body = null, maxRetries = 2) {
    let lastError;
    
    for (let i = 0; i < maxRetries; i++) {
        try {
            const { token } = await getAuthToken();
            
            const options = {
                method,
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            };
            
            if (body) options.body = JSON.stringify(body);
            
            const response = await fetch(endpoint, options);
            if (!response.ok) throw new Error(`APIエラー: ${response.status}`);
            return await response.json();
        } catch (error) {
            lastError = error;
            if (i < maxRetries - 1) {
                await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
            }
        }
    }
    
    throw lastError;
}

// ローディング表示制御
function showLoading(show = true) {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = show ? 'flex' : 'none';
    }
}

// エラーメッセージ表示（詳細モード対応）
function showError(title, message, details = '') {
    const errorContainer = document.getElementById('errorContainer');
    if (!errorContainer) return;
    
    errorContainer.innerHTML = `
        <div class="d-flex align-items-start">
            <i class="fas fa-exclamation-triangle fa-2x me-3 mt-1"></i>
            <div>
                <h4 class="mb-2">${title}</h4>
                <p class="mb-2">${message}</p>
                ${details ? `<details class="mb-2"><summary>詳細</summary><pre>${details}</pre></details>` : ''}
                <div class="d-flex gap-2">
                    <button class="btn btn-danger btn-sm" onclick="location.reload()">
                        <i class="fas fa-sync-alt me-1"></i>再読み込み
                    </button>
                    <button class="btn btn-outline-secondary btn-sm" onclick="document.getElementById('errorContainer').style.display='none'">
                        <i class="fas fa-times me-1"></i>閉じる
                    </button>
                </div>
            </div>
        </div>
    `;
    errorContainer.style.display = 'block';
}

// テーブルデータ表示共通関数（リスクレベル対応版）
function renderTable(data, tableId, columns, riskColumn = '') {
    const tbody = document.getElementById(tableId);
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach(item => {
        const row = document.createElement('tr');
        if (riskColumn && item[riskColumn]) {
            row.className = riskStyles[item[riskColumn]] || '';
            row.title = riskDescriptions[item[riskColumn]] || '';
        }
        
        columns.forEach(col => {
            const cell = document.createElement('td');
            cell.textContent = item[col] || '';
            row.appendChild(cell);
        });
        
        tbody.appendChild(row);
    });
}

// CSVエクスポート共通関数（エスケープ処理強化）
function exportToCsv(data, filename, columns) {
    try {
        const escapeCsv = (str) => {
            if (str == null) return '';
            return `"${String(str).replace(/"/g, '""')}"`;
        };
        
        const headers = columns.map(escapeCsv).join(',');
        const rows = data.map(item => 
            columns.map(col => escapeCsv(item[col])).join(',')
        );
        
        const csvContent = [headers, ...rows].join('\n');
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        
        setTimeout(() => URL.revokeObjectURL(url), 100);
    } catch (error) {
        showError('CSVエクスポートエラー', 'CSVファイルの生成に失敗しました', error.message);
    }
}

// フィルタリング処理（パフォーマンス改善版）
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

// フィルターメニュー初期化（パフォーマンス改善版）
function initFilters(data, columns) {
    columns.forEach((col, index) => {
        const select = document.querySelectorAll('.filter-row select')[index];
        if (!select) return;
        
        // 既存のオプションを保持（選択値を維持）
        const currentValue = select.value;
        select.innerHTML = '';
        
        // デフォルトオプション
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = `すべての${col}`;
        select.appendChild(defaultOption);
        
        // ユニークな値を取得してオプション追加
        const uniqueValues = [...new Set(data.map(item => item[col]).filter(v => v != null))];
        uniqueValues.sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            option.selected = value === currentValue;
            select.appendChild(option);
        });
    });
}

// ページネーション設定（改良版）
function setupPagination(data, itemsPerPage = 10, renderCallback) {
    let currentPage = 1;
    const totalPages = Math.ceil(data.length / itemsPerPage);
    const pagination = document.getElementById('pagination');
    if (!pagination) return;
    
    function updatePagination() {
        pagination.innerHTML = '';
        
        // 前へボタン
        const prevItem = document.createElement('li');
        prevItem.className = `page-item ${currentPage === 1 ? 'disabled' : ''}`;
        prevItem.innerHTML = `<a class="page-link" href="#" aria-label="Previous"><span aria-hidden="true">&laquo;</span></a>`;
        prevItem.addEventListener('click', (e) => {
            e.preventDefault();
            if (currentPage > 1) changePage(currentPage - 1);
        });
        pagination.appendChild(prevItem);
        
        // ページ番号（最大10ページ表示）
        const startPage = Math.max(1, Math.min(currentPage - 5, totalPages - 9));
        const endPage = Math.min(totalPages, startPage + 9);
        
        for (let i = startPage; i <= endPage; i++) {
            const pageItem = document.createElement('li');
            pageItem.className = `page-item ${i === currentPage ? 'active' : ''}`;
            pageItem.innerHTML = `<a class="page-link" href="#">${i}</a>`;
            pageItem.addEventListener('click', (e) => {
                e.preventDefault();
                changePage(i);
            });
            pagination.appendChild(pageItem);
        }
        
        // 次へボタン
        const nextItem = document.createElement('li');
        nextItem.className = `page-item ${currentPage === totalPages ? 'disabled' : ''}`;
        nextItem.innerHTML = `<a class="page-link" href="#" aria-label="Next"><span aria-hidden="true">&raquo;</span></a>`;
        nextItem.addEventListener('click', (e) => {
            e.preventDefault();
            if (currentPage < totalPages) changePage(currentPage + 1);
        });
        pagination.appendChild(nextItem);
    }
    
    function changePage(page) {
        currentPage = page;
        const start = (page - 1) * itemsPerPage;
        const end = start + itemsPerPage;
        renderCallback(data.slice(start, end));
        updatePagination();
    }
    
    updatePagination();
    changePage(1);
}

// ブラウザで使用するためmodule.exportsは不要