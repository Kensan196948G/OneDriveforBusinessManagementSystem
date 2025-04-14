param(
    [string]$CsvPath,
    [string]$OutputHtmlPath = "$(Get-Location)\OneDriveQuota_FromCsv.html"
)

if (-not (Test-Path $CsvPath)) {
    Write-Host "CSVファイルが見つかりません: $CsvPath" -ForegroundColor Red
    exit 1
}

$data = Import-Csv -Path $CsvPath -Encoding UTF8

$executionDateFormatted = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
$executorName = "管理者"
$userType = "Member"
$isAdmin = $true

# 列名リスト
$columns = @(
    "ユーザー名",
    "メールアドレス",
    "ログインユーザー名",
    "ユーザー種別",
    "アカウント状態",
    "OneDrive対応",
    "総容量(GB)",
    "使用容量(GB)",
    "残り容量(GB)",
    "使用率(%)",
    "状態"
)

# 各列のユニーク値を抽出
$uniqueValues = @{}
foreach ($col in $columns) {
    $uniqueValues[$col] = $data | Select-Object -ExpandProperty $col | Sort-Object -Unique
}

# JS用データ配列
$json = $data | ConvertTo-Json -Compress

# 1. HTML本体（tbodyまで）
$html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>OneDrive 基本データレポート</title>
<style>
body { font-family: Arial, sans-serif; margin:20px; background:#f5f5f5; }
.container { background:#fff; padding:20px; border-radius:5px; box-shadow:0 2px 5px rgba(0,0,0,0.1); }
.header { background:#0078d4; color:#fff; padding:15px; border-radius:5px; display:flex; align-items:center; gap:20px; }
.header-icon { font-size:24px; }
h1 { margin:0; font-size:24px; }
.info-section { background:#f0f0f0; padding:10px; margin:15px 0; border-radius:5px; font-size:14px; }
.info-label { font-weight:bold; margin-right:5px; }
.toolbar { margin-bottom:15px; display:flex; gap:10px; align-items:center; position:relative; }
#searchInput { padding:8px; border:1px solid #ddd; border-radius:4px; flex-grow:1; }
#searchSuggestions { position:absolute; top:100%; left:0; width:100%; max-height:200px; overflow-y:auto; background:#fff; border:1px solid #ddd; border-radius:0 0 4px 4px; z-index:1000; display:none; }
.suggestion-item { padding:8px; border-bottom:1px solid #eee; cursor:pointer; }
.suggestion-item:hover { background:#f0f0f0; }
.suggestion-item.no-results { color:#999; font-style:italic; cursor:default; }
button { padding:8px 12px; background:#0078d4; color:#fff; border:none; border-radius:4px; cursor:pointer; display:flex; align-items:center; }
button:hover { background:#106ebe; }
button:disabled { background:#ccc; cursor:not-allowed; }
.button-icon { margin-right:5px; }
#pagination { display:flex; justify-content:center; align-items:center; gap:10px; flex-wrap:wrap; margin:15px 0; }
.page-info { margin:0 10px; }
.rows-per-page { margin-left:20px; display:flex; align-items:center; }
.total-items { margin-left:15px; }
table { width:100%; border-collapse:collapse; margin-bottom:20px; }
th, td { padding:12px 15px; text-align:left; border-bottom:1px solid #ddd; }
th { background:#f2f2f2; font-weight:bold; }
.filter-row th { padding:5px; }
.column-filter { width:100%; padding:5px; border:1px solid #ddd; border-radius:4px; }
tr.danger { background:#ffebee; }
tr.warning { background:#fff8e1; }
tr.normal { background:#f1f8e9; }
tr.admin { background:#e3f2fd; }
tr.disabled { color:#999; font-style:italic; }
.status-icon { margin-right:5px; }
@media print {
.toolbar, button, #pagination, .filter-row { display:none; }
body { background:#fff; margin:0; }
.container { box-shadow:none; padding:0; }
.header { background:#000 !important; color:#fff !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
th { background:#f2f2f2 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
tr.danger { background:#ffebee !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
tr.warning { background:#fff8e1 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
tr.normal { background:#f1f8e9 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
tr.admin { background:#e3f2fd !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
}
</style>
<script>
window.quotaData = $json;
</script>
</head>
<body>
<div class="container">
<div class="header">
  <div class="header-icon">📊</div>
  <h1>OneDrive 基本データレポート <span style="font-size:22px;">🗂️</span></h1>
  <div style="margin-left:auto;">
    <label for="userSelect">👤 ユーザー選択：</label>
    <select id="userSelect"><option value="">-- ユーザーを選択してください --</option></select>
  </div>
</div>

<div class="info-section">
  <p><span class="info-label">📅 実行日時:</span> $executionDateFormatted</p>
  <p><span class="info-label">👤 実行者:</span> $executorName</p>
  <p><span class="info-label">🧑‍💼 実行者の種別:</span> $userType</p>
  <p><span class="info-label">🔑 実行モード:</span> $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})</p>
  <p><span class="info-label">📂 出力フォルダ:</span> $OutputDir</p>
</div>

<!-- 色の凡例 -->
<div class="info-section" style="margin-bottom:15px;">
  <span class="info-label">🎨 色の凡例：</span>
  <span style="display:inline-block;width:18px;height:18px;background:#e3f2fd;border:1px solid #90caf9;vertical-align:middle;margin-right:5px;"></span>
  <span style="vertical-align:middle;">管理者 <span style="font-size:18px;">👑</span></span>
  <span style="display:inline-block;width:18px;height:18px;background:#ffebee;border:1px solid #ef9a9a;vertical-align:middle;margin:0 5px 0 20px;"></span>
  <span style="vertical-align:middle;">危険（使用率90%以上）<span style="font-size:18px;">🔴</span></span>
  <span style="display:inline-block;width:18px;height:18px;background:#fff8e1;border:1px solid #ffe082;vertical-align:middle;margin:0 5px 0 20px;"></span>
  <span style="vertical-align:middle;">警告（使用率70%以上90%未満）<span style="font-size:18px;">🟡</span></span>
  <span style="display:inline-block;width:18px;height:18px;background:#f1f8e9;border:1px solid #aed581;vertical-align:middle;margin:0 5px 0 20px;"></span>
  <span style="vertical-align:middle;">正常（使用率70%未満）<span style="font-size:18px;">🟢</span></span>
  <span style="color:#999;margin-left:20px;">無効アカウント <span style="font-size:18px;">🚫</span></span>
</div>

<div class="toolbar">
<input type="text" id="searchInput" placeholder="🔍 検索...">
<div id="searchSuggestions"></div>
<button id="exportBtn"><span class="button-icon">📥</span>CSVエクスポート</button>
<button id="printBtn"><span class="button-icon">🖨️</span>印刷</button>
</div>

<!-- 凡例＋表示件数 -->
<div style="margin-bottom:10px; display:flex; align-items:center; gap:20px;">
  <span style="display:inline-block;width:18px;height:18px;background:#e3f2fd;border:1px solid #90caf9;vertical-align:middle;margin-right:5px;"></span>
  <span style="vertical-align:middle;">管理者（Administrator）</span>
<div class="rows-per-page" style="margin-left:20px; display:flex; align-items:center;">
    <label for="rowsPerPageSelect">表示件数:</label>
    <select id="rowsPerPageSelect">
      <option value="25" selected>25件</option>
      <option value="50">50件</option>
      <option value="100">100件</option>
    </select>
    <span id="currentCountInfo" style="margin-left:15px; font-weight:bold;"></span>
</div>
</div>

<div id="pagination"></div>

<table id="quotaTable">
<thead>
<tr>
"@
foreach ($col in $columns) {
    $html += "<th>$col</th>"
}
$html += "</tr>`n<tr class='filter-row'>"
foreach ($col in $columns) {
    $html += "<th><select class='column-filter'><option value=''>すべて</option>"
    foreach ($val in $uniqueValues[$col]) {
        if ($val -ne $null -and $val -ne "") {
            $escaped = [System.Net.WebUtility]::HtmlEncode($val)
            $html += "<option value='$escaped'>$escaped</option>"
        }
    }
    $html += "</select></th>"
}
$html += "</tr>
</thead>
<tbody>
</tbody>
</table>
"

# 2. 追加HTML/JS部分は@' ... '@で連結
$html += @'
<!-- 取得情報表示 -->
<div id="reportInfo" class="info-section" style="margin-top:10px;"></div>
</div>
<script>
(function() {
  // --- グローバル変数 ---
  let quotaData = window.quotaData || [];
  let filteredData = [];
  let currentPage = 1;
  let rowsPerPage = 25;
  let headers = [];

  // --- ヘッダー取得 ---
  function getHeaders() {
    const table = document.getElementById('quotaTable');
    if (!table) return [];
    return Array.from(table.tHead.rows[0].cells).map(cell => cell.textContent.trim());
  }

  // --- テーブル再描画 ---
  function renderTable(data) {
    const table = document.getElementById('quotaTable');
    const tbody = table ? table.tBodies[0] : null;
    if (!tbody) return;
    tbody.innerHTML = "";
    const total = data.length;
    const totalPages = Math.ceil(total / rowsPerPage);
    if (currentPage > totalPages) currentPage = 1;
    const start = (currentPage - 1) * rowsPerPage;
    const end = start + rowsPerPage;
    const pageData = data.slice(start, end);

    pageData.forEach(row => {
      const tr = document.createElement("tr");
      headers.forEach(header => {
        const td = document.createElement("td");
        td.textContent = row[header] || '';
        tr.appendChild(td);
      });
      if (row["ユーザー種別"] === "管理者") tr.classList.add("admin");
      if (row["状態"] === "危険") tr.classList.add("danger");
      else if (row["状態"] === "警告") tr.classList.add("warning");
      else tr.classList.add("normal");
      tbody.appendChild(tr);
    });
    updateOutputCountInfo(data);
    renderPagination(data);
  }

  // --- 件数・ページ情報 ---
  function updateOutputCountInfo(data) {
    const currentCountInfo = document.getElementById('currentCountInfo');
    const total = data.length;
    const totalPages = Math.ceil(total / rowsPerPage);
    const start = (currentPage - 1) * rowsPerPage + 1;
    const end = Math.min(currentPage * rowsPerPage, total);
    currentCountInfo.textContent =
      `表示中: ${start}～${end}件 / 総件数: ${total}件（${totalPages}ページ中${currentPage}ページ）`;
    document.getElementById('reportInfo').innerHTML =
      `<span class="info-label">表示中:</span> ${start}～${end}件 / <span class="info-label">総件数:</span> ${total}件（${totalPages}ページ中${currentPage}ページ）`;
  }

  // --- ページネーション ---
  function renderPagination(data) {
    const pagination = document.getElementById('pagination');
    const total = data.length;
    const totalPages = Math.ceil(total / rowsPerPage);
    pagination.innerHTML = "";
    if (totalPages > 1) {
      const prevBtn = document.createElement('button');
      prevBtn.textContent = '前へ';
      prevBtn.disabled = currentPage === 1;
      prevBtn.onclick = () => { if (currentPage > 1) { currentPage--; renderTable(data); }};
      pagination.appendChild(prevBtn);

      const startPage = Math.max(1, currentPage - 5);
      const endPage = Math.min(totalPages, startPage + 9);
      for (let i = startPage; i <= endPage; i++) {
        const pageBtn = document.createElement('button');
        pageBtn.textContent = i;
        pageBtn.disabled = i === currentPage;
        pageBtn.onclick = () => { currentPage = i; renderTable(data); };
        pagination.appendChild(pageBtn);
      }

      const nextBtn = document.createElement('button');
      nextBtn.textContent = '次へ';
      nextBtn.disabled = currentPage === totalPages;
      nextBtn.onclick = () => { if (currentPage < totalPages) { currentPage++; renderTable(data); }};
      pagination.appendChild(nextBtn);
    }
  }

  // --- 検索候補生成 ---
  function updateSearchSuggestions(searchTerm, results) {
    const searchSuggestions = document.getElementById('searchSuggestions');
    searchSuggestions.innerHTML = '';
    if (!searchTerm || results.length === 0) {
      searchSuggestions.style.display = 'none';
      return;
    }
    const uniqueValues = new Set();
    results.forEach(row => {
      Object.values(row).forEach(val => {
        const strVal = String(val);
        if (strVal.toLowerCase().includes(searchTerm.toLowerCase()) && !uniqueValues.has(strVal)) {
          uniqueValues.add(strVal);
          const suggestion = document.createElement('div');
          suggestion.className = 'suggestion-item';
          suggestion.textContent = strVal;
          suggestion.addEventListener('click', () => {
            document.getElementById('searchInput').value = strVal;
            searchSuggestions.style.display = 'none';
            filteredData = quotaData.filter(row =>
              Object.values(row).some(v => String(v).toLowerCase().includes(strVal.toLowerCase()))
            );
            currentPage = 1;
            renderTable(filteredData);
          });
          searchSuggestions.appendChild(suggestion);
        }
      });
    });
    if (uniqueValues.size > 0) {
      searchSuggestions.style.display = 'block';
    } else {
      searchSuggestions.innerHTML = '<div class="suggestion-item no-results">該当する結果がありません</div>';
      searchSuggestions.style.display = 'block';
    }
  }

  // --- 検索実行 ---
  function executeSearch() {
    const searchInput = document.getElementById('searchInput');
    const searchTerm = searchInput.value.trim();
    if (!searchTerm) {
      filteredData = [];
      currentPage = 1;
      renderTable(quotaData);
      updateSearchSuggestions('', []);
      return;
    }
    filteredData = quotaData.filter(row =>
      Object.values(row).some(val =>
        String(val).toLowerCase().includes(searchTerm.toLowerCase())
      )
    );
    currentPage = 1;
    renderTable(filteredData);
    updateSearchSuggestions(searchTerm, filteredData);
  }

  // --- フィルター実行 ---
  function executeFilter() {
    const filters = document.querySelectorAll('.column-filter');
    let data = quotaData;
    filters.forEach((filter, idx) => {
      const value = filter.value;
      if (value) {
        const col = headers[idx];
        data = data.filter(row => String(row[col]) === value);
      }
    });
    filteredData = data;
    currentPage = 1;
    renderTable(filteredData.length > 0 ? filteredData : quotaData);
    // フィルター情報表示
    const info = Array.from(filters).map((f, i) => f.value ? `${headers[i]}=${f.value}` : '').filter(Boolean).join(', ');
    document.getElementById('reportInfo').innerHTML += info ? `<br><span class="info-label">フィルター:</span> ${info}` : '';
  }

  // --- イベント初期化 ---
  function setupEventListeners() {
    // 検索
    const searchInput = document.getElementById('searchInput');
    const searchSuggestions = document.getElementById('searchSuggestions');
    searchInput.addEventListener('input', () => {
      executeSearch();
    });
    searchInput.addEventListener('keyup', (e) => {
      if (e.key === 'Enter') executeSearch();
    });
    // 検索候補外クリックで閉じる
    document.addEventListener('click', (e) => {
      if (!searchInput.contains(e.target) && !searchSuggestions.contains(e.target)) {
        searchSuggestions.style.display = 'none';
      }
    });

    // プルダウンフィルター
    document.querySelectorAll('.column-filter').forEach(filter => {
      filter.addEventListener('change', () => {
        executeFilter();
      });
    });

    // 表示件数
    const rowsPerPageSelect = document.getElementById('rowsPerPageSelect');
    rowsPerPageSelect.addEventListener('change', function() {
      rowsPerPage = parseInt(this.value, 10);
      currentPage = 1;
      renderTable(filteredData.length > 0 ? filteredData : quotaData);
    });
  }

  // --- 検索リセット ---
  function setupResetButton() {
    const searchInput = document.getElementById('searchInput');
    const searchSuggestions = document.getElementById('searchSuggestions');
    const resetSearch = document.createElement('button');
    resetSearch.innerHTML = '<span class="button-icon">↻</span>検索リセット';
    resetSearch.style.marginLeft = '10px';
    resetSearch.onclick = function() {
      searchInput.value = '';
      searchSuggestions.style.display = 'none';
      filteredData = [];
      currentPage = 1;
      renderTable(quotaData);
    };
    document.querySelector('.toolbar').appendChild(resetSearch);
  }

  // --- 初期化 ---
  function initialize() {
    headers = getHeaders();
    filteredData = [];
    currentPage = 1;
    const rowsPerPageSelect = document.getElementById('rowsPerPageSelect');
    rowsPerPage = parseInt(rowsPerPageSelect.value, 10) || 25;
    renderTable(quotaData);
    setupEventListeners();
    setupResetButton();
  }

  // --- 即時初期化 ---
  initialize();
})();
</script>
</body>
</html>
'@

$html | Out-File -FilePath $OutputHtmlPath -Encoding UTF8
Write-Host "HTMLファイルを作成しました: $OutputHtmlPath" -ForegroundColor Green
