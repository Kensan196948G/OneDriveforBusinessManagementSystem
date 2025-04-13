# OneDrive for Business 運用ツール - ITSM準拠
# GetOneDriveQuotaBackUp.ps1 - ストレージクォータ取得スクリプト（ユニーク値プルダウン・出力同期対応）

param (
    [string]$OutputDir = "$(Get-Location)\Output",
    [string]$LogDir = "$(Get-Location)\Log",
    [string]$AccessToken = $null
)

# 出力ディレクトリとログディレクトリ作成
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

# 実行日時・タイムスタンプ（HTML/CSVで完全同期）
$executionDateFormatted = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# 実行者情報（仮の値、必要に応じて取得処理を追加）
$executorName = "管理者"
$userType = "Member"
$isAdmin = $true

# ユーザーリスト取得処理
$userList = @()
$users = Get-MgUser -All -Property DisplayName,Mail,UserPrincipalName,AccountEnabled,UserType,onPremisesSamAccountName

foreach ($u in $users) {
    $userTypeValue = if ($u.UserType -eq "Guest") { "Guest" } else { "Member" }

    # 管理者判定
    try {
        $roles = Get-MgUserMemberOf -UserId $u.Id -ErrorAction Stop
        foreach ($role in $roles) {
            if ($role.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole') {
                $roleDetail = Get-MgDirectoryRole -DirectoryRoleId $role.Id
                if ($roleDetail.RoleTemplateId -eq '62e90394-69f5-4237-9190-012177145e10') {
                    $userTypeValue = "Administrator"
                    break
                }
            }
        }
    } catch {
        # 取得失敗時は既定値のまま
    }
    $accountStatus = if ($u.AccountEnabled) { "有効" } else { "無効" }
    $loginName = if ($u.onPremisesSamAccountName) { $u.onPremisesSamAccountName } else { "同期なし" }

    $onedriveStatus = "未対応"
    $totalGB = "取得不可"
    $usedGB = "取得不可"
    $remainingGB = "取得不可"
    $usagePercent = "取得不可"
    $status = "不明"

    try {
        $drive = Get-MgUserDrive -UserId $u.Id -ErrorAction Stop
        $onedriveStatus = "対応"
        $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
        $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
        $remainingGB = [math]::Round(($drive.Quota.Total - $drive.Quota.Used) / 1GB, 2)
        $usagePercent = if ($drive.Quota.Total -ne 0) { [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 2) } else { 0 }
        $status = if ($usagePercent -ge 90) { "危険" } elseif ($usagePercent -ge 70) { "警告" } else { "正常" }
    } catch {
        # OneDrive未対応や取得失敗時は既定値のまま
    }

    $userList += [PSCustomObject]@{
        "ユーザー名" = $u.DisplayName
        "メールアドレス" = $u.Mail
        "ログインユーザー名" = $loginName
        "ユーザー種別" = $userTypeValue
        "アカウント状態" = $accountStatus
        "OneDrive対応" = $onedriveStatus
        "総容量(GB)" = $totalGB
        "使用容量(GB)" = $usedGB
        "残り容量(GB)" = $remainingGB
        "使用率(%)" = $usagePercent
        "状態" = $status
    }
}

# 列名リスト
$columns = @(
    "ユーザー名", "メールアドレス", "ログインユーザー名", "ユーザー種別", "アカウント状態",
    "OneDrive対応", "総容量(GB)", "使用容量(GB)", "残り容量(GB)", "使用率(%)", "状態"
)

# 各列のユニーク値リストを作成
$uniqueValues = @{}
foreach ($col in $columns) {
    $uniqueValues[$col] = $userList | Select-Object -ExpandProperty $col | Sort-Object | Get-Unique
}

# 出力ファイル名（同期）
$htmlPath = Join-Path $OutputDir "OneDriveQuota.$timestamp.html"
$csvPath = Join-Path $OutputDir "OneDriveQuota.$timestamp.csv"

# HTMLテンプレート
$html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>OneDriveクォータレポート</title>
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
</head>
<body>
<div class="container">
<div class="header">
<div class="header-icon">💾</div>
<h1>OneDriveクォータレポート</h1>
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

<div class="toolbar">
<input type="text" id="searchInput" placeholder="🔍 検索...">
<div id="searchSuggestions"></div>
<button id="exportBtn"><span class="button-icon">📥</span>CSVエクスポート</button>
<button id="printBtn"><span class="button-icon">🖨️</span>印刷</button>
<div class="rows-per-page">
<label for="rowsPerPageSelect">表示件数:</label>
<select id="rowsPerPageSelect">
<option value="10" selected>10件</option>
<option value="50">50件</option>
<option value="100">100件</option>
</select>
</div>
</div>

<div id="pagination"></div>

<!-- 管理者凡例 -->
<div style="margin-bottom:10px;">
  <span style="display:inline-block;width:18px;height:18px;background:#e3f2fd;border:1px solid #90caf9;vertical-align:middle;margin-right:5px;"></span>
  <span style="vertical-align:middle;">管理者（Administrator）</span>
</div>

<table id="quotaTable">
<thead>
<tr>
"@

# ヘッダー
foreach ($col in $columns) {
    $html += "<th>$col</th>"
}
$html += "</tr>`n<tr class='filter-row'>"
# 各列のユニーク値でプルダウン生成
foreach ($col in $columns) {
    $html += "<th><select class='column-filter'><option value=''>すべて</option>"
    foreach ($val in $uniqueValues[$col]) {
        if ($val -ne $null -and $val -ne "") {
            $html += "<option value='$val'>$val</option>"
        }
    }
    $html += "</select></th>"
}
$html += "</tr>
</thead>
<tbody>
"

# データ行
foreach ($user in $userList) {
    $rowClass = ""
    if ($user.'ユーザー種別' -eq "Administrator") { $rowClass = " class='admin'" }
    $html += "<tr$rowClass>"
    foreach ($col in $columns) {
        $html += "<td>$($user.$col)</td>"
    }
    $html += "</tr>"
}

$html += @"
</tbody>
</table>
</div>
</body>
</html>
"@

# デバッグ: ユーザー件数とHTML長さ
Write-Host "デバッグ: 取得ユーザー数 = $($userList.Count)" -ForegroundColor Yellow
Write-Host "デバッグ: HTML文字数 = $($html.Length)" -ForegroundColor Yellow

# HTML出力
try {
    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "HTMLファイルを作成しました: $htmlPath" -ForegroundColor Green
} catch {
    Write-Error "HTMLファイルの出力に失敗しました: $_"
}

# CSVも必ず同じタイムスタンプで出力
try {
    $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    Write-Host "CSVファイルを作成しました: $csvPath" -ForegroundColor Green
} catch {
    Write-Error "CSVファイルの出力に失敗しました: $_"
}
