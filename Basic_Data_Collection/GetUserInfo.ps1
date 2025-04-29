# OneDrive for Business 運用ツール - ITSM準拠
# GetUserInfo.ps1 - ユーザー情報取得スクリプト（管理者判定強化＋高機能HTML版）

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$ConfigPath = "$PSScriptRoot\..\config.json"
)

# SafeExitModuleをインポート (ルートディレクトリ直下にあるためパス修正)
$modulePath = "$PSScriptRoot\..\SafeExitModule.psm1"
if (-not (Test-Path $modulePath)) {
    Write-Error "SafeExitModuleが見つかりません: $modulePath"
    exit 1
}
Import-Module $modulePath -Force

$executionTime = Get-Date

# 管理者ロールID辞書
$adminRoleIds = @{
    "GlobalAdministrator" = "62e90394-69f5-4237-9190-012177145e10"
    "UserAccountAdministrator" = "fe930be7-5e62-47db-91af-98c3a49a38b1"
    "ExchangeAdministrator" = "29232cdf-9323-42fd-ade2-1d097af3e4de"
    "SharePointAdministrator" = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
    "TeamsAdministrator" = "69091246-20e8-4a56-aa4d-066075b2a7a8"
    "SecurityAdministrator" = "194ae4cb-b126-40b2-bd5b-6091b380977d"
}

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
$reportDir = Join-Path -Path $dataCollectionDir -ChildPath "GetUserInfo"

# ディレクトリ作成（自動生成フォルダを含む）
try {
    # Basic_Data_Collection.日付フォルダが存在しない場合は作成
    if (-not (Test-Path -Path $dataCollectionDir)) {
        New-Item -Path $dataCollectionDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "Basic_Data_Collectionフォルダを作成しました: $dataCollectionDir" "INFO"
    }

    # GetUserInfoフォルダが存在しない場合は作成
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -Path $reportDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "GetUserInfoフォルダを作成しました: $reportDir" "INFO"
    }
    
    # GetUserInfoサブフォルダが存在しない場合は作成
    if (-not (Test-Path -Path $reportDir)) {
        New-Item -Path $reportDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "レポート出力ディレクトリを作成しました: $reportDir" "INFO"
    }
} catch {
    Write-Log "ディレクトリ作成エラー: $_" "ERROR"
    throw
}

Write-Log "ユーザー情報取得を開始します" "INFO"
Write-Log "ベースディレクトリ: $BaseDir" "INFO"
Write-Log "出力ルートディレクトリ: $outputRootDir" "INFO"
Write-Log "レポートディレクトリ: $reportDir" "INFO"
Write-Host "出力先ディレクトリ: $outputRootDir" -ForegroundColor Cyan

# Graph API接続（リトライ付き）
$maxRetries = 3
$retryCount = 0
$connected = $false

do {
    try {
        $config = Get-Content -Path $ConfigPath | ConvertFrom-Json
        $tokenUrl = "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token"
        $tokenBody = @{
            client_id     = $config.ClientId
            client_secret = $config.ClientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }

        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
        $script:AccessToken = $tokenResponse.access_token
        $connected = $true
        Write-Log "Microsoft Graphにクライアントシークレット認証で接続しました" -Level "SUCCESS"
    } catch {
        $retryCount++
        if ($retryCount -ge $maxRetries) {
            Write-Log "Microsoft Graph認証に失敗しました（リトライ上限）: $_" -Level "ERROR"
            Write-Log "スタックトレース: $($_.ScriptStackTrace)" -Level "DEBUG"
            exit 1
        }
        Write-Log "Microsoft Graph認証に失敗しました（リトライ $retryCount/$maxRetries）: $_" -Level "WARNING"
        Start-Sleep -Seconds (5 * $retryCount)  # 指数バックオフ
    }
} while (-not $connected -and $retryCount -lt $maxRetries)

$userList = @()

# ユーザー情報取得（ページネーション対応）
try {
    # メイン処理開始 - ユーザー情報取得
    Write-Log "Microsoft Graph REST APIで全ユーザー情報を取得します..." "INFO"
    $headers = @{ Authorization = "Bearer $script:AccessToken" }
    $url = "https://graph.microsoft.com/v1.0/users`?$top=999&`$select=displayName,mail,onPremisesSamAccountName,accountEnabled,onPremisesLastSyncDateTime,userType,userPrincipalName,id"
    $users = @()
    $retryCount = 0
    $maxApiRetries = 3

    do {
        try {
            $response = Invoke-RestMethod -Headers $headers -Uri $url -Method Get -ErrorAction Stop
            $users += $response.value
            $url = $response.'@odata.nextLink'
            $retryCount = 0  # 成功したらリトライカウントをリセット
        } catch {
            $retryCount++
            if ($retryCount -ge $maxApiRetries) {
                Write-Log "ユーザー情報取得に失敗しました（リトライ上限）: $_" -Level "ERROR"
                throw
            }
            Write-Log "ユーザー情報取得に失敗しました（リトライ $retryCount/$maxApiRetries）: $_" -Level "WARNING"
            Start-Sleep -Seconds (2 * $retryCount)
        }
    } while ($url)

    if ($users.Count -eq 0) {
        Write-Log "取得したユーザー情報が0件です" -Level "ERROR"
        exit 1
    }

    $totalUsers = $users.Count
    $processed = 0
    
    # 取得したユーザーデータを処理
    foreach ($user in $users) {
        try {
            Write-Log "ユーザー処理中: $($user.userPrincipalName)" "DEBUG"
            $processed++
            $progressPercent = [math]::Round(($processed / $totalUsers * 100), 2)
            $progressParams = @{
                Activity = "ユーザー情報を処理中"
                Status = "$processed/$totalUsers 完了 ($progressPercent%)"
                PercentComplete = $progressPercent
                CurrentOperation = "処理中: $($user.userPrincipalName)"
            }
            Write-Progress @progressParams
            
            if ($processed % 10 -eq 0 -or $processed -eq $totalUsers) {
                Write-Host "[$processed/$totalUsers] $($user.userPrincipalName)" -ForegroundColor DarkGray
            }

            $userTypeValue = "Member"
            if ($user.userPrincipalName -match "#EXT#" -or $user.userType -eq "Guest") {
                $userTypeValue = "Guest"
            } elseif ([string]::IsNullOrEmpty($user.id)) {
                $userTypeValue = "未設定"
            } else {
                try {
                    $memberOfUrl = "https://graph.microsoft.com/v1.0/users/$($user.id)/memberOf"
                    $memberOfResponse = Invoke-RestMethod -Headers $headers -Uri $memberOfUrl -Method Get
                    $memberOf = $memberOfResponse.value
                    $isAdmin = $false
                    foreach ($role in $memberOf) {
                        if ($role.'@odata.type' -eq "#microsoft.graph.directoryRole") {
                            $roleTemplateUrl = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)"
                            $roleDetail = Invoke-RestMethod -Headers $headers -Uri $roleTemplateUrl -Method Get
                            if ($adminRoleIds.Values -contains $roleDetail.roleTemplateId) {
                                $isAdmin = $true
                                break
                            }
                        }
                    }
                    if ($isAdmin) { $userTypeValue = "Administrator" }
                } catch {
                    Write-Log "ユーザー種別確認エラー: $($user.userPrincipalName) - $_" "WARNING"
                    $userTypeValue = "確認エラー"
                }
            }

            $userList += [PSCustomObject]@{
                "ユーザー名"       = $user.displayName
                "メールアドレス"   = $user.mail
                "ログインユーザー名" = if($user.onPremisesSamAccountName){$user.onPremisesSamAccountName}else{"同期なし"}
                "ユーザー種別"     = $userTypeValue
                "アカウント状態"   = if($user.accountEnabled){"有効"}else{"無効"}
                "最終同期日時"     = if($user.onPremisesLastSyncDateTime){$user.onPremisesLastSyncDateTime}else{"同期情報なし"}
            }
        } catch {
            Write-Log "ユーザー処理中にエラーが発生しました: $($user.userPrincipalName) - $_" "WARNING"
            continue
        }
    }

    Write-Progress -Activity "ユーザー情報処理" -Completed
    Write-Log "ユーザー情報の取得が完了しました。取得件数: $($userList.Count)" "SUCCESS"
    Write-Host "`nユーザー情報処理が完了しました: $($userList.Count)件のレコードを取得" -ForegroundColor Green

    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $csvFile = "UserInfo.$timestamp.csv"
    $htmlFile = "UserInfo.$timestamp.html"
    $jsFile = "UserInfo.$timestamp.js"

    $csvPath = Join-Path -Path $reportDir -ChildPath $csvFile
    $htmlPath = Join-Path -Path $reportDir -ChildPath $htmlFile
    $jsPath = Join-Path -Path $reportDir -ChildPath $jsFile

    Write-Log "CSV出力先: $csvPath" "DEBUG"
    Write-Log "HTML出力先: $htmlPath" "DEBUG"
    Write-Log "JS出力先: $jsPath" "DEBUG"
    Write-Host "CSV出力先: $csvPath" -ForegroundColor Cyan
    Write-Host "HTML出力先: $htmlPath" -ForegroundColor Cyan
    # ファイル出力処理
    try {
        # CSVファイル出力（BOM付きUTF8を確実に適用）
        try {
            # メモリストリームを使用して確実にBOMを付与
            $utf8WithBom = [System.Text.Encoding]::UTF8
            $memoryStream = New-Object System.IO.MemoryStream
            $streamWriter = New-Object System.IO.StreamWriter($memoryStream, $utf8WithBom)
            
            # CSVヘッダー書き込み（ダブルクォート囲み）
            $headers = $userList[0].PSObject.Properties |
                       Select-Object -ExpandProperty Name |
                       ForEach-Object { '"{0}"' -f $_ }
            $streamWriter.WriteLine(($headers -join ","))
            
            # データ行書き込み（値のエスケープ処理）
            foreach ($user in $userList) {
                $line = $user.PSObject.Properties |
                        Select-Object -ExpandProperty Value |
                        ForEach-Object {
                            $val = $_ -replace '"', '""'  # ダブルクォートをエスケープ
                            '"{0}"' -f $val  # 値をダブルクォートで囲む
                        }
                $streamWriter.WriteLine(($line -join ","))
            }
            
            $streamWriter.Flush()
            $memoryStream.Position = 0
            
            # ファイルに書き込み
            $fileStream = [System.IO.File]::Create($csvPath)
            $memoryStream.WriteTo($fileStream)
            
            $streamWriter.Close()
            $memoryStream.Close()
            $fileStream.Close()
            
            Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"
        } catch {
            Write-Log "CSVファイル出力エラー: $_" "ERROR"
            if ($streamWriter) { $streamWriter.Dispose() }
            if ($memoryStream) { $memoryStream.Dispose() }
            if ($fileStream) { $fileStream.Dispose() }
            throw
        }
        
        try {
            # HTMLファイル生成（ユーザーデータをJSONで埋め込み）
            $jsonData = $userList | ConvertTo-Json -Depth 5 -Compress
            $htmlTemplate = @"
<!DOCTYPE html>
<html>
<head>
<title>ユーザー情報レポート</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<script>
const userData = $jsonData;
</script>
<script src="common.js"></script>
</head>
<body>
<div class="container">
    <h1>ユーザー情報レポート</h1>
    <div id="userTableContainer"></div>
</div>
</body>
</html>
"@
            $htmlTemplate | Out-File -FilePath $htmlPath -Encoding UTF8 -Force
            Write-Log "HTMLファイルを作成しました: $htmlPath" "SUCCESS"
        } catch {
            Write-Log "HTMLファイル生成エラー: $_" "ERROR"
            throw
        }

# HTMLテンプレート生成と保存
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
<title>ユーザー情報レポート</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
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
  --login-color: #ba68c8;
  --type-color: #81c784;
  --status-color: #ffb74d;
  --lastsync-color: #4dd0e1;
}

.column-user {
  background: linear-gradient(135deg, var(--user-color), #42a5f5);
  color: white;
}
.column-email {
  background: linear-gradient(135deg, var(--email-color), #ffca28);
  color: #333;
}
.column-login {
  background: linear-gradient(135deg, var(--login-color), #9c27b0);
  color: white;
}
.column-type {
  background: linear-gradient(135deg, var(--type-color), #43a047);
  color: white;
}
.column-status {
  background: linear-gradient(135deg, var(--status-color), #ff8a65);
  color: white;
}
.column-lastsync {
  background: linear-gradient(135deg, var(--lastsync-color), #26c6da);
  color: white;
}

.status-active {
  background: linear-gradient(135deg, #69f0ae, #00c853);
  color: white;
  font-weight: bold;
}
.status-inactive {
  background: linear-gradient(135deg, #ffd740, #ffa000);
  color: #333;
  font-weight: bold;
}
.status-disabled {
  background: linear-gradient(135deg, #ff5252, #c62828);
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
  --login-color: #ba68c8;
  --type-color: #81c784;
  --status-color: #ffb74d;
  --lastsync-color: #4dd0e1;
}

.column-user {
  background: linear-gradient(135deg, var(--user-color), #42a5f5);
  color: white;
}
.column-email {
  background: linear-gradient(135deg, var(--email-color), #ffca28);
  color: #333;
}
.column-login {
  background: linear-gradient(135deg, var(--login-color), #9c27b0);
  color: white;
}
.column-type {
  background: linear-gradient(135deg, var(--type-color), #43a047);
  color: white;
}
.column-status {
  background: linear-gradient(135deg, var(--status-color), #ff8a65);
  color: white;
}
.column-lastsync {
  background: linear-gradient(135deg, var(--lastsync-color), #26c6da);
  color: white;
}

.status-active {
  background: linear-gradient(135deg, #69f0ae, #00c853);
  color: white;
  font-weight: bold;
}
.status-inactive {
  background: linear-gradient(135deg, #ffd740, #ffa000);
  color: #333;
  font-weight: bold;
}
.status-disabled {
  background: linear-gradient(135deg, #ff5252, #c62828);
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
th, td { width: 16.66%; }
#pagination { display: flex; justify-content: center; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 20px; }
.page-info { margin: 0 10px; }
@media print {
  .toolbar, button, #pagination, .filter-row { display: none; }
  .container { box-shadow: none; padding: 0; }
}
</style>
<script>
let currentPage = 1;
let rowsPerPage = 10;
let filteredRows = [];

function searchTable() {
    var input = document.getElementById('searchInput').value.toLowerCase();
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    filteredRows = [];
    for (var i = 0; i < rows.length; i++) {
        var found = false;
        var cells = rows[i].getElementsByTagName('td');
        for (var j = 0; j < cells.length; j++) {
            var cellText = cells[j].textContent || cells[j].innerText;
            if (cellText.toLowerCase().indexOf(input) > -1) { found = true; }
        }
        if (found) { filteredRows.push(rows[i]); }
    }
    currentPage = 1;
    updatePagination();
}

function updatePagination() {
    var table = document.getElementById('userTable');
    var tbody = table.getElementsByTagName('tbody')[0];
    var rows = tbody.getElementsByTagName('tr');
    for (var i = 0; i < rows.length; i++) { rows[i].style.display = 'none'; }
    var startIndex = (currentPage - 1) * rowsPerPage;
    var endIndex = Math.min(startIndex + rowsPerPage, filteredRows.length);
    for (var i = startIndex; i < endIndex; i++) { filteredRows[i].style.display = ''; }
    updatePaginationControls();
}

function updatePaginationControls() {
    var paginationDiv = document.getElementById('pagination');
    paginationDiv.innerHTML = '';
    var totalPages = Math.ceil(filteredRows.length / rowsPerPage);
    var prevButton = document.createElement('button');
    prevButton.innerHTML = '前へ';
    prevButton.disabled = currentPage === 1;
    prevButton.onclick = function() { if(currentPage>1){currentPage--;updatePagination();} };
    paginationDiv.appendChild(prevButton);
    var pageInfo = document.createElement('span');
    pageInfo.textContent = currentPage + ' / ' + (totalPages || 1) + 'ページ';
    paginationDiv.appendChild(pageInfo);
    var nextButton = document.createElement('button');
    nextButton.innerHTML = '次へ';
    nextButton.disabled = currentPage === totalPages || totalPages === 0;
    nextButton.onclick = function() { if(currentPage<totalPages){currentPage++;updatePagination();} };
    paginationDiv.appendChild(nextButton);

    var select = document.createElement('select');
    [10,20,50,100].forEach(function(n){
        var opt = document.createElement('option');
        opt.value = n;
        opt.text = n + '件';
        if(n===rowsPerPage) opt.selected = true;
        select.appendChild(opt);
    });
    select.onchange = function() {
        rowsPerPage = parseInt(this.value);
        currentPage = 1;
        updatePagination();
    };
    paginationDiv.appendChild(select);

    var totalCount = document.createElement('span');
    totalCount.textContent = ' 全 ' + filteredRows.length + ' 件';
    paginationDiv.appendChild(totalCount);
}

function exportCSV() {
    var csv = [];
    var rows = document.querySelectorAll("table tr");
    for (var i = 0; i < rows.length; i++) {
        var row = [], cols = rows[i].querySelectorAll("td, th");
        for (var j = 0; j < cols.length; j++)
            row.push('"' + cols[j].innerText.replace(/"/g, '""') + '"');
        csv.push(row.join(","));
    }
    var csvFile = new Blob(["\uFEFF" + csv.join("\n")], { type: "text/csv;charset=utf-8;" });
    var link = document.createElement("a");
    link.href = URL.createObjectURL(csvFile);
    link.download = "UserInfo_Export.csv";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function printTable() {
    window.print();
}

function createColumnFilters() {
    var table = document.getElementById('userTable');
    var headers = table.getElementsByTagName('thead')[0].getElementsByTagName('th');
    var filterRow = table.getElementsByClassName('filter-row')[0];
    for (var col = 0; col < headers.length; col++) {
        var select = filterRow.children[col].querySelector('select');
        var uniqueValues = new Set();
        var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
        for (var i = 0; i < rows.length; i++) {
            var cellText = rows[i].getElementsByTagName('td')[col].textContent.trim();
            uniqueValues.add(cellText);
        }
        uniqueValues = Array.from(uniqueValues).sort();
        uniqueValues.forEach(function(val){
            var opt = document.createElement('option');
            opt.value = val;
            opt.text = val;
            select.appendChild(opt);
        });
        select.onchange = applyColumnFilters;
    }
}

function applyColumnFilters() {
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var selects = document.getElementsByClassName('column-filter');
    filteredRows = [];
    for (var i = 0; i < rows.length; i++) {
        var show = true;
        for (var j = 0; j < selects.length; j++) {
            var filterVal = selects[j].value;
            var cellText = rows[i].getElementsByTagName('td')[j].textContent.trim();
            if (filterVal !== "" && cellText !== filterVal) {
                show = false;
                break;
            }
        }
        if (show) filteredRows.push(rows[i]);
    }
    currentPage = 1;
    updatePagination();
}

window.onload = function() {
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    filteredRows = Array.from(rows);
    createColumnFilters();
    updatePagination();
    document.getElementById('searchInput').addEventListener('keyup', searchTable);
    document.getElementById('exportBtn').addEventListener('click', exportCSV);
    document.getElementById('printBtn').addEventListener('click', printTable);
};
</script>
</head>
<body>
<div class="container">
<div class="header"><div class="header-icon">👥</div><h1>ユーザー情報レポート</h1></div>
<div class="info-section">
<p><span class="info-label">実行日時:</span> $($executionTime.ToString("yyyy/MM/dd HH:mm:ss"))</p>
<p><span class="info-label">出力フォルダ:</span> $OutputDir</p>
</div>
<div class="toolbar">
<input type="text" id="searchInput" placeholder="検索...">
<button id="exportBtn">CSVエクスポート</button>
<button id="printBtn">印刷</button>
</div>
<div id="pagination"></div>
<table id="userTable">
<thead>
<tr class="filter-row">
<th style="min-width: 80px; padding-right: 40px;"><select class="column-filter" data-column="0"><option value="" style="font-weight:bold; font-size:16px;">表示名</option></select></th>
<th style="min-width: 150px; padding-right: 40px;"><select class="column-filter" data-column="1"><option value="" style="font-weight:bold; font-size:16px;">メール</option></select></th>
<th style="min-width: 100px; padding-right: 40px;"><select class="column-filter" data-column="2"><option value="" style="font-weight:bold; font-size:16px;">ログインユーザー名</option></select></th>
<th style="padding-right: 20px;"><select class="column-filter" data-column="3"><option value="" style="font-weight:bold; font-size:16px;">ユーザー種別</option></select></th>
<th style="padding-right: 20px;"><select class="column-filter" data-column="4"><option value="" style="font-weight:bold; font-size:16px;">アカウント状態</option></select></th>
<th style="padding-right: 20px;"><select class="column-filter" data-column="5"><option value="" style="font-weight:bold; font-size:16px;">最終同期日時</option></select></th>
</tr>
</thead>
<tbody>
"@

foreach ($user in $userList) {
    $statusClass = switch ($user.'アカウント状態') {
        "有効" { "status-active" }
        "無効" { "status-inactive" }
        default { "" }
    }
    $htmlContent += @"
<tr class="$statusClass">
<td>$($user.'ユーザー名')</td>
<td>$($user.'メールアドレス')</td>
<td>$($user.'ログインユーザー名')</td>
<td>$($user.'ユーザー種別')</td>
<td>$($user.'アカウント状態')</td>
<td>$($user.'最終同期日時')</td>
</tr>
"@
    } # メインtryブロック終了
} catch {
    Write-Log "メイン処理でエラーが発生しました: $_" -Level "ERROR"
    Write-Log "スタックトレース: $($_.ScriptStackTrace)" -Level "DEBUG"
    exit 1
}

try {
    $htmlContent += @"
</tbody>
</table>
</div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Log "HTMLファイルを作成しました: $htmlPath" "SUCCESS"
} catch {
    Write-Log "HTMLファイル生成エラー: $_" "ERROR"
    throw
}
        } catch {
            Write-Log "HTMLファイル生成エラー: $_" "ERROR"
            throw
        }