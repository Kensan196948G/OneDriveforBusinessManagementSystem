# OneDrive for Business 運用ツール - ITSM準拠
# GetUserInfoNew.ps1 - ユーザー情報取得スクリプト（管理者判定強化版）

param (
    [string]$BaseDir = "$(Get-Location)",
    [string]$DateFolder = (Get-Date -Format "yyyyMMdd")
)

$executionTime = Get-Date

$outputRootDir = Join-Path -Path $BaseDir -ChildPath "OneDriveManagement.$DateFolder"
$logDir = Join-Path -Path $outputRootDir -ChildPath "Log"
$reportDir = Join-Path -Path $outputRootDir -ChildPath "Report"

if (-not (Test-Path -Path $outputRootDir)) { New-Item -ItemType Directory -Path $outputRootDir -Force | Out-Null }
if (-not (Test-Path -Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
if (-not (Test-Path -Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }

$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $logDir -ChildPath "GetUserInfo.$timestamp.log"
$errorLogPath = Join-Path -Path $logDir -ChildPath "GetUserInfo.Error.$timestamp.log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
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
    Write-Host $errorMessage -ForegroundColor Red
    Add-Content -Path $logFilePath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorDetails -Encoding UTF8
}

# 管理者ロールID辞書
$adminRoleIds = @{
    "GlobalAdministrator" = "62e90394-69f5-4237-9190-012177145e10"
    "UserAccountAdministrator" = "fe930be7-5e62-47db-91af-98c3a49a38b1"
    "ExchangeAdministrator" = "29232cdf-9323-42fd-ade2-1d097af3e4de"
    "SharePointAdministrator" = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
    "TeamsAdministrator" = "69091246-20e8-4a56-aa4d-066075b2a7a8"
    "SecurityAdministrator" = "194ae4cb-b126-40b2-bd5b-6091b380977d"
}

Write-Log "ユーザー情報取得を開始します" "INFO"
Write-Log "ベースディレクトリ: $BaseDir" "INFO"
Write-Log "出力ルートディレクトリ: $outputRootDir" "INFO"
Write-Log "レポートディレクトリ: $reportDir" "INFO"

try {
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath "..\config.json"
    $config = Get-Content -Path $configPath | ConvertFrom-Json

    $tokenUrl = "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token"
    $tokenBody = @{
        client_id     = $config.ClientId
        client_secret = $config.ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
    $script:AccessToken = $tokenResponse.access_token

    Write-Log "Microsoft Graphにクライアントシークレット認証で接続しました" "SUCCESS"
} catch {
    Write-ErrorLog $_ "Microsoft Graph認証に失敗しました"
    exit
}

$userList = @()

try {
    Write-Log "Microsoft Graph REST APIで全ユーザー情報を取得します..." "INFO"
    $headers = @{ Authorization = "Bearer $script:AccessToken" }
    $url = "https://graph.microsoft.com/v1.0/users`?$top=999&`$select=displayName,mail,onPremisesSamAccountName,accountEnabled,onPremisesLastSyncDateTime,userType,userPrincipalName,id"
    $users = @()

    do {
        $response = Invoke-RestMethod -Headers $headers -Uri $url -Method Get
        $users += $response.value
        $url = $response.'@odata.nextLink'
    } while ($url)

    foreach ($user in $users) {
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
    }

    Write-Log "ユーザー情報の取得が完了しました。取得件数: $($userList.Count)" "SUCCESS"
} catch {
    Write-ErrorLog $_ "ユーザー情報の取得中にエラーが発生しました"
}

$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$csvFile = "UserInfo.$timestamp.csv"
$htmlFile = "UserInfo.$timestamp.html"
$jsFile = "UserInfo.$timestamp.js"

$csvPath = Join-Path -Path $reportDir -ChildPath $csvFile
$htmlPath = Join-Path -Path $reportDir -ChildPath $htmlFile
$jsPath = Join-Path -Path $reportDir -ChildPath $jsFile

try {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    } else {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        $content = [System.IO.File]::ReadAllText($csvPath)
        [System.IO.File]::WriteAllText($csvPath, $content, [System.Text.Encoding]::UTF8)
    }
    Write-Log "CSVファイルを作成しました: $csvPath" "SUCCESS"

    try {
        Write-Log "Excelでファイルを開いて列幅の調整とフィルターの適用を行います..." "INFO"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Open($csvPath)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        $usedRange.AutoFilter() | Out-Null
        $excel.ActiveWindow.WindowState = -4143
        $workbook.Save()
        Write-Log "Excelでの処理が完了しました。" "SUCCESS"
    } catch {
        Write-Log "Excelでの処理中にエラーが発生しました: $_" "WARNING"
        Write-Log "CSVファイルは正常に作成されましたが、Excel処理はスキップされました。" "WARNING"
    }
} catch {
    Write-ErrorLog $_ "CSVファイルの作成中にエラーが発生しました"
}

$jsContent = @"
// UserInfo データ操作用 JavaScript
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
        var rowData = {};
        for (var j = 0; j < cells.length; j++) {
            var cellText = cells[j].textContent || cells[j].innerText;
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[j].textContent;
            rowData[headerText] = cellText;
            if (cellText.toLowerCase().indexOf(input) > -1) { found = true; }
        }
        if (found) { filteredRows.push({row: rows[i], data: rowData}); }
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
    for (var i = startIndex; i < endIndex; i++) { filteredRows[i].row.style.display = ''; }
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
}

window.onload = function() {
    var table = document.getElementById('userTable');
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
    document.getElementById('searchInput').addEventListener('keyup', searchTable);
};
"@

$jsContent | Out-File -FilePath $jsPath -Encoding UTF8
Write-Log "JavaScriptファイルを作成しました: $jsPath" "SUCCESS"

$executionDateFormatted = $executionTime.ToString("yyyy/MM/dd HH:mm:ss")
$executorName = "システム管理者"
$userType = "管理者"

$htmlContent = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>ユーザー情報レポート</title>
<script src="$jsFile"></script>
<style>
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
.container { background-color: white; padding: 20px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
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
table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #ddd; }
th { background-color: #f2f2f2; font-weight: bold; }
#pagination { display: flex; justify-content: center; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 20px; }
.page-info { margin: 0 10px; }
</style>
</head>
<body>
<div class="container">
<div class="header"><div class="header-icon">👥</div><h1>ユーザー情報レポート</h1></div>
<div class="info-section">
<p><span class="info-label">実行日時:</span> $executionDateFormatted</p>
<p><span class="info-label">実行者:</span> $executorName</p>
<p><span class="info-label">実行者の種別:</span> $userType</p>
<p><span class="info-label">出力フォルダ:</span> $reportDir</p>
</div>
<div class="toolbar">
<input type="text" id="searchInput" placeholder="検索...">
</div>
<div id="pagination"></div>
<table id="userTable">
<thead>
<tr>
<th>ユーザー名</th>
<th>メールアドレス</th>
<th>ログインユーザー名</th>
<th>ユーザー種別</th>
<th>アカウント状態</th>
<th>最終同期日時</th>
</tr>
</thead>
<tbody>
"@

foreach ($user in $userList) {
    $statusIcon = if ($user.'アカウント状態' -eq "有効") { "✅" } else { "❌" }
    $htmlContent += @"
<tr>
<td>$($user.'ユーザー名')</td>
<td>$($user.'メールアドレス')</td>
<td>$($user.'ログインユーザー名')</td>
<td>$($user.'ユーザー種別')</td>
<td><span class="status-icon">$ + ($user.'アカウント状態')</span>$ + ($user.'アカウント状態')</td>
<td>$ + ($user.'最終同期日時')</td>
</tr>
"@
}

$htmlContent += @"
</tbody>
</table>
</div>
</body>
</html>
"@

$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Log "HTMLファイルを作成しました: $htmlPath" "SUCCESS"
