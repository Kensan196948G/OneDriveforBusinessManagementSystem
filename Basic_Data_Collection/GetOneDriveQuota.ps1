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
    # 既存のセッション確認
    $context = Get-MgContext
    if (-not $context) {
        Write-Log "Microsoft Graphに接続されていません。Connect-MgGraphを実行します。" "WARNING"
        Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All","Directory.ReadWrite.All"
        $context = Get-MgContext
        if (-not $context) {
            Write-Log "Microsoft Graphへの接続に失敗しました。" "ERROR"
            exit
        }
    }

    # ログインユーザーのUPN（メールアドレス）を自動取得
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
                    Write-Log "ユーザー $($user.UserPrincipalName) のroleTemplateId: $($roleDetail.roleTemplateId)" "INFO"
                    if ($adminRoleIds.Values -contains $roleDetail.roleTemplateId) {
                        Write-Log "ユーザー $($user.UserPrincipalName) は管理者ロール: $($roleDetail.roleTemplateId)" "INFO"
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
            var cellValue = rows[j].getElementsBy
        }
    }
}
"@
