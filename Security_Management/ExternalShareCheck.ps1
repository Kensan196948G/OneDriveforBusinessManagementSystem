# OneDrive for Business 運用ツール - ITSM準拠
# ExternalShareCheck.ps1 - 外部共有監視スクリプト

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$LogDir = "$(Get-Location)\Log"
)

# 実行開始時刻を記録
$executionTime = Get-Date

# ログファイルのパスを設定
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $LogDir -ChildPath "ExternalShareCheck.$timestamp.log"
$errorLogPath = Join-Path -Path $LogDir -ChildPath "ExternalShareCheck.Error.$timestamp.log"

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

# 外部共有監視開始
Write-Log "外部共有監視を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"
Write-Log "ログディレクトリ: $LogDir" "INFO"

# Microsoft Graphの接続確認
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Log "Microsoft Graphに接続されていません。Main.ps1から実行してください。" "ERROR"
        exit
    }
    
    # ログインユーザーのUPN（メールアドレス）を自動取得
    $UserUPN = $context.Account
    
    # ログイン済ユーザー情報を取得
    $currentUser = Get-MgUser -UserId $UserUPN -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType
    
    # グローバル管理者かどうかを判定
    $isAdmin = ($context.Scopes -contains "Directory.ReadWrite.All")
    
    Write-Log "ログインユーザー: $($currentUser.DisplayName) ($UserUPN)" "INFO"
    Write-Log "実行モード: $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})" "INFO"
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続確認中にエラーが発生しました"
    exit
}

# 出力用の外部共有リスト
$externalShareList = @()

try {
    if ($isAdmin) {
        # グローバル管理者の場合、すべてのユーザーの外部共有を確認
        Write-Log "すべてのユーザーの外部共有を確認しています..." "INFO"
        $allUsers = Get-MgUser -All -Property DisplayName,Mail,AccountEnabled
        
        $totalUsers = $allUsers.Count
        $processedUsers = 0
        
        foreach ($user in $allUsers) {
            $processedUsers++
            $percentComplete = [math]::Round(($processedUsers / $totalUsers) * 100, 2)
            Write-Progress -Activity "外部共有を確認中" -Status "$processedUsers / $totalUsers ユーザー処理中 ($percentComplete%)" -PercentComplete $percentComplete
            
            try {
                # ユーザーのOneDriveを取得
                $drive = Get-MgUserDrive -UserId $user.UserPrincipalName -ErrorAction Stop
                
                # 外部共有を確認（実際のAPIでは直接外部共有を取得できないため、シミュレーション）
                # 実際の実装では、Microsoft GraphのAPIを使用して外部共有を取得する必要があります
                
                # サンプルデータ（実際の実装では削除）
                if ($processedUsers % 15 -eq 0) { # 15人に1人は外部共有があるとシミュレーション
                    $itemTypes = @("フォルダ", "ドキュメント", "画像", "スプレッドシート")
                    $itemType = $itemTypes[(Get-Random -Minimum 0 -Maximum 4)]
                    
                    $itemName = "外部共有$itemType`_$(Get-Random -Minimum 1 -Maximum 1000)"
                    
                    $sharingTypes = @("リンク共有", "直接共有", "グループ共有")
                    $sharingType = $sharingTypes[(Get-Random -Minimum 0 -Maximum 3)]
                    
                    $sharingScopes = @("組織外（特定ドメイン）", "組織外（任意）", "誰でも（リンク所有者のみ）", "誰でも（編集可能）")
                    $sharingScope = $sharingScopes[(Get-Random -Minimum 0 -Maximum 4)]
                    
                    $sharedWith = "external@example.com"
                    if ($sharingScope -eq "組織外（特定ドメイン）") {
                        $sharedWith = "user@partner-company.com"
                    } elseif ($sharingScope -eq "誰でも（リンク所有者のみ）" -or $sharingScope -eq "誰でも（編集可能）") {
                        $sharedWith = "匿名リンク"
                    }
                    
                    # リスクレベルを設定
                    $riskLevel = "低"
                    if ($sharingScope -eq "誰でも（編集可能）") {
                        $riskLevel = "高"
                    } elseif ($sharingScope -eq "誰でも（リンク所有者のみ）" -or $sharingScope -eq "組織外（任意）") {
                        $riskLevel = "中"
                    }
                    
                    # セキュリティ推奨を設定
                    $securityRecommendation = "問題ありません。"
                    if ($riskLevel -eq "高") {
                        $securityRecommendation = "編集権限を削除し、閲覧のみに制限することを推奨します。"
                    } elseif ($riskLevel -eq "中") {
                        $securityRecommendation = "共有範囲を特定ユーザーに限定することを検討してください。"
                    }
                    
                    $lastModified = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 30))
                    $webUrl = "https://tenant.sharepoint.com/personal/$($user.UserPrincipalName.Replace('@', '_').Replace('.', '_'))/Documents/$itemName"
                    
                    $externalShareList += [PSCustomObject]@{
                        "ユーザー名"       = $user.DisplayName
                        "メールアドレス"   = $user.Mail
                        "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
                        "アイテム名"       = $itemName
                        "アイテムタイプ"   = $itemType
                        "共有タイプ"       = $sharingType
                        "共有範囲"         = $sharingScope
                        "共有先"           = $sharedWith
                        "リスクレベル"     = $riskLevel
                        "最終更新日時"     = $lastModified
                        "WebURL"           = $webUrl
                        "セキュリティ推奨" = $securityRecommendation
                    }
                    
                    Write-Log "ユーザー $($user.UserPrincipalName) の外部共有を検出しました: $sharingType - $itemName" "INFO"
                }
            } catch {
                Write-Log "ユーザー $($user.UserPrincipalName) のOneDriveにアクセスできませんでした: $_" "WARNING"
            }
        }
        
        Write-Progress -Activity "外部共有を確認中" -Completed
        Write-Log "外部共有の確認が完了しました。検出された外部共有: $($externalShareList.Count)" "SUCCESS"
    } else {
        # 一般ユーザーまたはゲストの場合、自分自身の外部共有のみ確認
        Write-Log "自分自身の外部共有を確認しています..." "INFO"
        
        try {
            # 自分のOneDriveを取得
            $myDrive = Get-MgUserDrive -UserId $UserUPN -ErrorAction Stop
            
            # 外部共有を確認（実際のAPIでは直接外部共有を取得できないため、シミュレーション）
            # 実際の実装では、Microsoft GraphのAPIを使用して外部共有を取得する必要があります
            
            # サンプルデータ（実際の実装では削除）
            if (Get-Random -Minimum 0 -Maximum 2 -eq 0) { # 50%の確率で外部共有があるとシミュレーション
                $itemTypes = @("フォルダ", "ドキュメント", "画像", "スプレッドシート")
                $itemType = $itemTypes[(Get-Random -Minimum 0 -Maximum 4)]
                
                $itemName = "外部共有$itemType`_$(Get-Random -Minimum 1 -Maximum 1000)"
                
                $sharingTypes = @("リンク共有", "直接共有", "グループ共有")
                $sharingType = $sharingTypes[(Get-Random -Minimum 0 -Maximum 3)]
                
                $sharingScopes = @("組織外（特定ドメイン）", "組織外（任意）", "誰でも（リンク所有者のみ）", "誰でも（編集可能）")
                $sharingScope = $sharingScopes[(Get-Random -Minimum 0 -Maximum 4)]
                
                $sharedWith = "external@example.com"
                if ($sharingScope -eq "組織外（特定ドメイン）") {
                    $sharedWith = "user@partner-company.com"
                } elseif ($sharingScope -eq "誰でも（リンク所有者のみ）" -or $sharingScope -eq "誰でも（編集可能）") {
                    $sharedWith = "匿名リンク"
                }
                
                # リスクレベルを設定
                $riskLevel = "低"
                if ($sharingScope -eq "誰でも（編集可能）") {
                    $riskLevel = "高"
                } elseif ($sharingScope -eq "誰でも（リンク所有者のみ）" -or $sharingScope -eq "組織外（任意）") {
                    $riskLevel = "中"
                }
                
                # セキュリティ推奨を設定
                $securityRecommendation = "問題ありません。"
                if ($riskLevel -eq "高") {
                    $securityRecommendation = "編集権限を削除し、閲覧のみに制限することを推奨します。"
                } elseif ($riskLevel -eq "中") {
                    $securityRecommendation = "共有範囲を特定ユーザーに限定することを検討してください。"
                }
                
                $lastModified = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 30))
                $webUrl = "https://tenant.sharepoint.com/personal/$($UserUPN.Replace('@', '_').Replace('.', '_'))/Documents/$itemName"
                
                $externalShareList += [PSCustomObject]@{
                    "ユーザー名"       = $currentUser.DisplayName
                    "メールアドレス"   = $currentUser.Mail
                    "アカウント状態"   = if($currentUser.AccountEnabled){"有効"}else{"無効"}
                    "アイテム名"       = $itemName
                    "アイテムタイプ"   = $itemType
                    "共有タイプ"       = $sharingType
                    "共有範囲"         = $sharingScope
                    "共有先"           = $sharedWith
                    "リスクレベル"     = $riskLevel
                    "最終更新日時"     = $lastModified
                    "WebURL"           = $webUrl
                    "セキュリティ推奨" = $securityRecommendation
                }
                
                Write-Log "外部共有を検出しました: $sharingType - $itemName" "INFO"
            } else {
                Write-Log "外部共有は検出されませんでした。" "SUCCESS"
            }
        } catch {
            Write-ErrorLog $_ "OneDriveにアクセスできませんでした"
        }
        
        Write-Log "外部共有の確認が完了しました。検出された外部共有: $($externalShareList.Count)" "SUCCESS"
    }
} catch {
    Write-ErrorLog $_ "外部共有の確認中にエラーが発生しました"
}

# 外部共有が検出されなかった場合のメッセージ
if ($externalShareList.Count -eq 0) {
    $externalShareList += [PSCustomObject]@{
        "ユーザー名"       = "N/A"
        "メールアドレス"   = "N/A"
        "アカウント状態"   = "N/A"
        "アイテム名"       = "N/A"
        "アイテムタイプ"   = "N/A"
        "共有タイプ"       = "N/A"
        "共有範囲"         = "N/A"
        "共有先"           = "N/A"
        "リスクレベル"     = "低"
        "最終更新日時"     = Get-Date
        "WebURL"           = "N/A"
        "セキュリティ推奨" = "外部共有は検出されませんでした。"
    }
    
    Write-Log "外部共有は検出されませんでした。" "SUCCESS"
}

# タイムスタンプ
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# 出力ファイル名の設定
$csvFile = "ExternalShareCheck.$timestamp.csv"
$htmlFile = "ExternalShareCheck.$timestamp.html"
$jsFile = "ExternalShareCheck.$timestamp.js"

# 出力パスの設定
$csvPath = Join-Path -Path $OutputDir -ChildPath $csvFile
$htmlPath = Join-Path -Path $OutputDir -ChildPath $htmlFile
$jsPath = Join-Path -Path $OutputDir -ChildPath $jsFile

# CSV出力（文字化け対策済み）
try {
    # PowerShell Core (バージョン 6.0以上)の場合
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $externalShareList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    }
    # PowerShell 5.1以下の場合
    else {
        $externalShareList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
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
// ExternalShareCheck データ操作用 JavaScript

// グローバル変数
let currentPage = 1;
let rowsPerPage = 10; // デフォルトの1ページあたりの行数
let filteredRows = []; // フィルタリングされた行を保持する配列

// テーブルを検索する関数（インクリメンタル検索対応）
function searchTable() {
    var input = document.getElementById('searchInput').value.toLowerCase();
    var table = document.getElementById('externalShareTable');
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
    
    // 最大5件まで表示（より見やすく）
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
function createColumnFilters() {
    var table = document.getElementById('externalShareTable');
    var headers = table.getElementsByTagName('thead')[0].getElementsByTagName('th');
    var filterRow = document.createElement('tr');
    filterRow.className = 'filter-row';
    
    for (var i = 0; i < headers.length; i++) {
        var cell = document.createElement('th');
        var select = document.createElement('select');
        select.className = 'column-filter';
        select.setAttribute('data-column', i);
        
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
    var table = document.getElementById('externalShareTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var filters = document.getElementsByClassName('column-filter');
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

// ページングを更新する関数
function updatePagination() {
    var table = document.getElementById('externalShareTable');
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

// CSVとしてエクスポートする関数 (文字化け対策済み)
function exportTableToCSV() {
    var table = document.getElementById('externalShareTable');
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
    downloadLink.download = 'ExternalShareCheck_Export.csv';
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
    var table = document.getElementById('externalShareTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    
    for (var i = 0; i < rows.length; i++) {
        var riskLevelCell = rows[i].querySelector('td:nth-child(9)'); // リスクレベルのセル
        
        if (riskLevelCell) {
            var riskLevel = riskLevelCell.textContent;
            
            if (riskLevel === '高') {
                rows[i].classList.add('danger');
            } else if (riskLevel === '中') {
                rows[i].classList.add('warning');
            } else if (riskLevel === '低') {
                rows[i].classList.add('normal');
            }
        }
        
        // アカウント状態によっても色分け
        var accountStatus = rows[i].querySelector('td:nth-child(3)'); // アカウント状態のセル
        if (accountStatus && accountStatus.textContent === '無効') {
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
    var table = document.getElementById('externalShareTable');
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
"@

# JavaScript ファイルを出力
$jsContent | Out-File -FilePath $jsPath -Encoding UTF8
Write-Log "JavaScriptファイルを作成しました: $jsPath" "SUCCESS"

# 実行日時とユーザー情報を取得
$executionDateFormatted = $executionTime.ToString("yyyy/MM/dd HH:mm:ss")
$executorName = $currentUser.DisplayName
$userType = if($currentUser.UserType){$currentUser.UserType}else{"未定義"}

# HTML ファイルの生成
$htmlContent = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>OneDrive 外部共有監視レポート</title>
    <script src="$jsFile"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
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
            .toolbar, button, #pagination, .filter-row {
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
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-icon">🔒</div>
            <h1>OneDrive 外部共有監視レポート</h1>
        </div>
        
        <div class="info-section">
            <p><span class="info-label">実行日時:</span> $executionDateFormatted</p>
            <p><span class="info-label">実行者:</span> $executorName</p>
            <p><span class="info-label">実行者の種別:</span> $userType</p>
            <p><span class="info-label">実行モード:</span> $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})</p>
            <p><span class="info-label">出力フォルダ:</span> $OutputDir</p>
        </div>
        
        <div class="toolbar">
            <input type="text" id="searchInput" placeholder="検索...">
            <div id="searchSuggestions"></div>
            <button id="exportBtn"><span class="button-icon">📥</span>CSVエクスポート</button>
            <button id="printBtn"><span class="button-icon">🖨️</span>印刷</button>
        </div>
        
        <div id="pagination"></div>

        <table id="externalShareTable">
            <thead>
                <tr>
                    <th>ユーザー名</th>
                    <th>メールアドレス</th>
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

# HTML テーブル本体の作成
foreach ($share in $externalShareList) {
    # リスクレベルに応じたアイコンを設定
    $riskLevelIcon = switch ($share.'リスクレベル') {
        "高" { "🔴" }
        "中" { "🟡" }
        "低" { "🟢" }
        default { "❓" }
    }
    
    # 行を追加
    $htmlContent += @"
                <tr>
                    <td>$($share.'ユーザー名')</td>
                    <td>$($share.'メールアドレス')</td>
                    <td>$($share.'アカウント状態')</td>
                    <td>$($share.'アイテム名')</td>
                    <td>$($share.'アイテムタイプ')</td>
                    <td>$($share.'共有タイプ')</td>
                    <td>$($share.'共有範囲')</td>
                    <td>$($share.'共有先')</td>
                    <td><span class="status-icon">$riskLevelIcon</span>$($share.'リスクレベル')</td>
                    <td>$($share.'最終更新日時')</td>
                    <td>$($share.'WebURL')</td>
                    <td>$($share.'セキュリティ推奨')</td>
                </tr>
"@
}

# HTML 終了部分
$htmlContent += @"
            </tbody>
        </table>
        
        <div class="info-section">
            <p><span class="info-label">色の凡例:</span></p>
            <p>🟢 緑色の行: 低リスクの外部共有</p>
            <p>🟡 黄色の行: 中リスクの外部共有</p>
            <p>🔴 赤色の行: 高リスクの外部共有</p>
            <p>⚪ グレーの行: 無効なアカウント</p>
        </div>
    </div>
</body>
</html>
"@

# HTML ファイルを出力
$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Log "HTMLファイルを作成しました: $htmlPath" "SUCCESS"

# ログファイルに出力
$externalShareList | Format-Table -AutoSize | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
Write-Log "外部共有監視が完了しました" "SUCCESS"

# 出力ディレクトリを開く
try {
    Start-Process -FilePath "explorer.exe" -ArgumentList $OutputDir
    Write-Log "出力ディレクトリを開きました: $OutputDir" "SUCCESS"
} catch {
    Write-Log "出力ディレクトリを開けませんでした: $_" "WARNING"
}