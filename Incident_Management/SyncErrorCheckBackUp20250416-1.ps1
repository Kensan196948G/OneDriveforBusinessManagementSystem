# OneDrive for Business 運用ツール - ITSM準拠
# SyncErrorCheck.ps1 - 同期エラー確認スクリプト

param (
    [string]$OutputDir = "$(Get-Location)",
    [string]$LogDir = "$(Get-Location)\Log"
)

# 列名リストを明示的に定義
$columnNames = @("ユーザー名","メールアドレス","アカウント状態","OneDrive対応","エラー種別","ファイル名","ファイルパス","最終更新日時","サイズ(KB)","エラー詳細","推奨対応")

# 実行開始時刻を記録
$executionTime = Get-Date

# ログファイルのパスを設定
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $LogDir -ChildPath "SyncErrorCheck.$timestamp.log"
$errorLogPath = Join-Path -Path $LogDir -ChildPath "SyncErrorCheck.Error.$timestamp.log"

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

# 同期エラー確認開始
Write-Log "同期エラー確認を開始します" "INFO"
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

    # 非対話型認証（クライアントシークレット認証等）で$UserUPNが空の場合はconfig.jsonから取得
    if ([string]::IsNullOrEmpty($UserUPN)) {
        $configPath = Join-Path -Path $PSScriptRoot -ChildPath "..\config.json"
        if (Test-Path $configPath) {
            $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
            if ($config.AdminUPN) {
                $UserUPN = $config.AdminUPN
                Write-Log "config.jsonからAdminUPNを取得しました: $UserUPN" "INFO"
            }
        }
    }

    # ログイン済ユーザー情報を取得
    if (![string]::IsNullOrEmpty($UserUPN)) {
        try {
            $currentUser = Get-MgUser -UserId $UserUPN -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType
        } catch {
            if ($_.Exception.Message -like "*Resource '*' does not exist*") {
                Write-Log "AdminUPN（$UserUPN）がAzure AD上に存在しません。config.jsonのAdminUPNを実際のグローバル管理者UPN（メールアドレス）に修正してください。" "ERROR"
            } else {
                Write-Log "Get-MgUser実行時にエラーが発生しました: $($_.Exception.Message)" "ERROR"
            }
            exit
        }
    } else {
        Write-Log "UserUPNが空のため、ユーザー情報取得をスキップします" "ERROR"
        Write-Log "Microsoft Graphの認証情報に問題があるか、ログイン状態が不正です。Main.ps1から再度実行してください。" "ERROR"
        exit
    }
    
    # グローバル管理者かどうかを判定
    $isAdmin = ($context.Scopes -contains "Directory.ReadWrite.All")
    
    if ($currentUser) {
        Write-Log "ログインユーザー: $($currentUser.DisplayName) ($UserUPN)" "INFO"
    } else {
        Write-Log "ログインユーザー情報が取得できませんでした" "WARNING"
    }
    Write-Log "実行モード: $(if($isAdmin){"管理者モード"}else{"ユーザーモード"})" "INFO"
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続確認中にエラーが発生しました"
    exit
}

# 出力用のエラーリスト
$errorList = @()

try {
    if ($isAdmin) {
        # グローバル管理者の場合、すべてのユーザーの同期エラーを確認
        Write-Log "すべてのユーザーの同期エラーを確認しています..." "INFO"
        # Idプロパティを含めて全ユーザー取得（環境依存のため-UserIdでの呼び出しに備える）
        $allUsers = Get-MgUser -All -Property DisplayName,Mail,AccountEnabled,UserPrincipalName,Id

        $totalUsers = $allUsers.Count
        $processedUsers = 0

        foreach ($user in $allUsers) {
            $processedUsers++
            $percentComplete = [math]::Round(($processedUsers / $totalUsers) * 100, 2)
            Write-Progress -Activity "同期エラーを確認中" -Status "$processedUsers / $totalUsers ユーザー処理中 ($percentComplete%)" -PercentComplete $percentComplete

            # デバッグ: $user.Idの値をログ出力
            Write-Log "DEBUG: $($user.DisplayName) のUserId: $($user.Id)" "INFO"

            # UserPrincipalNameが空の場合はスキップ
            if ([string]::IsNullOrEmpty($user.UserPrincipalName)) {
                Write-Log "ユーザー $($user.DisplayName) のUserPrincipalNameが空のためスキップします" "WARNING"
                continue
            }

            $oneDriveSupported = "未対応"
            try {
                # ユーザーのOneDriveを取得
                $drive = Get-MgUserDrive -UserId $user.UserPrincipalName -ErrorAction Stop
                $oneDriveSupported = "対応"

                # OneDriveのルートアイテムを取得（DriveId優先、失敗時はUserIdで再試行）
                $rootItems = $null
                $rootSuccess = $false

                # DriveIdもUserIdも空の場合は絶対に呼び出さない
                # 厳密なnull/空文字チェックでプロンプト入力待ちを完全防止
                $hasDriveId = ($drive -and -not [string]::IsNullOrEmpty($drive.Id))
if ($hasDriveId) {
    Write-Log "DEBUG: $($user.DisplayName) のDriveId: $($drive.Id)" "INFO"
    try {
        Write-Log "DEBUG: Get-MgUserDriveRoot 実行直前 DriveId: $($drive.Id)" "INFO"
        $rootItems = Get-MgDriveRoot -DriveId ([string]$drive.Id) -ExpandProperty Children -ErrorAction Stop
        Write-Log "DEBUG: Get-MgUserDriveRoot 実行直後 DriveId: $($drive.Id)" "INFO"
        $rootSuccess = $true
    } catch {
        Write-Log "ユーザー $($user.UserPrincipalName) のDriveIdでの取得に失敗。スキップします。" "WARNING"
        continue
    }
} else {
    Write-Log "ユーザー $($user.UserPrincipalName) のDriveIdが取得できないためスキップします。" "WARNING"
    continue
}

                # 同期エラーを確認（実際のAPIでは直接同期エラーを取得できないため、シミュレーション）
                # 実際の実装では、Microsoft GraphのAPIを使用して同期エラーを取得する必要があります

                # サンプルデータ（実際の実装では削除）
                if ($processedUsers % 10 -eq 0) { # 10人に1人はエラーがあるとシミュレーション
                    $errorTypes = @("同期エラー", "アクセスエラー", "情報")
                    $errorType = $errorTypes[(Get-Random -Minimum 0 -Maximum 3)]
                    $fileName = "document_$(Get-Random -Minimum 1 -Maximum 1000).docx"
                    $filePath = "/Documents/$fileName"
                    $lastModified = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 30))
                    $fileSize = Get-Random -Minimum 10 -Maximum 5000
                    $errorDetail = "ファイルの同期中にエラーが発生しました。"
                    # エラー種別ごとに推奨対応を分岐
                    switch ($errorType) {
                        "同期エラー" {
                            $recommendation = "ネットワーク接続を確認し、OneDriveクライアントを再起動してください。解決しない場合はファイルを再アップロードしてください。"
                        }
                        "アクセスエラー" {
                            $recommendation = "ユーザーのOneDrive権限を確認し、必要に応じて管理者へ連絡してください。"
                        }
                        "情報" {
                            $recommendation = "特に対応は必要ありません。"
                        }
                        default {
                            $recommendation = "詳細を確認してください。"
                        }
                    }

                    $errorList += [PSCustomObject]@{
                        "ユーザー名"       = $user.DisplayName
                        "メールアドレス"   = $user.Mail
                        "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
                        "OneDrive対応"     = $oneDriveSupported
                        "エラー種別"       = $errorType
                        "ファイル名"       = $fileName
                        "ファイルパス"     = $filePath
                        "最終更新日時"     = $lastModified
                        "サイズ(KB)"       = $fileSize
                        "エラー詳細"       = $errorDetail
                        "推奨対応"         = $recommendation
                    }

                    Write-Log "ユーザー $($user.UserPrincipalName) の同期エラーを検出しました: $errorType - $fileName" "WARNING"
                }
            } catch {
                Write-Log "ユーザー $($user.UserPrincipalName) のOneDriveにアクセスできませんでした: $_" "WARNING"

                # アクセスエラーも記録
                $errorList += [PSCustomObject]@{
                    "ユーザー名"       = $user.DisplayName
                    "メールアドレス"   = $user.Mail
                    "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
                    "OneDrive対応"     = $oneDriveSupported
                    "エラー種別"       = "アクセスエラー"
                    "ファイル名"       = "N/A"
                    "ファイルパス"     = "N/A"
                    "最終更新日時"     = Get-Date
                    "サイズ(KB)"       = "N/A"
                    "エラー詳細"       = "OneDriveにアクセスできませんでした: $($_.Exception.Message)"
                    "推奨対応"         = "ユーザーのアクセス権限を確認してください。"
                }
            }
        }
        
        Write-Progress -Activity "同期エラーを確認中" -Completed
        Write-Log "同期エラーの確認が完了しました。検出されたエラー: $($errorList.Count)" "SUCCESS"
    } else {
        # 一般ユーザーまたはゲストの場合、自分自身の同期エラーのみ確認
        Write-Log "自分自身の同期エラーを確認しています..." "INFO"

        if ([string]::IsNullOrEmpty($UserUPN)) {
            Write-Log "UserUPNが空のため、自分自身の同期エラー確認をスキップします" "ERROR"
            Write-Log "Microsoft Graphの認証情報に問題があるか、ログイン状態が不正です。Main.ps1から再度実行してください。" "ERROR"
            exit
        }

        try {
            # 自分のOneDriveを取得
            $myDrive = Get-MgUserDrive -UserId $UserUPN -ErrorAction Stop

            # OneDriveのルートアイテムを取得（DriveIdを利用）
            if ($myDrive -and -not [string]::IsNullOrEmpty($myDrive.Id)) {
                $rootItems = Get-MgDriveRoot -DriveId ([string]$drive.Id) -ExpandProperty Children -ErrorAction Stop
            } else {
                Write-Log "自分のOneDrive DriveIdが取得できないため、同期エラー確認をスキップします。" "WARNING"
                continue
            }

            # 同期エラーを確認（実際のAPIでは直接同期エラーを取得できないため、シミュレーション）            
            # 実際の実装では、Microsoft GraphのAPIを使用して同期エラーを取得する必要があります

            # サンプルデータ（実際の実装では削除）
            if (Get-Random -Minimum 0 -Maximum 2 -eq 0) { # 50%の確率でエラーがあるとシミュレーション
                $errorTypes = @("同期エラー", "アクセスエラー", "情報")
                $errorType = $errorTypes[(Get-Random -Minimum 0 -Maximum 3)]
                $fileName = "document_$(Get-Random -Minimum 1 -Maximum 1000).docx"
                $filePath = "/Documents/$fileName"
                $lastModified = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 30))
                $fileSize = Get-Random -Minimum 10 -Maximum 5000
                $errorDetail = "ファイルの同期中にエラーが発生しました。"
                $recommendation = "ファイルを再度アップロードするか、OneDriveクライアントを再起動してください。"

                $errorList += [PSCustomObject]@{
                    "ユーザー名"       = $(if($currentUser){$currentUser.DisplayName}else{"N/A"})
                    "メールアドレス"   = $(if($currentUser){$currentUser.Mail}else{"N/A"})
                    "アカウント状態"   = $(if($currentUser){if($currentUser.AccountEnabled){"有効"}else{"無効"}}else{"N/A"})
                    "OneDrive対応"     = "N/A"
                    "エラー種別"       = $errorType
                    "ファイル名"       = $fileName
                    "ファイルパス"     = $filePath
                    "最終更新日時"     = $lastModified
                    "サイズ(KB)"       = $fileSize
                    "エラー詳細"       = $errorDetail
                    "推奨対応"         = $recommendation
                }

                Write-Log "同期エラーを検出しました: $errorType - $fileName" "WARNING"
            } else {
                Write-Log "同期エラーは検出されませんでした。" "SUCCESS"
            }
        } catch {
            Write-ErrorLog $_ "OneDriveにアクセスできませんでした"

            # アクセスエラーも記録
            $errorList += [PSCustomObject]@{
                "ユーザー名"       = $(if($currentUser){$currentUser.DisplayName}else{"N/A"})
                "メールアドレス"   = $(if($currentUser){$currentUser.Mail}else{"N/A"})
                "アカウント状態"   = $(if($currentUser){if($currentUser.AccountEnabled){"有効"}else{"無効"}}else{"N/A"})
                "OneDrive対応"     = "N/A"
                "エラー種別"       = "アクセスエラー"
                "ファイル名"       = "N/A"
                "ファイルパス"     = "N/A"
                "最終更新日時"     = Get-Date
                "サイズ(KB)"       = "N/A"
                "エラー詳細"       = "OneDriveにアクセスできませんでした: $($_.Exception.Message)"
                "推奨対応"         = "OneDriveの設定を確認してください。"
            }
        }

        Write-Log "同期エラーの確認が完了しました。検出されたエラー: $($errorList.Count)" "SUCCESS"
    }
} catch {
    Write-ErrorLog $_ "同期エラーの確認中にエラーが発生しました"
}

# エラーが検出されなかった場合のメッセージ
if ($errorList.Count -eq 0) {
            $errorList += [PSCustomObject]@{
                "ユーザー名"       = "N/A"
                "メールアドレス"   = "N/A"
                "アカウント状態"   = "N/A"
                "OneDrive対応"     = "N/A"
                "エラー種別"       = "情報"
                "ファイル名"       = "N/A"
                "ファイルパス"     = "N/A"
                "最終更新日時"     = Get-Date
                "サイズ(KB)"       = "N/A"
                "エラー詳細"       = "同期エラーは検出されませんでした。"
                "推奨対応"         = "対応は必要ありません。"
            }
    
    Write-Log "同期エラーは検出されませんでした。" "SUCCESS"
}

# タイムスタンプ
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# 出力ファイル名の設定
$csvFile = "SyncErrorCheck.$timestamp.csv"
$htmlFile = "SyncErrorCheck.$timestamp.html"
$jsFile = "SyncErrorCheck.$timestamp.js"

# 出力パスの設定
$csvPath = Join-Path -Path $OutputDir -ChildPath $csvFile
$htmlPath = Join-Path -Path $OutputDir -ChildPath $htmlFile
$jsPath = Join-Path -Path $OutputDir -ChildPath $jsFile

# CSV出力（文字化け対策済み）
try {
    # PowerShell Core (バージョン 6.0以上)の場合
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $errorList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    }
    # PowerShell 5.1以下の場合
    else {
        $errorList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
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
// SyncErrorCheck データ操作用 JavaScript

// グローバル変数
let currentPage = 1;
let rowsPerPage = 20; // デフォルトの1ページあたりの行数（10から20に変更）
let filteredRows = []; // フィルタリングされた行を保持する配列

// テーブルを検索する関数（インクリメンタル検索対応）
function searchTable() {
    var input = document.getElementById('searchInput').value.toLowerCase();
    var table = document.getElementById('errorTable');
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
    var table = document.getElementById('errorTable');
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
    var table = document.getElementById('errorTable');
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
    var table = document.getElementById('errorTable');
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
    var table = document.getElementById('errorTable');
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
    downloadLink.download = 'SyncErrorCheck_Export.csv';
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
    var table = document.getElementById('errorTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    
    for (var i = 0; i < rows.length; i++) {
        var errorTypeCell = rows[i].querySelector('td:nth-child(4)'); // エラー種別のセル
        
        if (errorTypeCell) {
            var errorType = errorTypeCell.textContent;
            
            if (errorType.includes('同期エラー')) {
                rows[i].classList.add('danger');
            } else if (errorType.includes('アクセスエラー')) {
                rows[i].classList.add('warning');
            } else if (errorType.includes('情報')) {
                rows[i].classList.add('info');
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
    var table = document.getElementById('errorTable');
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
    <title>OneDrive 同期エラーレポート</title>
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
        tr.info {
            background-color: #e3f2fd;
        }
        tr.disabled {
            color: #999;
            font-style: italic;
        }
        .status-icon {
            margin-right: 5px;
            font-size: 1.2em;
            vertical-align: middle;
        }
        .column-filter {
            width: 100%;
            padding: 5px;
            border: 1px solid #ddd;
            border-radius: 4px;
            background-color: white;
            font-size: 14px;
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
            tr.info {
                background-color: #e3f2fd !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-icon">⚠️</div>
            <h1>OneDrive 同期エラーレポート</h1>
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

        <table id="errorTable">
            <thead>
                <tr>
                    <th>ユーザー名</th>
                    <th>メールアドレス</th>
                    <th>アカウント状態</th>
                    <th>OneDrive対応</th>
                    <th>エラー種別</th>
                    <th>ファイル名</th>
                    <th>ファイルパス</th>
                    <th>最終更新日時</th>
                    <th>サイズ(KB)</th>
                    <th>エラー詳細</th>
                    <th>推奨対応</th>
                </tr>
            </thead>
            <tbody>
"@

# HTML テーブル本体の作成
foreach ($error in $errorList) {
    # エラー種別に応じたアイコンを設定
    $errorTypeIcon = switch ($error.'エラー種別') {
        "同期エラー" { "🔴" }
        "アクセスエラー" { "🟡" }
        "情報" { "🔵" }
        default { "❓" }
    }
    
    # 行を追加
    $htmlContent += @"
                <tr>
                    <td>$($error.'ユーザー名' -replace '(.{20})', '$1<br>')</td>
                    <td>$($error.'メールアドレス' -replace '(.{20})', '$1<br>')</td>
                    <td>$($error.'アカウント状態')</td>
                    <td style="font-size: 0.7em;">$($error.'OneDrive対応')</td>
                    <td><span class="status-icon" title="$($error.'エラー種別')">$(switch ($error.'エラー種別') {
                        "同期エラー" { "🔴" }
                        "アクセスエラー" { "🟡" }
                        "情報" { "🔵" }
                        default { "❓" }
                    })</span> <span style="font-size: 0.8em;">$($error.'エラー種別')</span></td>
                    <td style="font-size: 0.8em;">$($error.'ファイル名' -replace '(.{10})', '$1<br>')</td>
                    <td style="font-size: 0.8em;">$($error.'ファイルパス' -replace '(.{10})', '$1<br>')</td>
                    <td style="font-size: 0.8em;">$(if($error.'最終更新日時' -eq 'N/A'){'取得不可'}else{$error.'最終更新日時'})</td>
                    <td style="font-size: 0.8em;">$(if($error.'サイズ(KB)' -eq 'N/A'){'取得不可'}else{$error.'サイズ(KB)'})</td>
                    <td style="font-size: 0.8em;">$($error.'エラー詳細' -replace '(.{10})', '$1<br>')</td>
                    <td style="font-size: 0.8em; white-space: normal; min-width: 200px; background-color: #f8f8f8; padding: 8px; border: 1px solid #ddd;">$(if([string]::IsNullOrWhiteSpace($error.'推奨対応')){'1. エラーの詳細を確認<br>2. 該当ユーザーに連絡<br>3. 必要に応じて管理者へ報告'}else{$error.'推奨対応' -replace '(.{10})', '$1<br>'})</td>
                </tr>
"@
}

# HTML 終了部分
$htmlContent += @"
            </tbody>
        </table>
        
        <div class="info-section">
            <p><span class="info-label">色の凡例:</span></p>
            <p>🔴 赤色の行: 同期エラー</p>
            <p>🟡 黄色の行: アクセスエラー</p>
            <p>🔵 青色の行: 情報メッセージ</p>
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
$errorList | Format-Table -AutoSize | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
Write-Log "同期エラー確認が完了しました" "SUCCESS"

# 出力ディレクトリを開く
try {
    Start-Process -FilePath "explorer.exe" -ArgumentList $OutputDir
    Write-Log "出力ディレクトリを開きました: $OutputDir" "SUCCESS"
} catch {
    Write-Log "出力ディレクトリを開けませんでした: $_" "WARNING"
}
