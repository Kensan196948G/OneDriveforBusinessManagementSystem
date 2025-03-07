# OneDrive for Business 運用ツール - ITSM準拠
# Main.ps1 - メインスクリプト

param (
    [string]$OutputDir = "$(Get-Location)"
)

# 実行開始時刻を記録
$executionTime = Get-Date

# 日付ベースのフォルダ名とログフォルダを作成
$dateFolderName = "OneDriveManagement." + $executionTime.ToString("yyyyMMdd")
$dateFolderPath = Join-Path -Path $OutputDir -ChildPath $dateFolderName
$reportFolderPath = Join-Path -Path $dateFolderPath -ChildPath "Report"
$logFolderPath = Join-Path -Path $dateFolderPath -ChildPath "Log"

# 出力ディレクトリが存在しない場合は作成
if (-not (Test-Path -Path $dateFolderPath)) {
    New-Item -Path $dateFolderPath -ItemType Directory | Out-Null
    Write-Output "出力用フォルダを作成しました: $dateFolderPath"
}

# ログフォルダが存在しない場合は作成
if (-not (Test-Path -Path $logFolderPath)) {
    New-Item -Path $logFolderPath -ItemType Directory | Out-Null
    Write-Output "ログフォルダを作成しました: $logFolderPath"
}

# レポートフォルダが存在しない場合は作成
if (-not (Test-Path -Path $reportFolderPath)) {
    New-Item -Path $reportFolderPath -ItemType Directory | Out-Null
    Write-Output "レポートフォルダを作成しました: $reportFolderPath"
}

# ログファイルのパスを設定
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$executionLogPath = Join-Path -Path $logFolderPath -ChildPath "ExecutionLog.$timestamp.log"
$errorLogPath = Join-Path -Path $logFolderPath -ChildPath "ErrorLog.$timestamp.log"

# ログ関数
function Write-ExecutionLog {
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
    Add-Content -Path $executionLogPath -Value $logMessage -Encoding UTF8
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
    
    # 実行ログに出力
    Add-Content -Path $executionLogPath -Value $errorMessage -Encoding UTF8
    
    # エラーログに詳細を出力
    Add-Content -Path $errorLogPath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $errorLogPath -Value $errorDetails -Encoding UTF8
}

# Microsoft Graphモジュールのインストール確認と実施
function Install-RequiredModules {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-ExecutionLog "Microsoft.Graph モジュールが未インストールのためインストールします..." "INFO"
        try {
            Install-Module Microsoft.Graph -Scope CurrentUser -Force
            Write-ExecutionLog "Microsoft.Graph モジュールのインストールが完了しました" "SUCCESS"
        } catch {
            Write-ErrorLog $_ "Microsoft.Graph モジュールのインストールに失敗しました"
            return $false
        }
    } else {
        Write-ExecutionLog "Microsoft.Graph モジュールは既にインストールされています" "INFO"
    }
    
    return $true
}

# Microsoft Graphに接続
function Connect-ToMicrosoftGraph {
    try {
        Write-ExecutionLog "Microsoft Graphに接続しています..." "INFO"
        
        # グローバル管理者向けの事前設定情報
        <#
        ===== グローバル管理者向け事前設定手順 =====
        
        以下の手順は、グローバル管理者が初回のみ実行する必要があります。
        
        1. Azure Active Directoryポータル(https://aad.portal.azure.com/)にグローバル管理者でログイン
        
        2. [エンタープライズアプリケーション] → [新しいアプリケーション] → [独自のアプリケーションを作成]
           - 名前: "OneDrive運用ツール"
           - タイプ: "この組織ディレクトリのみに存在するアプリケーション"
        
        3. [APIのアクセス許可] → [アクセス許可の追加] → [Microsoft Graph] → [委任されたアクセス許可]
           - User.Read.All
           - Directory.Read.All
           - Sites.Read.All
           - Files.Read.All
        
        4. [管理者の同意を付与] ボタンをクリック
        #>
        
        Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All","Sites.Read.All","Files.Read.All"
        
        # ログインユーザーのUPN（メールアドレス）を自動取得
        $context = Get-MgContext
        $UserUPN = $context.Account
        
        # ログイン済ユーザー情報を取得
        $currentUser = Get-MgUser -UserId $UserUPN -Property DisplayName,Mail,UserType
        
        Write-ExecutionLog "Microsoft Graphに接続しました。ログインユーザー: $($currentUser.DisplayName) ($UserUPN)" "SUCCESS"
        return $true
    } catch {
        Write-ErrorLog $_ "Microsoft Graphへの接続に失敗しました"
        return $false
    }
}

# メインプログラム開始
Clear-Host
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  OneDrive for Business 運用ツール - ITSM準拠" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

Write-ExecutionLog "OneDrive for Business 運用ツール - ITSM準拠 を開始しました"
Write-ExecutionLog "出力フォルダ: $dateFolderPath"
Write-ExecutionLog "ログフォルダ: $logFolderPath"
Write-ExecutionLog "レポートフォルダ: $reportFolderPath"

# スクリプトのルートパスを取得
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# 必要なモジュールをインストール
$modulesInstalled = Install-RequiredModules
if (-not $modulesInstalled) {
    Write-ExecutionLog "必要なモジュールのインストールに失敗したため、プログラムを終了します。" "ERROR"
    exit
}

# Microsoft Graphに接続
$connected = Connect-ToMicrosoftGraph
if (-not $connected) {
    Write-ExecutionLog "Microsoft Graphへの接続に失敗したため、プログラムを終了します。" "ERROR"
    exit
}

# 基本データ収集メニュー
function Show-BasicDataMenu {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "  基本データ収集 (Basic Data Collection)" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] ユーザー情報取得 (GetUserInfo.ps1)"
    Write-Host "  [2] ストレージクォータ取得 (GetOneDriveQuota.ps1)"
    Write-Host "  [3] すべての基本データ収集 (GetAllBasicData.ps1)"
    Write-Host ""
    Write-Host "  [0] メインメニューに戻る"
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "選択してください"
    
    switch ($choice) {
        "1" { return "$scriptRoot\Basic_Data_Collection\GetUserInfo.ps1" }
        "2" { return "$scriptRoot\Basic_Data_Collection\GetOneDriveQuota.ps1" }
        "3" { return "$scriptRoot\Basic_Data_Collection\GetAllBasicData.ps1" }
        "0" { return $null }
        default { 
            Write-Host "無効な選択です。もう一度お試しください。" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return Show-BasicDataMenu
        }
    }
}

# インシデント管理メニュー
function Show-IncidentManagementMenu {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "  インシデント管理 (Incident Management)" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] 同期エラー確認 (SyncErrorCheck.ps1)"
    Write-Host ""
    Write-Host "  [0] メインメニューに戻る"
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "選択してください"
    
    switch ($choice) {
        "1" { return "$scriptRoot\Incident_Management\SyncErrorCheck.ps1" }
        "0" { return $null }
        default { 
            Write-Host "無効な選択です。もう一度お試しください。" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return Show-IncidentManagementMenu
        }
    }
}

# 変更管理メニュー
function Show-ChangeManagementMenu {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "  変更管理 (Change Management)" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] 共有設定確認 (SharingCheck.ps1)"
    Write-Host ""
    Write-Host "  [0] メインメニューに戻る"
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "選択してください"
    
    switch ($choice) {
        "1" { return "$scriptRoot\Change_Management\SharingCheck.ps1" }
        "0" { return $null }
        default { 
            Write-Host "無効な選択です。もう一度お試しください。" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return Show-ChangeManagementMenu
        }
    }
}

# セキュリティ管理メニュー
function Show-SecurityManagementMenu {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "  セキュリティ管理 (Security Management)" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] 外部共有監視 (ExternalShareCheck.ps1)"
    Write-Host ""
    Write-Host "  [0] メインメニューに戻る"
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "選択してください"
    
    switch ($choice) {
        "1" { return "$scriptRoot\Security_Management\ExternalShareCheck.ps1" }
        "0" { return $null }
        default { 
            Write-Host "無効な選択です。もう一度お試しください。" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return Show-SecurityManagementMenu
        }
    }
}

# スクリプト実行関数
function Invoke-SelectedScript {
    param (
        [string]$ScriptPath,
        [string]$BaseFolder,
        [string]$CategoryName,
        [string]$LogFolder
    )
    
    if (-not (Test-Path -Path $ScriptPath)) {
        Write-ExecutionLog "スクリプトが見つかりません: $ScriptPath" "ERROR"
        return
    }
    
    # カテゴリフォルダを作成
    $categoryFolderName = "$CategoryName." + $executionTime.ToString("yyyyMMdd")
    $categoryFolderPath = Join-Path -Path $BaseFolder -ChildPath $categoryFolderName
    
    if (-not (Test-Path -Path $categoryFolderPath)) {
        New-Item -Path $categoryFolderPath -ItemType Directory | Out-Null
        Write-ExecutionLog "カテゴリフォルダを作成しました: $categoryFolderPath" "INFO"
    }
    
    Write-ExecutionLog "スクリプト実行開始: $ScriptPath ($CategoryName)" "INFO"
    
    try {
        & $ScriptPath -OutputDir $categoryFolderPath -LogDir $LogFolder
        Write-ExecutionLog "スクリプト実行完了: $ScriptPath" "SUCCESS"
    } catch {
        Write-ErrorLog $_ "スクリプト実行中にエラーが発生しました: $ScriptPath"
    }
    
    Write-Host ""
    Write-Host "続行するには何かキーを押してください..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# メインメニュー
try {
    $exit = $false
    
    while (-not $exit) {
        Clear-Host
        Write-Host "=================================================" -ForegroundColor Cyan
        Write-Host "  OneDrive for Business 運用ツール - ITSM準拠" -ForegroundColor Cyan
        Write-Host "=================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  [0] 基本データ収集 (Basic Data Collection)"
        Write-Host "  [1] インシデント管理 (Incident Management)"
        Write-Host "  [2] 変更管理 (Change Management)"
        Write-Host "  [3] セキュリティ管理 (Security Management)"
        Write-Host "  [4] 総合レポート生成 (Generate Report)"
        Write-Host "  [5] Microsoft Graph再接続"
        Write-Host ""
        Write-Host "  [Q] 終了"
        Write-Host ""
        Write-Host "=================================================" -ForegroundColor Cyan
        
        $choice = Read-Host "選択してください"
        
        switch ($choice.ToUpper()) {
            "0" {
                $scriptPath = Show-BasicDataMenu
                if ($scriptPath) {
                    Invoke-SelectedScript -ScriptPath $scriptPath -BaseFolder $dateFolderPath -CategoryName "Basic_Data_Collection" -LogFolder $logFolderPath
                }
            }
            "1" {
                $scriptPath = Show-IncidentManagementMenu
                if ($scriptPath) {
                    Invoke-SelectedScript -ScriptPath $scriptPath -BaseFolder $dateFolderPath -CategoryName "Incident_Management" -LogFolder $logFolderPath
                }
            }
            "2" {
                $scriptPath = Show-ChangeManagementMenu
                if ($scriptPath) {
                    Invoke-SelectedScript -ScriptPath $scriptPath -BaseFolder $dateFolderPath -CategoryName "Change_Management" -LogFolder $logFolderPath
                }
            }
            "3" {
                $scriptPath = Show-SecurityManagementMenu
                if ($scriptPath) {
                    Invoke-SelectedScript -ScriptPath $scriptPath -BaseFolder $dateFolderPath -CategoryName "Security_Management" -LogFolder $logFolderPath
                }
            }
            "5" {
                Disconnect-MgGraph | Out-Null
                $connected = Connect-ToMicrosoftGraph
                if (-not $connected) {
                    Write-ExecutionLog "Microsoft Graphへの再接続に失敗したため、プログラムを終了します。" "ERROR"
                    exit
                }
            }
            "4" {
                # 総合レポート生成
                $reportScriptPath = "$scriptRoot\GenerateReport.ps1"
                if (Test-Path -Path $reportScriptPath) {
                    Invoke-SelectedScript -ScriptPath $reportScriptPath -BaseFolder $dateFolderPath -CategoryName "Report" -LogFolder $logFolderPath
                }
            }
            "Q" {
                $exit = $true
                Write-ExecutionLog "プログラムを終了します。" "INFO"
            }
            default {
                Write-Host "無効な選択です。もう一度お試しください。" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
    }
} catch {
    Write-ErrorLog $_ "メインプログラムの実行中にエラーが発生しました"
    
    Write-Host ""
    Write-Host "エラーが発生しました。詳細はログファイルを確認してください。" -ForegroundColor Red
    Write-Host "ログファイル: $executionLogPath" -ForegroundColor Yellow
    Write-Host "エラーログ: $errorLogPath" -ForegroundColor Yellow
    
    Write-ExecutionLog "エラーによりプログラムを終了します。" "ERROR"
}