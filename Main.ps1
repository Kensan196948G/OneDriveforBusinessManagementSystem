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
if ($null -eq $dateFolderPath -or -not (Test-Path -Path $dateFolderPath)) {
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
        [string]$Level = "INFO",
        [string]$ScriptName = $MyInvocation.ScriptName
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # スクリプト名からベース名を取得
    $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)
    
    # ログファイルパスを動的に生成
    $logTimestamp = Get-Date -Format "yyyyMMddHHmmss"
    $dynamicLogPath = Join-Path -Path $logFolderPath -ChildPath "$scriptBaseName.$logTimestamp.log"
    
    # コンソールに出力
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "DEBUG" { Write-Host $logMessage -ForegroundColor Cyan }
        default { Write-Host $logMessage }
    }
    
    # ログファイルに出力
    Add-Content -Path $dynamicLogPath -Value $logMessage -Encoding UTF8
    
    # グローバルな実行ログにも出力
    if ($executionLogPath -ne $dynamicLogPath) {
        Add-Content -Path $executionLogPath -Value $logMessage -Encoding UTF8
    }
}

function Write-ErrorLog {
    param (
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$Message,
        [string]$ScriptName = $MyInvocation.ScriptName
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
    
    # スクリプト名からベース名を取得
    $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)
    
    # エラーログファイルパスを動的に生成
    $logTimestamp = Get-Date -Format "yyyyMMddHHmmss"
    $dynamicErrorLogPath = Join-Path -Path $logFolderPath -ChildPath "$scriptBaseName.Error.$logTimestamp.log"
    
    # コンソールに出力
    Write-Host $errorMessage -ForegroundColor Red
    
    # 実行ログに出力
    Add-Content -Path $executionLogPath -Value $errorMessage -Encoding UTF8
    
    # エラーログに詳細を出力
    Add-Content -Path $dynamicErrorLogPath -Value $errorMessage -Encoding UTF8
    Add-Content -Path $dynamicErrorLogPath -Value $errorDetails -Encoding UTF8
}

# 共通ログ関数をエクスポート
Export-ModuleMember -Function Write-ExecutionLog, Write-ErrorLog

# Microsoft Graphモジュールのインストール確認と実施
function Install-RequiredModules {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-ExecutionLog "Microsoft.Graph モジュールが未インストールのためインストールします..." "INFO"
        Write-Host "[INFO] Microsoft.Graph モジュールのインストールを開始します..." -ForegroundColor Cyan
        try {
            Install-Module Microsoft.Graph -Scope CurrentUser -Force
            Write-ExecutionLog "Microsoft.Graph モジュールのインストールが完了しました" "SUCCESS"
            Write-Host "[SUCCESS] Microsoft.Graph モジュールのインストールが完了しました" -ForegroundColor Green
        } catch {
            Write-ErrorLog $_ "Microsoft.Graph モジュールのインストールに失敗しました"
            Write-Host "[ERROR] Microsoft.Graph モジュールのインストールに失敗しました" -ForegroundColor Red
            return $false
        }
    } else {
        Write-ExecutionLog "Microsoft.Graph モジュールは既にインストールされています" "INFO"
        Write-Host "[INFO] Microsoft.Graph モジュールは既にインストール済みです" -ForegroundColor Cyan
    }
    
    return $true
}

# Microsoft Graphに接続
function Connect-ToMicrosoftGraph {
    try {
        Write-ExecutionLog "config.jsonから認証情報を読み込みます..." "INFO"
        $configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
        if (-not (Test-Path $configPath)) {
            Write-ExecutionLog "config.jsonが見つかりません: $configPath" "ERROR"
            return $false
        }
        $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json

        $tenantId = $config.TenantId
        $clientId = $config.ClientId
        $clientSecret = $config.ClientSecret
        $scopes = "https://graph.microsoft.com/.default"
        $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

        Write-ExecutionLog "Microsoft Graph用アクセストークンを取得します..." "INFO"

        $body = @{
            client_id     = $clientId
            scope         = $scopes
            client_secret = $clientSecret
            grant_type    = "client_credentials"
        }

        try {
            $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
            $global:AccessToken = $response.access_token

            if (-not $global:AccessToken) {
                Write-ExecutionLog "アクセストークンの取得に失敗しました。レスポンス: $($response | ConvertTo-Json -Compress)" "ERROR"
                return $false
            }

            # TEMPフォルダにトークンを保存
            $tempDir = Join-Path -Path $PSScriptRoot -ChildPath "TEMP"
            if (-not (Test-Path -Path $tempDir)) {
                New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
            }
            $tokenFile = Join-Path -Path $tempDir -ChildPath "graph_token.txt"
            Set-Content -Path $tokenFile -Value $global:AccessToken -Encoding UTF8

            # PowerShell Graph SDKの認証状態をセット
            $secureToken = ConvertTo-SecureString $global:AccessToken -AsPlainText -Force
            Connect-MgGraph -AccessToken $secureToken | Out-Null

            Write-ExecutionLog "Microsoft Graphにクライアントシークレット認証で接続しました" "SUCCESS"
            return $true
        } catch {
            Write-ExecutionLog "トークン取得時のエラー詳細: $($_.Exception.Message)" "ERROR"
            if ($_.ErrorDetails) {
                Write-ExecutionLog "エラーディテール: $($_.ErrorDetails.Message)" "ERROR"
            }
            if ($_.Exception.Response -ne $null) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $responseBody = $reader.ReadToEnd()
                Write-ExecutionLog "HTTPレスポンス: $responseBody" "ERROR"
            }
            return $false
        }
    } catch {
        Write-ErrorLog $_ "Microsoft Graphへのクライアントシークレット認証に失敗しました"
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
        [string]$LogFolder,
        [array]$ArgumentList
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
        if ($ArgumentList -and $ArgumentList.Count -gt 0) {
            & $ScriptPath -OutputDir $categoryFolderPath -LogDir $LogFolder @ArgumentList
        } else {
            & $ScriptPath -OutputDir $categoryFolderPath -LogDir $LogFolder
        }
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
                    # OneDriveクォータ取得前に毎回トークン再取得
                    $connected = Connect-ToMicrosoftGraph
                    if (-not $connected) {
                        Write-ExecutionLog "Microsoft Graphへの再接続に失敗したため、スクリプトをスキップします。" "ERROR"
                    } else {
                        Invoke-SelectedScript -ScriptPath $scriptPath -BaseFolder $dateFolderPath -CategoryName "Basic_Data_Collection" -LogFolder $logFolderPath
                    }
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
                    Write-ExecutionLog "変更管理スクリプトを実行します: $scriptPath" "DEBUG"
                    
                    # サービスプリンシパル情報を取得
                    try {
                        $context = Get-MgContext
                        if (-not $context) {
                            Write-ExecutionLog "Microsoft Graphコンテキストを取得できませんでした" "ERROR"
                            Write-Host "エラー: Graph接続情報を取得できません" -ForegroundColor Red
                            return
                        }
                        
                        # サービスプリンシパル名を取得
                        $servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$($context.ClientId)'" -Top 1
                        if (-not $servicePrincipal) {
                            Write-ExecutionLog "サービスプリンシパル情報を取得できませんでした" "ERROR"
                            Write-Host "エラー: サービスプリンシパル情報が取得できません" -ForegroundColor Red
                            return
                        }
                        
                        $currentUser = $servicePrincipal.DisplayName
                        Write-ExecutionLog "サービスプリンシパル: $currentUser" "DEBUG"
                        
                        # 出力ディレクトリを作成
                        $changeMgmtDir = "$dateFolderPath\Change_Management.$($executionTime.ToString('yyyyMMdd'))"
                        if (-not (Test-Path -Path $changeMgmtDir)) {
                            New-Item -Path $changeMgmtDir -ItemType Directory | Out-Null
                        }

                        # スクリプトを実行 (明示的にPowerShellで実行)
                        try {
                            $arguments = "-OutputDir `"$changeMgmtDir`" -LogDir `"$logFolderPath`" -TargetUser `"$currentUser`" -AccessToken `"$global:AccessToken`""
                            Start-Process "powershell.exe" -ArgumentList "-File `"$scriptPath`" $arguments" -Wait -NoNewWindow
                            Write-ExecutionLog "スクリプト実行が完了しました" "SUCCESS"
                        } catch {
                            Write-ErrorLog $_ "SharingCheckスクリプト実行中にエラーが発生しました"
                            Write-Host "エラー: SharingCheckスクリプトの実行に失敗しました" -ForegroundColor Red
                        }
                    } catch {
                        Write-ErrorLog $_ "変更管理スクリプトの実行中にエラーが発生しました"
                        Write-Host "エラー: スクリプト実行中に問題が発生しました" -ForegroundColor Red
                    }
                }
            }
            "3" {
                $scriptPath = Show-SecurityManagementMenu
                if ($scriptPath) {
                    Invoke-SelectedScript -ScriptPath $scriptPath -BaseFolder $dateFolderPath -CategoryName "Security_Management" -LogFolder $logFolderPath
                }
            }
            "4" {
                # 総合レポート生成
                $reportScriptPath = "$scriptRoot\GenerateReport.ps1"
                if (Test-Path -Path $reportScriptPath) {
                    Invoke-SelectedScript -ScriptPath $reportScriptPath -BaseFolder $dateFolderPath -CategoryName "Report" -LogFolder $logFolderPath
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
