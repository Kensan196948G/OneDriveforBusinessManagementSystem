<# 
.SYNOPSIS
安全なスクリプト終了とログ記録を行うモジュール

.DESCRIPTION
スクリプトの安全な終了処理とログ記録機能を提供するPowerShellモジュール
#>

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO","WARNING","ERROR","DEBUG","SUCCESS")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory=$false)]
        [string]$LogPath = ".\Logs\Execution.log"
    )

    $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $logEntry = "[$timestamp][$Level] $Message"
    
    try {
        if (-not (Test-Path -Path (Split-Path -Path $LogPath -Parent))) {
            New-Item -Path (Split-Path -Path $LogPath -Parent) -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8
    }
    catch {
        Write-Host "ログ書き込みエラー: $_" -ForegroundColor Red
    }
    
    # コンソールにも出力
    switch ($Level) {
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
}

function Write-ErrorLog {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        
        [Parameter(Mandatory=$false)]
        [string]$CustomMessage,
        
        [Parameter(Mandatory=$false)]
        [string]$LogPath = ".\Logs\Error.log"
    )

    $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $errorMessage = if ($CustomMessage) { $CustomMessage } else { $ErrorRecord.Exception.Message }
    
    $logEntry = @"
[$timestamp][ERROR] $errorMessage
Exception Type: $($ErrorRecord.Exception.GetType().FullName)
StackTrace: $($ErrorRecord.ScriptStackTrace)
"@

    try {
        if (-not (Test-Path -Path (Split-Path -Path $LogPath -Parent))) {
            New-Item -Path (Split-Path -Path $LogPath -Parent) -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8
    }
    catch {
        Write-Host "エラーログ書き込みエラー: $_" -ForegroundColor Red
    }
    
    Write-Host $logEntry -ForegroundColor Red
}

function SafeExit {
    param(
        [Parameter(Mandatory=$false)]
        [int]$ExitCode = 0,
        
        [Parameter(Mandatory=$false)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$LogPath = ".\Logs\Execution.log"
    )

    if ($Message) {
        Write-Log -Message $Message -Level "INFO" -LogPath $LogPath
    }

    Write-Log -Message "スクリプトを終了します (終了コード: $ExitCode)" -Level "INFO" -LogPath $LogPath
    
    exit $ExitCode
}

Export-ModuleMember -Function Write-Log, Write-ErrorLog, SafeExit