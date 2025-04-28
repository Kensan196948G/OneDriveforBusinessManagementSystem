# OneDrive for Business 運用ツール - ITSM準拠
# SharingCheck.ps1 - 共有設定確認スクリプト

param (
    [Parameter(Mandatory=$true)]
    [string]$OutputDir,
    [Parameter(Mandatory=$true)] 
    [string]$LogDir,
    [string]$TargetUser = "",
    [Parameter(Mandatory=$true)]
    [string]$AccessToken
)

# 出力ディレクトリの存在確認
if (-not (Test-Path -Path $OutputDir)) {
    throw "出力ディレクトリが存在しません: $OutputDir"
}

# 実行開始時刻を記録
$executionTime = Get-Date

# ログファイルのパスを設定 (Main.ps1で作成済みのLogフォルダを使用)
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFilePath = Join-Path -Path $LogDir -ChildPath "SharingCheck.$timestamp.log"
$errorLogPath = Join-Path -Path $LogDir -ChildPath "SharingCheck.Error.$timestamp.log"

# ログ関数
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

Write-Log "共有設定確認を開始します" "INFO"
Write-Log "出力ディレクトリ: $OutputDir" "INFO"
Write-Log "ログディレクトリ: $LogDir" "INFO"

# 必要なモジュールをインポート
try {
    Import-Module Microsoft.PowerShell.Security -ErrorAction Stop
    Write-Log "Microsoft.PowerShell.Securityモジュールをインポートしました" "DEBUG"
} catch {
    Write-ErrorLog $_ "Microsoft.PowerShell.Securityモジュールのインポートに失敗しました"
    exit 1
}

# Microsoft Graphの接続確認
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Log "Microsoft Graphに接続されていません。Main.ps1から実行してください。" "ERROR"
        exit 1
    }
} catch {
    Write-ErrorLog $_ "Microsoft Graphの接続確認中にエラーが発生しました"
    Write-Log "Microsoft Graphへの接続に失敗しました。Main.ps1から再実行してください。" "ERROR"
    exit 1
}

    # サービスプリンシパル情報を取得
    try {
        $servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$($context.ClientId)'" -Top 1
        if (-not $servicePrincipal) {
            throw "サービスプリンシパル情報を取得できませんでした"
        }
    } catch {
        Write-ErrorLog $_ "サービスプリンシパル情報の取得中にエラーが発生しました"
        exit 1
    }

    # サービスプリンシパルモード判定
    $isAdmin = $false
    $currentUser = $servicePrincipal
    
    # サービスプリンシパル判定条件
    $isServicePrincipal = ($context.AppName -ne $null) -or 
                         ($context.Account -match "appId=") -or
                         ($context.Account -eq $servicePrincipal.DisplayName)
    
    if (-not $isServicePrincipal) {
        # ユーザーモードの場合のみユーザー情報取得を試みる
        try {
            if (-not [string]::IsNullOrEmpty($context.Account)) {
                $me = Get-MgUser -UserId $context.Account -ErrorAction Stop
                $isAdmin = $true
                $currentUser = $me
            }
        } catch {
            Write-Log "ユーザー情報取得に失敗しました。サービスプリンシパルモードで続行します。" "WARNING"
        }
    }
    
    Write-Log "実行モード: $(if($isAdmin){"管理者モード"}else{"サービスプリンシパルモード"})" "INFO"
    
    Write-Log "サービスプリンシパル: $($servicePrincipal.DisplayName)" "INFO"
    Write-Log "実行モード: $(if($isAdmin){"管理者モード"}else{"標準モード"})" "INFO"

    # 管理者モードでない場合、対象ユーザー/サービスプリンシパルをパラメータで受け取る
    if (-not $isAdmin) {
        if (-not $PSBoundParameters.ContainsKey('TargetUser')) {
            throw "管理者モードでない場合、-TargetUserパラメータで対象ユーザーまたはサービスプリンシパルを指定してください"
        }
        $UserUPN = $TargetUser
        Write-Log "対象ユーザー/サービスプリンシパル: $UserUPN" "INFO"
    }

# 出力用の共有設定リスト
$sharingList = @()
# 全ユーザーの処理ステータスを保持するリスト
$allUserStatuses = @()

try {
    # サービスプリンシパルまたは管理者アカウントの場合、全ユーザーの共有設定を確認する
    if ($true) { # サービスプリンシパルでの実行を前提とする
        # 全ユーザーの共有設定を確認
        Write-Log "すべてのユーザーの共有設定を確認しています (サービスプリンシパルモード)..." "INFO"
        if ($context) {
            $allUsers = Get-MgUser -All -Property Id,DisplayName,Mail,UserPrincipalName,AccountEnabled -ErrorAction Stop
        } else {
            Write-Log "Microsoft Graphコンテキストが利用できません。処理をスキップします。" "WARNING"
            return
        }
        
        $totalUsers = $allUsers.Count
        $processedUsers = 0
        
        foreach ($user in $allUsers) {
            $processedUsers++
            $percentComplete = [math]::Round(($processedUsers / $totalUsers) * 100, 2)
            Write-Progress -Activity "共有設定を確認中" -Status "$processedUsers / $totalUsers ユーザー処理中 ($percentComplete%)" -PercentComplete $percentComplete
            
            # このユーザーの処理ステータスを初期化
            $oneDriveStatus = "不明"
            $userAccountStatus = if($user.AccountEnabled){"有効"}else{"無効"}
            $userHasShares = $false # このユーザーが共有設定を持っているかのフラグ

            try {
                # アカウント無効ユーザーの処理
                if (-not $user.AccountEnabled) {
                    $oneDriveStatus = "未対応 (アカウント無効)"
                    $allUserStatuses += [PSCustomObject]@{
                        "ユーザー名"       = $user.DisplayName
                        "メールアドレス"   = $user.Mail
                        "アカウント状態"   = $userAccountStatus
                        "OneDrive対応状況" = $oneDriveStatus
                        "共有方向"         = "N/A"
                        "アイテム名"       = "N/A"
                        "アイテムタイプ"   = "N/A"
                        "共有タイプ"       = "N/A"
                        "共有範囲"         = "N/A"
                        "共有相手"         = "N/A"
                        "リスクレベル"     = "N/A"
                        "最終更新日時"     = "N/A"
                        "WebURL"           = "N/A"
                        "権限ID"           = "N/A"
                    }
                    continue # 次のユーザーへ
                }
                }

                # 有効なユーザーのOneDriveを取得
                if (-not [string]::IsNullOrEmpty($user.UserPrincipalName)) {
                    try {
                        Write-Log "ユーザー $($user.UserPrincipalName) のOneDriveドライブを取得中..." "DEBUG"
                        $driveInfo = Get-MgUserDrive -UserId $user.Id -ErrorAction SilentlyContinue # UserPrincipalNameではなくIdを使用
                        if (-not $driveInfo) {
                            $oneDriveStatus = "未対応 (アクセス不可/存在せず)"
                            Write-Log "ユーザー $($user.UserPrincipalName) のOneDriveドライブが見つからないか、アクセス権がありません。" "WARNING"
                            $allUserStatuses += [PSCustomObject]@{ "ユーザー名" = $user.DisplayName; "メールアドレス" = $user.Mail; "アカウント状態" = $userAccountStatus; "OneDrive対応状況" = $oneDriveStatus; "共有方向"="N/A"; "アイテム名"="N/A"; "アイテムタイプ"="N/A"; "共有タイプ"="N/A"; "共有範囲"="N/A"; "共有相手"="N/A"; "リスクレベル"="N/A"; "最終更新日時"="N/A"; "WebURL"="N/A"; "権限ID"="N/A" }
                            continue # 次のユーザーへ
                        }
                        # 複数のドライブが返る可能性を考慮し、最初のものを取得
                        $drive = $driveInfo[0]
                        # ドライブオブジェクト自体またはそのIDが取得できたか確認し、空または空白ならスキップ
                        if ($drive -eq $null -or [string]::IsNullOrWhiteSpace($drive.Id)) {
                            $oneDriveStatus = "未対応 (ドライブID取得失敗)"
                            Write-Log "ユーザー $($user.UserPrincipalName) のドライブ情報またはドライブIDが取得できませんでした (空または空白)。スキップします。" "WARNING"
                            $allUserStatuses += [PSCustomObject]@{ "ユーザー名" = $user.DisplayName; "メールアドレス" = $user.Mail; "アカウント状態" = $userAccountStatus; "OneDrive対応状況" = $oneDriveStatus; "共有方向"="N/A"; "アイテム名"="N/A"; "アイテムタイプ"="N/A"; "共有タイプ"="N/A"; "共有範囲"="N/A"; "共有相手"="N/A"; "リスクレベル"="N/A"; "最終更新日時"="N/A"; "WebURL"="N/A"; "権限ID"="N/A" }
                            continue # 次のユーザーへ
                        }
                        # DriveId が有効な場合のみログ出力と後続処理へ
                        Write-Log "ユーザー $($user.UserPrincipalName) のドライブID: $($drive.Id)" "DEBUG"

                        # ドライブ内の全アイテムを取得 (権限情報は別途取得)
                        Write-Log "ユーザー $($user.UserPrincipalName) のドライブ $($drive.Id) 内のアイテムを取得中..." "DEBUG"
                        $items = $null # items変数を初期化
                        try {
                            # Get-MgDriveItem 呼び出しを try-catch で囲む
                            try {
                                # -Filter "id ne ''" を追加して filter 要求エラーを回避試行
                                # -ErrorAction Stop を追加してエラーを捕捉可能にする
                                $items = Get-MgDriveItem -DriveId $drive.Id -All -Property Id,Name,Folder,File,LastModifiedDateTime,WebUrl -Filter "id ne ''" -ErrorAction Stop 
                            } catch { # Get-MgDriveItem のエラーを捕捉
                                $oneDriveStatus = "未対応 (アイテム取得エラー: $($_.Exception.Message))"
                                Write-Log "ユーザー $($user.UserPrincipalName) のドライブアイテム取得中にエラーが発生しました: $($_.Exception.Message)" "WARNING"
                                # Filter 要求エラーの場合、-All なしで再試行
                                if ($_.Exception.Message -like "*The 'filter' query option must be provided*") {
                                    Write-Log "  -> Filterエラーのため、ルートアイテムのみ取得して再試行します..." "WARNING"
                                    try {
                                        # 再試行時も -ErrorAction Stop を追加
                                        $items = Get-MgDriveItem -DriveId $drive.Id -Property Id,Name,Folder,File,LastModifiedDateTime,WebUrl -ErrorAction Stop # -All を削除
                                        $oneDriveStatus = "対応 (ルートアイテムのみ)" # 재시도 성공 시에는 상태 변경
                                    } catch { 
                                        # 재시도 시에도 -ErrorAction Stop 을 추가
                                        $items = Get-MgDriveItem -DriveId $drive.Id -Property Id,Name,Folder,File,LastModifiedDateTime,WebUrl -ErrorAction Stop # -All 을 삭제
                                        $oneDriveStatus = "대응 (루트 아이템만)" # 재시도 성공 시에는 상태 변경
                                    }
                                } else {
                                    # Filter요구 이외의 에러의 경우는 상세를 로그에 기록해 스킵
                                    Write-Log "  -> Filter요구 이외의 에러입니다. 스킵합니다。" "WARNING"
                                    Write-Log "  상세: $($_.ScriptStackTrace)" "DEBUG"
                                    $allUserStatuses += [PSCustomObject]@{ "사용자 이름" = $user.DisplayName; "메일 주소" = $user.Mail; "계정 상태" = $userAccountStatus; "OneDrive 대응 상황" = $oneDriveStatus; "공유 방향"="N/A"; "아이템 이름"="N/A"; "아이템 유형"="N/A"; "공유 유형"="N/A"; "공유 범위"="N/A"; "공유 상대"="N/A"; "위험 수준"="N/A"; "最終更新時間"="N/A"; "WebURL"="N/A"; "権限ID"="N/A" }
                                    continue # 다음 사용자로
                                }
                            }
                        } catch { # 바깥쪽의 try-catch (예기치 않은 에러용)
                             $oneDriveStatus = "미대응 (아이템 획득 처리 에러: $($_.Exception.Message))"
                             Write-Log "사용자 $($user.UserPrincipalName) 의 드라이브 아이템 획득 처리 중에 예상치 못한 오류가 발생했습니다: $($_.Exception.Message)" "WARNING"
                             Write-Log "상세 정보: $($_.ScriptStackTrace)" "DEBUG"
                             $allUserStatuses += [PSCustomObject]@{ "사용자 이름" = $user.DisplayName; "메일 주소" = $user.Mail; "계정 상태" = $userAccountStatus; "OneDrive 대응 상황" = $oneDriveStatus; "공유 방향"="N/A"; "아이템 이름"="N/A"; "아이템 유형"="N/A"; "공유 유형"="N/A"; "공유 범위"="N/A"; "공유 상대"="N/A"; "위험 수준"="N/A"; "最終更新時間"="N/A"; "WebURL"="N/A"; "権限ID"="N/A" }
                             continue # 다음 사용자로
                        }

                        # アイテム 획득에 성공한 경우 (에러로부터의 복귀 포함)
                        if ($oneDriveStatus -notlike "미대응*") {
                            $oneDriveStatus = "대응" # 기본 상태를 「대응」으로 설정
                        }

                        # アイテム을 획득할 수 없었던 경우는 스킵 (에러 처리 완료되었을 것이지만 만약을 위해)
                        if ($items -eq $null) {
                            $oneDriveStatus = "대응 (아이템 없음)"
                            Write-Log "사용자 $($user.UserPrincipalName) 의 드라이브에 아이템이 발견되지 않았습니다。" "DEBUG"
                            # 공유는 없지만, 상태는 기록
                            $allUserStatuses += [PSCustomObject]@{ "사용자 이름" = $user.DisplayName; "메일 주소" = $user.Mail; "계정 상태" = $userAccountStatus; "OneDrive 대응 상황" = $oneDriveStatus; "공유 방향"="N/A"; "아이템 이름"="N/A"; "아이템 유형"="N/A"; "공유 유형"="N/A"; "공유 범위"="N/A"; "공유 상대"="N/A"; "위험 수준"="N/A"; "最終更新時間"="N/A"; "WebURL"="N/A"; "権限ID"="N/A" }
                            continue # 다음 사용자로
                        }
                        Write-Log "사용자 $($user.UserPrincipalName) 의 드라이브로부터 $($items.Count) 개의 아이템을 획득했습니다。" "DEBUG"

                        # アイテム마다 권한을 획득해 처리
                        foreach ($item in $items) {
                            # Write-Log "  아이템 '$($item.Name)' (ID: $($item.Id)) 의 권한을 획득 중..." "DEBUG" # 상세 로그 삭제
                            try {
                                # アイテム마다 권한을 획득
                                $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $item.Id -All -ErrorAction Stop 
                                if ($permissions -eq $null -or $permissions.Count -eq 0) {
                                    # Write-Log "  아이템 '$($item.Name)' 에는 명시적인 권한 정보가 없습니다。" "DEBUG"
                                    continue # 권한 정보가 없는 아이템은 스킵 (다음 아이템으로)
                                }
                                # Write-Log "  아이템 '$($item.Name)' 로부터 $($permissions.Count) 개의 권한을 획득했습니다。" "DEBUG" # 상세 로그 삭제

                                foreach ($permission in $permissions) {
                                    # 1. 소유자 권한은 스킵
                                    if ($permission.Roles -contains 'owner') {
                                        # Write-Log "    -> Skipping (Owner)" "DEBUG"
                                        continue
                                    }

                                    # 2. 링크 공유인지 직접 공유인지를 판정
                                    $isLinkShare = $permission.Link
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
