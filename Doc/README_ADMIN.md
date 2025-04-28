# OneDrive for Business 運用ツール - グローバル管理者向け設定ガイド

このドキュメントは、OneDrive for Business 運用ツールを使用するためにグローバル管理者が事前に行う必要がある設定について説明しています。

## 必要なアクセス許可

このツールは以下のMicrosoft Graph APIのアクセス許可を使用します：

- **User.Read.All** - すべてのユーザーの完全なプロファイルを読み取る権限
- **Directory.Read.All** - ディレクトリデータの読み取り権限
- **Sites.Read.All** - すべてのサイトコレクションのアイテムを読み取る権限
- **Files.Read.All** - すべてのファイルの読み取り権限

これらのアクセス許可はすべて「管理者の同意」が必要です。

## 一般ユーザーへの権限委任について

グローバル管理者が上記のパーミッションに対して「管理者の同意」を付与すれば、**一般ユーザーでもこれらの情報を取得できるようになります**。これは「委任されたアクセス許可」として設定されるため、以下のような仕組みになります：

1. グローバル管理者がアプリケーション（OneDrive運用ツール）に対してこれらのパーミッションを付与し、「管理者の同意」を行います
（**一度設定すれば、グローバル管理者が明示的に取り消さない限り有効です**）
2. 一般ユーザーがツールを実行すると、自分のアカウントでログインしますが、グローバル管理者が事前に同意したパーミッションの範囲内で操作できます
3. 一般ユーザーは自分の権限を超えた操作（例：他のユーザーのOneDriveデータの閲覧）ができるようになります

ただし、以下の点に注意が必要です：

- 一般ユーザーは、あくまでもツールを通じてのみこれらの権限を行使できます
- 監査ログには、実際に操作を行ったユーザーの情報が記録されます
- グローバル管理者は、必要に応じて特定のユーザーにのみツールの使用を許可するなど、アクセス制御を行うことができます

このような設定は、IT管理者やヘルプデスク担当者など、業務上必要なユーザーに対して、グローバル管理者権限を付与せずに特定の管理タスクを委任するために有効です。

### 設定の有効期間について

グローバル管理者が行った設定（管理者の同意）は、以下の場合を除いて**永続的に有効**です：

- グローバル管理者が明示的に同意を取り消した場合
- アプリケーション（OneDrive運用ツール）が削除された場合
- Microsoft 365テナントのセキュリティポリシーが変更された場合（例：条件付きアクセスポリシーの適用）

## 設定手順

### 方法1: Azure Active Directoryポータルでの設定（推奨）

1. [Azure Active Directoryポータル](https://aad.portal.azure.com/)にグローバル管理者でログインします。

2. 左側のメニューから「エンタープライズアプリケーション」を選択します。

3. 「新しいアプリケーション」をクリックします。

4. 「独自のアプリケーションを作成」を選択します。
   - 名前: "OneDrive運用ツール"
   - タイプ: "この組織ディレクトリのみに存在するアプリケーション"

5. アプリケーションが作成されたら、「APIのアクセス許可」を選択します。

6. 「アクセス許可の追加」をクリックし、以下の手順で権限を追加します：
   - 「Microsoft Graph」を選択
   - 「委任されたアクセス許可」を選択
   - 以下の権限を検索して追加：
     - User.Read.All
     - Directory.Read.All
     - Sites.Read.All
     - Files.Read.All

7. すべての権限を追加したら、「管理者の同意を付与」ボタンをクリックします。

8. アプリケーションIDとテナントIDをメモしておきます（オプション）。

### 方法2: PowerShellでの設定

グローバル管理者は、以下のPowerShellコマンドを使用して必要な設定を行うこともできます：

```powershell
# Microsoft Graph PowerShellモジュールをインストール
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# 管理者権限でサインイン
Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

# アプリケーション登録
$app = New-MgApplication -DisplayName "OneDrive運用ツール" -SignInAudience "AzureADMyOrg"

# 必要なAPIアクセス許可を追加
$graphResourceId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph

# User.Read.All
$userReadAllId = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"
New-MgApplicationPermission -ApplicationId $app.Id -ResourceId $graphResourceId -ApiId $graphResourceId -PermissionId $userReadAllId

# Directory.Read.All
$directoryReadAllId = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"
New-MgApplicationPermission -ApplicationId $app.Id -ResourceId $graphResourceId -ApiId $graphResourceId -PermissionId $directoryReadAllId

# Sites.Read.All
$sitesReadAllId = "01d4889c-1287-42c6-ac1f-5d1e02578ef6"
New-MgApplicationPermission -ApplicationId $app.Id -ResourceId $graphResourceId -ApiId $graphResourceId -PermissionId $sitesReadAllId

# Files.Read.All
$filesReadAllId = "01d4889c-1287-42c6-ac1f-5d1e02578ef6"
New-MgApplicationPermission -ApplicationId $app.Id -ResourceId $graphResourceId -ApiId $graphResourceId -PermissionId $filesReadAllId

# 管理者の同意を付与
New-MgOauth2PermissionGrant -ClientId $app.AppId -ConsentType "AllPrincipals" -ResourceId $graphResourceId -Scope "User.Read.All Directory.Read.All Sites.Read.All Files.Read.All"

Write-Host "アプリケーションID: $($app.AppId)"
Write-Host "テナントID: $((Get-MgContext).TenantId)"
```

## 確認方法

設定が正しく行われたかを確認するには、以下の手順を実行します：

1. OneDrive運用ツールのMain.ps1を実行します。
2. Microsoft Graphへの接続が成功することを確認します。
3. 各機能（ユーザー情報取得、ストレージクォータ取得など）が正常に動作することを確認します。

## トラブルシューティング

### アクセス許可の問題

「アクセスが拒否されました」というエラーが表示される場合：

1. グローバル管理者として再度ログインしていることを確認します。
2. 必要なすべてのアクセス許可が追加され、管理者の同意が付与されていることを確認します。
3. Azure Active Directoryポータルで、アプリケーションのアクセス許可を確認します。

### 接続の問題

Microsoft Graphへの接続に問題がある場合：

1. インターネット接続を確認します。
2. Microsoft 365サービスの状態を確認します。
3. 必要に応じて、Main.ps1の「Microsoft Graph再接続」オプションを使用します。

## サポート

問題が解決しない場合は、IT管理者またはMicrosoft 365管理者にお問い合わせください。