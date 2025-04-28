# OneDrive for Business 運用ツール 仕様と利用手順

---

## 1. 概要

本ツールは、Microsoft 365 OneDrive for Businessの運用管理を  
**ITSM（ITサービスマネジメント）準拠**で効率的に行うためのPowerShellスクリプト集です。

Microsoft Graph APIを利用し、  
- ユーザー情報  
- OneDriveストレージ使用状況  
- 共有設定  
- 同期エラー  
- 外部共有状況  
を自動収集・レポート化します。

---

## 2. 認証方式

- **Azure ADアプリ登録のクライアントID・シークレット・テナントIDを`config.json`に格納**
- **クライアントシークレット方式（client_credentials grant）でアクセストークンを取得**
- **Microsoft Graph REST APIに対してBearerトークンでアクセス**

---

## 3. 構成

```
OneDrive運用ツール/
├── Main.ps1                 # メインメニュー・統合実行
├── GenerateReport.ps1       # 総合レポート生成
├── config.json              # 認証情報・設定
├── Basic_Data_Collection/   # 基本情報収集
│   ├── GetUserInfo.ps1
│   ├── GetOneDriveQuota.ps1
│   ├── GetAllBasicData.ps1
│   └── GetUserInfoScriptBackUp/
├── Change_Management/
│   └── SharingCheck.ps1
├── Incident_Management/
│   └── SyncErrorCheck.ps1
├── Security_Management/
│   └── ExternalShareCheck.ps1
├── OneDriveManagement.YYYYMMDD/  # 出力先（自動生成）
│   ├── Log/
│   └── Report/
├── README.md
├── README_ADMIN.md
└── OneDrive for Business 運用ツールの ITSM 準拠管理項目Final.txt
```

---

## 4. 利用手順

### 4.1 事前準備

- Azure ADにアプリ登録し、必要なGraph API権限を付与
- `config.json`に  
  - TenantId  
  - ClientId  
  - ClientSecret  
  - 必要なスコープ  
  を設定

### 4.2 実行方法

1. PowerShellを管理者権限で起動
2. 実行ポリシーを許可  
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. 作業ディレクトリに移動し  
   ```powershell
   .\Main.ps1
   ```
4. メニューから  
   - 基本データ収集  
   - インシデント管理  
   - 変更管理  
   - セキュリティ管理  
   - 総合レポート生成  
   を選択

### 4.3 出力物

- CSV, HTML, JavaScriptファイル
- 日付付きフォルダに自動保存
- HTMLは検索・フィルタ・色分け・エクスポート・印刷対応

---

## 5. 権限要件

- **User.Read.All**  
- **Directory.Read.All**  
- **Sites.Read.All**  
- **Files.Read.All**  
（すべて管理者の同意が必要）

---

## 6. 注意事項

- 一部の共有設定・同期エラー・外部共有は**シミュレーションデータ**であり、  
  実装時はGraph APIでの実データ取得が必要
- 一般ユーザーも利用可能（管理者の事前同意が必要）
- 操作ログ・エラーログは全て保存
- **現在のシステム動作状況 (2025/04/28時点)**
  - Main.ps1: 認証テスト実行中
  - GetUserInfo.ps1: テストユーザーでの動作確認中
  - WebUI: ポート8005でHTTPサーバー稼働中

---

生成日時：20250428
作成者  
問い合わせ：help@mirai-const.co.jp
