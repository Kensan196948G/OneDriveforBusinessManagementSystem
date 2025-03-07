# OneDrive for Business 運用ツール - ITSM準拠

このツールは、OneDrive for Businessの運用管理をITSM（ITサービスマネジメント）の原則に基づいて行うためのPowerShellスクリプト集です。

## 概要

OneDrive for Business 運用ツールは、以下の4つの主要カテゴリに分かれた機能を提供します：

1. **基本データ収集** - ユーザー情報やストレージクォータなどの基本データを収集
2. **インシデント管理** - 同期エラーなどの問題を検出・管理
3. **変更管理** - 共有設定の変更を管理
4. **セキュリティ管理** - 外部共有などのセキュリティリスクを監視

各機能は、Microsoft Graph APIを使用してOneDriveのデータにアクセスし、CSV、HTML、JavaScriptファイルを生成して結果を表示します。

## 前提条件

- Windows PowerShell 5.1以上またはPowerShell Core 6.0以上
- Microsoft 365アカウント（管理者権限推奨）
- Microsoft Graph PowerShellモジュール
- インターネット接続

## インストール方法

1. このリポジトリをダウンロードまたはクローンします。
2. PowerShellを管理者権限で実行します。
3. 必要に応じて実行ポリシーを変更します：
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## 使用方法

1. PowerShellを開き、ダウンロードしたフォルダに移動します。
2. Main.ps1を実行します：
   ```powershell
   .\Main.ps1
   ```
3. メインメニューから実行したい機能を選択します。
4. 画面の指示に従って操作します。

## 機能詳細

### 基本データ収集 (Basic Data Collection)

- **ユーザー情報取得 (GetUserInfo.ps1)**
  - ユーザー名、メールアドレス、アカウント状態などの基本情報を取得
  - CSV、HTML形式でレポート出力

- **ストレージクォータ取得 (GetOneDriveQuota.ps1)**
  - 各ユーザーのOneDriveストレージ使用状況を取得
  - 総容量、使用容量、残り容量、使用率を表示

- **すべての基本データ収集 (GetAllBasicData.ps1)**
  - 上記の機能をまとめて実行

### インシデント管理 (Incident Management)

- **同期エラー確認 (SyncErrorCheck.ps1)**
  - OneDriveの同期エラーを検出
  - エラーの種類、発生日時、影響を受けるユーザーを特定

### 変更管理 (Change Management)

- **共有設定確認 (SharingCheck.ps1)**
  - OneDriveの共有設定を確認
  - 共有方法、共有先、アクセス権限を表示

### セキュリティ管理 (Security Management)

- **外部共有監視 (ExternalShareCheck.ps1)**
  - 組織外部との共有を監視
  - リスクレベルに応じた警告を表示

### 総合レポート生成 (Generate Report)

- すべてのカテゴリのデータを統合した総合レポートを生成
- タブ形式のHTMLレポートで、カテゴリごとにデータを表示

## 出力ファイル

すべての機能は、以下の形式でデータを出力します：

- **CSV** - データの詳細な分析やExcelでの加工に適したフォーマット
- **HTML** - インタラクティブな表示、検索、フィルタリング機能を備えたレポート
- **JavaScript** - HTMLレポートの動的機能を提供

出力ファイルは日付ベースのフォルダ（OneDriveManagement.YYYYMMDD）に保存され、各カテゴリごとにサブフォルダが作成されます。

## 権限要件

このツールは、Microsoft Graph APIを使用してOneDriveのデータにアクセスします。以下の権限が必要です：

- User.Read.All - ユーザー情報の取得に必要
- Directory.Read.All - ディレクトリ情報の取得に必要
- Sites.Read.All - OneDriveサイトへのアクセスに必要
- Files.Read.All - OneDriveファイルの読み取りに必要

これらの権限はグローバル管理者の同意が必要です。詳細は [README_ADMIN.md](./README_ADMIN.md) を参照してください。

## 一般ユーザーの利用について

グローバル管理者が事前に必要な権限設定を行うことで、一般ユーザーもこのツールを使用できます。一般ユーザーは自分のアカウントでログインしますが、グローバル管理者が同意した権限の範囲内で操作が可能です。

## ファイル構成

```
OneDrive運用ツール/
├── Main.ps1                           # メインスクリプト
├── GenerateReport.ps1                 # 総合レポート生成
├── README.md                          # このファイル
├── README_ADMIN.md                    # 管理者向け設定ガイド
├── Basic_Data_Collection/             # 基本データ収集スクリプト
│   ├── GetUserInfo.ps1                # ユーザー情報取得
│   ├── GetOneDriveQuota.ps1           # ストレージクォータ取得
│   └── GetAllBasicData.ps1            # すべての基本データ収集
├── Incident_Management/               # インシデント管理スクリプト
│   └── SyncErrorCheck.ps1             # 同期エラー確認
├── Change_Management/                 # 変更管理スクリプト
│   └── SharingCheck.ps1               # 共有設定確認
└── Security_Management/               # セキュリティ管理スクリプト
    └── ExternalShareCheck.ps1         # 外部共有監視
```

## トラブルシューティング

- **Microsoft Graphへの接続エラー**
  - メインメニューの「Microsoft Graph再接続」オプションを使用してください。
  - グローバル管理者に必要な権限が付与されているか確認してください。

- **スクリプト実行エラー**
  - PowerShellの実行ポリシーを確認してください。
  - 最新のPowerShellバージョンを使用しているか確認してください。

- **データ取得エラー**
  - インターネット接続を確認してください。
  - Microsoft 365サービスの状態を確認してください。

## 注意事項

- このツールは読み取り専用の操作のみを行います。OneDriveのデータを変更・削除することはありません。
- 大量のユーザーがいる環境では、データ収集に時間がかかる場合があります。
- すべての操作はログに記録され、監査可能です。

## サポート

問題が解決しない場合は、IT管理者またはMicrosoft 365管理者にお問い合わせください。
