# 📄 OneDrive for Business 運用ツール - ITSM準拠

このツールは、**OneDrive for Business** の運用管理を **ITSM（ITサービスマネジメント）** の原則に基づいて効率化する **PowerShellスクリプト集** です！

🚀 作業を切り開く！精彩なIT運用を支えます！

---

## 📚 概要

OneDrive運用ツールは、以下の4つの主要機能を提供します。

| 分類               | 内容                                       |
|:-------------------|:-------------------------------------------|
| 💡 基本データ収集     | ユーザー情報・ストレージ情報の収集         |
| ⚡ インシデント管理   | 同期エラーなど障害情報の検出・管理         |
| 🔄 変更管理           | 共有設定の変更・監査                       |
| 🔒 セキュリティ管理   | 外部共有リスクの監視・レポート       |

---

## ⚙️ 前提条件

- Windows PowerShell 5.1以上 または PowerShell Core 6.0以上
- Microsoft 365アカウント（管理者権限推奨）
- Microsoft Graph PowerShell SDK
- インターネット接続

---

## 📥 インストール方法

1. このリポジトリをダウンロードまたはクローン
2. PowerShellを「管理者権限」で起動
3. 必要に応じ実行ポリシーを変更
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

---

## 🚀 使用方法

1. PowerShellを開きダウンロードフォルダへ移動
2. `Main.ps1`を実行
   ```powershell
   .\Main.ps1
   ```
3. メニューから機能を選択
4. 案内に従って操作

---

## 📊 出力ファイルと保存先

出力データ（CSV・HTML・JS）および実行ログは、**実行日ベース** のフォルダに自動保存されます。さらに、リアルタイム実行ログとエラーログも自動生成されます。

📂 保存先構成：
```
OneDrive運用ツール/
  └— OneDriveManagement.YYYYMMDD/
       ├— Basic_Data_Collection-YYYYMMDD/
       ├— Incident_Management-YYYYMMDD/
       ├— Change_Management-YYYYMMDD/
       ├— Security_Management-YYYYMMDD/
       └— LogYYYYMMDD/
            ├— ExecutionLogYYYYMMDDMMSS.txt
            └— ErrorLogYYYYMMDDMMSS.txt
```

🔹 カテゴリ別に明確に保存！
🔹 実行日つきで調査も楽々！

---

## 🗂️ フォルダ構成

```
OneDrive運用ツール/
│   .gitignore                  # GitHub用除外設定
│   GenerateReport.ps1           # 総合レポート生成
│   Main.ps1                     # メインスクリプト
│   README.md                    # ユーザー向けガイド
│   README_ADMIN.md              # 管理者向け設定ガイド
│   WebUI_Specification.md       # WebUI作成仕様書
│
├───Basic_Data_Collection        # 基本データ収集
│       GetAllBasicData.ps1       # 全て一括収集
│       GetOneDriveQuota.ps1      # ストレージ収集
│       GetUserInfo.ps1           # ユーザー情報収集
│
├───Change_Management            # 変更管理
│       SharingCheck.ps1          # 共有設定確認
│
├───Doc                           # ドキュメント保管
│
├───Incident_Management          # インシデント管理
│       SyncErrorCheck.ps1        # 同期エラー確認
│
├───Security_Management          # セキュリティ管理
│       ExternalShareCheck.ps1    # 外部共有監視
│
├───TEMP                          # 一時ファイル保存
│
├───WebUI_Sample                  # WebUIサンプル
│
└───WebUI_Template                # WebUIテンプレート
        common.js                  # 共通JavaScript
        ExternalShareCheck.html    # 外部共有チェックUI
        GenerateReport.html        # レポート生成UI
        GetAllBasicData.html       # 基本データ収集UI
        GetOneDriveQuota.html      # ストレージ情報UI
        GetOneDriveQuota.ps1       # ストレージ情報スクリプト
        GetUserInfo.html           # ユーザー情報UI
        GetUserInfo.ps1            # ユーザー情報スクリプト
        SharingCheck.html          # 共有設定UI
        SyncErrorCheck.html        # 同期エラーUI
        HTMLコピー/                # HTMLテンプレートバックアップ
            ExternalShareCheck.html
            GenerateReport.html
            GetAllBasicData.html
            GetOneDriveQuota.html
            GetUserInfo.html
            SharingCheck.html
            SyncErrorCheck.html
```

---

## 🛡️ 必要権限

このツールの利用には，次の Microsoft Graph API 権限が必要です。

- `User.Read.All`
- `Directory.Read.All`
- `Sites.Read.All`
- `Files.Read.All`

🔹 詳細は [README_ADMIN.md](./README_ADMIN.md) をご参照ください！

---

## 💡 特記事項

- 本ツールは **OneDriveデータの読み取りのみ** を実行します。
- データの変更や削除はしません。
- 操作記録ログを残し、監査跡象を確保します。
