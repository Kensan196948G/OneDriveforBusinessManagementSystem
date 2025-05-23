# OneDrive for Business 運用ツール 運用手順書

---

## 1. 運用の目的

- OneDrive利用状況の可視化  
- セキュリティリスクの早期発見  
- ITSMに準拠した継続的な改善  
- インシデントの迅速な検知と対応

---

## 2. 役割分担

| 役割             | 主な作業内容                                         |
|------------------|----------------------------------------------------|
| グローバル管理者 | 初期設定、権限付与、全体監視、レポート確認、改善指示 |
| 一般ユーザー     | 自身のOneDrive状況確認、軽微な対応                   |
| ヘルプデスク     | 一般ユーザー支援、一次対応、管理者へのエスカレーション |

---

## 3. 日常運用フロー

### 3.1 データ取得方法の選択
- **CSVから生成する場合**:
  - 既存のCSVデータを使用
  - 認証不要でオフライン利用可能
- **Graph APIから直接取得する場合**:
  - Azure AD認証が必要
  - 最新データをリアルタイムで取得可能

### 3.2 日次運用

- `Main.ps1`を起動し、  
  - **インシデント管理 → 同期エラー確認**  
  - **セキュリティ管理 → 外部共有監視**  
  を実行
- エラー・高リスク共有があれば即時対応・通知
- 必要に応じて**基本データ収集**も実施

### 3.2 週次

- `Main.ps1`で**変更管理 → 共有設定確認**を実行
- **総合レポート生成**を行い、全体状況を把握
- 共有設定の見直し・改善策の検討

### 3.3 月次

- **全カテゴリのスクリプトを実行し、総合レポートを保存・共有**
- 利用状況・リスクの傾向を分析
- 改善計画の策定・関係者への報告

---

## 4. レポートの確認ポイント

- **同期エラー**  
  - 頻発ユーザーの特定  
  - エラー種別別の対応策実施
- **共有設定**  
  - 「誰でも（編集可能）」の高リスク共有の削減  
  - 不要な共有の解除
- **外部共有**  
  - 匿名リンクの制限  
  - 不審な共有先の確認・遮断
- **容量超過**  
  - 90%以上のユーザーに対し容量追加・整理指導

---

## 5. トラブル対応

| 事象                           | 対応策                                                         |
|--------------------------------|----------------------------------------------------------------|
| Graph認証失敗                  | `config.json`の見直し、Azure AD設定確認、権限の再付与を実施   |
| CSVデータ不整合                | データソース確認、CSV再生成、文字コード(UTF-8)を確認          |
| スクリプトエラー               | ログファイル確認、権限・接続状況確認                           |
| データ取得不可                 | ネットワーク・API制限・権限設定確認                            |
| レポートに異常値・空欄が多い   | API制限・権限不足・一時的障害の可能性、再試行                  |
| 高リスク共有・外部共有が多発   | 共有ポリシーの見直し、ユーザー教育、必要に応じて共有制限設定   |

---

## 6. 補足

- **Main.ps1**から全ての操作が可能  
- 出力ファイルは日付付きで自動管理  
- HTMLレポートは検索・フィルタ・色分け・エクスポート・印刷対応  
- **Graph APIの仕様変更時はスクリプトの更新が必要**
- **現在のシステム動作状況 (2025/04/28時点)**
  - Main.ps1: 認証テスト実行中
  - GetUserInfo.ps1: テストユーザーでの動作確認中
  - WebUI: ポート8005でHTTPサーバー稼働中

---

生成日時：20250428
作成者  
問い合わせ：help@mirai-const.co.jp
