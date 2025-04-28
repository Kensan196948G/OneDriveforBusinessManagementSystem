# OneDrive管理ツール WebUI 作成仕様書

## 1. 概要
OneDrive for Businessの管理ツールで使用するWebUIの仕様書です。各スクリプトの出力を表示するHTMLレポートの要件を定義します。

## 2. 共通仕様
- レスポンシブデザイン対応
- Bootstrap 5.3.0を使用
- Font Awesome 6.4.0を使用
- 以下の共通機能を実装:
  - 検索/フィルタ機能
  - CSVエクスポート
  - 印刷機能
  - ページネーション

### 2.1 認証方式
各レポートは以下の2つの方法でデータを取得可能:

1. **CSVファイルから生成**:
   - 既存のCSVデータを使用
   - 認証不要
   - オフライン環境で利用可能

2. **Microsoft Graph APIから直接取得**:
   - Azure ADアプリ登録が必要
   - 必要な権限:
     - User.Read.All
     - Files.Read.All  
     - Sites.Read.All
   - config.jsonに認証情報を設定
   - 非対話型認証(client_credentials grant)を使用
   
### 2.2 ローディングオーバーレイ仕様
- 以下の操作時に表示:
  - ページ初期表示
  - データ再取得
  - リロード操作
  - 手動更新操作
  
- 仕様詳細:
  - 全画面に半透明の背景（rgba(255, 255, 255, 0.9)）を表示
  - 中央にスピナーアニメーションとメッセージを表示
  - 最低1.5秒間表示を保証（高速なデータ取得時も）
  - エラー時はオーバーレイを閉じてからエラー表示
  - ユーザー操作をブロックする仕様

例：
```html
<div id="loading-overlay" style="display:none;">
  <div class="spinner-border text-primary" role="status">
    <span class="visually-hidden">Loading...</span>
  </div>
  <p>読み込み中...</p>
</div>
```

### 2.3 エラー表示仕様
- 初期状態: 表示（"データ読み込み中"メッセージ）
- 成功時: 非表示
- 失敗時:
  - エラーメッセージ表示
  - 再試行/再読み込みボタン表示
  - ITサポート連絡先表示

### 2.4 操作ボタン標準仕様
- **再読み込みボタン**:
  - ページ全体をリロード
  - ローディング表示をトリガー
- **再試行ボタン**:
  - データ取得のみ再実行
  - ローディング表示をトリガー

制御例：
```javascript
// ローディング表示（最低1.5秒保証）
function showLoading() {
  const overlay = document.getElementById('loadingOverlay');
  overlay.style.display = 'flex';
  return Date.now(); // 開始時刻記録
}

function hideLoading(startTime) {
  const elapsed = Date.now() - startTime;
  const remaining = Math.max(0, 1500 - elapsed);
  setTimeout(() => {
    document.getElementById('loadingOverlay').style.display = 'none';
  }, remaining);
}

// エラー表示
function showError(message) {
  document.getElementById('errorContainer').style.display = 'block';
  document.getElementById('errorMessage').textContent = message;
}
```

## 3. 各レポート仕様

### 3.1 ユーザー情報取得 (GetUserInfo.html)
- **表示項目**:
  - ユーザー名
  - メールアドレス
  - ログインユーザー名
  - ユーザー種別
  - アカウント状態
  - 最終同期日時

- **特別な機能**:
  - アカウント状態による色分け
  - 最終同期日時によるソート

### 3.2 ストレージクォータ取得 (GetOneDriveQuota.html)
- **表示項目**:
  - ユーザー名
  - メールアドレス
  - 総容量
  - 使用容量
  - 残り容量
  - 使用率
  - 状態

- **特別な機能**:
  - 使用率による色分け
  - クォータ警告表示

### 3.3 基本データ収集 (GetAllBasicData.html)
- **表示項目**:
  - ユーザー名
  - メールアドレス
  - ログインユーザー名
  - ユーザー種別
  - アカウント状態
  - OneDrive対応
  - クォータ
  - 最終アクセス日

- **特別な機能**:
  - 複合検索
  - 状態によるフィルタ

### 3.4 同期エラー確認 (SyncErrorCheck.html)
- **表示項目**:
  - ユーザー名
  - メールアドレス
  - エラー内容
  - 発生日時
  - ステータス

- **特別な機能**:
  - エラー重要度による色分け
  - 緊急度表示

### 3.5 共有設定確認 (SharingCheck.html)
- **表示項目**:
  - 所有者
  - 共有アイテム
  - 権限
  - 共有先
  - 共有日
  - リスクレベル

- **特別な機能**:
  - リスクレベルによる色分け
  - 外部共有警告

### 3.6 外部共有監視 (ExternalShareCheck.html)
- **表示項目**:
  - 所有者
  - 共有アイテム
  - 外部共有先
  - 共有日
  - 推奨アクション

- **特別な機能**:
  - 外部ドメイン検出
  - 機密ファイル警告

### 3.7 総合レポート (GenerateReport.html)
- **表示項目**:
  - サマリー情報
  - グラフ表示
  - 直近のアラート
  - 統計データ

- **特別な機能**:
  - ダッシュボード表示
  - PDF出力
  - インタラクティブなグラフ

## 4. 画面レイアウト
各レポートの画面レイアウトは以下の構成とする:
1. ヘッダー (タイトル)
2. フィルタ/操作エリア
3. データ表示エリア (テーブルまたはダッシュボード)
4. フッター (ページネーション/操作ボタン)

## 5. 技術要件
- HTML5, CSS3, JavaScript (ES6)
- ブラウザ互換性: Chrome最新版, Edge最新版, Firefox最新版
- 印刷用CSSを別途定義
- パフォーマンス: 1000件以上のデータでも快適に操作可能

## 6. セキュリティ要件
- 機密情報はマスク表示
- CSVエクスポート時はパスワード保護
- 印刷時は機密情報を省略可能

## 7. 参考資料
- [OneDrive運用ツール_運用手順.md](Doc/OneDrive運用ツール_運用手順.md)
- [OneDrive運用ツール_仕様と利用手順.md](Doc/OneDrive運用ツール_仕様と利用手順.md)
