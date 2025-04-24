# 📝 Context7 操作手順マニュアル

## 🔄 標準操作フロー
1. <span style="color: #9C27B0;">🔌 接続開始</span>
   ```javascript
   const mcp = new MCPConnector();
   mcp.start();
   ```

2. <span style="color: #3F51B5;">🛠️ ツール選択</span>
   ```javascript
   const tools = await mcp.sendCommand('list-tools');
   ```

3. <span style="color: #4CAF50;">🚀 コマンド実行</span>
   ```javascript
   const result = await mcp.sendCommand('analyze-data');
   ```

4. <span style="color: #FF9800;">📊 結果解析</span>
   ```javascript
   const analysis = analyzeMCPResult(result);
   ```

## 🚨 エラー対応
<div style="background-color: #FFEBEE; padding: 10px; border-left: 4px solid #F44336;">
❗ <strong>接続エラー時の対応</strong>:
1. サーバーが起動しているか確認
2. ネットワーク接続を確認
3. ログを確認
</div>

## ⏱️ タイムアウト設定
```javascript
// 30秒でタイムアウト
setTimeout(() => {
  mcp.stop();
}, 30000);
```

<div style="color: #607D8B; font-style: italic;">
💡 ヒント: コマンドは必ずawaitで待機してください
</div>