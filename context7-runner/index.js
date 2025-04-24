const MCPConnector = require('./mcp-connector');
const { analyzeMCPResult } = require("../reasoning/analyze");

async function runMCP() {
  console.log("🚀 新しいMCP接続方法を開始します");
  
  const mcp = new MCPConnector();
  mcp.start();

  // メッセージ受信処理
  mcp.on('message', (msg) => {
    console.log("📢 MCPメッセージ:", msg);
  });

  try {
    // サーバー準備待ち
    await new Promise(resolve => mcp.once('ready', resolve));
    console.log("✅ MCPサーバー準備完了");

    // ツール一覧取得
    const toolsResponse = await mcp.sendCommand('list-tools');
    console.log("🔧 ツール一覧レスポンス:", toolsResponse);

    // レスポンス解析
    try {
      const toolsData = JSON.parse(toolsResponse);
      if (toolsData.tools) {
        console.log("🛠️ 利用可能なツール:", toolsData.tools.join(', '));
      }
    } catch (err) {
      console.error("JSON解析エラー:", err.message);
    }

    // ダミーデータで分析処理をテスト
    const dummyResult = {
      status: "success",
      message: "MCP接続が確立されました",
      tools: ["context-analyzer", "data-processor"]
    };

    const analysis = await analyzeMCPResult(dummyResult);
    console.log("🧠 DeepSeek Reasoner分析結果:", analysis);

  } catch (err) {
    console.error("❌ 実行中にエラー:", err.message);
  } finally {
    // 30秒後に接続を閉じる
    setTimeout(() => {
      mcp.stop();
      console.log("🛑 MCP接続を終了しました");
    }, 30000);
  }
}

runMCP();