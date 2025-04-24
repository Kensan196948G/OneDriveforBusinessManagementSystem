const { MCPClient } = require("@modelcontextprotocol/sdk");

const run = async () => {
  console.log("MCPクライアントを初期化しています...");
  const client = new MCPClient({
    server: "stdio",
  });

  try {
    console.log("MCPサーバーに接続中...");
    await client.connect();
    
    console.log("利用可能なツールを取得中...");
    const tools = await client.listTools();
    console.log("利用可能なツール:", tools);
  } catch (error) {
    console.error("接続エラー:", error);
  }
};

run();