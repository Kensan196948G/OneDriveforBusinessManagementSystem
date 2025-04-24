const { spawn } = require('child_process');
const os = require('os');
const EventEmitter = require('events');

class MCPConnector extends EventEmitter {
  constructor() {
    super();
    this.process = null;
    this.buffer = '';
    this.ready = false;
    this.responseHandlers = [];
  }

  start() {
    const command = os.platform() === 'win32' ? 'npx.cmd' : 'npx';
    this.process = spawn(command, ['-y', '@upstash/context7-mcp@latest'], {
      stdio: ['pipe', 'pipe', 'pipe'],
      shell: true
    });

    this.process.stdout.on('data', (data) => {
      this.buffer += data.toString();
      
      // 改行ごとにメッセージを処理
      const messages = this.buffer.split('\n');
      this.buffer = messages.pop(); // 最後の不完全な行を保持
      
      messages.forEach(msg => {
        msg = msg.trim();
        if (msg) {
          if (msg.includes('MCP Server ready')) {
            this.ready = true;
            this.emit('ready');
          } else if (this.responseHandlers.length > 0) {
            const handler = this.responseHandlers.shift();
            handler(msg);
          } else {
            console.log(`MCP出力: ${msg}`);
            this.emit('message', msg);
          }
        }
      });
    });

    this.process.stderr.on('data', (data) => {
      console.error(`MCPエラー: ${data.toString().trim()}`);
    });

    this.process.on('close', (code) => {
      console.log(`MCPプロセス終了、コード: ${code}`);
      this.emit('close', code);
    });
  }

  async sendCommand(command) {
    if (!this.process) {
      throw new Error('MCPプロセスが起動していません');
    }

    if (!this.ready) {
      await new Promise(resolve => this.once('ready', resolve));
    }

    return new Promise((resolve) => {
      this.responseHandlers.push(resolve);
      this.process.stdin.write(`${command}\n`);
    });
  }

  stop() {
    if (this.process) {
      this.process.kill();
    }
  }
}

module.exports = MCPConnector;