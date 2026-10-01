/**
 * Smoke test: stdout of the stdio server is the JSON-RPC stream, so every line
 * written there must be a JSON-RPC message (diagnostics go to stderr).
 */
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

jest.setTimeout(20000);

test('initialize + tools/list yield only JSON-RPC lines on stdout', async () => {
  // Empty home directory: no token file is read or written from the real profile.
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-smoke-'));
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'index.js')], {
    env: { ...process.env, HOME: home, USERPROFILE: home, USE_TEST_MODE: 'false' },
    stdio: ['pipe', 'pipe', 'pipe']
  });

  let stdout = '';
  server.stdout.on('data', (chunk) => { stdout += chunk; });
  server.stderr.resume();

  const send = (msg) => server.stdin.write(JSON.stringify(msg) + '\n');
  const waitFor = (id) => new Promise((resolve, reject) => {
    const timer = setInterval(() => {
      if (stdout.split('\n').some(line => line.includes(`"id":${id}`))) {
        clearInterval(timer);
        resolve();
      }
    }, 50);
    server.on('exit', (code) => { clearInterval(timer); reject(new Error(`server exited with ${code}`)); });
  });

  try {
    send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'smoke', version: '1' } } });
    await waitFor(1);
    send({ jsonrpc: '2.0', method: 'notifications/initialized' });
    send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} });
    await waitFor(2);
    // check-auth-status reads the (missing) token file: its logging must not reach stdout either.
    send({ jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'check-auth-status', arguments: {} } });
    await waitFor(3);
  } finally {
    server.kill();
    fs.rmSync(home, { recursive: true, force: true });
  }

  const lines = stdout.split('\n').filter(line => line.trim() !== '');
  const messages = lines.map(line => JSON.parse(line)); // throws on any non-JSON line
  for (const msg of messages) {
    expect(msg.jsonrpc).toBe('2.0');
  }
  const byId = Object.fromEntries(messages.map(m => [m.id, m]));
  expect(byId[1].result.serverInfo).toBeDefined();
  expect(byId[2].result.tools.map(t => t.name)).toEqual(expect.arrayContaining(['authenticate', 'check-auth-status', 'read-email', 'send-email']));
  expect(byId[3].result.content[0].text).toMatch(/^Not authenticated/);
});
