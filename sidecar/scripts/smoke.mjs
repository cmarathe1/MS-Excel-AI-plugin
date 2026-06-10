/**
 * End-to-end smoke test, no Excel required:
 *
 *   mock OpenAI-compatible model  <--HTTP--  sidecar  --WS-->  fake add-in
 *                                                              (emulator as the workbook)
 *
 * Exercises: pairing, auth, provider settings API, the agent loop through a
 * real HTTP provider adapter, tool execution over the WS bridge protocol,
 * change-set staging/preview/approval, and formula recalculation.
 *
 * Run: pnpm build && node sidecar/scripts/smoke.mjs
 */
import { createServer } from 'node:http';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import WebSocket from 'ws';

process.env.EXCELAI_DATA_DIR = mkdtempSync(join(tmpdir(), 'excelai-smoke-'));

const { openDb } = await import('../dist/db.js');
const { startServer } = await import('../dist/server.js');
const { WorkbookEmulator } = await import('../dist/workbook/emulator.js');

const log = (msg) => console.log(`[smoke] ${msg}`);
const fail = (msg) => {
  console.error(`[smoke] FAIL: ${msg}`);
  process.exit(1);
};

/* 1. Mock OpenAI-compatible model: reads A1:A2, writes a SUM, then summarizes. */
let modelTurn = 0;
const model = createServer((req, res) => {
  let body = '';
  req.on('data', (c) => (body += c));
  req.on('end', () => {
    modelTurn++;
    const reply = (message) => {
      res.writeHead(200, { 'content-type': 'application/json' });
      res.end(
        JSON.stringify({
          choices: [
            {
              message,
              finish_reason: message.tool_calls ? 'tool_calls' : 'stop',
            },
          ],
          usage: { prompt_tokens: 10, completion_tokens: 5 },
        }),
      );
    };
    if (modelTurn === 1) {
      reply({
        content: null,
        tool_calls: [
          {
            id: 'c1',
            function: { name: 'read_range', arguments: '{"range":"Sheet1!A1:A2"}' },
          },
        ],
      });
    } else if (modelTurn === 2) {
      reply({
        content: null,
        tool_calls: [
          {
            id: 'c2',
            function: {
              name: 'write_range',
              arguments: JSON.stringify({
                range: 'Sheet1!A3',
                cells: [[{ formula: '=SUM(A1:A2)' }]],
                reason: 'total requested by user',
              }),
            },
          },
        ],
      });
    } else {
      reply({ content: 'I staged a SUM formula in A3 for your approval.' });
    }
  });
});
await new Promise((r) => model.listen(0, '127.0.0.1', r));
const modelPort = model.address().port;
log(`mock model listening on :${modelPort}`);

/* 2. Start the sidecar. */
const sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
const base = `http://127.0.0.1:${sidecar.port}`;
log(`sidecar listening on :${sidecar.port}`);

/* 3. Pair and configure the provider over the real HTTP API. */
const pairRes = await fetch(`${base}/api/pair`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ code: sidecar.auth.currentPairingCode }),
});
const { token } = await pairRes.json();
if (!token) fail('pairing failed');
log('paired');

const unauth = await fetch(`${base}/api/ledger/summary`);
if (unauth.status !== 401) fail('unauthenticated request was not rejected');

const put = await fetch(`${base}/api/settings/provider`, {
  method: 'PUT',
  headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
  body: JSON.stringify({
    kind: 'openai-compatible',
    model: 'mock-model',
    baseUrl: `http://127.0.0.1:${modelPort}/v1`,
  }),
});
if (put.status !== 200) fail(`provider settings rejected: ${put.status}`);
log('provider configured (openai-compatible -> mock)');

/* 4. Fake add-in: emulator answers tool requests over the WS protocol. */
const workbook = new WorkbookEmulator(['Sheet1']);
workbook.setCell('Sheet1!A1', 5);
workbook.setCell('Sheet1!A2', 7);

const ws = new WebSocket(`ws://127.0.0.1:${sidecar.port}/ws?token=${token}`);
const send = (type, payload) =>
  ws.send(JSON.stringify({ v: 1, id: crypto.randomUUID(), type, payload }));

let stagedId = null;
let sawStagedPreview = false;
let applied = false;

const finished = new Promise((resolve, reject) => {
  const timeout = setTimeout(() => reject(new Error('smoke test timed out')), 20_000);
  ws.on('open', () => {
    send('hello', { token, workbookId: 'wb-smoke', workbookName: 'Smoke.xlsx', client: 'addin' });
  });
  ws.on('message', async (raw) => {
    const msg = JSON.parse(String(raw));
    switch (msg.type) {
      case 'hello_ack':
        if (!msg.payload.ok) return reject(new Error(msg.payload.error));
        log('WS hello acknowledged');
        send('chat', { message: 'Total A1:A2 into A3 please' });
        break;
      case 'tool_exec': {
        const { requestId, tool, args } = msg.payload;
        try {
          const result =
            tool === 'get_workbook_map'
              ? await workbook.getWorkbookMap()
              : await workbook.readRange(args.range);
          send('tool_result', { requestId, ok: true, result });
        } catch (e) {
          send('tool_result', { requestId, ok: false, error: String(e) });
        }
        break;
      }
      case 'apply_ops': {
        const { requestId, ops } = msg.payload;
        try {
          await workbook.applyOps(ops);
          send('tool_result', { requestId, ok: true });
        } catch (e) {
          send('tool_result', { requestId, ok: false, error: String(e) });
        }
        break;
      }
      case 'agent_event': {
        const ev = msg.payload;
        if (ev.kind === 'changeset_staged') {
          stagedId = ev.changeSetId;
          sawStagedPreview = Array.isArray(ev.preview) && ev.preview.length === 1;
          if (workbook.getValue('Sheet1!A3') !== null) {
            return reject(new Error('change applied before approval!'));
          }
          log('change-set staged with preview; cell untouched before approval ✓');
        }
        if (ev.kind === 'changeset_applied') {
          applied = true;
          clearTimeout(timeout);
          resolve();
        }
        if (ev.kind === 'error') reject(new Error(`agent error: ${ev.message}`));
        if (ev.kind === 'done' && stagedId && !applied) {
          send('approve_changeset', { changeSetId: stagedId });
        }
        break;
      }
    }
  });
  ws.on('error', reject);
});

await finished;

const a3 = workbook.getValue('Sheet1!A3');
if (a3 !== 12) fail(`A3 should be 12 after apply, got ${a3}`);
if (!sawStagedPreview) fail('preview missing');
log('change-set approved over WS; formula recalculated to 12 ✓');

const ledger = await (
  await fetch(`${base}/api/ledger/summary`, { headers: { authorization: `Bearer ${token}` } })
).json();
if (ledger.totalCalls < 3) fail(`ledger should have recorded model calls, got ${ledger.totalCalls}`);
log(`ledger recorded ${ledger.totalCalls} model calls ✓`);

ws.close();
await sidecar.close();
model.close();
log('ALL CHECKS PASSED');
