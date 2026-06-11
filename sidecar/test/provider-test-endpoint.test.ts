import { afterEach, describe, expect, it } from 'vitest';
import { createServer, type Server } from 'node:http';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

let sidecar: SidecarServer | null = null;
let mockModel: Server | null = null;

afterEach(async () => {
  await sidecar?.close();
  sidecar = null;
  await new Promise<void>((r) => (mockModel ? mockModel.close(() => r()) : r()));
  mockModel = null;
});

async function boot(): Promise<{ base: string; token: string }> {
  sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
  const base = `http://127.0.0.1:${sidecar.port}`;
  const res = await fetch(`${base}/api/pair`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ code: sidecar.auth.currentPairingCode }),
  });
  const { token } = (await res.json()) as { token: string };
  return { base, token };
}

function startMockModel(reply: { status: number; body: unknown }): Promise<number> {
  return new Promise((resolve) => {
    mockModel = createServer((_req, res) => {
      res.writeHead(reply.status, { 'content-type': 'application/json' });
      res.end(JSON.stringify(reply.body));
    });
    mockModel.listen(0, '127.0.0.1', () => {
      resolve((mockModel!.address() as { port: number }).port);
    });
  });
}

describe('provider test endpoint', () => {
  it('reports success when the endpoint answers', async () => {
    const port = await startMockModel({
      status: 200,
      body: {
        choices: [{ message: { content: 'OK' }, finish_reason: 'stop' }],
        usage: { prompt_tokens: 3, completion_tokens: 1 },
      },
    });
    const { base, token } = await boot();
    const res = await fetch(`${base}/api/settings/provider/test`, {
      method: 'POST',
      headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
      body: JSON.stringify({
        kind: 'openai-compatible',
        model: 'llama3',
        baseUrl: `http://127.0.0.1:${port}/v1`,
      }),
    });
    const body = (await res.json()) as { ok: boolean; reply?: string };
    expect(body.ok).toBe(true);
    expect(body.reply).toBe('OK');

    // probe call landed in the ledger
    const ledger = (await (
      await fetch(`${base}/api/ledger/summary`, { headers: { authorization: `Bearer ${token}` } })
    ).json()) as { byFeature: { feature: string }[] };
    expect(ledger.byFeature.some((f) => f.feature === 'probe')).toBe(true);
  });

  it('reports a clean failure for unreachable/erroring endpoints', async () => {
    const port = await startMockModel({ status: 401, body: { error: 'bad key' } });
    const { base, token } = await boot();
    const res = await fetch(`${base}/api/settings/provider/test`, {
      method: 'POST',
      headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
      body: JSON.stringify({
        kind: 'openai-compatible',
        model: 'llama3',
        baseUrl: `http://127.0.0.1:${port}/v1`,
      }),
    });
    const body = (await res.json()) as { ok: boolean; error?: string };
    expect(body.ok).toBe(false);
    expect(body.error).toContain('401');
  });
});

describe('settings key preservation', () => {
  it('keeps the stored API key when an update of the same kind omits it', async () => {
    const { base, token } = await boot();
    const headers = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };

    await fetch(`${base}/api/settings/provider`, {
      method: 'PUT',
      headers,
      body: JSON.stringify({ kind: 'anthropic', model: 'claude-a', apiKey: 'sk-original-9876' }),
    });
    // user changes only the model; apiKey field left blank in the UI
    await fetch(`${base}/api/settings/provider`, {
      method: 'PUT',
      headers,
      body: JSON.stringify({ kind: 'anthropic', model: 'claude-b' }),
    });

    const body = (await (
      await fetch(`${base}/api/settings/provider`, { headers })
    ).json()) as { provider: { model: string; apiKey?: string } };
    expect(body.provider.model).toBe('claude-b');
    expect(body.provider.apiKey).toContain('9876'); // key survived (redacted)
  });

  it('drops the key when the provider kind changes', async () => {
    const { base, token } = await boot();
    const headers = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };
    await fetch(`${base}/api/settings/provider`, {
      method: 'PUT',
      headers,
      body: JSON.stringify({ kind: 'anthropic', model: 'claude-a', apiKey: 'sk-original-9876' }),
    });
    await fetch(`${base}/api/settings/provider`, {
      method: 'PUT',
      headers,
      body: JSON.stringify({ kind: 'openai-compatible', model: 'llama3', baseUrl: 'http://localhost:11434/v1' }),
    });
    const body = (await (
      await fetch(`${base}/api/settings/provider`, { headers })
    ).json()) as { provider: { apiKey?: string } };
    expect(body.provider.apiKey).toBeUndefined();
  });
});
