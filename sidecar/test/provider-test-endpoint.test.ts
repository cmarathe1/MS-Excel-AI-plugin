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

async function boot(): Promise<{ base: string; token: string; headers: Record<string, string> }> {
  sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
  const base = `http://127.0.0.1:${sidecar.port}`;
  const res = await fetch(`${base}/api/pair`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ code: sidecar.auth.currentPairingCode }),
  });
  const { token } = (await res.json()) as { token: string };
  return {
    base,
    token,
    headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
  };
}

function startMockModel(
  handler: (path: string) => { status: number; body: unknown },
): Promise<number> {
  return new Promise((resolve) => {
    mockModel = createServer((req, res) => {
      const out = handler(req.url ?? '/');
      res.writeHead(out.status, { 'content-type': 'application/json' });
      res.end(JSON.stringify(out.body));
    });
    mockModel.listen(0, '127.0.0.1', () => {
      resolve((mockModel!.address() as { port: number }).port);
    });
  });
}

const okChat = {
  status: 200,
  body: {
    choices: [{ message: { content: 'OK' }, finish_reason: 'stop' }],
    usage: { prompt_tokens: 3, completion_tokens: 1 },
  },
};

describe('auth check', () => {
  it('answers 200 with a valid token and 401 without', async () => {
    const { base, headers } = await boot();
    expect((await fetch(`${base}/api/auth/check`, { headers })).status).toBe(200);
    expect((await fetch(`${base}/api/auth/check`)).status).toBe(401);
    expect(
      (await fetch(`${base}/api/auth/check`, { headers: { authorization: 'Bearer stale' } })).status,
    ).toBe(401);
  });
});

describe('candidate provider test endpoint', () => {
  it('reports success when the endpoint answers', async () => {
    const port = await startMockModel(() => okChat);
    const { base, headers } = await boot();
    const res = await fetch(`${base}/api/settings/providers/test`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        kind: 'openai-compatible',
        model: 'llama3',
        baseUrl: `http://127.0.0.1:${port}/v1`,
      }),
    });
    const body = (await res.json()) as { ok: boolean; reply?: string };
    expect(body.ok).toBe(true);
    expect(body.reply).toBe('OK');

    const ledger = (await (
      await fetch(`${base}/api/ledger/summary`, { headers })
    ).json()) as { byFeature: { feature: string }[] };
    expect(ledger.byFeature.some((f) => f.feature === 'probe')).toBe(true);
  });

  it('reports a clean failure for erroring endpoints', async () => {
    const port = await startMockModel(() => ({ status: 401, body: { error: 'bad key' } }));
    const { base, headers } = await boot();
    const res = await fetch(`${base}/api/settings/providers/test`, {
      method: 'POST',
      headers,
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

describe('saved model configurations', () => {
  it('add, auto-activate first, switch, test by id, delete with active fallback', async () => {
    const port = await startMockModel(() => okChat);
    const { base, headers } = await boot();

    const add = async (model: string): Promise<string> => {
      const res = await fetch(`${base}/api/settings/providers`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          kind: 'openai-compatible',
          model,
          baseUrl: `http://127.0.0.1:${port}/v1`,
          apiKey: `key-for-${model}-1234`,
        }),
      });
      return ((await res.json()) as { id: string }).id;
    };

    const idA = await add('model-a');
    const idB = await add('model-b');

    // first added is active; list redacts keys
    let list = (await (
      await fetch(`${base}/api/settings/providers`, { headers })
    ).json()) as { providers: { id: string; apiKey?: string }[]; activeProviderId: string };
    expect(list.activeProviderId).toBe(idA);
    expect(list.providers).toHaveLength(2);
    expect(list.providers[0]!.apiKey).toContain('••••');
    expect(list.providers[0]!.apiKey).not.toContain('key-for');

    // switch active
    await fetch(`${base}/api/settings/providers/${idB}/activate`, { method: 'POST', headers });
    list = (await (await fetch(`${base}/api/settings/providers`, { headers })).json()) as typeof list;
    expect(list.activeProviderId).toBe(idB);

    // test a stored configuration (uses the stored, unredacted key)
    const test = (await (
      await fetch(`${base}/api/settings/providers/${idA}/test`, { method: 'POST', headers })
    ).json()) as { ok: boolean };
    expect(test.ok).toBe(true);

    // deleting the active config falls back to the remaining one
    await fetch(`${base}/api/settings/providers/${idB}`, { method: 'DELETE', headers });
    list = (await (await fetch(`${base}/api/settings/providers`, { headers })).json()) as typeof list;
    expect(list.providers).toHaveLength(1);
    expect(list.activeProviderId).toBe(idA);

    // unknown ids 404
    expect(
      (await fetch(`${base}/api/settings/providers/nope`, { method: 'DELETE', headers })).status,
    ).toBe(404);
  });

  it('rejects invalid configurations', async () => {
    const { base, headers } = await boot();
    const res = await fetch(`${base}/api/settings/providers`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ kind: 'nonsense', model: '' }),
    });
    expect(res.status).toBe(400);
  });
});

describe('model discovery', () => {
  it('lists models from an OpenAI-compatible endpoint', async () => {
    const port = await startMockModel((path) =>
      path.endsWith('/models')
        ? { status: 200, body: { data: [{ id: 'llama3.1' }, { id: 'qwen3' }] } }
        : { status: 404, body: {} },
    );
    const { base, headers } = await boot();
    const res = await fetch(`${base}/api/providers/models`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ kind: 'openai-compatible', baseUrl: `http://127.0.0.1:${port}/v1` }),
    });
    const body = (await res.json()) as { models: string[] };
    expect(body.models).toEqual(['llama3.1', 'qwen3']);
  });

  it('fails gracefully when the endpoint is down', async () => {
    const { base, headers } = await boot();
    const res = await fetch(`${base}/api/providers/models`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ kind: 'openai-compatible', baseUrl: 'http://127.0.0.1:9/v1' }),
    });
    const body = (await res.json()) as { models: string[]; error?: string };
    expect(body.models).toEqual([]);
    expect(body.error).toBeTruthy();
  });
});
