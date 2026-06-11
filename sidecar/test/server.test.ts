import { afterEach, describe, expect, it } from 'vitest';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

let sidecar: SidecarServer | null = null;

afterEach(async () => {
  await sidecar?.close();
  sidecar = null;
});

async function boot(): Promise<{ base: string; token: string }> {
  sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
  const base = `http://127.0.0.1:${sidecar.port}`;
  const code = sidecar.auth.currentPairingCode!;
  const res = await fetch(`${base}/api/pair`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ code }),
  });
  const { token } = (await res.json()) as { token: string };
  return { base, token };
}

describe('sidecar HTTP API', () => {
  it('serves health without auth', async () => {
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const res = await fetch(`http://127.0.0.1:${sidecar.port}/api/health`);
    expect(res.status).toBe(200);
    const body = (await res.json()) as { ok: boolean; protocol: number };
    expect(body.ok).toBe(true);
    expect(body.protocol).toBe(1);
  });

  it('rejects API calls without a valid token', async () => {
    const { base } = await boot();
    const res = await fetch(`${base}/api/ledger/summary`);
    expect(res.status).toBe(401);
    const forged = await fetch(`${base}/api/ledger/summary`, {
      headers: { authorization: 'Bearer forged' },
    });
    expect(forged.status).toBe(401);
  });

  it('pairs once only', async () => {
    const { base } = await boot();

    // pairing code is single-use
    const again = await fetch(`${base}/api/pair`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ code: 'ABCDEF' }),
    });
    expect(again.status).toBe(403);
  });

  it('manages memory entries over HTTP', async () => {
    const { base, token } = await boot();
    const auth = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };
    const add = await fetch(`${base}/api/memory`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ scope: 'user', content: 'prefers ISO dates' }),
    });
    expect(add.status).toBe(200);
    const list = await fetch(`${base}/api/memory?scope=user`, { headers: auth });
    const body = (await list.json()) as { entries: { content: string }[] };
    expect(body.entries[0]?.content).toBe('prefers ISO dates');
  });
});
