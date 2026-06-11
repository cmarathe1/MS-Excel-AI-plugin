import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import { createServer as createHttpsServer } from 'node:https';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { extname, join, normalize } from 'node:path';
import { WebSocketServer, type WebSocket } from 'ws';
import { PROTOCOL_VERSION, type ClientMessage } from '@excelai/shared';
import { z } from 'zod';
import { Auth } from './auth.js';
import {
  activeProvider,
  loadSettings,
  redactProvider,
  saveSettings,
  type StoredProvider,
} from './config.js';
import type { Db } from './db.js';
import { Ledger } from './ledger.js';
import { MemoryStore } from './memory/store.js';
import { FunctionEngine } from './functions/batch.js';
import { createProvider, ProviderError, type Provider } from './providers/index.js';
import { BridgeExecutor } from './bridge.js';
import { Session } from './session.js';

const SIDECAR_VERSION = '0.1.0';

/** Origins allowed to talk to the sidecar (Office webviews + local dev). */
const ALLOWED_ORIGIN_RE =
  /^(https:\/\/[a-z0-9-]+\.officeapps\.live\.com|https:\/\/[a-z0-9-]+\.office\.com|https:\/\/[a-z0-9.-]+\.sharepoint\.com|https?:\/\/localhost(:\d+)?|https?:\/\/127\.0\.0\.1(:\d+)?)$/i;

export interface SidecarServer {
  server: Server;
  auth: Auth;
  port: number;
  close(): Promise<void>;
}

const providerConfigSchema = z.object({
  kind: z.enum(['anthropic', 'openai', 'openai-compatible']),
  model: z.string().min(1),
  apiKey: z.string().min(1).optional(),
  baseUrl: z.string().url().optional(),
});

export async function startServer(options: {
  db: Db;
  port?: number;
  staticDir?: string;
}): Promise<SidecarServer> {
  const { db } = options;
  const auth = new Auth(db);
  const ledger = new Ledger(db);
  const memory = new MemoryStore(db);
  const functions = new FunctionEngine(db);

  let cachedProvider: { provider: Provider; fingerprint: string } | null = null;
  const getProvider = (): Provider => {
    const active = activeProvider(loadSettings());
    if (!active) {
      throw new ProviderError('No model configured. Open Settings in the task pane and add one.');
    }
    const fingerprint = JSON.stringify(active);
    if (!cachedProvider || cachedProvider.fingerprint !== fingerprint) {
      cachedProvider = { provider: createProvider(active), fingerprint };
    }
    return cachedProvider.provider;
  };

  // Optional TLS (e.g. office-addin-dev-certs for hosts that require
  // https://localhost): set EXCELAI_TLS_CERT and EXCELAI_TLS_KEY.
  const tlsCert = process.env.EXCELAI_TLS_CERT;
  const tlsKey = process.env.EXCELAI_TLS_KEY;
  const handler = (req: IncomingMessage, res: ServerResponse): void => {
    void handleHttp(req, res).catch((e) => {
      sendJson(res, 500, { error: e instanceof Error ? e.message : 'Internal error' });
    });
  };
  const server: Server =
    tlsCert && tlsKey
      ? createHttpsServer({ cert: readFileSync(tlsCert), key: readFileSync(tlsKey) }, handler)
      : createServer(handler);

  async function handleHttp(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    const origin = req.headers.origin;

    // CORS for the Office webview / browser clients.
    if (origin && ALLOWED_ORIGIN_RE.test(origin)) {
      res.setHeader('access-control-allow-origin', origin);
      res.setHeader('access-control-allow-headers', 'authorization, content-type');
      res.setHeader('access-control-allow-methods', 'GET, POST, PUT, OPTIONS');
    }
    if (req.method === 'OPTIONS') {
      res.writeHead(204).end();
      return;
    }

    if (url.pathname === '/api/health') {
      sendJson(res, 200, { ok: true, version: SIDECAR_VERSION, protocol: PROTOCOL_VERSION });
      return;
    }

    if (url.pathname === '/api/pair' && req.method === 'POST') {
      const body = await readJson(req);
      const code = typeof (body as { code?: unknown })?.code === 'string' ? (body as { code: string }).code : '';
      const token = auth.pair(code);
      if (!token) {
        sendJson(res, 403, { error: 'Invalid or already-used pairing code' });
        return;
      }
      sendJson(res, 200, { token });
      return;
    }

    if (url.pathname.startsWith('/api/')) {
      const token = bearerToken(req);
      if (!auth.verify(token)) {
        sendJson(res, 401, { error: 'Unauthorized' });
        return;
      }

      if (url.pathname === '/api/auth/check' && req.method === 'GET') {
        sendJson(res, 200, { ok: true });
        return;
      }

      /* ----- saved model configurations ----- */

      if (url.pathname === '/api/settings/providers' && req.method === 'GET') {
        const settings = loadSettings();
        sendJson(res, 200, {
          providers: settings.providers.map(redactProvider),
          activeProviderId: settings.activeProviderId,
        });
        return;
      }
      if (url.pathname === '/api/settings/providers' && req.method === 'POST') {
        const parsed = providerConfigSchema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        const settings = loadSettings();
        const stored: StoredProvider = { ...parsed.data, id: crypto.randomUUID() };
        settings.providers.push(stored);
        // First saved model becomes active automatically.
        if (!settings.activeProviderId) settings.activeProviderId = stored.id;
        saveSettings(settings);
        cachedProvider = null;
        sendJson(res, 200, { id: stored.id, activeProviderId: settings.activeProviderId });
        return;
      }

      const providerRoute = /^\/api\/settings\/providers\/([\w-]+)(?:\/(activate|test))?$/.exec(
        url.pathname,
      );
      if (providerRoute) {
        const [, id, action] = providerRoute;
        const settings = loadSettings();
        const stored = settings.providers.find((p) => p.id === id);

        if (!action && req.method === 'DELETE') {
          if (!stored) {
            sendJson(res, 404, { error: 'No such model configuration' });
            return;
          }
          settings.providers = settings.providers.filter((p) => p.id !== id);
          if (settings.activeProviderId === id) {
            settings.activeProviderId = settings.providers[0]?.id ?? null;
          }
          saveSettings(settings);
          cachedProvider = null;
          sendJson(res, 200, { ok: true, activeProviderId: settings.activeProviderId });
          return;
        }
        if (action === 'activate' && req.method === 'POST') {
          if (!stored) {
            sendJson(res, 404, { error: 'No such model configuration' });
            return;
          }
          settings.activeProviderId = stored.id;
          saveSettings(settings);
          cachedProvider = null;
          sendJson(res, 200, { ok: true, activeProviderId: stored.id });
          return;
        }
        if (action === 'test' && req.method === 'POST') {
          if (!stored) {
            sendJson(res, 404, { ok: false, error: 'No such model configuration' });
            return;
          }
          sendJson(res, 200, await testProviderConfig(stored, ledger));
          return;
        }
      }

      /** Test a candidate configuration before saving it. */
      if (url.pathname === '/api/settings/providers/test' && req.method === 'POST') {
        const parsed = providerConfigSchema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { ok: false, error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        sendJson(res, 200, await testProviderConfig(parsed.data, ledger));
        return;
      }

      /** Best-effort model discovery for the add/edit form (and an Ollama
       *  readiness check: lists the models actually installed). */
      if (url.pathname === '/api/providers/models' && req.method === 'POST') {
        const schema = z.object({
          kind: z.enum(['anthropic', 'openai', 'openai-compatible']),
          apiKey: z.string().optional(),
          baseUrl: z.string().url().optional(),
        });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { models: [], error: 'Invalid request' });
          return;
        }
        sendJson(res, 200, await listModels(parsed.data));
        return;
      }
      if (url.pathname === '/api/functions/run' && req.method === 'POST') {
        const schema = z.object({
          kind: z.enum(['ai', 'classify', 'extract', 'translate']),
          prompt: z.string().default(''),
          input: z.string().optional(),
          categories: z.array(z.string()).optional(),
          targetLang: z.string().optional(),
        });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { error: 'Invalid function request' });
          return;
        }
        try {
          const provider = getProvider();
          const value = await functions.run(provider, parsed.data, (usage) =>
            ledger.record({ provider: provider.id, model: provider.model, feature: 'functions', ...usage }),
          );
          sendJson(res, 200, { value });
        } catch (e) {
          sendJson(res, 502, { error: e instanceof Error ? e.message : 'Function failed' });
        }
        return;
      }
      if (url.pathname === '/api/ledger/summary' && req.method === 'GET') {
        sendJson(res, 200, ledger.summary());
        return;
      }
      if (url.pathname === '/api/memory' && req.method === 'GET') {
        const scope = url.searchParams.get('scope') ?? undefined;
        sendJson(res, 200, { entries: memory.list(scope) });
        return;
      }
      if (url.pathname === '/api/memory' && req.method === 'POST') {
        const schema = z.object({ scope: z.string().min(1), content: z.string().min(1) });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { error: 'scope and content are required' });
          return;
        }
        sendJson(res, 200, memory.add(parsed.data.scope, parsed.data.content));
        return;
      }
      sendJson(res, 404, { error: 'Not found' });
      return;
    }

    // Static add-in assets (same-origin serving avoids mixed-content issues).
    if (options.staticDir) {
      serveStatic(options.staticDir, url.pathname, res);
      return;
    }
    sendJson(res, 404, { error: 'Not found' });
  }

  /* ---------------- WebSocket ---------------- */

  const wss = new WebSocketServer({ noServer: true });

  server.on('upgrade', (req, socket, head) => {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    const origin = req.headers.origin;
    if (url.pathname !== '/ws' || (origin && !ALLOWED_ORIGIN_RE.test(origin))) {
      socket.destroy();
      return;
    }
    if (!auth.verify(url.searchParams.get('token'))) {
      socket.destroy();
      return;
    }
    wss.handleUpgrade(req, socket, head, (ws) => wss.emit('connection', ws, req));
  });

  wss.on('connection', (ws: WebSocket) => {
    let session: Session | null = null;
    let bridge: BridgeExecutor | null = null;

    ws.on('message', (raw) => {
      void (async () => {
        let msg: ClientMessage;
        try {
          msg = JSON.parse(String(raw)) as ClientMessage;
        } catch {
          return;
        }
        if (typeof msg.v !== 'number' || Math.floor(msg.v) !== PROTOCOL_VERSION) {
          ws.send(
            JSON.stringify({
              v: PROTOCOL_VERSION,
              id: msg.id ?? 'unknown',
              type: 'hello_ack',
              payload: {
                ok: false,
                error: `Protocol version mismatch (sidecar ${PROTOCOL_VERSION}, client ${String(msg.v)}). Update both sides.`,
                sidecarVersion: SIDECAR_VERSION,
              },
            }),
          );
          ws.close();
          return;
        }
        switch (msg.type) {
          case 'hello': {
            if (!auth.verify(msg.payload.token)) {
              ws.close();
              return;
            }
            bridge = new BridgeExecutor(ws);
            session = new Session(
              ws,
              msg.payload.workbookId,
              msg.payload.workbookName,
              bridge,
              getProvider,
              ledger,
              memory,
            );
            ws.send(
              JSON.stringify({
                v: PROTOCOL_VERSION,
                id: msg.id,
                type: 'hello_ack',
                payload: { ok: true, sidecarVersion: SIDECAR_VERSION },
              }),
            );
            break;
          }
          case 'tool_result':
            bridge?.handleToolResult(msg.payload);
            break;
          case 'chat':
            await session?.handleChat(msg.payload.message);
            break;
          case 'approve_changeset':
            await session?.handleApprove(msg.payload.changeSetId);
            break;
          case 'reject_changeset':
            session?.handleReject(msg.payload.changeSetId);
            break;
          case 'undo_changeset':
            await session?.handleUndo(msg.payload.changeSetId);
            break;
        }
      })();
    });

    ws.on('close', () => bridge?.abortAll('Add-in disconnected'));
  });

  const port = options.port ?? 8923;
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, '127.0.0.1', resolve);
  });
  const actualPort = (server.address() as { port: number }).port;

  return {
    server,
    auth,
    port: actualPort,
    close: () =>
      new Promise<void>((resolve, reject) => {
        wss.close();
        server.close((e) => (e ? reject(e) : resolve()));
      }),
  };
}

/* ---------------- helpers ---------------- */

async function testProviderConfig(
  config: Parameters<typeof createProvider>[0],
  ledger: Ledger,
): Promise<{ ok: boolean; reply?: string; error?: string }> {
  try {
    const provider = createProvider(config);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 20_000);
    const response = await provider
      .chat({
        messages: [{ role: 'user', content: 'Reply with exactly: OK' }],
        maxTokens: 16,
        temperature: 0,
        abortSignal: controller.signal,
      })
      .finally(() => clearTimeout(timer));
    ledger.record({
      provider: provider.id,
      model: provider.model,
      feature: 'probe',
      ...response.usage,
    });
    return { ok: true, reply: response.text.trim().slice(0, 80) };
  } catch (e) {
    const message =
      e instanceof Error && e.name === 'AbortError'
        ? 'Timed out after 20s — is the endpoint reachable?'
        : e instanceof Error
          ? e.message
          : String(e);
    return { ok: false, error: message };
  }
}

async function listModels(config: {
  kind: 'anthropic' | 'openai' | 'openai-compatible';
  apiKey?: string;
  baseUrl?: string;
}): Promise<{ models: string[]; error?: string }> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8_000);
  try {
    let endpoint: string;
    const headers: Record<string, string> = {};
    if (config.kind === 'anthropic') {
      endpoint = `${config.baseUrl ?? 'https://api.anthropic.com'}/v1/models`;
      if (config.apiKey) headers['x-api-key'] = config.apiKey;
      headers['anthropic-version'] = '2023-06-01';
    } else {
      const base = config.baseUrl ?? 'https://api.openai.com/v1';
      endpoint = `${base.replace(/\/$/, '')}/models`;
      if (config.apiKey) headers.authorization = `Bearer ${config.apiKey}`;
    }
    const res = await fetch(endpoint, { headers, signal: controller.signal });
    if (!res.ok) return { models: [], error: `Endpoint answered ${res.status}` };
    const data = (await res.json()) as { data?: { id?: string }[]; models?: { id?: string }[] };
    const items = data.data ?? data.models ?? [];
    const models = items
      .map((m) => m.id)
      .filter((id): id is string => typeof id === 'string')
      .sort();
    return { models };
  } catch (e) {
    return {
      models: [],
      error:
        e instanceof Error && e.name === 'AbortError'
          ? 'Timed out — endpoint unreachable'
          : e instanceof Error
            ? e.message
            : String(e),
    };
  } finally {
    clearTimeout(timer);
  }
}

function bearerToken(req: IncomingMessage): string | null {
  const header = req.headers.authorization;
  if (header?.startsWith('Bearer ')) return header.slice(7);
  return null;
}

function sendJson(res: ServerResponse, status: number, body: unknown): void {
  const json = JSON.stringify(body);
  res.writeHead(status, { 'content-type': 'application/json' });
  res.end(json);
}

async function readJson(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > 5_000_000) throw new Error('Request body too large');
    chunks.push(chunk as Buffer);
  }
  const text = Buffer.concat(chunks).toString('utf8');
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return {};
  }
}

const MIME: Record<string, string> = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

function serveStatic(dir: string, pathname: string, res: ServerResponse): void {
  const rel = pathname === '/' ? '/index.html' : pathname;
  const file = normalize(join(dir, rel));
  if (!file.startsWith(normalize(dir)) || !existsSync(file) || !statSync(file).isFile()) {
    sendJson(res, 404, { error: 'Not found' });
    return;
  }
  res.writeHead(200, { 'content-type': MIME[extname(file)] ?? 'application/octet-stream' });
  res.end(readFileSync(file));
}
