import {
  makeEnvelope,
  PROTOCOL_VERSION,
  type ChangeOp,
  type ServerMessage,
} from '@excelai/shared';
import { applyOps, getWorkbookMap, readRange } from './officeBridge.js';

export const SIDECAR_BASE = `http://127.0.0.1:8923`;

const TOKEN_KEY = 'excelai-token';

export function getToken(): string | null {
  return localStorage.getItem(TOKEN_KEY);
}

export async function pair(code: string): Promise<void> {
  const res = await fetch(`${SIDECAR_BASE}/api/pair`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ code }),
  });
  if (!res.ok) {
    const body = (await res.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error ?? `Pairing failed (${res.status})`);
  }
  const { token } = (await res.json()) as { token: string };
  localStorage.setItem(TOKEN_KEY, token);
}

export async function api<T>(path: string, init?: RequestInit): Promise<T> {
  const token = getToken();
  const res = await fetch(`${SIDECAR_BASE}${path}`, {
    ...init,
    headers: {
      'content-type': 'application/json',
      ...(token ? { authorization: `Bearer ${token}` } : {}),
      ...(init?.headers ?? {}),
    },
  });
  const body = (await res.json().catch(() => ({}))) as T & { error?: string };
  if (!res.ok) throw new Error(body.error ?? `Request failed (${res.status})`);
  return body;
}

export type AgentEventPayload = Extract<ServerMessage, { type: 'agent_event' }>['payload'];

export interface SidecarConnection {
  sendChat(message: string): void;
  approve(changeSetId: string): void;
  reject(changeSetId: string): void;
  undo(changeSetId: string): void;
  close(): void;
}

/**
 * Connects to the sidecar, identifies the workbook, then serves two duties:
 * answering tool-execution requests against Excel and forwarding agent
 * events to the UI.
 */
export function connect(options: {
  workbookId: string;
  workbookName: string;
  onEvent: (ev: AgentEventPayload) => void;
  onStatus: (status: 'connected' | 'disconnected' | 'error', detail?: string) => void;
}): SidecarConnection {
  const token = getToken();
  if (!token) throw new Error('Not paired');

  const ws = new WebSocket(`${SIDECAR_BASE.replace('http', 'ws')}/ws?token=${token}`);

  ws.onopen = () => {
    ws.send(
      JSON.stringify(
        makeEnvelope('hello', {
          token,
          workbookId: options.workbookId,
          workbookName: options.workbookName,
          client: 'addin' as const,
        }),
      ),
    );
  };

  ws.onclose = () => options.onStatus('disconnected');
  ws.onerror = () => options.onStatus('error', 'WebSocket error — is the sidecar running?');

  ws.onmessage = (event) => {
    void (async () => {
      let msg: ServerMessage;
      try {
        msg = JSON.parse(String(event.data)) as ServerMessage;
      } catch {
        return;
      }
      if (msg.v !== PROTOCOL_VERSION) return;

      switch (msg.type) {
        case 'hello_ack':
          if (msg.payload.ok) options.onStatus('connected');
          else options.onStatus('error', msg.payload.error);
          break;
        case 'agent_event':
          options.onEvent(msg.payload);
          break;
        case 'tool_exec': {
          const { requestId, tool, args } = msg.payload;
          try {
            let result: unknown;
            if (tool === 'get_workbook_map') result = await getWorkbookMap();
            else if (tool === 'read_range') result = await readRange((args as { range: string }).range);
            else throw new Error(`Add-in cannot execute tool: ${tool}`);
            ws.send(JSON.stringify(makeEnvelope('tool_result', { requestId, ok: true, result })));
          } catch (e) {
            ws.send(
              JSON.stringify(
                makeEnvelope('tool_result', {
                  requestId,
                  ok: false,
                  error: e instanceof Error ? e.message : String(e),
                }),
              ),
            );
          }
          break;
        }
        case 'apply_ops': {
          const { requestId, ops } = msg.payload;
          try {
            await applyOps(ops as ChangeOp[]);
            ws.send(JSON.stringify(makeEnvelope('tool_result', { requestId, ok: true })));
          } catch (e) {
            ws.send(
              JSON.stringify(
                makeEnvelope('tool_result', {
                  requestId,
                  ok: false,
                  error: e instanceof Error ? e.message : String(e),
                }),
              ),
            );
          }
          break;
        }
      }
    })();
  };

  const send = (type: string, payload: unknown): void => {
    if (ws.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify(makeEnvelope(type, payload)));
    }
  };

  return {
    sendChat: (message) => send('chat', { message }),
    approve: (changeSetId) => send('approve_changeset', { changeSetId }),
    reject: (changeSetId) => send('reject_changeset', { changeSetId }),
    undo: (changeSetId) => send('undo_changeset', { changeSetId }),
    close: () => ws.close(),
  };
}
