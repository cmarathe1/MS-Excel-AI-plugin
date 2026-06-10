import { useCallback, useEffect, useRef, useState } from 'react';
import type { ChangePreviewItem } from '@excelai/shared';
import {
  api,
  connect,
  getToken,
  pair,
  type AgentEventPayload,
  type SidecarConnection,
} from '../connection.js';
import { getWorkbookIdentity } from '../officeBridge.js';

type ChatItem =
  | { kind: 'user'; text: string }
  | { kind: 'assistant'; text: string }
  | { kind: 'tool'; text: string }
  | { kind: 'error'; text: string }
  | { kind: 'changeset'; id: string; preview: ChangePreviewItem[]; status: 'staged' | 'applied' | 'rejected' | 'undone' };

interface ProviderForm {
  kind: 'anthropic' | 'openai' | 'openai-compatible';
  model: string;
  apiKey: string;
  baseUrl: string;
}

export function App(): JSX.Element {
  const [paired, setPaired] = useState<boolean>(() => getToken() !== null);
  const [status, setStatus] = useState<'connecting' | 'connected' | 'disconnected' | 'error'>('connecting');
  const [statusDetail, setStatusDetail] = useState('');
  const [items, setItems] = useState<ChatItem[]>([]);
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const connection = useRef<SidecarConnection | null>(null);

  const handleEvent = useCallback((ev: AgentEventPayload) => {
    switch (ev.kind) {
      case 'text':
        setItems((prev) => [...prev, { kind: 'assistant', text: ev.text }]);
        break;
      case 'tool_use':
        setItems((prev) => [...prev, { kind: 'tool', text: `${ev.tool} ${ev.summary}`.trim() }]);
        break;
      case 'changeset_staged':
        setItems((prev) => [
          ...prev,
          { kind: 'changeset', id: ev.changeSetId, preview: ev.preview, status: 'staged' },
        ]);
        break;
      case 'changeset_applied':
        setItems((prev) =>
          prev.map((it) =>
            it.kind === 'changeset' && it.id === ev.changeSetId ? { ...it, status: 'applied' } : it,
          ),
        );
        break;
      case 'changeset_undone':
        setItems((prev) =>
          prev.map((it) =>
            it.kind === 'changeset' && it.id === ev.changeSetId ? { ...it, status: 'undone' } : it,
          ),
        );
        break;
      case 'error':
        setItems((prev) => [...prev, { kind: 'error', text: ev.message }]);
        break;
      case 'done':
        setBusy(false);
        break;
    }
  }, []);

  useEffect(() => {
    if (!paired) return;
    let conn: SidecarConnection | null = null;
    void (async () => {
      try {
        const { id, name } = await getWorkbookIdentity();
        conn = connect({
          workbookId: id,
          workbookName: name,
          onEvent: handleEvent,
          onStatus: (s, detail) => {
            setStatus(s === 'connected' ? 'connected' : s);
            setStatusDetail(detail ?? '');
          },
        });
        connection.current = conn;
      } catch (e) {
        setStatus('error');
        setStatusDetail(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => conn?.close();
  }, [paired, handleEvent]);

  const sendMessage = (): void => {
    const text = input.trim();
    if (!text || busy || status !== 'connected') return;
    setItems((prev) => [...prev, { kind: 'user', text }]);
    setInput('');
    setBusy(true);
    connection.current?.sendChat(text);
  };

  if (!paired) {
    return <PairingScreen onPaired={() => setPaired(true)} />;
  }

  return (
    <div className="app">
      <header>
        <span className={`dot ${status}`} title={statusDetail} />
        <strong>Excel AI</strong>
        <button className="ghost" onClick={() => setShowSettings((s) => !s)}>
          {showSettings ? 'Chat' : 'Settings'}
        </button>
      </header>

      {showSettings ? (
        <SettingsScreen />
      ) : (
        <>
          <main className="chat">
            {items.length === 0 && (
              <p className="hint">
                Ask anything about this workbook. Edits are staged for your approval — nothing
                changes without your OK.
              </p>
            )}
            {items.map((item, i) => (
              <ChatBubble
                key={i}
                item={item}
                onApprove={(id) => connection.current?.approve(id)}
                onReject={(id) => {
                  connection.current?.reject(id);
                  setItems((prev) =>
                    prev.map((it) =>
                      it.kind === 'changeset' && it.id === id ? { ...it, status: 'rejected' } : it,
                    ),
                  );
                }}
                onUndo={(id) => connection.current?.undo(id)}
              />
            ))}
            {busy && <p className="hint">Working…</p>}
          </main>
          <footer>
            <textarea
              value={input}
              placeholder="e.g. Sum column B into B20 and explain what this sheet does"
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  sendMessage();
                }
              }}
            />
            <button onClick={sendMessage} disabled={busy || status !== 'connected'}>
              Send
            </button>
          </footer>
        </>
      )}
    </div>
  );
}

function ChatBubble(props: {
  item: ChatItem;
  onApprove: (id: string) => void;
  onReject: (id: string) => void;
  onUndo: (id: string) => void;
}): JSX.Element {
  const { item } = props;
  switch (item.kind) {
    case 'user':
      return <div className="bubble user">{item.text}</div>;
    case 'assistant':
      return <div className="bubble assistant">{item.text}</div>;
    case 'tool':
      return <div className="toolnote">🔧 {item.text}</div>;
    case 'error':
      return <div className="bubble error">{item.text}</div>;
    case 'changeset':
      return (
        <div className="changeset">
          <div className="changeset-head">
            Proposed changes <span className={`badge ${item.status}`}>{item.status}</span>
          </div>
          <ul>
            {item.preview.map((p, i) => (
              <li key={i}>
                <OpSummary item={p} />
              </li>
            ))}
          </ul>
          {item.status === 'staged' && (
            <div className="actions">
              <button onClick={() => props.onApprove(item.id)}>Apply</button>
              <button className="ghost" onClick={() => props.onReject(item.id)}>
                Reject
              </button>
            </div>
          )}
          {item.status === 'applied' && (
            <div className="actions">
              <button className="ghost" onClick={() => props.onUndo(item.id)}>
                Undo
              </button>
            </div>
          )}
        </div>
      );
  }
}

function OpSummary({ item }: { item: ChangePreviewItem }): JSX.Element {
  const { op } = item;
  switch (op.kind) {
    case 'write_range': {
      const cellCount = op.cells.length * (op.cells[0]?.length ?? 0);
      return (
        <span>
          <code>{op.range}</code> — write {cellCount} cell{cellCount === 1 ? '' : 's'}
          <em> · {op.reason}</em>
        </span>
      );
    }
    case 'clear_range':
      return (
        <span>
          <code>{op.range}</code> — clear
          <em> · {op.reason}</em>
        </span>
      );
    case 'add_sheet':
      return (
        <span>
          new sheet <code>{op.name}</code>
          <em> · {op.reason}</em>
        </span>
      );
  }
}

function PairingScreen({ onPaired }: { onPaired: () => void }): JSX.Element {
  const [code, setCode] = useState('');
  const [error, setError] = useState('');
  return (
    <div className="app center">
      <h2>Pair with the companion app</h2>
      <p className="hint">
        Start the sidecar (<code>pnpm dev:sidecar</code>) and enter the pairing code it prints.
      </p>
      <input
        value={code}
        maxLength={6}
        placeholder="ABC123"
        onChange={(e) => setCode(e.target.value.toUpperCase())}
      />
      <button
        onClick={() => {
          void pair(code).then(onPaired, (e: unknown) =>
            setError(e instanceof Error ? e.message : String(e)),
          );
        }}
        disabled={code.length !== 6}
      >
        Pair
      </button>
      {error && <p className="bubble error">{error}</p>}
    </div>
  );
}

function SettingsScreen(): JSX.Element {
  const [form, setForm] = useState<ProviderForm>({
    kind: 'anthropic',
    model: '',
    apiKey: '',
    baseUrl: '',
  });
  const [message, setMessage] = useState('');

  useEffect(() => {
    void api<{ provider?: Partial<ProviderForm> }>('/api/settings/provider').then(
      (s) => {
        if (s.provider) {
          setForm((f) => ({
            ...f,
            kind: s.provider!.kind ?? 'anthropic',
            model: s.provider!.model ?? '',
            baseUrl: s.provider!.baseUrl ?? '',
          }));
        }
      },
      () => {},
    );
  }, []);

  const save = (): void => {
    const body: Record<string, string> = { kind: form.kind, model: form.model };
    if (form.apiKey) body.apiKey = form.apiKey;
    if (form.baseUrl) body.baseUrl = form.baseUrl;
    void api('/api/settings/provider', { method: 'PUT', body: JSON.stringify(body) }).then(
      () => setMessage('Saved.'),
      (e: unknown) => setMessage(e instanceof Error ? e.message : String(e)),
    );
  };

  return (
    <main className="settings">
      <h3>Model</h3>
      <label>
        Provider
        <select
          value={form.kind}
          onChange={(e) => setForm({ ...form, kind: e.target.value as ProviderForm['kind'] })}
        >
          <option value="anthropic">Anthropic</option>
          <option value="openai">OpenAI</option>
          <option value="openai-compatible">OpenAI-compatible (Ollama, LM Studio, …)</option>
        </select>
      </label>
      <label>
        Model
        <input
          value={form.model}
          placeholder={form.kind === 'openai-compatible' ? 'llama3.1' : 'model id'}
          onChange={(e) => setForm({ ...form, model: e.target.value })}
        />
      </label>
      {form.kind === 'openai-compatible' && (
        <label>
          Base URL
          <input
            value={form.baseUrl}
            placeholder="http://localhost:11434/v1"
            onChange={(e) => setForm({ ...form, baseUrl: e.target.value })}
          />
        </label>
      )}
      <label>
        API key {form.kind === 'openai-compatible' && <em>(optional)</em>}
        <input
          type="password"
          value={form.apiKey}
          placeholder="stored locally by the sidecar, never sent elsewhere"
          onChange={(e) => setForm({ ...form, apiKey: e.target.value })}
        />
      </label>
      <button onClick={save} disabled={!form.model}>
        Save
      </button>
      {message && <p className="hint">{message}</p>}
    </main>
  );
}
