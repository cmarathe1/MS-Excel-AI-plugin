import { useCallback, useEffect, useRef, useState } from 'react';
import type { ChangePreviewItem } from '@excelai/shared';
import {
  api,
  checkAuth,
  clearToken,
  connect,
  getToken,
  pair,
  SIDECAR_BASE,
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

type ProviderKind = 'anthropic' | 'openai' | 'openai-compatible';

interface SavedProvider {
  id: string;
  kind: ProviderKind;
  model: string;
  baseUrl?: string;
  /** masked, presence means a key is stored */
  apiKey?: string;
}

interface ProviderList {
  providers: SavedProvider[];
  activeProviderId: string | null;
}

const KIND_LABELS: Record<ProviderKind, string> = {
  anthropic: 'Anthropic',
  openai: 'OpenAI',
  'openai-compatible': 'OpenAI-compatible (Ollama, LM Studio, …)',
};

export function App(): JSX.Element {
  const [paired, setPaired] = useState<boolean>(() => getToken() !== null);
  const [status, setStatus] = useState<'connecting' | 'connected' | 'disconnected' | 'error'>('connecting');
  const [statusDetail, setStatusDetail] = useState('');
  const [items, setItems] = useState<ChatItem[]>([]);
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [connectAttempt, setConnectAttempt] = useState(0);
  const [activeModel, setActiveModel] = useState<string | null>(null);
  const connection = useRef<SidecarConnection | null>(null);
  const chatRef = useRef<HTMLElement | null>(null);

  // Keep the latest message in view.
  useEffect(() => {
    chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight });
  }, [items, busy]);

  const unpair = useCallback((): void => {
    clearToken();
    connection.current?.close();
    setPaired(false);
  }, []);

  // Stale token (sidecar data reset) -> back to pairing automatically.
  useEffect(() => {
    if (!paired) return;
    void checkAuth().then((ok) => {
      if (!ok) unpair();
    });
  }, [paired, unpair]);

  const refreshActiveModel = useCallback((): void => {
    void api<ProviderList>('/api/settings/providers').then(
      (list) => {
        const active = list.providers.find((p) => p.id === list.activeProviderId);
        setActiveModel(active ? active.model : null);
        if (!active) setShowSettings(true); // first run: configure a model
      },
      () => {},
    );
  }, []);

  useEffect(() => {
    if (paired) refreshActiveModel();
  }, [paired, refreshActiveModel]);

  // Auto-reconnect with a fixed backoff while the sidecar is unreachable.
  useEffect(() => {
    if (!paired || status === 'connected' || status === 'connecting') return;
    const timer = setTimeout(() => {
      setStatus('connecting');
      setConnectAttempt((n) => n + 1);
    }, 8000);
    return () => clearTimeout(timer);
  }, [paired, status]);

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
  }, [paired, connectAttempt, handleEvent]);

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
        {activeModel && !showSettings && (
          <button className="chip" title="Change model" onClick={() => setShowSettings(true)}>
            {activeModel}
          </button>
        )}
        {status !== 'connected' && (
          <button
            className="ghost"
            onClick={() => {
              setStatus('connecting');
              setConnectAttempt((n) => n + 1);
            }}
          >
            Reconnect
          </button>
        )}
        <button
          className="ghost"
          onClick={() => {
            if (showSettings) refreshActiveModel();
            setShowSettings((s) => !s);
          }}
        >
          {showSettings ? 'Chat' : 'Settings'}
        </button>
      </header>
      {status !== 'connected' && (
        <p className="hint banner">
          {status === 'connecting'
            ? 'Connecting to the sidecar…'
            : `Not connected${statusDetail ? ` — ${statusDetail}` : ''}. Is the sidecar running? (pnpm dev:sidecar)`}{' '}
          <button className="link" onClick={unpair}>
            Re-pair
          </button>
        </p>
      )}

      {showSettings ? (
        <ModelsScreen onChanged={refreshActiveModel} />
      ) : (
        <>
          <main className="chat" ref={chatRef}>
            {items.length === 0 && (
              <div className="empty">
                <div className="empty-title">Ask anything about this workbook</div>
                <div className="empty-sub">Edits are proposed first — nothing changes until you approve.</div>
              </div>
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
            {busy && (
              <div className="typing" aria-label="Working">
                <span /><span /><span />
              </div>
            )}
          </main>
          <footer>
            <div className="composer">
              <textarea
                value={input}
                rows={1}
                placeholder="Ask about this sheet… (Shift+Enter for a new line)"
                onChange={(e) => setInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    sendMessage();
                  }
                }}
              />
              <button
                className="send"
                aria-label="Send"
                onClick={sendMessage}
                disabled={busy || status !== 'connected' || !input.trim()}
              >
                ↑
              </button>
            </div>
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
    case 'changeset': {
      const n = item.preview.length;
      return (
        <div className={`cs ${item.status}`}>
          <div className="cs-head">
            <span className="cs-title">
              {item.status === 'staged'
                ? `${n} proposed change${n === 1 ? '' : 's'}`
                : item.status === 'applied'
                  ? 'Applied'
                  : item.status === 'undone'
                    ? 'Undone'
                    : 'Dismissed'}
            </span>
            <span className="cs-actions">
              {item.status === 'staged' && (
                <>
                  <button className="sm" onClick={() => props.onApprove(item.id)}>
                    Apply
                  </button>
                  <button className="sm subtle" onClick={() => props.onReject(item.id)}>
                    Dismiss
                  </button>
                </>
              )}
              {item.status === 'applied' && (
                <button className="sm subtle" onClick={() => props.onUndo(item.id)}>
                  Undo
                </button>
              )}
            </span>
          </div>
          <ul className="cs-ops">
            {item.preview.map((p, i) => (
              <OpRow key={i} item={p} />
            ))}
          </ul>
        </div>
      );
    }
  }
}

function OpRow({ item }: { item: ChangePreviewItem }): JSX.Element {
  const { op } = item;
  let target: string;
  let desc: string;
  switch (op.kind) {
    case 'write_range': {
      const cells = op.cells.length * (op.cells[0]?.length ?? 0);
      target = op.range;
      desc = cells === 1 ? 'write' : `write ${cells} cells`;
      break;
    }
    case 'clear_range':
      target = op.range;
      desc = 'clear';
      break;
    case 'format_range':
      target = op.range;
      desc = `format: ${Object.keys(op.format).join(', ')}`;
      break;
    case 'add_sheet':
      target = op.name;
      desc = 'new sheet';
      break;
  }
  return (
    <li title={`${target} · ${op.reason}`}>
      <code>{target.replace(/^[^!]*!/, '')}</code>
      <span className="cs-desc">{desc}</span>
      <span className="cs-reason">{op.reason}</span>
    </li>
  );
}

function PairingScreen({ onPaired }: { onPaired: () => void }): JSX.Element {
  const [code, setCode] = useState('');
  const [error, setError] = useState('');
  const [sidecarUp, setSidecarUp] = useState<boolean | null>(null);

  // Poll health so the user can see the sidecar come alive while pairing.
  useEffect(() => {
    let stop = false;
    const check = async (): Promise<void> => {
      try {
        const res = await fetch(`${SIDECAR_BASE}/api/health`);
        if (!stop) setSidecarUp(res.ok);
      } catch {
        if (!stop) setSidecarUp(false);
      }
    };
    void check();
    const timer = setInterval(() => void check(), 3000);
    return () => {
      stop = true;
      clearInterval(timer);
    };
  }, []);

  return (
    <div className="app center">
      <h2>Pair with the companion app</h2>
      <p className="hint">
        {sidecarUp === true
          ? '✓ Sidecar detected. Enter the pairing code from its console.'
          : sidecarUp === false
            ? 'Sidecar not detected — start it with: pnpm dev:sidecar'
            : 'Looking for the sidecar…'}
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

/* ----------------------- model management ----------------------- */

interface AddForm {
  kind: ProviderKind;
  model: string;
  apiKey: string;
  baseUrl: string;
}

const EMPTY_FORM: AddForm = { kind: 'anthropic', model: '', apiKey: '', baseUrl: '' };

function ModelsScreen({ onChanged }: { onChanged: () => void }): JSX.Element {
  const [list, setList] = useState<ProviderList>({ providers: [], activeProviderId: null });
  const [form, setForm] = useState<AddForm>(EMPTY_FORM);
  const [showAdd, setShowAdd] = useState(false);
  const [message, setMessage] = useState('');
  const [rowMessage, setRowMessage] = useState<Record<string, string>>({});
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [busyAction, setBusyAction] = useState('');

  const load = useCallback((): void => {
    void api<ProviderList>('/api/settings/providers').then(
      (l) => {
        setList(l);
        if (l.providers.length === 0) setShowAdd(true);
      },
      (e: unknown) => setMessage(String(e instanceof Error ? e.message : e)),
    );
  }, []);

  useEffect(load, [load]);

  const formBody = (): Record<string, unknown> => {
    const body: Record<string, unknown> = { kind: form.kind, model: form.model.trim() };
    if (form.apiKey.trim()) body.apiKey = form.apiKey.trim();
    if (form.baseUrl.trim()) body.baseUrl = form.baseUrl.trim();
    return body;
  };

  const addModel = (): void => {
    setBusyAction('add');
    void api('/api/settings/providers', { method: 'POST', body: JSON.stringify(formBody()) })
      .then(() => {
        setForm(EMPTY_FORM);
        setShowAdd(false);
        setSuggestions([]);
        setMessage('Model added.');
        load();
        onChanged();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const testCandidate = (): void => {
    setBusyAction('test');
    setMessage('Testing…');
    void api<{ ok: boolean; reply?: string; error?: string }>('/api/settings/providers/test', {
      method: 'POST',
      body: JSON.stringify(formBody()),
    })
      .then((r) =>
        setMessage(r.ok ? `✓ Connected — model replied: "${r.reply ?? ''}"` : `✗ ${r.error ?? 'Failed'}`),
      )
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const loadModels = (): void => {
    setBusyAction('models');
    setMessage('Fetching available models…');
    const body: Record<string, unknown> = { kind: form.kind };
    if (form.apiKey.trim()) body.apiKey = form.apiKey.trim();
    if (form.baseUrl.trim()) body.baseUrl = form.baseUrl.trim();
    void api<{ models: string[]; error?: string }>('/api/providers/models', {
      method: 'POST',
      body: JSON.stringify(body),
    })
      .then((r) => {
        setSuggestions(r.models);
        setMessage(
          r.models.length > 0
            ? `${r.models.length} models available — pick one from the Model field.`
            : `✗ ${r.error ?? 'No models found'}`,
        );
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const activate = (id: string): void => {
    void api(`/api/settings/providers/${id}/activate`, { method: 'POST' }).then(() => {
      load();
      onChanged();
    });
  };

  const remove = (id: string): void => {
    void api(`/api/settings/providers/${id}`, { method: 'DELETE' }).then(() => {
      load();
      onChanged();
    });
  };

  const testSaved = (id: string): void => {
    setRowMessage((m) => ({ ...m, [id]: 'Testing…' }));
    void api<{ ok: boolean; reply?: string; error?: string }>(
      `/api/settings/providers/${id}/test`,
      { method: 'POST' },
    )
      .then((r) =>
        setRowMessage((m) => ({
          ...m,
          [id]: r.ok ? `✓ "${r.reply ?? ''}"` : `✗ ${r.error ?? 'Failed'}`,
        })),
      )
      .catch((e: unknown) =>
        setRowMessage((m) => ({ ...m, [id]: `✗ ${e instanceof Error ? e.message : String(e)}` })),
      );
  };

  return (
    <main className="settings">
      <h3>Your models</h3>
      {list.providers.length === 0 && <p className="hint">No models yet — add one below.</p>}
      {list.providers.map((p) => {
        const active = p.id === list.activeProviderId;
        return (
          <div key={p.id} className={`model-card${active ? ' active' : ''}`}>
            <div className="model-card-main">
              <strong>{p.model}</strong> {active && <span className="badge applied">active</span>}
              <div className="hint">
                {KIND_LABELS[p.kind]}
                {p.baseUrl ? ` · ${p.baseUrl}` : ''}
                {p.apiKey ? ` · key ${p.apiKey}` : ' · no key'}
              </div>
              {rowMessage[p.id] && <div className="hint">{rowMessage[p.id]}</div>}
            </div>
            <div className="actions">
              {!active && <button onClick={() => activate(p.id)}>Use</button>}
              <button className="ghost" onClick={() => testSaved(p.id)}>
                Test
              </button>
              <button className="ghost danger" onClick={() => remove(p.id)}>
                Delete
              </button>
            </div>
          </div>
        );
      })}

      {!showAdd ? (
        <button className="ghost" onClick={() => setShowAdd(true)}>
          + Add a model
        </button>
      ) : (
        <div className="add-form">
          <h3>Add a model</h3>
          <label>
            Provider
            <select
              value={form.kind}
              onChange={(e) => {
                setSuggestions([]);
                setForm({ ...form, kind: e.target.value as ProviderKind });
              }}
            >
              <option value="anthropic">Anthropic</option>
              <option value="openai">OpenAI</option>
              <option value="openai-compatible">OpenAI-compatible (Ollama, LM Studio, …)</option>
            </select>
          </label>
          {form.kind === 'openai-compatible' && (
            <>
              <label>
                Base URL
                <input
                  value={form.baseUrl}
                  placeholder="http://localhost:11434/v1"
                  onChange={(e) => setForm({ ...form, baseUrl: e.target.value })}
                />
              </label>
              <button
                className="ghost"
                onClick={() =>
                  setForm({ ...form, baseUrl: 'http://localhost:11434/v1', model: form.model || 'llama3.1' })
                }
              >
                Use Ollama defaults
              </button>
            </>
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
          <label>
            Model
            <input
              value={form.model}
              list="model-suggestions"
              placeholder={form.kind === 'openai-compatible' ? 'llama3.1' : 'model id'}
              onChange={(e) => setForm({ ...form, model: e.target.value })}
            />
            <datalist id="model-suggestions">
              {suggestions.map((m) => (
                <option key={m} value={m} />
              ))}
            </datalist>
          </label>
          <div className="actions">
            <button
              className="ghost"
              onClick={loadModels}
              disabled={busyAction !== '' || (form.kind !== 'openai' && form.kind !== 'anthropic' && !form.baseUrl)}
            >
              {busyAction === 'models' ? 'Fetching…' : 'List available models'}
            </button>
            <button className="ghost" onClick={testCandidate} disabled={!form.model || busyAction !== ''}>
              {busyAction === 'test' ? 'Testing…' : 'Test'}
            </button>
            <button onClick={addModel} disabled={!form.model || busyAction !== ''}>
              {busyAction === 'add' ? 'Adding…' : 'Add model'}
            </button>
            {list.providers.length > 0 && (
              <button className="ghost" onClick={() => setShowAdd(false)}>
                Cancel
              </button>
            )}
          </div>
        </div>
      )}
      {message && <p className="hint">{message}</p>}
    </main>
  );
}
