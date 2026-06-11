import { afterEach, describe, expect, it } from 'vitest';
import { createServer, type Server } from 'node:http';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';
import { OpenAIProvider } from '../src/providers/openai.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { runAgentTurn, type AgentEvent } from '../src/agent/loop.js';
import type { ChatMessage } from '../src/providers/types.js';

let server: Server | null = null;

afterEach(async () => {
  await new Promise<void>((r) => (server ? server.close(() => r()) : r()));
  server = null;
});

interface OAIMessage {
  role: string;
  content?: string | null;
  tool_calls?: { id: string; function: { name: string; arguments: string } }[];
  tool_call_id?: string;
}

/**
 * A mock OpenAI server that validates requests the way the real API does:
 * every `tool` message must answer a tool_call id from the immediately
 * preceding assistant message. Catches malformed history replay.
 */
function startStrictOpenAI(
  script: ((messages: OAIMessage[]) => OAIMessage)[],
): Promise<{ port: number; requests: OAIMessage[][] }> {
  const requests: OAIMessage[][] = [];
  let turn = 0;
  return new Promise((resolve) => {
    server = createServer((req, res) => {
      let body = '';
      req.on('data', (c) => (body += c));
      req.on('end', () => {
        const { messages } = JSON.parse(body) as { messages: OAIMessage[] };
        requests.push(messages);

        // strict validation, like the real API
        const knownCallIds = new Set<string>();
        for (const m of messages) {
          if (m.role === 'assistant' && m.tool_calls) {
            for (const tc of m.tool_calls) knownCallIds.add(tc.id);
          }
          if (m.role === 'tool') {
            if (!m.tool_call_id || !knownCallIds.has(m.tool_call_id)) {
              res.writeHead(400, { 'content-type': 'application/json' });
              res.end(
                JSON.stringify({
                  error: { message: `Invalid parameter: tool message without matching tool_call_id` },
                }),
              );
              return;
            }
          }
        }

        const fn = script[Math.min(turn, script.length - 1)]!;
        turn++;
        const message = fn(messages);
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(
          JSON.stringify({
            choices: [
              { message, finish_reason: message.tool_calls ? 'tool_calls' : 'stop' },
            ],
            usage: { prompt_tokens: 10, completion_tokens: 5 },
          }),
        );
      });
    });
    server.listen(0, '127.0.0.1', () => {
      resolve({ port: (server!.address() as { port: number }).port, requests });
    });
  });
}

async function turnWith(
  provider: OpenAIProvider | ScriptedProvider,
  wb: WorkbookEmulator,
  changeSets: ChangeSetManager,
  history: ChatMessage[],
  userMessage: string,
): Promise<{ events: AgentEvent[]; history: ChatMessage[]; stagedId?: string }> {
  const events: AgentEvent[] = [];
  const result = await runAgentTurn({
    provider,
    executor: wb,
    changeSets,
    workbookName: 'Test.xlsx',
    history,
    userMessage,
    onEvent: (ev) => events.push(ev),
  });
  return { events, history: result.history, ...(result.stagedChangeSetId ? { stagedId: result.stagedChangeSetId } : {}) };
}

describe('multi-turn conversation through the real OpenAI adapter', () => {
  it('second and third turns get responses after a tool-using first turn', async () => {
    const { port } = await startStrictOpenAI([
      // turn 1, step 1: tool call
      () => ({
        role: 'assistant',
        content: null,
        tool_calls: [
          {
            id: 'call_a',
            function: {
              name: 'write_range',
              arguments: JSON.stringify({
                range: 'Sheet1!A1',
                cells: [[42]],
                reason: 'user asked',
              }),
            },
          },
        ],
      }),
      // turn 1, step 2: final text
      () => ({ role: 'assistant', content: 'Staged 42 into A1.' }),
      // turn 2+: plain replies
      (messages) => ({
        role: 'assistant',
        content: `Reply ${messages.filter((m) => m.role === 'user').length}`,
      }),
    ]);

    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb');
    const provider = new OpenAIProvider('gpt-test', 'key', `http://127.0.0.1:${port}/v1`);

    const t1 = await turnWith(provider, wb, changeSets, [], 'put 42 in A1');
    expect(t1.events.some((e) => e.kind === 'changeset_staged')).toBe(true);
    expect(t1.events.some((e) => e.kind === 'text')).toBe(true);

    // approve, as the user did between messages
    await changeSets.apply(t1.stagedId!);

    const t2 = await turnWith(provider, wb, changeSets, t1.history, 'thanks, what did you do?');
    const t2Text = t2.events.filter((e) => e.kind === 'text');
    expect(t2Text.length).toBeGreaterThan(0); // <-- the reported bug: silence

    const t3 = await turnWith(
      provider,
      wb,
      changeSets,
      t2.history,
      'line one\nline two\nwith multiple lines',
    );
    expect(t3.events.some((e) => e.kind === 'text')).toBe(true);
  });

  it('emits an explicit event instead of silence when the model returns empty', async () => {
    const { port } = await startStrictOpenAI([() => ({ role: 'assistant', content: '' })]);
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb');
    const provider = new OpenAIProvider('gpt-test', 'key', `http://127.0.0.1:${port}/v1`);

    const t = await turnWith(provider, wb, changeSets, [], 'hello?');
    // Must never end a turn with zero user-visible events.
    expect(t.events.length).toBeGreaterThan(0);
  });

  it('multi-line user input flows through the scripted provider too', async () => {
    const provider = new ScriptedProvider([{ text: 'got it' }]);
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb');
    const t = await turnWith(provider, wb, changeSets, [], 'a\nb\nc');
    expect(t.events.some((e) => e.kind === 'text' && e.text === 'got it')).toBe(true);
    const sent = provider.requests[0]!.messages.findLast((m) => m.role === 'user');
    expect(sent && 'content' in sent ? sent.content : '').toContain('a\nb\nc');
  });
});
