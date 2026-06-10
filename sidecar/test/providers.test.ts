import { afterEach, describe, expect, it, vi } from 'vitest';
import { AnthropicProvider } from '../src/providers/anthropic.js';
import { OpenAIProvider } from '../src/providers/openai.js';
import { withRetry } from '../src/providers/index.js';
import { ProviderError } from '../src/providers/types.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

function stubFetch(handler: (url: string, init: RequestInit) => { status: number; body: unknown }) {
  const calls: { url: string; body: unknown }[] = [];
  vi.stubGlobal('fetch', async (url: string | URL, init?: RequestInit) => {
    const parsedBody = init?.body ? JSON.parse(String(init.body)) : undefined;
    calls.push({ url: String(url), body: parsedBody });
    const res = handler(String(url), init ?? {});
    return new Response(JSON.stringify(res.body), {
      status: res.status,
      headers: { 'content-type': 'application/json' },
    });
  });
  return calls;
}

describe('Anthropic adapter', () => {
  it('maps tool definitions, tool calls and usage', async () => {
    const calls = stubFetch(() => ({
      status: 200,
      body: {
        content: [
          { type: 'text', text: 'Reading the range.' },
          { type: 'tool_use', id: 'tu_1', name: 'read_range', input: { range: 'Sheet1!A1' } },
        ],
        stop_reason: 'tool_use',
        usage: { input_tokens: 11, output_tokens: 7 },
      },
    }));

    const p = new AnthropicProvider('claude-test', 'sk-test');
    const res = await p.chat({
      messages: [
        { role: 'system', content: 'sys' },
        { role: 'user', content: 'hello' },
      ],
      tools: [{ name: 'read_range', description: 'd', inputSchema: { type: 'object' } }],
    });

    expect(res.toolCalls).toEqual([
      { id: 'tu_1', name: 'read_range', argsJson: '{"range":"Sheet1!A1"}' },
    ]);
    expect(res.stopReason).toBe('tool_use');
    expect(res.usage).toEqual({ inputTokens: 11, outputTokens: 7 });

    const sent = calls[0]!.body as Record<string, unknown>;
    expect(sent.system).toBe('sys');
    expect((sent.tools as unknown[])[0]).toMatchObject({ name: 'read_range' });
    expect(calls[0]!.url).toContain('/v1/messages');
  });

  it('throws a retryable ProviderError on 429/5xx', async () => {
    stubFetch(() => ({ status: 429, body: { error: 'rate limited' } }));
    const p = new AnthropicProvider('claude-test', 'sk-test');
    await expect(p.chat({ messages: [{ role: 'user', content: 'x' }] })).rejects.toMatchObject({
      retryable: true,
      status: 429,
    });
  });
});

describe('OpenAI / OpenAI-compatible adapter', () => {
  it('maps tool calls and respects custom base URLs (Ollama etc.)', async () => {
    const calls = stubFetch(() => ({
      status: 200,
      body: {
        choices: [
          {
            message: {
              content: null,
              tool_calls: [
                { id: 'c1', function: { name: 'read_range', arguments: '{"range":"Sheet1!A1"}' } },
              ],
            },
            finish_reason: 'tool_calls',
          },
        ],
        usage: { prompt_tokens: 5, completion_tokens: 3 },
      },
    }));

    const p = new OpenAIProvider('llama3', undefined, 'http://localhost:11434/v1', 'openai-compatible');
    const res = await p.chat({ messages: [{ role: 'user', content: 'x' }] });

    expect(calls[0]!.url).toBe('http://localhost:11434/v1/chat/completions');
    expect(res.toolCalls[0]).toMatchObject({ name: 'read_range' });
    expect(res.stopReason).toBe('tool_use');
  });

  it('round-trips assistant tool calls and tool results in history', async () => {
    const calls = stubFetch(() => ({
      status: 200,
      body: { choices: [{ message: { content: 'ok' }, finish_reason: 'stop' }] },
    }));
    const p = new OpenAIProvider('gpt-test', 'key');
    await p.chat({
      messages: [
        { role: 'user', content: 'q' },
        {
          role: 'assistant',
          content: '',
          toolCalls: [{ id: 'c1', name: 't', argsJson: '{}' }],
        },
        { role: 'tool', toolCallId: 'c1', content: '{"ok":true}' },
      ],
    });
    const sent = calls[0]!.body as { messages: Record<string, unknown>[] };
    expect(sent.messages[1]).toMatchObject({ role: 'assistant' });
    expect(sent.messages[2]).toMatchObject({ role: 'tool', tool_call_id: 'c1' });
  });
});

describe('withRetry', () => {
  it('retries retryable errors with backoff then succeeds', async () => {
    let attempts = 0;
    const result = await withRetry(
      async () => {
        attempts++;
        if (attempts < 3) throw new ProviderError('rate', 429, true);
        return 'ok';
      },
      3,
      1,
    );
    expect(result).toBe('ok');
    expect(attempts).toBe(3);
  });

  it('does not retry non-retryable errors', async () => {
    let attempts = 0;
    await expect(
      withRetry(
        async () => {
          attempts++;
          throw new ProviderError('bad key', 401, false);
        },
        3,
        1,
      ),
    ).rejects.toThrow('bad key');
    expect(attempts).toBe(1);
  });
});
