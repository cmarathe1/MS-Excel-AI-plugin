import {
  ProviderError,
  type ChatRequestOptions,
  type ChatResponse,
  type Provider,
  type ToolCall,
} from './types.js';

interface AnthropicContentBlock {
  type: string;
  text?: string;
  id?: string;
  name?: string;
  input?: unknown;
}

/** Anthropic Messages API adapter (plain fetch, no SDK). */
export class AnthropicProvider implements Provider {
  readonly id = 'anthropic';
  readonly supportsTools = true;

  constructor(
    readonly model: string,
    private readonly apiKey: string,
    private readonly baseUrl = 'https://api.anthropic.com',
  ) {}

  async chat(options: ChatRequestOptions): Promise<ChatResponse> {
    const system = options.messages.find((m) => m.role === 'system');
    const messages: unknown[] = [];
    for (const m of options.messages) {
      if (m.role === 'system') continue;
      if (m.role === 'user') {
        messages.push({ role: 'user', content: m.content });
      } else if (m.role === 'assistant') {
        const content: AnthropicContentBlock[] = [];
        if (m.content) content.push({ type: 'text', text: m.content });
        for (const tc of m.toolCalls ?? []) {
          content.push({
            type: 'tool_use',
            id: tc.id,
            name: tc.name,
            input: safeParse(tc.argsJson),
          });
        }
        messages.push({ role: 'assistant', content });
      } else {
        messages.push({
          role: 'user',
          content: [{ type: 'tool_result', tool_use_id: m.toolCallId, content: m.content }],
        });
      }
    }

    const body: Record<string, unknown> = {
      model: this.model,
      max_tokens: options.maxTokens ?? 4096,
      messages,
    };
    if (system) body.system = system.content;
    if (options.temperature !== undefined) body.temperature = options.temperature;
    if (options.tools && options.tools.length > 0) {
      body.tools = options.tools.map((t) => ({
        name: t.name,
        description: t.description,
        input_schema: t.inputSchema,
      }));
    }

    const res = await fetch(`${this.baseUrl}/v1/messages`, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': this.apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
      signal: options.abortSignal ?? null,
    });

    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new ProviderError(
        `Anthropic API error ${res.status}: ${truncate(text)}`,
        res.status,
        res.status === 429 || res.status >= 500,
      );
    }

    const data = (await res.json()) as {
      content: AnthropicContentBlock[];
      stop_reason: string;
      usage?: { input_tokens?: number; output_tokens?: number };
    };

    let text = '';
    const toolCalls: ToolCall[] = [];
    for (const block of data.content ?? []) {
      if (block.type === 'text' && block.text) text += block.text;
      if (block.type === 'tool_use' && block.id && block.name) {
        toolCalls.push({ id: block.id, name: block.name, argsJson: JSON.stringify(block.input ?? {}) });
      }
    }

    return {
      text,
      toolCalls,
      usage: {
        inputTokens: data.usage?.input_tokens ?? 0,
        outputTokens: data.usage?.output_tokens ?? 0,
      },
      stopReason:
        data.stop_reason === 'tool_use'
          ? 'tool_use'
          : data.stop_reason === 'max_tokens'
            ? 'max_tokens'
            : data.stop_reason === 'end_turn'
              ? 'end'
              : 'other',
    };
  }
}

function safeParse(json: string): unknown {
  try {
    return JSON.parse(json);
  } catch {
    return {};
  }
}

function truncate(s: string, n = 400): string {
  return s.length > n ? `${s.slice(0, n)}…` : s;
}
