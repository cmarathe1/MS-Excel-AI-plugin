import {
  ProviderError,
  type ChatRequestOptions,
  type ChatResponse,
  type Provider,
  type ToolCall,
} from './types.js';

/**
 * OpenAI Chat Completions adapter. With a custom baseUrl this same adapter
 * speaks to Ollama, LM Studio, vLLM, llama.cpp server, LiteLLM and any other
 * OpenAI-compatible endpoint — that's what makes "any model" true.
 *
 * Newer OpenAI models renamed `max_tokens` to `max_completion_tokens` and
 * some reject `temperature` overrides entirely, while most compatible
 * servers still expect the classic parameters. Rather than hard-coding model
 * lists, the adapter retries a 400 that names an offending parameter with an
 * adjusted body and remembers the working shape for the rest of the session.
 */
export class OpenAIProvider implements Provider {
  readonly supportsTools = true;

  private maxTokensParam: 'max_tokens' | 'max_completion_tokens' = 'max_tokens';
  private sendTemperature = true;

  constructor(
    readonly model: string,
    private readonly apiKey: string | undefined,
    private readonly baseUrl = 'https://api.openai.com/v1',
    readonly id: string = 'openai',
  ) {}

  async chat(options: ChatRequestOptions): Promise<ChatResponse> {
    // At most one retry per adaptable parameter.
    for (let attempt = 0; attempt < 3; attempt++) {
      const res = await this.send(options);
      if (res.ok) return this.parse(res);

      const text = await res.text().catch(() => '');
      if (res.status === 400 && this.adaptParams(text)) continue;

      throw new ProviderError(
        `${this.id} API error ${res.status}: ${truncate(text)}`,
        res.status,
        res.status === 429 || res.status >= 500,
      );
    }
    throw new ProviderError(`${this.id}: request failed after parameter adaptation`);
  }

  /** Returns true if the 400 named a parameter we can adjust. */
  private adaptParams(errorText: string): boolean {
    const lower = errorText.toLowerCase();
    const unsupported = /unsupported|not supported|does not support|invalid/.test(lower);
    if (!unsupported) return false;
    if (this.maxTokensParam === 'max_tokens' && lower.includes('max_tokens')) {
      this.maxTokensParam = 'max_completion_tokens';
      return true;
    }
    if (this.sendTemperature && lower.includes('temperature')) {
      this.sendTemperature = false;
      return true;
    }
    return false;
  }

  private async send(options: ChatRequestOptions): Promise<Response> {
    const messages: unknown[] = options.messages.map((m) => {
      switch (m.role) {
        case 'system':
          return { role: 'system', content: m.content };
        case 'user':
          return { role: 'user', content: m.content };
        case 'assistant': {
          const out: Record<string, unknown> = { role: 'assistant', content: m.content || null };
          if (m.toolCalls && m.toolCalls.length > 0) {
            out.tool_calls = m.toolCalls.map((tc) => ({
              id: tc.id,
              type: 'function',
              function: { name: tc.name, arguments: tc.argsJson },
            }));
          }
          return out;
        }
        case 'tool':
          return { role: 'tool', tool_call_id: m.toolCallId, content: m.content };
      }
    });

    const body: Record<string, unknown> = {
      model: this.model,
      messages,
      [this.maxTokensParam]: options.maxTokens ?? 4096,
    };
    if (this.sendTemperature && options.temperature !== undefined) {
      body.temperature = options.temperature;
    }
    if (options.tools && options.tools.length > 0) {
      body.tools = options.tools.map((t) => ({
        type: 'function',
        function: { name: t.name, description: t.description, parameters: t.inputSchema },
      }));
    }

    const headers: Record<string, string> = { 'content-type': 'application/json' };
    if (this.apiKey) headers.authorization = `Bearer ${this.apiKey}`;

    return fetch(`${this.baseUrl}/chat/completions`, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
      signal: options.abortSignal ?? null,
    });
  }

  private async parse(res: Response): Promise<ChatResponse> {
    const data = (await res.json()) as {
      choices?: {
        message?: {
          content?: string | null;
          tool_calls?: { id: string; function: { name: string; arguments: string } }[];
        };
        finish_reason?: string;
      }[];
      usage?: { prompt_tokens?: number; completion_tokens?: number };
    };

    const choice = data.choices?.[0];
    if (!choice?.message) throw new ProviderError(`${this.id}: empty response`);

    const toolCalls: ToolCall[] = (choice.message.tool_calls ?? []).map((tc) => ({
      id: tc.id,
      name: tc.function.name,
      argsJson: tc.function.arguments,
    }));

    return {
      text: choice.message.content ?? '',
      toolCalls,
      usage: {
        inputTokens: data.usage?.prompt_tokens ?? 0,
        outputTokens: data.usage?.completion_tokens ?? 0,
      },
      stopReason:
        choice.finish_reason === 'tool_calls'
          ? 'tool_use'
          : choice.finish_reason === 'length'
            ? 'max_tokens'
            : choice.finish_reason === 'stop'
              ? 'end'
              : 'other',
    };
  }
}

function truncate(s: string, n = 400): string {
  return s.length > n ? `${s.slice(0, n)}…` : s;
}
