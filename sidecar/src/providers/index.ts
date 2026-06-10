import { AnthropicProvider } from './anthropic.js';
import { OpenAIProvider } from './openai.js';
import { ProviderError, type Provider, type ProviderConfig } from './types.js';

export * from './types.js';
export { AnthropicProvider } from './anthropic.js';
export { OpenAIProvider } from './openai.js';
export { ScriptedProvider, type ScriptedTurn } from './scripted.js';

export function createProvider(config: ProviderConfig): Provider {
  switch (config.kind) {
    case 'anthropic': {
      if (!config.apiKey) throw new ProviderError('Anthropic requires an API key');
      return new AnthropicProvider(config.model, config.apiKey, config.baseUrl);
    }
    case 'openai': {
      if (!config.apiKey) throw new ProviderError('OpenAI requires an API key');
      return new OpenAIProvider(config.model, config.apiKey, config.baseUrl);
    }
    case 'openai-compatible': {
      if (!config.baseUrl) {
        throw new ProviderError('openai-compatible providers require a baseUrl (e.g. http://localhost:11434/v1)');
      }
      return new OpenAIProvider(config.model, config.apiKey, config.baseUrl, 'openai-compatible');
    }
  }
}

/** Retry with exponential backoff on retryable provider errors. */
export async function withRetry<T>(
  fn: () => Promise<T>,
  attempts = 3,
  baseDelayMs = 500,
): Promise<T> {
  let lastError: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (e) {
      lastError = e;
      const retryable = e instanceof ProviderError ? e.retryable : false;
      if (!retryable || i === attempts - 1) throw e;
      await new Promise((r) => setTimeout(r, baseDelayMs * 2 ** i));
    }
  }
  throw lastError;
}
