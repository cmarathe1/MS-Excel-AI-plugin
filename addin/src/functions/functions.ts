/**
 * =AI() custom-function family. Each call forwards to the sidecar's batch
 * engine, which dedupes, caches and rate-limits — a column fill of identical
 * prompts bills once.
 *
 * Failures surface as Excel #VALUE! errors with a message, never as
 * silently wrong values.
 */

const SIDECAR_BASE = 'http://127.0.0.1:8923';

async function callFunction(body: Record<string, unknown>): Promise<string> {
  const token = localStorage.getItem('excelai-token');
  if (!token) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      'Not paired with the sidecar. Open the task pane first.',
    );
  }
  const res = await fetch(`${SIDECAR_BASE}/api/functions/run`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', authorization: `Bearer ${token}` },
    body: JSON.stringify(body),
  });
  const data = (await res.json().catch(() => ({}))) as { value?: string; error?: string };
  if (!res.ok || data.value === undefined) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      data.error ?? `Sidecar error (${res.status})`,
    );
  }
  return data.value;
}

export async function ai(prompt: string, input?: string): Promise<string> {
  const body: Record<string, unknown> = { kind: 'ai', prompt };
  if (input !== undefined && input !== null) body.input = String(input);
  return callFunction(body);
}

export async function aiClassify(text: string, categories: string): Promise<string> {
  return callFunction({
    kind: 'classify',
    prompt: '',
    input: String(text),
    categories: String(categories)
      .split(',')
      .map((c) => c.trim())
      .filter(Boolean),
  });
}

export async function aiExtract(text: string, field: string): Promise<string> {
  return callFunction({ kind: 'extract', prompt: String(field), input: String(text) });
}

export async function aiTranslate(text: string, targetLang: string): Promise<string> {
  return callFunction({ kind: 'translate', prompt: '', input: String(text), targetLang });
}

/* global CustomFunctions */
CustomFunctions.associate('AI', ai);
CustomFunctions.associate('AI_CLASSIFY', aiClassify);
CustomFunctions.associate('AI_EXTRACT', aiExtract);
CustomFunctions.associate('AI_TRANSLATE', aiTranslate);
