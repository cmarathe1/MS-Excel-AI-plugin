import { zodToJsonSchema } from 'zod-to-json-schema';
import {
  toolSchemas,
  isWriteTool,
  type ChangeOp,
  type ToolName,
  type ChangePreviewItem,
} from '@excelai/shared';
import type { ChatMessage, Provider, ToolSpec } from '../providers/types.js';
import { withRetry } from '../providers/index.js';
import type { WorkbookExecutor } from '../workbook/executor.js';
import type { ChangeSetManager } from '../changeset/manager.js';
import { buildSystemPrompt, emulatedToolInstructions } from './prompts.js';

export type AgentEvent =
  | { kind: 'text'; text: string }
  | { kind: 'tool_use'; tool: string; summary: string }
  | { kind: 'changeset_staged'; changeSetId: string; preview: ChangePreviewItem[] }
  | { kind: 'error'; message: string };

export interface AgentTurnOptions {
  provider: Provider;
  executor: WorkbookExecutor;
  changeSets: ChangeSetManager;
  workbookName: string;
  history: ChatMessage[];
  userMessage: string;
  onEvent: (ev: AgentEvent) => void;
  onUsage?: (usage: { inputTokens: number; outputTokens: number }) => void;
  maxIterations?: number;
}

export interface AgentTurnResult {
  history: ChatMessage[];
  stagedChangeSetId?: string;
}

const MAX_TOOL_RESULT_CHARS = 20_000;
const MAX_REPAIR_ATTEMPTS_PER_CALL = 1; // a malformed call gets one structured error back per attempt

export function buildToolSpecs(): ToolSpec[] {
  return (Object.keys(toolSchemas) as ToolName[]).map((name) => ({
    name,
    description: toolSchemas[name].description,
    inputSchema: zodToJsonSchema(toolSchemas[name].parameters, { $refStrategy: 'none' }) as Record<
      string,
      unknown
    >,
  }));
}

/**
 * One user-visible agent turn: a bounded tool-calling loop in which every
 * tool argument is schema-validated before it does anything, write tools are
 * staged into a change-set (never applied), and malformed calls get a
 * structured error fed back so the model can repair itself.
 */
export async function runAgentTurn(options: AgentTurnOptions): Promise<AgentTurnResult> {
  const {
    provider,
    executor,
    changeSets,
    workbookName,
    onEvent,
    onUsage,
    maxIterations = 16,
  } = options;

  let system = buildSystemPrompt(workbookName);
  const specs = buildToolSpecs();
  if (!provider.supportsTools) {
    const toolList = specs
      .map((s) => `- ${s.name}: ${s.description}\n  args schema: ${JSON.stringify(s.inputSchema)}`)
      .join('\n');
    system += `\n\n${emulatedToolInstructions(toolList)}`;
  }

  const messages: ChatMessage[] = [
    { role: 'system', content: system },
    ...options.history,
    { role: 'user', content: options.userMessage },
  ];

  const pendingOps: ChangeOp[] = [];

  for (let iteration = 0; iteration < maxIterations; iteration++) {
    const response = await withRetry(() =>
      provider.chat({
        messages,
        tools: provider.supportsTools ? specs : undefined,
        temperature: 0,
      }),
    );
    onUsage?.(response.usage);

    let toolCalls = response.toolCalls;
    let assistantText = response.text;

    // Emulated tool calling: the "tool call" arrives as a JSON text body.
    if (!provider.supportsTools && toolCalls.length === 0) {
      const parsed = tryParseEmulatedCall(response.text);
      if (parsed) {
        toolCalls = [{ id: `emu_${iteration}`, name: parsed.tool, argsJson: JSON.stringify(parsed.args) }];
        assistantText = '';
      }
    }

    if (assistantText.trim()) onEvent({ kind: 'text', text: assistantText });

    messages.push({ role: 'assistant', content: assistantText, toolCalls });

    if (toolCalls.length === 0) break;

    for (const call of toolCalls) {
      const result = await executeToolCall(call.name, call.argsJson, executor, pendingOps, onEvent);
      messages.push({ role: 'tool', toolCallId: call.id, content: result });
    }
  }

  const result: AgentTurnResult = {
    // Persist everything after the system prompt as conversation history.
    history: messages.slice(1),
  };

  if (pendingOps.length > 0) {
    const cs = changeSets.stage(pendingOps);
    const preview = await changeSets.preview(cs.id);
    onEvent({ kind: 'changeset_staged', changeSetId: cs.id, preview });
    result.stagedChangeSetId = cs.id;
  }

  return result;
}

async function executeToolCall(
  name: string,
  argsJson: string,
  executor: WorkbookExecutor,
  pendingOps: ChangeOp[],
  onEvent: (ev: AgentEvent) => void,
): Promise<string> {
  if (!(name in toolSchemas)) {
    return errorResult(`Unknown tool "${name}". Available tools: ${Object.keys(toolSchemas).join(', ')}`);
  }
  const toolName = name as ToolName;

  let rawArgs: unknown;
  try {
    rawArgs = argsJson.trim() === '' ? {} : JSON.parse(argsJson);
  } catch {
    return errorResult('Tool arguments were not valid JSON. Re-send the call with valid JSON arguments.');
  }

  const parsed = toolSchemas[toolName].parameters.safeParse(rawArgs);
  if (!parsed.success) {
    const issues = parsed.error.issues
      .map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`)
      .join('; ');
    return errorResult(`Invalid arguments for ${toolName}: ${issues}`);
  }
  const args = parsed.data as Record<string, unknown>;

  onEvent({ kind: 'tool_use', tool: toolName, summary: summarize(toolName, args) });

  try {
    if (isWriteTool(toolName)) {
      pendingOps.push(toChangeOp(toolName, args));
      return JSON.stringify({
        ok: true,
        staged: true,
        note: 'Staged into the pending change-set. Not applied yet; reads will not reflect it until the user approves.',
      });
    }
    switch (toolName) {
      case 'get_workbook_map': {
        const map = await executor.getWorkbookMap();
        return JSON.stringify({ ok: true, result: map });
      }
      case 'read_range': {
        const cells = await executor.readRange(args.range as string);
        const json = JSON.stringify({ ok: true, result: cells });
        if (json.length > MAX_TOOL_RESULT_CHARS) {
          return errorResult(
            `Result too large (${json.length} chars). Read a smaller, more targeted range.`,
          );
        }
        return json;
      }
      default:
        return errorResult(`Tool ${toolName} is not executable here.`);
    }
  } catch (e) {
    return errorResult(e instanceof Error ? e.message : String(e));
  }
}

function toChangeOp(name: ToolName, args: Record<string, unknown>): ChangeOp {
  switch (name) {
    case 'write_range': {
      const cells = (args.cells as (string | number | boolean | null | { formula: string })[][]).map(
        (row) =>
          row.map((cell) =>
            typeof cell === 'object' && cell !== null && 'formula' in cell
              ? { formula: cell.formula }
              : { value: cell },
          ),
      );
      return {
        kind: 'write_range',
        range: args.range as string,
        cells,
        reason: args.reason as string,
      };
    }
    case 'clear_range':
      return { kind: 'clear_range', range: args.range as string, reason: args.reason as string };
    case 'add_sheet':
      return { kind: 'add_sheet', name: args.name as string, reason: args.reason as string };
    default:
      throw new Error(`Not a write tool: ${name}`);
  }
}

function summarize(name: ToolName, args: Record<string, unknown>): string {
  switch (name) {
    case 'read_range':
    case 'write_range':
    case 'clear_range':
      return String(args.range ?? '');
    case 'add_sheet':
      return String(args.name ?? '');
    default:
      return '';
  }
}

function errorResult(message: string): string {
  return JSON.stringify({ ok: false, error: message });
}

function tryParseEmulatedCall(text: string): { tool: string; args: unknown } | null {
  const trimmed = text.trim();
  // Tolerate code fences around the JSON.
  const body = trimmed.startsWith('```')
    ? trimmed.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '')
    : trimmed;
  if (!body.startsWith('{')) return null;
  try {
    const obj = JSON.parse(body) as Record<string, unknown>;
    if (typeof obj.tool === 'string' && 'args' in obj) {
      return { tool: obj.tool, args: obj.args };
    }
    return null;
  } catch {
    return null;
  }
}

export { MAX_REPAIR_ATTEMPTS_PER_CALL };
