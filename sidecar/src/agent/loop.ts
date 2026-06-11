import { zodToJsonSchema } from 'zod-to-json-schema';
import {
  toolSchemas,
  isWriteTool,
  colToIndex,
  parseA1,
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
  let emittedAnything = false;
  const emit = (ev: AgentEvent): void => {
    emittedAnything = true;
    onEvent(ev);
  };

  for (let iteration = 0; iteration < maxIterations; iteration++) {
    const response = await withRetry(() =>
      provider.chat({
        messages,
        tools: provider.supportsTools ? specs : undefined,
        temperature: 0,
        // Reasoning models spend completion tokens on internal reasoning
        // before any visible output; a small budget yields empty replies.
        maxTokens: 16384,
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

    if (assistantText.trim()) emit({ kind: 'text', text: assistantText });

    messages.push({ role: 'assistant', content: assistantText, toolCalls });

    if (toolCalls.length === 0) {
      // A turn must never end in silence: explain empty replies.
      if (!assistantText.trim()) {
        emit({
          kind: 'error',
          message:
            response.stopReason === 'max_tokens'
              ? 'The model ran out of output tokens before producing a reply (it may have spent them on internal reasoning). Try a shorter request, or a model with a larger output budget.'
              : 'The model returned an empty reply. Try rephrasing, or test the model in Settings.',
        });
      }
      break;
    }

    for (const call of toolCalls) {
      const result = await executeToolCall(call.name, call.argsJson, executor, pendingOps, emit);
      messages.push({ role: 'tool', toolCallId: call.id, content: result });
    }
  }

  // Loop exhausted while the model was still calling tools.
  if (messages[messages.length - 1]?.role === 'tool') {
    emit({
      kind: 'error',
      message: `Stopped after ${maxIterations} steps without a final answer. You can ask the model to continue.`,
    });
  }

  const result: AgentTurnResult = {
    // Persist everything after the system prompt as conversation history.
    history: messages.slice(1),
  };

  if (pendingOps.length > 0) {
    const cs = changeSets.stage(pendingOps);
    const preview = await changeSets.preview(cs.id);
    emit({ kind: 'changeset_staged', changeSetId: cs.id, preview });
    result.stagedChangeSetId = cs.id;
  }

  if (!emittedAnything) {
    // Belt and braces: whatever happened above, the user gets a signal.
    onEvent({
      kind: 'error',
      message: 'The model produced no visible output for this turn. Check the sidecar console for details.',
    });
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
      case 'find': {
        const matches = await executor.find(args.query as string, args.sheet as string | undefined);
        return JSON.stringify({ ok: true, result: matches });
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
    case 'format_range':
      return {
        kind: 'format_range',
        range: args.range as string,
        format: args.format as Extract<ChangeOp, { kind: 'format_range' }>['format'],
        reason: args.reason as string,
      };
    case 'add_sheet':
      return { kind: 'add_sheet', name: args.name as string, reason: args.reason as string };
    case 'rename_sheet':
      return {
        kind: 'rename_sheet',
        name: args.name as string,
        newName: args.newName as string,
        reason: args.reason as string,
      };
    case 'delete_sheet':
      return { kind: 'delete_sheet', name: args.name as string, reason: args.reason as string };
    case 'insert_rows':
    case 'delete_rows':
      return {
        kind: name,
        sheet: args.sheet as string,
        at: (args.at as number) - 1, // model speaks 1-based rows
        count: args.count as number,
        reason: args.reason as string,
      };
    case 'insert_cols':
    case 'delete_cols':
      return {
        kind: name,
        sheet: args.sheet as string,
        at: colToIndex((args.at as string).toUpperCase()),
        count: args.count as number,
        reason: args.reason as string,
      };
    case 'sort_range': {
      const range = args.range as string;
      const addr = parseA1(range);
      const keyOffset = colToIndex((args.keyColumn as string).toUpperCase()) - addr.startCol;
      if (keyOffset < 0 || keyOffset > addr.endCol - addr.startCol) {
        throw new Error(`keyColumn ${String(args.keyColumn)} is outside the range ${range}`);
      }
      return {
        kind: 'sort_range',
        range,
        keyOffset,
        ascending: args.ascending as boolean,
        hasHeader: args.hasHeader as boolean,
        reason: args.reason as string,
      };
    }
    default:
      throw new Error(`Not a write tool: ${name}`);
  }
}

function summarize(name: ToolName, args: Record<string, unknown>): string {
  switch (name) {
    case 'read_range':
    case 'write_range':
    case 'clear_range':
    case 'format_range':
    case 'sort_range':
      return String(args.range ?? '');
    case 'add_sheet':
    case 'rename_sheet':
    case 'delete_sheet':
      return String(args.name ?? '');
    case 'insert_rows':
    case 'delete_rows':
    case 'insert_cols':
    case 'delete_cols':
      return `${String(args.sheet ?? '')} @${String(args.at ?? '')}×${String(args.count ?? 1)}`;
    case 'find':
      return String(args.query ?? '');
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
