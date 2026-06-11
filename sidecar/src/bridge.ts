import type { WebSocket } from 'ws';
import {
  makeEnvelope,
  type CellData,
  type ChangeOp,
  type WorkbookMap,
} from '@excelai/shared';
import type { FindMatch, WorkbookExecutor } from './workbook/executor.js';

const READ_TIMEOUT_MS = 60_000;
const APPLY_TIMEOUT_MS = 120_000;

/**
 * WorkbookExecutor implementation that delegates to the add-in over
 * WebSocket: the sidecar is the brain, the add-in is the hands.
 */
export class BridgeExecutor implements WorkbookExecutor {
  private pending = new Map<
    string,
    { resolve: (v: unknown) => void; reject: (e: Error) => void; timer: NodeJS.Timeout }
  >();

  constructor(private readonly ws: WebSocket) {}

  /** Called by the server when a tool_result message arrives. */
  handleToolResult(payload: { requestId: string; ok: boolean; result?: unknown; error?: string }): void {
    const entry = this.pending.get(payload.requestId);
    if (!entry) return;
    this.pending.delete(payload.requestId);
    clearTimeout(entry.timer);
    if (payload.ok) entry.resolve(payload.result);
    else entry.reject(new Error(payload.error ?? 'Tool execution failed in the add-in'));
  }

  /** Reject all in-flight requests (connection closed). */
  abortAll(reason: string): void {
    for (const [, entry] of this.pending) {
      clearTimeout(entry.timer);
      entry.reject(new Error(reason));
    }
    this.pending.clear();
  }

  async getWorkbookMap(): Promise<WorkbookMap> {
    return (await this.request('get_workbook_map', {}, READ_TIMEOUT_MS)) as WorkbookMap;
  }

  async readRange(range: string): Promise<CellData[][]> {
    return (await this.request('read_range', { range }, READ_TIMEOUT_MS)) as CellData[][];
  }

  async find(query: string, sheet?: string): Promise<FindMatch[]> {
    return (await this.request('find', { query, sheet }, READ_TIMEOUT_MS)) as FindMatch[];
  }

  async applyOps(ops: ChangeOp[]): Promise<void> {
    const requestId = crypto.randomUUID();
    const env = makeEnvelope('apply_ops', { requestId, ops });
    await this.send(env, requestId, APPLY_TIMEOUT_MS);
  }

  private async request(tool: string, args: unknown, timeoutMs: number): Promise<unknown> {
    const requestId = crypto.randomUUID();
    const env = makeEnvelope('tool_exec', { requestId, tool, args });
    return this.send(env, requestId, timeoutMs);
  }

  private send(env: unknown, requestId: string, timeoutMs: number): Promise<unknown> {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId);
        reject(new Error(`Add-in did not answer within ${timeoutMs / 1000}s`));
      }, timeoutMs);
      this.pending.set(requestId, { resolve, reject, timer });
      this.ws.send(JSON.stringify(env), (err) => {
        if (err) {
          clearTimeout(timer);
          this.pending.delete(requestId);
          reject(err);
        }
      });
    });
  }
}
