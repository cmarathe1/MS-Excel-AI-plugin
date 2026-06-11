import { makeEnvelope } from '@excelai/shared';
import type { WebSocket } from 'ws';
import { runAgentTurn } from './agent/loop.js';
import { ChangeSetManager } from './changeset/manager.js';
import type { WorkbookExecutor } from './workbook/executor.js';
import type { ChatMessage, Provider } from './providers/types.js';
import type { Ledger } from './ledger.js';
import type { MemoryStore } from './memory/store.js';

/**
 * One session per connected workbook. Sessions are scoped by workbookId so
 * multiple open workbooks (or Excel instances) never share state.
 */
export class Session {
  readonly changeSets: ChangeSetManager;
  private history: ChatMessage[] = [];
  private busy = false;

  constructor(
    private readonly ws: WebSocket,
    readonly workbookId: string,
    private readonly workbookName: string,
    private readonly executor: WorkbookExecutor,
    private readonly getProvider: () => Provider,
    private readonly ledger: Ledger,
    private readonly memory: MemoryStore,
  ) {
    this.changeSets = new ChangeSetManager(executor, workbookId);
  }

  private emit(payload: Parameters<typeof makeEnvelope<'agent_event', unknown>>[1]): void {
    if (this.ws.readyState === this.ws.OPEN) {
      this.ws.send(JSON.stringify(makeEnvelope('agent_event', payload)));
    }
  }

  async handleChat(message: string): Promise<void> {
    if (this.busy) {
      this.emit({ kind: 'error', message: 'The agent is already working on a request.' });
      return;
    }
    this.busy = true;
    const startedAt = Date.now();
    let totalIn = 0;
    let totalOut = 0;
    try {
      const provider = this.getProvider();

      // Inject relevant memories (user + this workbook) into the turn.
      const memories = this.memory.search(message, ['user', this.workbookId], 5);
      const memoryNote =
        memories.length > 0
          ? `\n\nRelevant saved context:\n${memories.map((m) => `- ${m.content}`).join('\n')}`
          : '';

      const result = await runAgentTurn({
        provider,
        executor: this.executor,
        changeSets: this.changeSets,
        workbookName: this.workbookName,
        history: this.history,
        userMessage: message + memoryNote,
        onEvent: (ev) => this.emit(ev),
        onUsage: (usage) => {
          totalIn += usage.inputTokens;
          totalOut += usage.outputTokens;
          this.ledger.record({
            provider: provider.id,
            model: provider.model,
            feature: 'chat',
            ...usage,
          });
        },
      });
      this.history = result.history;
      // Bound history growth; compaction lands with the indexer work.
      if (this.history.length > 60) this.history = this.history.slice(-40);
      console.log(
        `[chat] ${this.workbookName} · ${provider.model} · ${Date.now() - startedAt}ms · tokens ${totalIn} in / ${totalOut} out${result.stagedChangeSetId ? ' · change-set staged' : ''}`,
      );
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      console.error(`[chat] turn failed after ${Date.now() - startedAt}ms: ${message}`);
      this.emit({ kind: 'error', message });
    } finally {
      this.busy = false;
      this.emit({ kind: 'done' });
    }
  }

  async handleApprove(changeSetId: string): Promise<void> {
    try {
      const result = await this.changeSets.apply(changeSetId);
      if (result.ok) {
        this.emit({ kind: 'changeset_applied', changeSetId });
      } else {
        this.emit({
          kind: 'error',
          message: `Change-set verification failed and was rolled back:\n${result.problems.join('\n')}`,
        });
      }
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  handleReject(changeSetId: string): void {
    try {
      this.changeSets.reject(changeSetId);
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  async handleUndo(changeSetId: string): Promise<void> {
    try {
      await this.changeSets.undo(changeSetId);
      this.emit({ kind: 'changeset_undone', changeSetId });
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }
}
