import { describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { runAgentTurn, type AgentEvent } from '../src/agent/loop.js';

async function runScript(
  wb: WorkbookEmulator,
  provider: ScriptedProvider,
): Promise<{ events: AgentEvent[]; changeSets: ChangeSetManager; stagedId?: string }> {
  const changeSets = new ChangeSetManager(wb, 'wb-test');
  const events: AgentEvent[] = [];
  const result = await runAgentTurn({
    provider,
    executor: wb,
    changeSets,
    workbookName: 'Test.xlsx',
    history: [],
    userMessage: 'do the thing',
    onEvent: (ev) => events.push(ev),
  });
  const out: { events: AgentEvent[]; changeSets: ChangeSetManager; stagedId?: string } = {
    events,
    changeSets,
  };
  if (result.stagedChangeSetId !== undefined) out.stagedId = result.stagedChangeSetId;
  return out;
}

describe('agent loop (full integration against the emulator)', () => {
  it('reads the workbook and stages a write as a change-set', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 5);
    wb.setCell('Sheet1!A2', 7);

    const provider = new ScriptedProvider([
      { toolCalls: [{ name: 'get_workbook_map', args: {} }] },
      { toolCalls: [{ name: 'read_range', args: { range: 'Sheet1!A1:A2' } }] },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A3',
              cells: [[{ formula: '=SUM(A1:A2)' }]],
              reason: 'total of A1:A2',
            },
          },
        ],
      },
      { text: 'I staged a SUM formula in A3.' },
    ]);

    const { events, changeSets, stagedId } = await runScript(wb, provider);

    expect(stagedId).toBeDefined();
    expect(wb.getValue('Sheet1!A3')).toBeNull(); // nothing applied yet

    const staged = events.find((e) => e.kind === 'changeset_staged');
    expect(staged).toBeDefined();

    const result = await changeSets.apply(stagedId!);
    expect(result.ok).toBe(true);
    expect(wb.getValue('Sheet1!A3')).toBe(12); // formula actually recalculates

    // The model saw real read results.
    const readResult = provider.requests.at(2)?.messages.findLast((m) => m.role === 'tool');
    expect(readResult && 'content' in readResult ? readResult.content : '').toContain('5');
  });

  it('feeds schema violations back for repair instead of executing them', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      // missing required "reason", wrong cells shape
      { toolCalls: [{ name: 'write_range', args: { range: 'Sheet1!A1', cells: 'oops' } }] },
      // repaired call (plain scalars are values; objects are formulas)
      {
        toolCalls: [
          {
            name: 'write_range',
            args: { range: 'Sheet1!A1', cells: [[1]], reason: 'fixed' },
          },
        ],
      },
      { text: 'done' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);

    // The invalid call produced an error tool-result that the model saw next turn.
    const repairInput = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(repairInput && 'content' in repairInput ? repairInput.content : '').toContain(
      'Invalid arguments',
    );

    // Only the valid op was staged.
    expect(stagedId).toBeDefined();
    expect(changeSets.get(stagedId!)?.ops).toHaveLength(1);
  });

  it('rejects unknown tools and malformed JSON gracefully', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      { toolCalls: [{ name: 'delete_everything', args: {} }] },
      { toolCalls: [{ name: 'read_range', args: '{not json' }] },
      { text: 'giving up politely' },
    ]);

    const { stagedId, events } = await runScript(wb, provider);
    expect(stagedId).toBeUndefined();
    expect(events.some((e) => e.kind === 'changeset_staged')).toBe(false);

    const unknownToolReply = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(unknownToolReply && 'content' in unknownToolReply ? unknownToolReply.content : '').toContain(
      'Unknown tool',
    );
  });

  it('stages formatting alongside values and applies both', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: { range: 'Sheet1!B1', cells: [[1234.5]], reason: 'value' },
          },
          {
            name: 'format_range',
            args: {
              range: 'Sheet1!B1',
              format: { bold: true, numberFormat: '#,##0.00', fillColor: '#FFF2CC' },
              reason: 'highlight the total',
            },
          },
        ],
      },
      { text: 'Wrote and formatted B1.' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);
    expect(stagedId).toBeDefined();
    const cs = changeSets.get(stagedId!)!;
    expect(cs.ops).toHaveLength(2);
    expect(cs.undoFidelity).toBe('partial'); // formatting is not snapshot-restorable yet

    const result = await changeSets.apply(stagedId!);
    expect(result.ok).toBe(true);
    expect(wb.getValue('Sheet1!B1')).toBe(1234.5);
    expect(wb.getFormat('Sheet1!B1')).toMatchObject({
      bold: true,
      numberFormat: '#,##0.00',
      fillColor: '#FFF2CC',
    });
  });

  it('rejects malformed format payloads', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'format_range',
            args: { range: 'Sheet1!B1', format: {}, reason: 'empty format' },
          },
        ],
      },
      { text: 'ok' },
    ]);
    const { stagedId } = await runScript(wb, provider);
    expect(stagedId).toBeUndefined(); // empty format object fails validation
  });

  it('supports emulated tool calling for models without native tools', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'hello');

    class NoToolsProvider extends ScriptedProvider {
      override readonly supportsTools = false;
    }
    const provider = new NoToolsProvider([
      { text: '{"tool": "read_range", "args": {"range": "Sheet1!A1"}}' },
      { text: 'The cell says hello.' },
    ]);

    const { events } = await runScript(wb, provider);
    expect(events.some((e) => e.kind === 'tool_use' && e.tool === 'read_range')).toBe(true);
    expect(events.some((e) => e.kind === 'text' && e.text.includes('hello'))).toBe(true);
  });

  it('stops at the iteration bound even if the model loops forever', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const looping = Array.from({ length: 50 }, () => ({
      toolCalls: [{ name: 'get_workbook_map', args: {} }],
    }));
    const provider = new ScriptedProvider(looping);

    const changeSets = new ChangeSetManager(wb, 'wb-test');
    await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Test.xlsx',
      history: [],
      userMessage: 'loop',
      onEvent: () => {},
      maxIterations: 5,
    });
    expect(provider.requests.length).toBe(5);
  });
});
