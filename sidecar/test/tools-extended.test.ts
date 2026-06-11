import { describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { runAgentTurn } from '../src/agent/loop.js';

async function stageAndApply(wb: WorkbookEmulator, toolCalls: { name: string; args: unknown }[]) {
  const changeSets = new ChangeSetManager(wb, 'wb');
  const provider = new ScriptedProvider([{ toolCalls }, { text: 'done' }]);
  const result = await runAgentTurn({
    provider,
    executor: wb,
    changeSets,
    workbookName: 'T.xlsx',
    history: [],
    userMessage: 'go',
    onEvent: () => {},
  });
  expect(result.stagedChangeSetId).toBeDefined();
  const apply = await changeSets.apply(result.stagedChangeSetId!);
  expect(apply.ok).toBe(true);
  return changeSets.get(result.stagedChangeSetId!)!;
}

describe('structural operations', () => {
  it('inserts and deletes rows with cell shifting', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'header');
    wb.setCell('Sheet1!A2', 'row2');
    wb.setCell('Sheet1!A3', 'row3');

    await stageAndApply(wb, [
      { name: 'insert_rows', args: { sheet: 'Sheet1', at: 2, count: 2, reason: 'space' } },
    ]);
    expect(wb.getValue('Sheet1!A1')).toBe('header');
    expect(wb.getValue('Sheet1!A2')).toBeNull();
    expect(wb.getValue('Sheet1!A4')).toBe('row2');
    expect(wb.getValue('Sheet1!A5')).toBe('row3');

    await stageAndApply(wb, [
      { name: 'delete_rows', args: { sheet: 'Sheet1', at: 2, count: 2, reason: 'undo space' } },
    ]);
    expect(wb.getValue('Sheet1!A2')).toBe('row2');
    expect(wb.getValue('Sheet1!A3')).toBe('row3');
  });

  it('inserts and deletes columns by letter', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'a');
    wb.setCell('Sheet1!B1', 'b');

    await stageAndApply(wb, [
      { name: 'insert_cols', args: { sheet: 'Sheet1', at: 'B', count: 1, reason: 'gap' } },
    ]);
    expect(wb.getValue('Sheet1!B1')).toBeNull();
    expect(wb.getValue('Sheet1!C1')).toBe('b');

    await stageAndApply(wb, [
      { name: 'delete_cols', args: { sheet: 'Sheet1', at: 'B', count: 1, reason: 'close gap' } },
    ]);
    expect(wb.getValue('Sheet1!B1')).toBe('b');
  });

  it('renames and deletes sheets with guards', async () => {
    const wb = new WorkbookEmulator(['Sheet1', 'Temp']);
    wb.setCell('Temp!A1', 'x');
    await stageAndApply(wb, [
      { name: 'rename_sheet', args: { name: 'Temp', newName: 'Data', reason: 'clearer' } },
    ]);
    expect(wb.getValue('Data!A1')).toBe('x');

    await stageAndApply(wb, [
      { name: 'delete_sheet', args: { name: 'Data', reason: 'no longer needed' } },
    ]);
    await expect(wb.readRange('Data!A1')).rejects.toThrow(/No such sheet/);

    // cannot delete the only remaining sheet
    await expect(
      wb.applyOps([{ kind: 'delete_sheet', name: 'Sheet1', reason: 'x' }]),
    ).rejects.toThrow(/only sheet/);
  });
});

describe('sorting', () => {
  it('sorts rows by a key column, keeps headers, blanks last, and is undoable', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const rows: [string, number | null][] = [
      ['Name', 0], // header (B header is numeric 0 just to fill)
      ['carol', 3],
      ['alice', 1],
      ['', null],
      ['bob', 2],
    ];
    rows.forEach(([a, b], i) => {
      wb.setCell(`Sheet1!A${i + 1}`, a);
      wb.setCell(`Sheet1!B${i + 1}`, b);
    });

    const cs = await stageAndApply(wb, [
      {
        name: 'sort_range',
        args: {
          range: 'Sheet1!A1:B5',
          keyColumn: 'A',
          ascending: true,
          hasHeader: true,
          reason: 'alphabetical',
        },
      },
    ]);
    expect(cs.undoFidelity).toBe('full'); // sort is snapshot-restorable

    expect(wb.getValue('Sheet1!A1')).toBe('Name'); // header pinned
    expect(wb.getValue('Sheet1!A2')).toBe('alice');
    expect(wb.getValue('Sheet1!A3')).toBe('bob');
    expect(wb.getValue('Sheet1!A4')).toBe('carol');
    expect(wb.getValue('Sheet1!B4')).toBe(3); // row moved together
    expect(wb.getValue('Sheet1!A5')).toBe(''); // blank key last
  });

  it('rejects a key column outside the range', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 1);
    const changeSets = new ChangeSetManager(wb, 'wb');
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'sort_range',
            args: { range: 'Sheet1!A1:B2', keyColumn: 'Z', ascending: true, hasHeader: false, reason: 'x' },
          },
        ],
      },
      { text: 'done' },
    ]);
    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'T.xlsx',
      history: [],
      userMessage: 'go',
      onEvent: () => {},
    });
    expect(result.stagedChangeSetId).toBeUndefined();
    // The model received the validation error and could repair.
    const reply = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(reply && 'content' in reply ? reply.content : '').toContain('outside the range');
  });
});

describe('borders and extended formatting', () => {
  it('stages and applies borders, font size and wrap', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'Title');
    await stageAndApply(wb, [
      {
        name: 'format_range',
        args: {
          range: 'Sheet1!A1:B2',
          format: {
            border: { edges: ['all'], style: 'thin', color: '#333333' },
            fontSize: 14,
            wrapText: true,
          },
          reason: 'table borders',
        },
      },
    ]);
    expect(wb.getFormat('Sheet1!A1')).toMatchObject({
      border: { edges: ['all'], style: 'thin', color: '#333333' },
      fontSize: 14,
      wrapText: true,
    });
    expect(wb.getFormat('Sheet1!B2')).toMatchObject({ fontSize: 14 });
  });
});

describe('find tool', () => {
  it('locates cells by substring across sheets, case-insensitive', async () => {
    const wb = new WorkbookEmulator(['Sheet1', 'Data']);
    wb.setCell('Sheet1!A1', 'Total Revenue');
    wb.setCell('Data!C3', 'revenue stream');
    wb.setCell('Data!D4', 12345);

    const all = await wb.find('REVENUE');
    expect(all.map((m) => m.address)).toEqual(['Sheet1!A1', 'Data!C3']);

    const scoped = await wb.find('revenue', 'Data');
    expect(scoped).toHaveLength(1);

    const numeric = await wb.find('234');
    expect(numeric[0]?.address).toBe('Data!D4');
  });

  it('is exposed to the agent as a read-only tool', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!B7', 'needle');
    const changeSets = new ChangeSetManager(wb, 'wb');
    const provider = new ScriptedProvider([
      { toolCalls: [{ name: 'find', args: { query: 'needle' } }] },
      { text: 'found it' },
    ]);
    await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'T.xlsx',
      history: [],
      userMessage: 'where is needle?',
      onEvent: () => {},
    });
    const reply = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(reply && 'content' in reply ? reply.content : '').toContain('Sheet1!B7');
  });
});
