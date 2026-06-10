import { describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';

describe('change-set lifecycle', () => {
  it('stages, previews, applies and undoes a write', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'old');
    const mgr = new ChangeSetManager(wb, 'wb-1');

    const cs = mgr.stage([
      { kind: 'write_range', range: 'Sheet1!A1', cells: [[{ value: 'new' }]], reason: 'test' },
    ]);
    expect(cs.status).toBe('staged');
    expect(cs.undoFidelity).toBe('full');

    const preview = await mgr.preview(cs.id);
    expect(preview[0]?.before?.[0]?.[0]?.value).toBe('old');
    expect(wb.getValue('Sheet1!A1')).toBe('old'); // staging never touches the sheet

    const result = await mgr.apply(cs.id);
    expect(result.ok).toBe(true);
    expect(wb.getValue('Sheet1!A1')).toBe('new');

    await mgr.undo(cs.id);
    expect(wb.getValue('Sheet1!A1')).toBe('old');
    expect(mgr.get(cs.id)?.status).toBe('undone');
  });

  it('restores formulas (not just values) on undo', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 2);
    wb.setCell('Sheet1!B1', { formula: '=A1*10' });
    const mgr = new ChangeSetManager(wb, 'wb-1');

    const cs = mgr.stage([
      { kind: 'write_range', range: 'Sheet1!B1', cells: [[{ value: 999 }]], reason: 'flatten' },
    ]);
    await mgr.apply(cs.id);
    expect(wb.getValue('Sheet1!B1')).toBe(999);

    await mgr.undo(cs.id);
    expect(wb.getValue('Sheet1!B1')).toBe(20);
    wb.setCell('Sheet1!A1', 5);
    expect(wb.getValue('Sheet1!B1')).toBe(50); // formula restored, still live
  });

  it('rolls back automatically when a change introduces formula errors', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 10);
    const mgr = new ChangeSetManager(wb, 'wb-1');

    const cs = mgr.stage([
      {
        kind: 'write_range',
        range: 'Sheet1!B1',
        cells: [[{ formula: '=A1/0' }]],
        reason: 'bad formula',
      },
    ]);
    const result = await mgr.apply(cs.id);
    expect(result.ok).toBe(false);
    expect(result.problems[0]).toContain('#DIV/0!');
    expect(wb.getValue('Sheet1!B1')).toBeNull(); // rolled back
    expect(mgr.get(cs.id)?.status).toBe('staged'); // back to staged for iteration
  });

  it('marks change-sets with structural ops as partial undo fidelity', () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = mgr.stage([{ kind: 'add_sheet', name: 'Report', reason: 'new output sheet' }]);
    expect(cs.undoFidelity).toBe('partial');
  });

  it('rejects mismatched cell matrix dimensions', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = mgr.stage([
      {
        kind: 'write_range',
        range: 'Sheet1!A1:B2',
        cells: [[{ value: 1 }]], // 1x1 into 2x2
        reason: 'bad dims',
      },
    ]);
    await expect(mgr.apply(cs.id)).rejects.toThrow(/does not match range/);
  });

  it('refuses to double-apply or undo a non-applied set', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = mgr.stage([
      { kind: 'write_range', range: 'Sheet1!A1', cells: [[{ value: 1 }]], reason: 'x' },
    ]);
    await expect(mgr.undo(cs.id)).rejects.toThrow();
    await mgr.apply(cs.id);
    await expect(mgr.apply(cs.id)).rejects.toThrow();
  });
});
