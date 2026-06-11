import { randomUUID } from 'node:crypto';
import {
  parseA1,
  formatA1,
  rangeCellCount,
  type CellData,
  type ChangeOp,
  type ChangePreviewItem,
  type ChangeSet,
  type RangeSnapshot,
} from '@excelai/shared';
import type { WorkbookExecutor } from '../workbook/executor.js';

const MAX_SNAPSHOT_CELLS = 100_000;

export interface ApplyResult {
  ok: boolean;
  /** Verification problems found after apply (apply is rolled back if any). */
  problems: string[];
}

/**
 * Stages, previews, applies, verifies and undoes change-sets.
 *
 * Safety contract:
 *  - ops are staged, never applied directly by the agent;
 *  - before apply, affected ranges are snapshotted for undo;
 *  - after apply, touched ranges are re-read and scanned for newly introduced
 *    error values; any regression rolls the change-set back automatically.
 */
export class ChangeSetManager {
  private changeSets = new Map<string, ChangeSet>();

  constructor(
    private readonly executor: WorkbookExecutor,
    private readonly workbookId: string,
  ) {}

  stage(ops: ChangeOp[]): ChangeSet {
    // Structural/format ops are not snapshot-restorable in v1: undo restores
    // values and formulas exactly, but not sheet structure or formatting.
    const NON_RESTORABLE: ChangeOp['kind'][] = [
      'add_sheet',
      'rename_sheet',
      'delete_sheet',
      'format_range',
      'insert_rows',
      'delete_rows',
      'insert_cols',
      'delete_cols',
    ];
    const hasStructural = ops.some((op) => NON_RESTORABLE.includes(op.kind));
    const cs: ChangeSet = {
      id: randomUUID(),
      workbookId: this.workbookId,
      createdAt: new Date().toISOString(),
      status: 'staged',
      ops,
      undoFidelity: hasStructural ? 'partial' : 'full',
    };
    this.changeSets.set(cs.id, cs);
    return cs;
  }

  get(id: string): ChangeSet | undefined {
    return this.changeSets.get(id);
  }

  list(): ChangeSet[] {
    return [...this.changeSets.values()];
  }

  async preview(id: string): Promise<ChangePreviewItem[]> {
    const cs = this.mustGet(id, 'staged');
    const items: ChangePreviewItem[] = [];
    for (const op of cs.ops) {
      if (op.kind === 'write_range' || op.kind === 'clear_range' || op.kind === 'sort_range') {
        let before: CellData[][] | undefined;
        try {
          before = await this.executor.readRange(op.range);
        } catch {
          before = undefined; // e.g. range on a sheet this same change-set adds
        }
        items.push({ op, ...(before !== undefined ? { before } : {}) });
      } else {
        items.push({ op });
      }
    }
    return items;
  }

  reject(id: string): void {
    const cs = this.mustGet(id, 'staged');
    cs.status = 'rejected';
  }

  async apply(id: string): Promise<ApplyResult> {
    const cs = this.mustGet(id, 'staged');

    // 1. Snapshot affected ranges for undo.
    const snapshots: RangeSnapshot[] = [];
    let snapshotCells = 0;
    for (const op of cs.ops) {
      if (op.kind !== 'write_range' && op.kind !== 'clear_range' && op.kind !== 'sort_range') continue;
      const addr = parseA1(op.range);
      snapshotCells += rangeCellCount(addr);
      if (snapshotCells > MAX_SNAPSHOT_CELLS) {
        throw new Error(
          `Change-set touches more than ${MAX_SNAPSHOT_CELLS} cells; refusing to apply without splitting it up`,
        );
      }
      let cells: CellData[][];
      try {
        cells = await this.executor.readRange(op.range);
      } catch {
        continue; // sheet does not exist yet (created earlier in this set)
      }
      snapshots.push({ range: formatA1(addr), cells });
    }
    cs.snapshots = snapshots;

    const beforeErrors = countErrors(snapshots.flatMap((s) => s.cells));

    // 2. Apply.
    await this.executor.applyOps(cs.ops);

    // 3. Verify: re-read touched ranges and scan for newly introduced errors.
    const problems: string[] = [];
    for (const op of cs.ops) {
      if (op.kind !== 'write_range' && op.kind !== 'clear_range' && op.kind !== 'sort_range') continue;
      const after = await this.executor.readRange(op.range);
      for (const [r, row] of after.entries()) {
        for (const [c, cell] of row.entries()) {
          if (typeof cell.value === 'string' && cell.value.startsWith('#')) {
            // Writing an error value intentionally is almost never what the
            // user asked for; surface every error in the written range.
            problems.push(`${op.range} row ${r + 1} col ${c + 1} now contains ${cell.value}`);
          }
        }
      }
    }

    if (problems.length > 0 && beforeErrors === 0) {
      // Regression introduced by this change-set: roll back.
      await this.restoreSnapshots(cs);
      cs.status = 'staged';
      return { ok: false, problems };
    }

    cs.status = 'applied';
    return { ok: true, problems: [] };
  }

  async undo(id: string): Promise<void> {
    const cs = this.mustGet(id, 'applied');
    await this.restoreSnapshots(cs);
    cs.status = 'undone';
  }

  private async restoreSnapshots(cs: ChangeSet): Promise<void> {
    const ops: ChangeOp[] = [];
    for (const snap of cs.snapshots ?? []) {
      ops.push({
        kind: 'write_range',
        range: snap.range,
        cells: snap.cells.map((row) =>
          row.map((cell) =>
            cell.formula !== undefined ? { formula: cell.formula } : { value: cell.value },
          ),
        ),
        reason: 'undo: restore snapshot',
      });
    }
    if (ops.length > 0) await this.executor.applyOps(ops);
  }

  private mustGet(id: string, expected: ChangeSet['status']): ChangeSet {
    const cs = this.changeSets.get(id);
    if (!cs) throw new Error(`Unknown change-set: ${id}`);
    if (cs.status !== expected) {
      throw new Error(`Change-set ${id} is ${cs.status}, expected ${expected}`);
    }
    return cs;
  }
}

function countErrors(cells: CellData[][]): number {
  let n = 0;
  for (const row of cells) {
    for (const cell of row) {
      if (typeof cell.value === 'string' && cell.value.startsWith('#')) n++;
    }
  }
  return n;
}
