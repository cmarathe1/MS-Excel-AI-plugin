import type { CellData, CellScalar } from './cells.js';

/**
 * Change-sets: every model-initiated mutation is staged, previewed,
 * explicitly approved, applied, and undoable.
 */

export interface CellFormat {
  numberFormat?: string;
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  autofitColumns?: boolean;
}

export type ChangeOp =
  | {
      kind: 'write_range';
      range: string;
      /** Row-major; strings starting with "=" are formulas. */
      cells: { value?: CellScalar; formula?: string }[][];
      reason: string;
    }
  | { kind: 'clear_range'; range: string; reason: string }
  | { kind: 'format_range'; range: string; format: CellFormat; reason: string }
  | { kind: 'add_sheet'; name: string; reason: string };

export type ChangeSetStatus = 'staged' | 'applied' | 'rejected' | 'undone';

export interface ChangeSet {
  id: string;
  workbookId: string;
  createdAt: string;
  status: ChangeSetStatus;
  ops: ChangeOp[];
  /** Snapshot of affected ranges taken just before apply, for undo. */
  snapshots?: RangeSnapshot[];
  /**
   * Undo fidelity. "full" = all ops are range-level and snapshot-restorable.
   * "partial" = contains structural ops (e.g. add_sheet) that undo can remove
   * but whose side effects may not be perfectly reversible.
   */
  undoFidelity: 'full' | 'partial';
}

export interface RangeSnapshot {
  range: string;
  cells: CellData[][];
}

export interface ChangePreviewItem {
  op: ChangeOp;
  /** Current cells in the affected range (before), for diff rendering. */
  before?: CellData[][];
}
