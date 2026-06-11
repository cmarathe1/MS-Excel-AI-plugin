import type { CellData, CellScalar } from './cells.js';

/**
 * Change-sets: every model-initiated mutation is staged, previewed,
 * explicitly approved, applied, and undoable.
 */

export type BorderEdge =
  | 'all'
  | 'outline'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'insideHorizontal'
  | 'insideVertical';

export interface BorderFormat {
  edges: BorderEdge[];
  style?: 'thin' | 'medium' | 'thick' | 'double';
  color?: string;
}

export interface CellFormat {
  numberFormat?: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: 'left' | 'center' | 'right';
  wrapText?: boolean;
  border?: BorderFormat;
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
  | { kind: 'add_sheet'; name: string; reason: string }
  | { kind: 'rename_sheet'; name: string; newName: string; reason: string }
  | { kind: 'delete_sheet'; name: string; reason: string }
  /** `at` is 0-based; rows/cols shift like in Excel. */
  | { kind: 'insert_rows'; sheet: string; at: number; count: number; reason: string }
  | { kind: 'delete_rows'; sheet: string; at: number; count: number; reason: string }
  | { kind: 'insert_cols'; sheet: string; at: number; count: number; reason: string }
  | { kind: 'delete_cols'; sheet: string; at: number; count: number; reason: string }
  /** keyOffset is the 0-based column offset of the sort key within the range. */
  | {
      kind: 'sort_range';
      range: string;
      keyOffset: number;
      ascending: boolean;
      hasHeader: boolean;
      reason: string;
    };

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
