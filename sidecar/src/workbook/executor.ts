import type { CellData, ChangeOp, WorkbookMap } from '@excelai/shared';

/**
 * The workbook executor: everything that touches a workbook goes through this
 * interface. Two implementations exist:
 *
 *  - the Office.js bridge (real Excel, executed in the add-in over WebSocket)
 *  - the headless emulator (tests, capability probe, verification sandbox)
 *
 * Contract tests run against both so the emulator cannot silently drift.
 */
export interface WorkbookExecutor {
  getWorkbookMap(): Promise<WorkbookMap>;
  readRange(range: string): Promise<CellData[][]>;
  /** Apply already-approved change ops, in order. Atomic per call where the host allows. */
  applyOps(ops: ChangeOp[]): Promise<void>;
}
