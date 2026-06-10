/**
 * Cell values and A1-style addressing, shared by the Office.js bridge,
 * the sidecar, and the headless emulator.
 *
 * All formulas in the protocol are canonical en-US ("formulas" in Office.js,
 * never "formulasLocal"), so parsing is locale-independent.
 */

export type CellScalar = string | number | boolean | null;

/** A cell as transported over the protocol. */
export interface CellData {
  /** Computed value. Errors are encoded as strings beginning with "#". */
  value: CellScalar;
  /** Canonical en-US formula including leading "=", if the cell has one. */
  formula?: string;
}

export interface RangeAddress {
  sheet: string;
  /** 0-based inclusive */
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

const COL_RE = /^[A-Z]{1,3}$/;

export function colToIndex(col: string): number {
  if (!COL_RE.test(col)) throw new Error(`Invalid column: ${col}`);
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function indexToCol(index: number): string {
  if (index < 0) throw new Error(`Invalid column index: ${index}`);
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const A1_RE =
  /^(?:'([^']+)'|([A-Za-z0-9_]+))!(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$/;
const BARE_A1_RE = /^(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$/;

/**
 * Parse an A1 reference such as "Sheet1!A1", "'My Sheet'!B2:C10" or, with a
 * default sheet, "A1:B2". Returns a normalized RangeAddress (top-left to
 * bottom-right).
 */
export function parseA1(ref: string, defaultSheet?: string): RangeAddress {
  const trimmed = ref.trim();
  let sheet: string | undefined;
  let m = A1_RE.exec(trimmed);
  let groups: (string | undefined)[];
  if (m) {
    sheet = m[1] ?? m[2];
    groups = [m[4], m[6], m[8], m[10]];
  } else {
    const bare = BARE_A1_RE.exec(trimmed.toUpperCase());
    if (!bare) throw new Error(`Invalid A1 reference: ${ref}`);
    sheet = defaultSheet;
    groups = [bare[2], bare[4], bare[6], bare[8]];
  }
  if (!sheet) throw new Error(`Reference is missing a sheet name: ${ref}`);
  const [c1, r1, c2, r2] = groups;
  if (c1 === undefined || r1 === undefined) throw new Error(`Invalid A1 reference: ${ref}`);
  const startCol = colToIndex(c1.toUpperCase());
  const startRow = parseInt(r1, 10) - 1;
  const endCol = c2 !== undefined ? colToIndex(c2.toUpperCase()) : startCol;
  const endRow = r2 !== undefined ? parseInt(r2, 10) - 1 : startRow;
  if (startRow < 0 || endRow < 0) throw new Error(`Invalid row in reference: ${ref}`);
  return {
    sheet,
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

export function formatA1(addr: RangeAddress): string {
  const sheet = /^[A-Za-z0-9_]+$/.test(addr.sheet) ? addr.sheet : `'${addr.sheet}'`;
  const start = `${indexToCol(addr.startCol)}${addr.startRow + 1}`;
  if (addr.startRow === addr.endRow && addr.startCol === addr.endCol) {
    return `${sheet}!${start}`;
  }
  return `${sheet}!${start}:${indexToCol(addr.endCol)}${addr.endRow + 1}`;
}

export function rangeWidth(addr: RangeAddress): number {
  return addr.endCol - addr.startCol + 1;
}

export function rangeHeight(addr: RangeAddress): number {
  return addr.endRow - addr.startRow + 1;
}

export function rangeCellCount(addr: RangeAddress): number {
  return rangeWidth(addr) * rangeHeight(addr);
}

export function rangesOverlap(a: RangeAddress, b: RangeAddress): boolean {
  return (
    a.sheet === b.sheet &&
    a.startRow <= b.endRow &&
    b.startRow <= a.endRow &&
    a.startCol <= b.endCol &&
    b.startCol <= a.endCol
  );
}
