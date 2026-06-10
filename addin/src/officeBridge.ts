import type { CellData, ChangeOp, WorkbookMap } from '@excelai/shared';

/**
 * The "hands": executes workbook tools against real Excel via Office.js.
 * Mirrors the sidecar's WorkbookEmulator behavior — the contract both
 * implementations share is defined by the shared tool schemas.
 */

export async function getWorkbookMap(): Promise<WorkbookMap> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();

    const usedRanges = sheets.items.map((ws) => {
      const used = ws.getUsedRangeOrNullObject();
      used.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
      return { name: ws.name, used };
    });
    await context.sync();

    return {
      sheets: usedRanges.map(({ name, used }) => ({
        name,
        rowCount: used.isNullObject ? 0 : used.rowCount,
        colCount: used.isNullObject ? 0 : used.columnCount,
        usedRange: used.isNullObject ? null : used.address,
      })),
    };
  });
}

export async function readRange(rangeRef: string): Promise<CellData[][]> {
  return Excel.run(async (context) => {
    const range = getRange(context, rangeRef);
    range.load(['values', 'formulas']);
    await context.sync();

    const out: CellData[][] = [];
    for (let r = 0; r < range.values.length; r++) {
      const row: CellData[] = [];
      for (let c = 0; c < range.values[r]!.length; c++) {
        const value = range.values[r]![c] as CellData['value'];
        const formula = range.formulas[r]![c];
        const isFormula = typeof formula === 'string' && formula.startsWith('=');
        row.push({ value, ...(isFormula ? { formula } : {}) });
      }
      out.push(row);
    }
    return out;
  });
}

export async function applyOps(ops: ChangeOp[]): Promise<void> {
  await Excel.run(async (context) => {
    for (const op of ops) {
      switch (op.kind) {
        case 'add_sheet': {
          context.workbook.worksheets.add(op.name);
          break;
        }
        case 'clear_range': {
          getRange(context, op.range).clear(Excel.ClearApplyTo.contents);
          break;
        }
        case 'write_range': {
          const range = getRange(context, op.range);
          // Setting `formulas` writes formulas for "=..." strings and treats
          // everything else as a literal value — one matrix handles both.
          const matrix = op.cells.map((row) =>
            row.map((cell) => (cell.formula !== undefined ? cell.formula : (cell.value ?? ''))),
          );
          range.formulas = matrix as unknown as string[][];
          break;
        }
      }
      // Apply ops in order so later ops can target sheets created earlier.
      await context.sync();
    }
  });
}

function getRange(context: Excel.RequestContext, ref: string): Excel.Range {
  const bang = ref.lastIndexOf('!');
  if (bang < 0) {
    return context.workbook.worksheets.getActiveWorksheet().getRange(ref);
  }
  const sheetName = ref.slice(0, bang).replace(/^'|'$/g, '');
  const address = ref.slice(bang + 1);
  return context.workbook.worksheets.getItem(sheetName).getRange(address);
}

/**
 * Stable per-workbook identity, stored in the document itself so memory
 * survives renames. Includes a random component generated on first use.
 */
export async function getWorkbookIdentity(): Promise<{ id: string; name: string }> {
  const settings = Office.context.document.settings;
  let id = settings.get('excelai-workbook-id') as string | null;
  if (!id) {
    id = crypto.randomUUID();
    settings.set('excelai-workbook-id', id);
    await new Promise<void>((resolve) => settings.saveAsync(() => resolve()));
  }
  const url = Office.context.document.url ?? 'Workbook';
  const name = url.split(/[\\/]/).pop() || 'Workbook';
  return { id, name };
}
