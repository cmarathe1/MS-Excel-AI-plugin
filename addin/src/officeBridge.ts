import { indexToCol, type BorderFormat, type CellData, type ChangeOp, type WorkbookMap } from '@excelai/shared';

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
        case 'rename_sheet': {
          context.workbook.worksheets.getItem(op.name).name = op.newName;
          break;
        }
        case 'delete_sheet': {
          context.workbook.worksheets.getItem(op.name).delete();
          break;
        }
        case 'insert_rows': {
          // Row addresses are 1-based in A1 notation; op.at is 0-based.
          const sheet = context.workbook.worksheets.getItem(op.sheet);
          sheet.getRange(`${op.at + 1}:${op.at + op.count}`).insert(Excel.InsertShiftDirection.down);
          break;
        }
        case 'delete_rows': {
          const sheet = context.workbook.worksheets.getItem(op.sheet);
          sheet.getRange(`${op.at + 1}:${op.at + op.count}`).delete(Excel.DeleteShiftDirection.up);
          break;
        }
        case 'insert_cols': {
          const sheet = context.workbook.worksheets.getItem(op.sheet);
          sheet
            .getRange(`${indexToCol(op.at)}:${indexToCol(op.at + op.count - 1)}`)
            .insert(Excel.InsertShiftDirection.right);
          break;
        }
        case 'delete_cols': {
          const sheet = context.workbook.worksheets.getItem(op.sheet);
          sheet
            .getRange(`${indexToCol(op.at)}:${indexToCol(op.at + op.count - 1)}`)
            .delete(Excel.DeleteShiftDirection.left);
          break;
        }
        case 'sort_range': {
          const range = getRange(context, op.range);
          range.sort.apply(
            [{ key: op.keyOffset, ascending: op.ascending }],
            false,
            op.hasHeader,
            Excel.SortOrientation.rows,
          );
          break;
        }
        case 'format_range': {
          const range = getRange(context, op.range);
          const f = op.format;
          if (f.numberFormat !== undefined) {
            // numberFormat must be a matrix matching the range dimensions.
            range.load(['rowCount', 'columnCount']);
            await context.sync();
            range.numberFormat = Array.from({ length: range.rowCount }, () =>
              Array.from({ length: range.columnCount }, () => f.numberFormat!),
            );
          }
          if (f.bold !== undefined) range.format.font.bold = f.bold;
          if (f.italic !== undefined) range.format.font.italic = f.italic;
          if (f.fontColor !== undefined) range.format.font.color = f.fontColor;
          if (f.fillColor !== undefined) range.format.fill.color = f.fillColor;
          if (f.horizontalAlignment !== undefined) {
            range.format.horizontalAlignment = {
              left: Excel.HorizontalAlignment.left,
              center: Excel.HorizontalAlignment.center,
              right: Excel.HorizontalAlignment.right,
            }[f.horizontalAlignment];
          }
          if (f.fontSize !== undefined) range.format.font.size = f.fontSize;
          if (f.wrapText !== undefined) range.format.wrapText = f.wrapText;
          if (f.border !== undefined) applyBorders(range, f.border);
          if (f.autofitColumns) range.format.autofitColumns();
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

const EDGE_MAP: Record<string, Excel.BorderIndex[]> = {
  top: [Excel.BorderIndex.edgeTop],
  bottom: [Excel.BorderIndex.edgeBottom],
  left: [Excel.BorderIndex.edgeLeft],
  right: [Excel.BorderIndex.edgeRight],
  insideHorizontal: [Excel.BorderIndex.insideHorizontal],
  insideVertical: [Excel.BorderIndex.insideVertical],
  outline: [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
  ],
  all: [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
    Excel.BorderIndex.insideHorizontal,
    Excel.BorderIndex.insideVertical,
  ],
};

function applyBorders(range: Excel.Range, border: BorderFormat): void {
  const style =
    border.style === 'double' ? Excel.BorderLineStyle.double : Excel.BorderLineStyle.continuous;
  const weight =
    border.style === 'thick'
      ? Excel.BorderWeight.thick
      : border.style === 'medium'
        ? Excel.BorderWeight.medium
        : Excel.BorderWeight.thin;
  const indexes = new Set<Excel.BorderIndex>();
  for (const edge of border.edges) for (const idx of EDGE_MAP[edge] ?? []) indexes.add(idx);
  for (const idx of indexes) {
    const b = range.format.borders.getItem(idx);
    b.style = style;
    b.weight = weight;
    if (border.color) b.color = border.color;
  }
}

/** Case-insensitive substring search across sheets, capped at 50 matches. */
export async function findCells(
  query: string,
  sheetName?: string,
): Promise<{ address: string; value: string }[]> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const targets = sheets.items.filter((ws) => !sheetName || ws.name === sheetName);
    if (sheetName && targets.length === 0) throw new Error(`No such sheet: ${sheetName}`);

    const matches: { address: string; value: string }[] = [];
    for (const ws of targets) {
      const found = ws.findAllOrNullObject(query, { completeMatch: false, matchCase: false });
      found.load(['address', 'isNullObject']);
      await context.sync();
      if (found.isNullObject) continue;
      for (const address of found.address.split(',')) {
        // Expand multi-cell areas conservatively: report the area address.
        const range = ws.getRange(address);
        range.load(['values', 'address']);
        await context.sync();
        const flat = range.values.flat();
        matches.push({ address: range.address, value: String(flat[0] ?? '') });
        if (matches.length >= 50) return matches;
      }
    }
    return matches;
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
