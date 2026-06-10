import { describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { extractReferences } from '../src/workbook/formula.js';

function emu(): WorkbookEmulator {
  return new WorkbookEmulator(['Sheet1', 'Data']);
}

describe('formula evaluator', () => {
  it('evaluates arithmetic and precedence', () => {
    const wb = emu();
    wb.setCell('Sheet1!A1', { formula: '=1+2*3' });
    wb.setCell('Sheet1!A2', { formula: '=(1+2)*3' });
    wb.setCell('Sheet1!A3', { formula: '=2^3+1' });
    wb.setCell('Sheet1!A4', { formula: '=-5+10' });
    expect(wb.getValue('Sheet1!A1')).toBe(7);
    expect(wb.getValue('Sheet1!A2')).toBe(9);
    expect(wb.getValue('Sheet1!A3')).toBe(9);
    expect(wb.getValue('Sheet1!A4')).toBe(5);
  });

  it('resolves cell references including cross-sheet and absolute', () => {
    const wb = emu();
    wb.setCell('Data!B2', 42);
    wb.setCell('Sheet1!A1', { formula: '=Data!B2*2' });
    wb.setCell('Sheet1!A2', { formula: '=$A$1+1' });
    expect(wb.getValue('Sheet1!A1')).toBe(84);
    expect(wb.getValue('Sheet1!A2')).toBe(85);
  });

  it('supports SUM/AVERAGE/COUNT over ranges', () => {
    const wb = emu();
    wb.setCell('Sheet1!A1', 1);
    wb.setCell('Sheet1!A2', 2);
    wb.setCell('Sheet1!A3', 3);
    wb.setCell('Sheet1!A4', 'text'); // ignored by numeric aggregates
    wb.setCell('Sheet1!B1', { formula: '=SUM(A1:A4)' });
    wb.setCell('Sheet1!B2', { formula: '=AVERAGE(A1:A3)' });
    wb.setCell('Sheet1!B3', { formula: '=COUNT(A1:A4)' });
    wb.setCell('Sheet1!B4', { formula: '=COUNTA(A1:A4)' });
    expect(wb.getValue('Sheet1!B1')).toBe(6);
    expect(wb.getValue('Sheet1!B2')).toBe(2);
    expect(wb.getValue('Sheet1!B3')).toBe(3);
    expect(wb.getValue('Sheet1!B4')).toBe(4);
  });

  it('handles IF, comparisons, string concat and text functions', () => {
    const wb = emu();
    wb.setCell('Sheet1!A1', 10);
    wb.setCell('Sheet1!B1', { formula: '=IF(A1>5,"big","small")' });
    wb.setCell('Sheet1!B2', { formula: '="x"&"y"' });
    wb.setCell('Sheet1!B3', { formula: '=UPPER(TRIM("  hi  "))' });
    wb.setCell('Sheet1!B4', { formula: '=LEN(B2)' });
    wb.setCell('Sheet1!B5', { formula: '=CONCATENATE("a",1,TRUE)' });
    expect(wb.getValue('Sheet1!B1')).toBe('big');
    expect(wb.getValue('Sheet1!B2')).toBe('xy');
    expect(wb.getValue('Sheet1!B3')).toBe('HI');
    expect(wb.getValue('Sheet1!B4')).toBe(2);
    expect(wb.getValue('Sheet1!B5')).toBe('a1TRUE');
  });

  it('produces Excel-style errors and propagates them', () => {
    const wb = emu();
    wb.setCell('Sheet1!A1', { formula: '=1/0' });
    wb.setCell('Sheet1!A2', { formula: '=A1+1' });
    wb.setCell('Sheet1!A3', { formula: '=NOSUCHFN(1)' });
    wb.setCell('Sheet1!A4', { formula: '=IFERROR(1/0,"fallback")' });
    wb.setCell('Sheet1!A5', { formula: '=Missing!A1' });
    expect(wb.getValue('Sheet1!A1')).toBe('#DIV/0!');
    expect(wb.getValue('Sheet1!A2')).toBe('#DIV/0!');
    expect(wb.getValue('Sheet1!A3')).toBe('#NAME?');
    expect(wb.getValue('Sheet1!A4')).toBe('fallback');
    expect(wb.getValue('Sheet1!A5')).toBe('#REF!');
  });

  it('detects circular references instead of hanging', () => {
    const wb = emu();
    wb.setCell('Sheet1!A1', { formula: '=A2+1' });
    wb.setCell('Sheet1!A2', { formula: '=A1+1' });
    expect(wb.getValue('Sheet1!A1')).toBe('#CYCLE!');
  });

  it('extracts references for dependency tracking', () => {
    const refs = extractReferences('=SUM(A1:B2)+Data!C3*2', 'Sheet1');
    expect(refs).toHaveLength(2);
    expect(refs[0]).toMatchObject({ sheet: 'Sheet1', startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
    expect(refs[1]).toMatchObject({ sheet: 'Data', startRow: 2, startCol: 2 });
  });
});
