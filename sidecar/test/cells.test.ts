import { describe, expect, it } from 'vitest';
import { colToIndex, indexToCol, parseA1, formatA1 } from '@excelai/shared';

describe('A1 addressing', () => {
  it('converts columns both ways', () => {
    expect(colToIndex('A')).toBe(0);
    expect(colToIndex('Z')).toBe(25);
    expect(colToIndex('AA')).toBe(26);
    expect(colToIndex('AZ')).toBe(51);
    for (const i of [0, 25, 26, 51, 701, 702]) {
      expect(colToIndex(indexToCol(i))).toBe(i);
    }
  });

  it('parses sheet-qualified references', () => {
    expect(parseA1('Sheet1!B2')).toEqual({
      sheet: 'Sheet1',
      startRow: 1,
      startCol: 1,
      endRow: 1,
      endCol: 1,
    });
    expect(parseA1("'My Sheet'!A1:C10")).toMatchObject({
      sheet: 'My Sheet',
      startRow: 0,
      startCol: 0,
      endRow: 9,
      endCol: 2,
    });
  });

  it('normalizes reversed ranges and applies default sheet', () => {
    expect(parseA1('C10:A1', 'S')).toMatchObject({ sheet: 'S', startRow: 0, endRow: 9 });
  });

  it('round-trips through formatA1', () => {
    for (const ref of ['Sheet1!A1', 'Sheet1!A1:C10', "'My Sheet'!B2:D4"]) {
      expect(formatA1(parseA1(ref))).toBe(ref);
    }
  });

  it('rejects invalid references', () => {
    expect(() => parseA1('NotARef!!')).toThrow();
    expect(() => parseA1('A1')).toThrow(); // no default sheet provided
  });
});
