import { describe, expect, it } from 'vitest';
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS } from '../models/workbook';
import { cellAddress, columnIndex, columnLabel, parseReference, rangeAddress } from './cellAddress';

describe('column labels', () => {
  it('converts between indexes and letters', () => {
    expect(columnLabel(0)).toBe('A');
    expect(columnLabel(25)).toBe('Z');
    expect(columnLabel(26)).toBe('AA');
    expect(columnLabel(701)).toBe('ZZ');
    expect(columnLabel(702)).toBe('AAA');
    expect(columnLabel(EXCEL_MAX_COLS - 1)).toBe('XFD');
    for (const index of [0, 25, 26, 51, 700, 16383]) expect(columnIndex(columnLabel(index))).toBe(index);
    expect(columnIndex('xfd')).toBe(16383);
    expect(columnIndex('XFE')).toBe(-1);
    expect(columnIndex('A1')).toBe(-1);
  });
});

describe('parseReference', () => {
  it('parses single cells in any case, with or without $', () => {
    expect(parseReference('A152')?.range).toEqual({ r0: 151, c0: 0, r1: 151, c1: 0 });
    expect(parseReference(' b7 ')?.range).toEqual({ r0: 6, c0: 1, r1: 6, c1: 1 });
    expect(parseReference('$C$3')?.range).toEqual({ r0: 2, c0: 2, r1: 2, c1: 2 });
  });

  it('parses ranges, whole columns and whole rows', () => {
    expect(parseReference('C10:A1')?.range).toEqual({ r0: 0, c0: 0, r1: 9, c1: 2 });
    expect(parseReference('B:D')?.range).toEqual({ r0: 0, c0: 1, r1: EXCEL_MAX_ROWS - 1, c1: 3 });
    expect(parseReference('3:7')?.range).toEqual({ r0: 2, c0: 0, r1: 6, c1: EXCEL_MAX_COLS - 1 });
  });

  it('parses sheet-qualified references', () => {
    expect(parseReference('Sheet2!A1')).toEqual({ sheetName: 'Sheet2', range: { r0: 0, c0: 0, r1: 0, c1: 0 } });
    expect(parseReference("'Q1 ''24'!B2")?.sheetName).toBe("Q1 '24");
  });

  it('rejects invalid input', () => {
    for (const input of ['', 'hello', 'A0', 'A1048577', 'XFE1', 'A1:B2:C3', '!A1', '1A']) {
      expect(parseReference(input)).toBeNull();
    }
  });
});

describe('addresses', () => {
  it('formats cells and ranges', () => {
    expect(cellAddress(151, 0)).toBe('A152');
    expect(rangeAddress({ r0: 0, c0: 0, r1: 0, c1: 0 })).toBe('A1');
    expect(rangeAddress({ r0: 0, c0: 0, r1: 4, c1: 2 })).toBe('A1:C5');
    expect(rangeAddress({ r0: 0, c0: 1, r1: EXCEL_MAX_ROWS - 1, c1: 1 })).toBe('B:B');
    expect(rangeAddress({ r0: 2, c0: 0, r1: 3, c1: EXCEL_MAX_COLS - 1 })).toBe('3:4');
  });
});
