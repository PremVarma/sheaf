import { describe, expect, it } from 'vitest';
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS } from '../../models/workbook';
import { createSheet } from '../../test/fixtures';
import { buildCopyPayload } from './clipboardService';

const rows = [
  ['Name', 'Qty', 'Note'],
  ['Pens', 12, ''],
  ['Tab\there', 3, 'line1\nline2'],
  ['Say "hi"', null, 'x'],
];
const sheet = createSheet(4, 3, (r, c) => rows[r][c] as string | number | null);

describe('buildCopyPayload', () => {
  it('copies a single cell', () => {
    expect(buildCopyPayload(sheet, { r0: 1, c0: 0, r1: 1, c1: 0 })).toEqual({ text: 'Pens', rows: 1, columns: 1 });
  });

  it('copies a range as tab-separated rows, keeping empty cells', () => {
    const payload = buildCopyPayload(sheet, { r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(payload?.text).toBe('Name\tQty\tNote\nPens\t12\t');
  });

  it('quotes cells containing tabs, line breaks or quotes', () => {
    const payload = buildCopyPayload(sheet, { r0: 2, c0: 0, r1: 3, c1: 2 });
    expect(payload?.text).toBe('"Tab\there"\t3\t"line1\nline2"\n"Say ""hi"""\t\tx');
  });

  it('limits whole-column and select-all copies to the used range', () => {
    expect(buildCopyPayload(sheet, { r0: 0, c0: 1, r1: EXCEL_MAX_ROWS - 1, c1: 1 })?.text).toBe('Qty\n12\n3\n');
    const all = buildCopyPayload(sheet, { r0: 0, c0: 0, r1: EXCEL_MAX_ROWS - 1, c1: EXCEL_MAX_COLS - 1 });
    expect(all).toMatchObject({ rows: 4, columns: 3 });
  });

  it('refuses selections too large to copy', () => {
    expect(buildCopyPayload(sheet, { r0: 0, c0: 0, r1: 99_999, c1: 99 })).toBeNull();
  });
});
