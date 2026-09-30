import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS } from '../../models/workbook';
import { buildWorkbookFile } from '../../test/fixtures';
import { parseWorkbookBytes } from '../workbook/parser';
import { getCell, sortRange } from '../workbook/sheetService';
import { readBlock } from './cells';
import { EditRefusedError, EditSession } from './editSession';
import { createBlankWorkbook } from './export';

const f = (v: number, formula: string) => ({ t: 'n', v, f: formula }) as XLSX.CellObject;

const invoice = () =>
  buildWorkbookFile({
    Invoice: [
      ['Item', 'Qty', 'Rate', 'Amount'],
      ['Pens', 10, 5, f(50, 'B2*C2')],
      ['Paper', 2, 100, f(200, 'B3*C3')],
      ['Total', null, null, f(250, 'SUM(D2:D3)')],
    ],
    Summary: [['Grand total', f(250, 'Invoice!D4')]],
  });

function open(bytes = invoice(), name = 'invoice.xlsx') {
  return new EditSession(parseWorkbookBytes(bytes, name), bytes);
}

const at = (session: EditSession, sheet: number, address: string) => {
  const { r, c } = XLSX.utils.decode_cell(address);
  return getCell(session.current.sheets[sheet], r, c);
};

const styleAt = (session: EditSession, sheet: number, address: string) => session.current.styles[at(session, sheet, address)?.styleId ?? 0];

const reopen = (session: EditSession, name = 'invoice.xlsx') => parseWorkbookBytes(session.buildFile('xlsx', 0), name);

const cellOf = (workbook: ReturnType<typeof parseWorkbookBytes>, sheet: number, address: string) => {
  const { r, c } = XLSX.utils.decode_cell(address);
  return getCell(workbook.sheets[sheet], r, c);
};

describe('formatting', () => {
  it('applies text formatting to a range and keeps it when saved', () => {
    const session = open();
    const base = session.current.defaultFont!.size; // SheetJS writes 12pt Calibri
    session.format(0, { r0: 0, c0: 0, r1: 0, c1: 3 }, { bold: true, fill: '#ffe699', fontScale: 14 / base });
    session.format(0, { r0: 1, c0: 3, r1: 3, c1: 3 }, { numberFormat: '#,##0.00', hAlign: 'center' });
    expect(styleAt(session, 0, 'A1')).toMatchObject({ bold: true, fill: '#ffe699' });
    expect(at(session, 0, 'D2')?.displayValue).toBe('50.00');

    const saved = reopen(session);
    const styles = saved.styles;
    expect(styles[cellOf(saved, 0, 'B1')!.styleId]).toMatchObject({ bold: true, fill: '#ffe699', fontScale: expect.closeTo(14 / base, 3) });
    expect(styles[cellOf(saved, 0, 'D3')!.styleId]).toMatchObject({ numberFormat: '#,##0.00', hAlign: 'center' });
    expect(cellOf(saved, 0, 'D3')?.displayValue).toBe('200.00');
  });

  it('makes rows taller for bigger or wrapped text', () => {
    const session = open();
    session.format(0, { r0: 0, c0: 0, r1: 0, c1: 0 }, { fontScale: 2 });
    const heights = session.current.sheets[0].rowHeights;
    expect([...heights.index]).toEqual([0]);
    expect(heights.size[0]).toBeGreaterThan(session.current.sheets[0].defaultRowHeight);
    session.undo();
    expect(session.current.sheets[0].rowHeights.index.length).toBe(0);
  });

  it('formats empty cells so borders and fills show', () => {
    const session = open(buildWorkbookFile({ S: [['a']] }), 's.xlsx');
    const edge = { width: 1 as const, style: 'solid' as const };
    session.format(0, { r0: 2, c0: 1, r1: 3, c1: 2 }, (row, col) => ({
      borders: { top: row === 2 ? edge : undefined, bottom: row === 3 ? edge : undefined, left: col === 1 ? edge : undefined, right: col === 2 ? edge : undefined },
    }));
    expect(styleAt(session, 0, 'B3')?.borders).toEqual({ top: edge, left: edge });
    expect(styleAt(session, 0, 'C4')?.borders).toEqual({ right: edge, bottom: edge });
    const saved = reopen(session, 's.xlsx');
    expect(saved.styles[cellOf(saved, 0, 'C4')!.styleId].borders).toEqual({ right: edge, bottom: edge });
  });

  it('clears formats and undoes formatting', () => {
    const session = open();
    session.format(0, { r0: 1, c0: 1, r1: 1, c1: 1 }, { italic: true, numberFormat: '0.0%' });
    expect(at(session, 0, 'B2')?.displayValue).toBe('1000.0%');
    session.undo();
    expect(styleAt(session, 0, 'B2')).toEqual({});
    expect(at(session, 0, 'B2')?.displayValue).toBe('10');
    session.redo();
    session.clearFormats(0, { r0: 0, c0: 0, r1: 3, c1: 3 });
    expect(at(session, 0, 'B2')?.displayValue).toBe('10');
    expect(styleAt(session, 0, 'B2')).toEqual({});
  });

  it('merges and centers, keeping only the top-left value', () => {
    const session = open();
    expect(session.mergeLosesData(0, { r0: 0, c0: 0, r1: 0, c1: 1 }, 'center')).toBe(true);
    session.merge(0, { r0: 0, c0: 0, r1: 0, c1: 3 }, 'center');
    expect(session.current.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 3 }]);
    expect(at(session, 0, 'B1')).toBeNull();
    expect(styleAt(session, 0, 'A1')?.hAlign).toBe('center');
    const saved = reopen(session);
    expect(saved.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 3 }]);
    session.unmerge(0, { r0: 0, c0: 1, r1: 0, c1: 1 });
    expect(session.current.sheets[0].merges).toEqual([]);
  });
});

describe('rows and columns', () => {
  it('inserts rows, moving cells and updating formulas on every sheet', () => {
    const session = open();
    session.insertRows(0, 2, 2); // above "Paper"
    expect(at(session, 0, 'A5')?.value).toBe('Paper');
    expect(at(session, 0, 'D5')?.formula).toBe('=B5*C5');
    expect(at(session, 0, 'D6')?.formula).toBe('=SUM(D2:D5)');
    expect(at(session, 1, 'B1')?.formula).toBe('=Invoice!D6');
    session.setText([{ sheet: 0, row: 2, col: 3, text: '7' }]);
    expect(at(session, 0, 'D6')?.value).toBe(257);
    expect(at(session, 1, 'B1')?.value).toBe(257);

    const saved = reopen(session);
    expect(cellOf(saved, 0, 'A5')?.value).toBe('Paper');
    expect(cellOf(saved, 0, 'D6')).toMatchObject({ formula: '=SUM(D2:D5)', value: 257 });
    expect(cellOf(saved, 1, 'B1')).toMatchObject({ formula: '=Invoice!D6', value: 257 });
  });

  it('deletes rows: references to them become #REF!', () => {
    const session = open();
    session.deleteRows(0, 1, 1); // "Pens"
    expect(at(session, 0, 'A2')?.value).toBe('Paper');
    expect(at(session, 0, 'D3')?.formula).toBe('=SUM(D2:D2)');
    expect(at(session, 0, 'D3')?.value).toBe(200);
    session.undo();
    expect(at(session, 0, 'A2')?.value).toBe('Pens');
    expect(at(session, 0, 'D4')?.value).toBe(250);

    const single = open(buildWorkbookFile({ S: [[1, f(2, 'A1*2')]] }), 's.xlsx');
    single.deleteColumns(0, 0, 1);
    expect(at(single, 0, 'A1')).toMatchObject({ formula: '=#REF!*2', type: 'error', value: '#REF!' });
  });

  it('inserts and deletes columns, with merges, widths and frozen panes following', () => {
    const session = open();
    session.setColumnWidths(0, [3], 120);
    session.merge(0, { r0: 3, c0: 0, r1: 3, c1: 2 }, 'merge');
    session.setFrozen(0, 1, 1);
    session.insertColumns(0, 1, 1);
    const sheet = session.current.sheets[0];
    expect(at(session, 0, 'E2')?.formula).toBe('=C2*D2');
    expect(sheet.merges).toEqual([{ r0: 3, c0: 0, r1: 3, c1: 3 }]);
    expect([...sheet.columnWidths.index]).toEqual([4]);
    expect(sheet.frozen).toEqual({ rows: 1, columns: 1 });

    const saved = reopen(session);
    expect(cellOf(saved, 0, 'E4')?.formula).toBe('=SUM(E2:E3)');
    expect(saved.sheets[0].merges).toEqual([{ r0: 3, c0: 0, r1: 3, c1: 3 }]);
    expect([...saved.sheets[0].columnWidths.index]).toEqual([4]);
    expect(saved.sheets[0].columnWidths.size[0]).toBe(120);
    expect(saved.sheets[0].frozen).toEqual({ rows: 1, columns: 1 });

    session.deleteColumns(0, 0, 2);
    expect(at(session, 0, 'A1')?.value).toBe('Qty');
    expect(session.current.sheets[0].merges).toEqual([{ r0: 3, c0: 0, r1: 3, c1: 1 }]);
  });

  it('gives inserted rows the formatting of the row above', () => {
    const session = open();
    session.format(0, { r0: 1, c0: 0, r1: 1, c1: 3 }, { fill: '#ddebf7' });
    session.insertRows(0, 2, 1);
    expect(styleAt(session, 0, 'B3')?.fill).toBe('#ddebf7');
    expect(at(session, 0, 'B3')?.value).toBeNull();
  });

  it('refuses inserts that would push data off the sheet', () => {
    const session = new EditSession(createBlankWorkbook(), null);
    session.setText([{ sheet: 0, row: EXCEL_MAX_ROWS - 1, col: 0, text: 'end' }]);
    expect(() => session.insertRows(0, 0, 1)).toThrow(EditRefusedError);
    session.setText([{ sheet: 0, row: 0, col: EXCEL_MAX_COLS - 1, text: 'edge' }]);
    expect(() => session.insertColumns(0, 5, 1)).toThrow(EditRefusedError);
  });

  it('resizes and hides rows and columns', () => {
    const session = open();
    session.setRowHeights(0, [0], 40);
    session.setColumnWidths(0, [2], 0);
    const saved = reopen(session);
    expect(saved.sheets[0].rowHeights.size[0]).toBe(40);
    expect(saved.sheets[0].columnWidths.size[0]).toBe(0);
  });
});

describe('sheets', () => {
  it('adds, renames, moves and duplicates sheets', () => {
    const session = open();
    session.addSheet(2);
    session.renameSheet(0, 'Q1 Invoice');
    expect(at(session, 1, 'B1')?.formula).toBe("='Q1 Invoice'!D4");
    session.moveSheet(2, 0);
    expect(session.current.sheets.map((s) => s.name)).toEqual(['Sheet3', 'Q1 Invoice', 'Summary']);
    session.duplicateSheet(1);
    expect(session.current.sheets.map((s) => s.name)).toEqual(['Sheet3', 'Q1 Invoice', 'Q1 Invoice (2)', 'Summary']);
    session.setText([{ sheet: 2, row: 1, col: 1, text: '1' }]);
    expect(at(session, 2, 'D4')?.value).toBe(205);
    expect(at(session, 1, 'D4')?.value).toBe(250);

    const saved = reopen(session);
    expect(saved.sheets.map((s) => s.name)).toEqual(['Sheet3', 'Q1 Invoice', 'Q1 Invoice (2)', 'Summary']);
    expect(cellOf(saved, 3, 'B1')).toMatchObject({ formula: "='Q1 Invoice'!D4", value: 250 });
    expect(cellOf(saved, 2, 'D2')?.value).toBe(5);
    expect(cellOf(saved, 2, 'A3')?.value).toBe('Paper');
  });

  it('deletes a sheet: formulas that used it show #REF!', () => {
    const session = open();
    session.deleteSheet(0);
    expect(session.current.sheets.map((s) => s.name)).toEqual(['Summary']);
    expect(at(session, 0, 'B1')).toMatchObject({ formula: '=#REF!', value: '#REF!' });
    expect(() => session.deleteSheet(0)).toThrow('at least one visible sheet');
    session.undo();
    expect(at(session, 1, 'B1')).toMatchObject({ formula: '=Invoice!D4', value: 250 });
    const saved = reopen(session);
    expect(saved.sheets.map((s) => s.name)).toEqual(['Invoice', 'Summary']);
  });

  it('validates sheet names', () => {
    const session = open();
    expect(() => session.renameSheet(0, 'summary')).toThrow('already a sheet');
    expect(() => session.renameSheet(0, 'a/b')).toThrow('can’t contain');
    expect(() => session.renameSheet(0, 'x'.repeat(32))).toThrow('31 characters');
  });
});

describe('copy, fill and sort', () => {
  it('pastes copied cells with formulas moved and formats kept', () => {
    const session = open();
    session.format(0, { r0: 1, c0: 3, r1: 1, c1: 3 }, { bold: true });
    const block = readBlock(session.current.sheets[0], { r0: 1, c0: 0, r1: 1, c1: 3 });
    session.paste(0, { r0: 5, c0: 0, r1: 5, c1: 0 }, block);
    expect(at(session, 0, 'D6')).toMatchObject({ formula: '=B6*C6', value: 50 });
    expect(styleAt(session, 0, 'D6')?.bold).toBe(true);
    session.paste(0, { r0: 6, c0: 0, r1: 6, c1: 0 }, block, 'values');
    expect(at(session, 0, 'D7')).toMatchObject({ value: 50 });
    expect(at(session, 0, 'D7')?.formula).toBeUndefined();
  });

  it('fills down, and continues series with the fill handle', () => {
    const session = new EditSession(createBlankWorkbook(), null);
    session.setText([
      { sheet: 0, row: 0, col: 0, text: '1' },
      { sheet: 0, row: 1, col: 0, text: '3' },
      { sheet: 0, row: 0, col: 1, text: 'Item 1' },
      { sheet: 0, row: 0, col: 2, text: 'Jan' },
      { sheet: 0, row: 0, col: 3, text: '2024-01-31' },
      { sheet: 0, row: 0, col: 4, text: '=A1*2' },
    ]);
    session.fill(0, { r0: 0, c0: 0, r1: 1, c1: 0 }, { r0: 2, c0: 0, r1: 3, c1: 0 }, true);
    expect([at(session, 0, 'A3')?.value, at(session, 0, 'A4')?.value]).toEqual([5, 7]);
    session.fill(0, { r0: 0, c0: 1, r1: 0, c1: 4 }, { r0: 1, c0: 1, r1: 2, c1: 4 }, true);
    expect(at(session, 0, 'B3')?.value).toBe('Item 3');
    expect(at(session, 0, 'C3')?.value).toBe('Mar');
    expect(at(session, 0, 'D2')?.displayValue).toBe('2024-02-01');
    expect(at(session, 0, 'E3')).toMatchObject({ formula: '=A3*2', value: 10 });
    // Fill down copies without a series.
    session.fill(0, { r0: 0, c0: 0, r1: 0, c1: 0 }, { r0: 1, c0: 0, r1: 1, c1: 0 }, false);
    expect(at(session, 0, 'A2')?.value).toBe(1);
  });

  it('sorts rows by a column, keeping formulas on their rows', () => {
    const session = open();
    session.sort(0, { r0: 1, c0: 0, r1: 2, c1: 3 }, 3, false);
    expect(at(session, 0, 'A2')?.value).toBe('Paper');
    expect(at(session, 0, 'D2')).toMatchObject({ formula: '=B2*C2', value: 200 });
    session.sort(0, { r0: 1, c0: 0, r1: 2, c1: 3 }, 0, true);
    expect(at(session, 0, 'A2')?.value).toBe('Paper');
    session.sort(0, { r0: 1, c0: 0, r1: 2, c1: 3 }, 0, false);
    expect(at(session, 0, 'A2')?.value).toBe('Pens');
  });

  it('moves cut cells, and formulas that used them follow', () => {
    const session = open();
    session.move(0, { r0: 1, c0: 1, r1: 2, c1: 2 }, 0, 10, 5);
    expect(at(session, 0, 'B2')).toBeNull();
    expect(at(session, 0, 'F11')?.value).toBe(10);
    expect(at(session, 0, 'D2')?.formula).toBe('=F11*G11');
    expect(at(session, 0, 'D4')?.value).toBe(250);
    session.undo();
    expect(at(session, 0, 'B2')?.value).toBe(10);
    expect(at(session, 0, 'D2')?.formula).toBe('=B2*C2');
  });
});

describe('sort range', () => {
  it('sorts the data block around a cell, not its title, heading or total', () => {
    const bytes = buildWorkbookFile(
      {
        S: [
          ['Quarterly report', null, null],
          ['City', 'Orders', 'Units'],
          ['Pune', 3, 30],
          ['Agra', 9, 90],
          ['Kochi', 5, 50],
          ['Total', f(17, 'SUM(B3:B5)'), f(170, 'SUM(C3:C5)')],
        ],
      },
      'xlsx',
      (book) => {
        book.Sheets.S['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
      },
    );
    const workbook = parseWorkbookBytes(bytes, 's.xlsx');
    expect(sortRange(workbook.sheets[0], { row: 3, col: 1 }, workbook.styles)).toEqual({ r0: 2, c0: 0, r1: 4, c1: 2 });

    const session = new EditSession(workbook, bytes);
    session.sort(0, { r0: 2, c0: 0, r1: 4, c1: 2 }, 1, false);
    expect(['A3', 'A4', 'A5', 'B6'].map((a) => at(session, 0, a)?.value)).toEqual(['Agra', 'Kochi', 'Pune', 17]);
  });
});
