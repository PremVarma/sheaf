import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { buildWorkbookFile, patchXlsx as patchParts } from '../../test/fixtures';
import { parseWorkbookBytes } from '../workbook/parser';
import { getCell } from '../workbook/sheetService';
import { EditSession } from './editSession';
import { createBlankWorkbook } from './export';

const invoice = () =>
  buildWorkbookFile({
    Invoice: [
      ['Item', 'Qty', 'Rate', 'Amount'],
      ['Pens', 10, 5, { t: 'n', v: 50, f: 'B2*C2' } as XLSX.CellObject],
      ['Paper', 2, 100, { t: 'n', v: 200, f: 'B3*C3' } as XLSX.CellObject],
      ['Subtotal', null, null, { t: 'n', v: 250, f: 'SUM(D2:D3)' } as XLSX.CellObject],
      ['GST 18%', null, null, { t: 'n', v: 45, f: 'D4*Rates!B1' } as XLSX.CellObject],
      ['Total', null, null, { t: 'n', v: 295, f: 'D4+D5' } as XLSX.CellObject],
    ],
    Rates: [['GST', 0.18]],
  });

function open(bytes: Uint8Array, name = 'invoice.xlsx') {
  const workbook = parseWorkbookBytes(bytes, name);
  return new EditSession(workbook, name.endsWith('.xlsx') ? bytes : null);
}

const value = (session: EditSession, sheet: number, address: string) => {
  const { r, c } = XLSX.utils.decode_cell(address);
  return getCell(session.current.sheets[sheet], r, c);
};

describe('EditSession', () => {
  it('recalculates dependent formulas, across sheets', () => {
    const session = open(invoice());
    session.setText([{ sheet: 0, row: 1, col: 1, text: '20' }]);
    expect(value(session, 0, 'D2')?.value).toBe(100);
    expect(value(session, 0, 'D4')?.value).toBe(300);
    expect(value(session, 0, 'D6')?.value).toBe(354);

    session.setText([{ sheet: 1, row: 0, col: 1, text: '12%' }]);
    expect(value(session, 1, 'B1')?.value).toBe(0.12);
    expect(value(session, 1, 'B1')?.displayValue).toBe('12%');
    expect(value(session, 0, 'D6')?.value).toBe(336);
    expect(session.staleFormulas).toBe(0);
  });

  it('parses typed values like Excel', () => {
    const session = new EditSession(createBlankWorkbook(), null);
    session.setText([
      { sheet: 0, row: 0, col: 0, text: '1,234.5' },
      { sheet: 0, row: 1, col: 0, text: '2024-03-15' },
      { sheet: 0, row: 2, col: 0, text: 'true' },
      { sheet: 0, row: 3, col: 0, text: "'00123" },
      { sheet: 0, row: 4, col: 0, text: '₹2,500' },
      { sheet: 0, row: 5, col: 0, text: '=A1*2' },
      { sheet: 0, row: 6, col: 0, text: 'hello' },
    ]);
    expect(value(session, 0, 'A1')).toMatchObject({ type: 'number', value: 1234.5 });
    expect(value(session, 0, 'A2')).toMatchObject({ type: 'date', value: 45366, displayValue: '2024-03-15' });
    expect(value(session, 0, 'A3')).toMatchObject({ type: 'boolean', value: true });
    expect(value(session, 0, 'A4')).toMatchObject({ type: 'string', value: '00123' });
    expect(value(session, 0, 'A5')).toMatchObject({ type: 'number', value: 2500, displayValue: '₹2,500.00' });
    expect(value(session, 0, 'A6')).toMatchObject({ type: 'number', value: 2469, formula: '=A1*2' });
    expect(value(session, 0, 'A7')).toMatchObject({ type: 'string', value: 'hello' });
  });

  it('undoes and redoes, including recalculated values', () => {
    const session = open(invoice());
    expect(session.dirty).toBe(false);
    session.setText([{ sheet: 0, row: 1, col: 2, text: '6' }]);
    expect(value(session, 0, 'D6')?.value).toBe(306.8);
    expect(session.dirty).toBe(true);
    session.undo();
    expect(value(session, 0, 'C2')?.value).toBe(5);
    expect(value(session, 0, 'D6')?.value).toBe(295);
    expect(session.dirty).toBe(false);
    session.redo();
    expect(value(session, 0, 'D6')?.value).toBe(306.8);
  });

  it('clears contents and keeps dependents consistent', () => {
    const session = open(invoice());
    session.clear(0, { r0: 2, c0: 1, r1: 2, c1: 2 });
    expect(value(session, 0, 'B3')).toBeNull();
    expect(value(session, 0, 'D3')?.value).toBe(0);
    expect(value(session, 0, 'D6')?.value).toBe(59);
  });

  it('flags circular references without hanging', () => {
    const session = new EditSession(createBlankWorkbook(), null);
    const outcome = session.setText([
      { sheet: 0, row: 0, col: 0, text: '=B1+1' },
      { sheet: 0, row: 0, col: 1, text: '=A1+1' },
    ]);
    expect(outcome.circular).toBe(true);
  });

  it('keeps saved values for formulas it cannot calculate, and counts them', () => {
    const bytes = buildWorkbookFile({ S: [[1, { t: 'n', v: 99, f: 'WEBSERVICE("x")+A1' } as XLSX.CellObject]] });
    const session = open(bytes, 's.xlsx');
    session.setText([{ sheet: 0, row: 0, col: 0, text: '5' }]);
    expect(value(session, 0, 'B1')?.value).toBe(99);
    expect(session.staleFormulas).toBe(1);
  });
});

describe('saving', () => {
  it('patches edits into the original .xlsx and keeps its formatting', () => {
    const styled = patchParts(invoice(), {
      'xl/styles.xml': (xml) =>
        xml
          .replace(/<fonts count="(\d+)">/, (_, n) => `<fonts count="${Number(n) + 1}">`)
          .replace('</fonts>', '<font><b/><sz val="12"/><name val="Calibri"/></font></fonts>')
          .replace(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/, (_m, n, inner: string) => {
            const fonts = Number(/<fonts count="(\d+)"/.exec(xml)![1]);
            return `<cellXfs count="${Number(n) + 1}">${inner}<xf numFmtId="0" fontId="${fonts}" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`;
          }),
      'xl/worksheets/sheet1.xml': (xml) => {
        const count = Number(/<cellXfs count="(\d+)">/.exec(xml)?.[1] ?? 0);
        void count;
        return xml.replace('<c r="A1"', '<c r="A1" s="__BOLD__"');
      },
    });
    // Point A1 at the bold format we appended (the last xf).
    const container = XLSX.CFB.read(styled, { type: 'array' });
    const stylesXml = new TextDecoder().decode(XLSX.CFB.find(container, '/xl/styles.xml')!.content as Uint8Array);
    const boldXf = Number(/<cellXfs count="(\d+)">/.exec(stylesXml)![1]) - 1;
    const source = patchParts(styled, { 'xl/worksheets/sheet1.xml': (xml) => xml.replace('s="__BOLD__"', `s="${boldXf}"`) });

    const session = open(source);
    session.setText([
      { sheet: 0, row: 1, col: 1, text: '20' },
      { sheet: 0, row: 7, col: 0, text: 'Thank you' },
      { sheet: 0, row: 1, col: 4, text: '=D2*2' },
    ]);
    const saved = session.buildFile('xlsx', 0);
    const reopened = parseWorkbookBytes(saved, 'invoice.xlsx');
    const sheet = reopened.sheets[0];
    expect(getCell(sheet, 1, 1)?.value).toBe(20);
    expect(getCell(sheet, 1, 3)).toMatchObject({ value: 100, formula: '=B2*C2' });
    expect(getCell(sheet, 5, 3)?.value).toBe(354);
    expect(getCell(sheet, 7, 0)?.value).toBe('Thank you');
    expect(getCell(sheet, 1, 4)).toMatchObject({ value: 200, formula: '=D2*2' });
    expect(reopened.styles[getCell(sheet, 0, 0)!.styleId]).toMatchObject({ bold: true });
    expect(reopened.sheets.map((s) => s.name)).toEqual(['Invoice', 'Rates']);

    const workbookXml = new TextDecoder().decode(
      XLSX.CFB.find(XLSX.CFB.read(saved, { type: 'array' }), '/xl/workbook.xml')!.content as Uint8Array,
    );
    expect(workbookXml).toContain('fullCalcOnLoad="1"');
  });

  it('writes number formats for typed percentages and dates into styles.xml', () => {
    const session = open(invoice());
    session.setText([
      { sheet: 1, row: 0, col: 1, text: '12.5%' },
      { sheet: 0, row: 8, col: 0, text: '2024-03-15' },
    ]);
    const reopened = parseWorkbookBytes(session.buildFile('xlsx', 0), 'invoice.xlsx');
    expect(getCell(reopened.sheets[1], 0, 1)?.displayValue).toBe('12.50%');
    expect(getCell(reopened.sheets[0], 8, 0)).toMatchObject({ type: 'date', displayValue: '2024-03-15' });
  });

  it('saves delimited text and new workbooks', () => {
    const csv = parseWorkbookBytes(new TextEncoder().encode('a;b\n1;"x;y"\n'), 'data.csv');
    const session = new EditSession(csv, null);
    session.setText([{ sheet: 0, row: 1, col: 0, text: '2' }]);
    expect(new TextDecoder().decode(session.buildFile('csv', 0))).toBe('a;b\r\n2;"x;y"\r\n');

    const blank = new EditSession(createBlankWorkbook(), null);
    blank.setText([
      { sheet: 0, row: 0, col: 0, text: '5' },
      { sheet: 0, row: 0, col: 1, text: '=A1*3' },
    ]);
    const reopened = parseWorkbookBytes(blank.buildFile('xlsx', 0), 'new.xlsx');
    expect(getCell(reopened.sheets[0], 0, 1)).toMatchObject({ value: 15, formula: '=A1*3' });
  });

  it('is clean again after saving', () => {
    const session = open(invoice());
    session.setText([{ sheet: 0, row: 1, col: 1, text: '3' }]);
    session.buildFile('xlsx', 0);
    session.markSaved();
    expect(session.dirty).toBe(false);
    session.setText([{ sheet: 0, row: 1, col: 1, text: '4' }]);
    const again = parseWorkbookBytes(session.buildFile('xlsx', 0), 'invoice.xlsx');
    expect(getCell(again.sheets[0], 1, 3)?.value).toBe(20);
  });
});
