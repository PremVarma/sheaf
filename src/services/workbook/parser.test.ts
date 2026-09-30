import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { buildWorkbookFile, patchXlsx } from '../../test/fixtures';
import { toUserFacingError, WorkbookError } from './errors';
import { parseWorkbookBytes } from './parser';
import { formulaBarText, getCell } from './sheetService';

const mixedRows = [
  ['Name', 'Amount', 'When', 'Paid', 'Formula', 'Error'],
  [
    'Mumbai',
    1234.5,
    { t: 'n', v: 45366, z: 'yyyy-mm-dd' } as XLSX.CellObject,
    true,
    { t: 'n', v: 3, f: 'B3+1' } as XLSX.CellObject,
    { t: 'e', v: 0x07 } as XLSX.CellObject,
  ],
  ['Delhi', 2, null, false, { t: 's', v: 'MUMBAI', f: 'UPPER(A2)' } as XLSX.CellObject, null],
];

function expectCode(run: () => unknown, code: string) {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(WorkbookError);
    expect((error as WorkbookError).code).toBe(code);
    return;
  }
  throw new Error(`Expected ${code} but parsing succeeded`);
}

describe('parseWorkbookBytes: xlsx', () => {
  const workbook = parseWorkbookBytes(buildWorkbookFile({ Data: mixedRows, Empty: [], Third: [['x']] }), 'report.xlsx');
  const data = workbook.sheets[0];

  it('reads every sheet in order', () => {
    expect(workbook.format).toBe('xlsx');
    expect(workbook.sheets.map((s) => s.name)).toEqual(['Data', 'Empty', 'Third']);
    expect(data.rowCount).toBe(3);
    expect(data.columnCount).toBe(6);
  });

  it('reads strings', () => {
    expect(getCell(data, 0, 0)).toMatchObject({ type: 'string', value: 'Name', displayValue: 'Name' });
  });

  it('reads numbers', () => {
    expect(getCell(data, 1, 1)).toMatchObject({ type: 'number', value: 1234.5, displayValue: '1234.5' });
  });

  it('reads dates using their number format', () => {
    expect(getCell(data, 1, 2)).toMatchObject({ type: 'date', value: 45366, displayValue: '2024-03-15' });
  });

  it('reads booleans', () => {
    expect(getCell(data, 1, 3)).toMatchObject({ type: 'boolean', value: true, displayValue: 'TRUE' });
    expect(getCell(data, 2, 3)).toMatchObject({ type: 'boolean', value: false, displayValue: 'FALSE' });
  });

  it('reads formulas with their cached results', () => {
    expect(getCell(data, 1, 4)).toMatchObject({ type: 'number', value: 3, formula: '=B3+1' });
    expect(getCell(data, 2, 4)).toMatchObject({ type: 'string', displayValue: 'MUMBAI', formula: '=UPPER(A2)' });
    expect(formulaBarText(data, 1, 4)).toBe('=B3+1');
    expect(formulaBarText(data, 1, 1)).toBe('1234.5');
  });

  it('reads error values', () => {
    expect(getCell(data, 1, 5)).toMatchObject({ type: 'error', displayValue: '#DIV/0!' });
  });

  it('treats missing cells as empty', () => {
    expect(getCell(data, 2, 2)).toBeNull();
    expect(getCell(data, 50, 50)).toBeNull();
  });

  it('handles empty sheets', () => {
    const empty = workbook.sheets[1];
    expect(empty.rowCount).toBe(0);
    expect(empty.columnCount).toBe(0);
    expect(empty.cellCount).toBe(0);
  });

  it('applies number formats to display values', () => {
    const formatted = parseWorkbookBytes(
      buildWorkbookFile({
        S: [[{ t: 'n', v: 1234567.891, z: '#,##0.00' } as XLSX.CellObject, { t: 'n', v: 0.256, z: '0.0%' } as XLSX.CellObject]],
      }),
      'f.xlsx',
    ).sheets[0];
    expect(getCell(formatted, 0, 0)?.displayValue).toBe('1,234,567.89');
    expect(getCell(formatted, 0, 1)?.displayValue).toBe('25.6%');
  });

  it('reads merges, column widths, row heights, hidden sheets and properties', () => {
    const bytes = buildWorkbookFile({ Main: [['Title'], ['a', 'b', 'c']], Secret: [['s']] }, 'xlsx', (wb) => {
      const sheet = wb.Sheets.Main;
      sheet['!merges'] = [XLSX.utils.decode_range('A1:C1')];
      sheet['!cols'] = [{ wch: 20 }, { hidden: true }];
      sheet['!rows'] = [{ hpt: 30 }];
      wb.Workbook = { Sheets: [{}, { Hidden: 1 }] };
      wb.Props = { Title: 'Quarterly', Author: 'Finance' };
    });
    const parsed = parseWorkbookBytes(bytes, 'layout.xlsx');
    const main = parsed.sheets[0];
    expect(main.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 2 }]);
    expect(Array.from(main.columnWidths.index)).toEqual([0, 1]);
    expect(main.columnWidths.size[0]).toBeGreaterThan(140);
    expect(main.columnWidths.size[1]).toBe(0);
    expect(main.rowHeights.size[0]).toBe(40);
    expect(parsed.sheets[1].visibility).toBe('hidden');
    expect(parsed.properties).toMatchObject({ title: 'Quarterly', author: 'Finance' });
  });

  it('reads freeze panes, hidden gridlines and cell formatting from the package XML', () => {
    const base = buildWorkbookFile({ S: [['Header', 'Total'], ['a', 1]] });
    const bytes = patchXlsx(base, {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml
          .replace(
            /<sheetView[^>]*\/>/,
            '<sheetView workbookViewId="0" showGridLines="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView>',
          )
          .replace('<c r="A1"', '<c r="A1" s="1"'),
      'xl/styles.xml': (xml) =>
        xml
          .replace(/<fonts count="(\d+)">/, (_, n) => `<fonts count="${Number(n) + 1}">`)
          .replace('</fonts>', '<font><b/><i/><sz val="16"/><color rgb="FFFF0000"/><name val="Calibri"/></font></fonts>')
          .replace(/<fills count="(\d+)">/, (_, n) => `<fills count="${Number(n) + 1}">`)
          .replace('</fills>', '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>')
          .replace(/<borders count="(\d+)">/, (_, n) => `<borders count="${Number(n) + 1}">`)
          .replace('</borders>', '<border><left/><right/><top/><bottom style="medium"><color rgb="FF0000FF"/></bottom><diagonal/></border></borders>')
          .replace(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/, (_match, _count, inner: string) => {
            const fonts = Number(/<fonts count="(\d+)"/.exec(xml)![1]);
            const fills = Number(/<fills count="(\d+)"/.exec(xml)![1]);
            const borders = Number(/<borders count="(\d+)"/.exec(xml)![1]);
            const first = /<xf [^>]*?(\/>|>[\s\S]*?<\/xf>)/.exec(inner)![0];
            const styled = `<xf numFmtId="0" fontId="${fonts}" fillId="${fills}" borderId="${borders}" xfId="0" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>`;
            return `<cellXfs count="2">${first}${styled}</cellXfs>`;
          }),
    });
    const parsed = parseWorkbookBytes(bytes, 'styled.xlsx');
    const sheet = parsed.sheets[0];
    expect(sheet.frozen).toEqual({ rows: 1, columns: 1 });
    expect(sheet.showGridlines).toBe(false);

    const style = parsed.styles[getCell(sheet, 0, 0)!.styleId];
    expect(style).toMatchObject({ bold: true, italic: true, color: '#ff0000', fill: '#ffff00', hAlign: 'center', wrap: true });
    expect(style.fontScale).toBeCloseTo(16 / 12, 2);
    expect(style.borders?.bottom).toEqual({ width: 2, style: 'solid', color: '#0000ff' });
    expect(getCell(sheet, 0, 1)!.styleId).toBe(0);
  });
});

describe('parseWorkbookBytes: xls', () => {
  it('reads legacy .xls workbooks with multiple sheets', () => {
    const bytes = buildWorkbookFile({ Regions: mixedRows.slice(0, 2), Other: [['x', 2]] }, 'xls');
    const parsed = parseWorkbookBytes(bytes, 'legacy.xls');
    expect(parsed.format).toBe('xls');
    expect(parsed.sheets.map((s) => s.name)).toEqual(['Regions', 'Other']);
    const sheet = parsed.sheets[0];
    expect(getCell(sheet, 1, 0)).toMatchObject({ type: 'string', value: 'Mumbai' });
    expect(getCell(sheet, 1, 1)).toMatchObject({ type: 'number', value: 1234.5 });
    expect(getCell(sheet, 1, 2)).toMatchObject({ type: 'date', value: 45366 });
    expect(getCell(sheet, 1, 3)).toMatchObject({ type: 'boolean', value: true });
    // SheetJS writes BIFF8 values but not formula records; the cached result is what matters here.
    expect(getCell(sheet, 1, 4)).toMatchObject({ type: 'number', value: 3 });
  });

  it('reads XML Spreadsheet 2003 files saved with an .xls extension, including formulas', () => {
    const xml = `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Report"><Table>
  <Row><Cell><Data ss:Type="String">Qty</Data></Cell><Cell><Data ss:Type="Number">2</Data></Cell></Row>
  <Row><Cell><Data ss:Type="String">Double</Data></Cell><Cell ss:Formula="=R1C2*2"><Data ss:Type="Number">4</Data></Cell></Row>
 </Table></Worksheet>
</Workbook>`;
    const parsed = parseWorkbookBytes(new TextEncoder().encode(xml), 'export.xls');
    expect(parsed.format).toBe('xls');
    expect(parsed.sheets[0].name).toBe('Report');
    expect(getCell(parsed.sheets[0], 1, 1)).toMatchObject({ type: 'number', value: 4, formula: '=$B$1*2' });
  });
});

describe('parseWorkbookBytes: csv', () => {
  const csv = (text: string, name = 'data.csv') => parseWorkbookBytes(new TextEncoder().encode(text), name).sheets[0];

  it('reads delimited text into one sheet named after the file', () => {
    const parsed = parseWorkbookBytes(new TextEncoder().encode('City,Pop\nMumbai,12442373\n'), 'cities.csv');
    expect(parsed.format).toBe('csv');
    expect(parsed.sheets).toHaveLength(1);
    expect(parsed.sheets[0].name).toBe('cities');
    expect(getCell(parsed.sheets[0], 1, 1)).toMatchObject({ type: 'number', value: 12442373, displayValue: '12442373' });
  });

  it('handles quotes, escaped quotes, embedded commas and line breaks', () => {
    const sheet = csv('a,b\n"Hello, ""world""","line1\nline2"\nlast,');
    expect(getCell(sheet, 1, 0)?.displayValue).toBe('Hello, "world"');
    expect(getCell(sheet, 1, 1)?.displayValue).toBe('line1\nline2');
    expect(getCell(sheet, 2, 0)?.displayValue).toBe('last');
    expect(sheet.rowCount).toBe(3);
  });

  it('keeps text exactly as written', () => {
    const sheet = csv('id,zip,big,flag\n1,00123,12345678901234567890,TRUE\n');
    expect(getCell(sheet, 1, 1)).toMatchObject({ type: 'string', displayValue: '00123' });
    expect(getCell(sheet, 1, 2)).toMatchObject({ type: 'string', displayValue: '12345678901234567890' });
    expect(getCell(sheet, 1, 3)).toMatchObject({ type: 'boolean', displayValue: 'TRUE' });
  });

  it('decodes UTF-8 with and without BOM, and Windows-1252', () => {
    expect(getCell(csv('﻿Zürich,मुंबई\n'), 0, 1)?.displayValue).toBe('मुंबई');
    expect(getCell(csv('Zürich\n'), 0, 0)?.displayValue).toBe('Zürich');
    const latin1 = parseWorkbookBytes(new Uint8Array([0x43, 0x61, 0x66, 0xe9]), 'cafe.csv').sheets[0];
    expect(getCell(latin1, 0, 0)?.displayValue).toBe('Café');
  });

  it('detects semicolon delimiters and reads .tsv as tab-separated', () => {
    expect(getCell(csv('a;b;c\n1;2;3\n'), 1, 2)?.displayValue).toBe('3');
    expect(getCell(csv('a,b\tc\n', 'x.tsv'), 0, 1)?.displayValue).toBe('c');
  });

  it('sizes columns to their content', () => {
    const sheet = csv('short,a much longer header value\n1,2\n');
    expect(sheet.columnWidths.size[1]).toBeGreaterThan(sheet.columnWidths.size[0]);
  });
});

describe('parseWorkbookBytes: errors', () => {
  const valid = buildWorkbookFile({ S: [['a', 1]] });

  it('rejects unsupported extensions', () => {
    expectCode(() => parseWorkbookBytes(valid, 'report.pdf'), 'UNSUPPORTED_FORMAT');
  });

  it('rejects empty files', () => {
    expectCode(() => parseWorkbookBytes(new Uint8Array(), 'empty.xlsx'), 'EMPTY_FILE');
  });

  it('rejects files that are not spreadsheets', () => {
    const binary = Uint8Array.from({ length: 4096 }, (_, i) => (i * 7919 + 13) % 256);
    expectCode(() => parseWorkbookBytes(binary, 'random.xlsx'), 'INVALID_FILE');
    expectCode(() => parseWorkbookBytes(new TextEncoder().encode('just some text'), 'notes.xlsx'), 'INVALID_FILE');
    expectCode(() => parseWorkbookBytes(binary, 'random.csv'), 'INVALID_FILE');
  });

  it('reports truncated workbooks as corrupted', () => {
    expectCode(() => parseWorkbookBytes(valid.slice(0, Math.floor(valid.length * 0.6)), 'cut.xlsx'), 'CORRUPTED');
  });

  it('reports ZIP files that are not workbooks', () => {
    const container = XLSX.CFB.utils.cfb_new();
    XLSX.CFB.utils.cfb_add(container, 'word/document.xml', new TextEncoder().encode('<w:document/>') as unknown as number[]);
    const docx = new Uint8Array(XLSX.CFB.write(container, { fileType: 'zip', type: 'array' }) as number[]);
    expectCode(() => parseWorkbookBytes(docx, 'renamed.xlsx'), 'NOT_A_WORKBOOK');
  });

  it('reports password-protected workbooks', () => {
    const container = XLSX.CFB.utils.cfb_new();
    XLSX.CFB.utils.cfb_add(container, '/EncryptionInfo', [4, 0, 4, 0, 0x40, 0, 0, 0]);
    XLSX.CFB.utils.cfb_add(container, '/EncryptedPackage', new Array(64).fill(0));
    const encrypted = new Uint8Array(XLSX.CFB.write(container, { type: 'array' }) as number[]);
    expectCode(() => parseWorkbookBytes(encrypted, 'secret.xlsx'), 'PASSWORD_PROTECTED');
  });

  it('turns failures into friendly messages without technical details', () => {
    const error = new WorkbookError('CORRUPTED', 'TypeError: x.charCodeAt is not a function at parse_zip (xlsx.mjs:123)');
    const message = toUserFacingError(error, 'budget.xlsx');
    expect(message.title).toBe('Unable to open “budget.xlsx”');
    expect(message.message).toBe('The file may be corrupted, password protected, or use an unsupported Excel feature.');
    expect(JSON.stringify(message)).not.toMatch(/TypeError|xlsx\.mjs|charCodeAt/);
    expect(toUserFacingError(new Error('boom')).message).not.toContain('boom');
  });
});
