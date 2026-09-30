import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { buildWorkbookFile } from '../../test/fixtures';
import { parseWorkbookBytes } from '../workbook/parser';
import { getCell } from '../workbook/sheetService';
import { EditSession } from './editSession';
import { createBlankWorkbook } from './export';

type Container = { FullPaths: string[]; FileIndex: { content?: Uint8Array | number[] }[] };

/** Every part of a package as text. */
function partsOf(bytes: Uint8Array): Record<string, string> {
  const container = XLSX.CFB.read(bytes, { type: 'array' }) as unknown as Container;
  const out: Record<string, string> = {};
  container.FullPaths.forEach((path, i) => {
    const content = container.FileIndex[i].content;
    if (!content || path.endsWith('/')) return;
    out[path.replace(/^Root Entry\//, '')] = new TextDecoder().decode(content instanceof Uint8Array ? content : Uint8Array.from(content));
  });
  return out;
}

/** Adds or changes parts of a package. */
function withParts(bytes: Uint8Array, changes: Record<string, string | ((xml: string) => string)>): Uint8Array {
  const container = XLSX.CFB.read(bytes, { type: 'array' });
  for (const [path, change] of Object.entries(changes)) {
    const entry = XLSX.CFB.find(container, `/${path}`);
    const current = entry ? new TextDecoder().decode(entry.content as Uint8Array) : '';
    const next = new TextEncoder().encode(typeof change === 'function' ? change(current) : change);
    if (entry) {
      entry.content = next as unknown as typeof entry.content;
      entry.size = next.length;
    } else {
      XLSX.CFB.utils.cfb_add(container, `/${path}`, next);
    }
  }
  return new Uint8Array(XLSX.CFB.write(container, { fileType: 'zip', type: 'array' }) as number[]);
}

function expectWellFormed(parts: Record<string, string>): void {
  for (const [path, xml] of Object.entries(parts)) {
    if (!/\.(xml|rels|vml)$/.test(path)) continue;
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(doc.getElementsByTagName('parsererror').length, `${path} is not well-formed`).toBe(0);
  }
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const f = (v: number, formula: string) => ({ t: 'n', v, f: formula }) as XLSX.CellObject;

/** An invoice with a table, conditional formatting, validation, names, a comment, a shape and a chart. */
function richWorkbook(): Uint8Array {
  const base = buildWorkbookFile({
    Invoice: [
      ['Item', 'Qty', 'Rate', 'Amount'],
      ['Pens', 10, 5, f(50, 'B2*C2')],
      ['Paper', 2, 100, f(200, 'B3*C3')],
      ['Total', null, null, f(250, 'SUM(D2:D3)')],
    ],
    Notes: [['Untouched']],
  });
  return withParts(base, {
    'xl/worksheets/sheet1.xml': (xml) =>
      xml
        .replace(
          '</sheetData>',
          '</sheetData><conditionalFormatting sqref="D2:D3"><cfRule type="expression" dxfId="0" priority="1"><formula>D2&gt;100</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="whole" sqref="B2:B3"><formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations>',
        )
        .replace('</worksheet>', '<drawing r:id="rIdD"/><legacyDrawing r:id="rIdV"/><tableParts count="1"><tablePart r:id="rIdT"/></tableParts></worksheet>'),
    'xl/worksheets/_rels/sheet1.xml.rels':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdT" Type="${REL}/table" Target="../tables/table1.xml"/>` +
      `<Relationship Id="rIdD" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/>` +
      `<Relationship Id="rIdV" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
      `<Relationship Id="rIdC" Type="${REL}/comments" Target="../comments1.xml"/></Relationships>`,
    'xl/tables/table1.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Items" displayName="Items" ref="A1:D3" totalsRowShown="0">' +
      '<autoFilter ref="A1:D3"/><tableColumns count="4"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Rate"/><tableColumn id="4" name="Amount"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>',
    'xl/drawings/drawing1.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Box"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
    'xl/drawings/vmlDrawing1.vml': '<xml xmlns:x="urn:schemas-microsoft-com:office:excel"><x:ClientData ObjectType="Note"><x:Row>2</x:Row><x:Column>1</x:Column></x:ClientData></xml>',
    'xl/comments1.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>me</author></authors><commentList><comment ref="B3" authorId="0"><text><r><t>check</t></r></text></comment></commentList></comments>',
    'xl/charts/chart1.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Invoice!$D$2:$D$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
    'xl/workbook.xml': (xml) =>
      xml.replace(
        '</sheets>',
        '</sheets><definedNames><definedName name="Rates">Invoice!$C$2:$C$3</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Invoice!$A$1:$D$4</definedName></definedNames>',
      ),
    '[Content_Types].xml': (xml) =>
      xml.replace(
        '</Types>',
        '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
          '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
          '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' +
          '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      ),
  });
}

const cellOf = (workbook: ReturnType<typeof parseWorkbookBytes>, sheet: number, address: string) => {
  const { r, c } = XLSX.utils.decode_cell(address);
  return getCell(workbook.sheets[sheet], r, c);
};

describe('writing .xlsx', () => {
  it('keeps tables, formats, names, comments, drawings and charts in step with inserted rows and columns', () => {
    const original = richWorkbook();
    const session = new EditSession(parseWorkbookBytes(original, 'rich.xlsx'), original);
    expect(session.current.sheets[0].tables).toEqual([
      { name: 'Items', range: { r0: 0, c0: 0, r1: 2, c1: 3 }, columns: ['Item', 'Qty', 'Rate', 'Amount'], headerRow: true },
    ]);
    session.insertRows(0, 2, 1); // inside the table, above "Paper"
    session.insertColumns(0, 2, 1); // between Qty and Rate
    session.setText([{ sheet: 0, row: 0, col: 2, text: 'Discount' }]);
    session.renameSheet(0, 'Q1 Invoice');

    const bytes = session.buildFile('xlsx', 0);
    const parts = partsOf(bytes);
    const originalParts = partsOf(original);
    expectWellFormed(parts);

    const sheet = parts['xl/worksheets/sheet1.xml'];
    expect(sheet).toContain('<conditionalFormatting sqref="E2:E4"><cfRule type="expression" dxfId="0" priority="1"><formula>E2&gt;100</formula>');
    expect(sheet).toContain('<dataValidation type="whole" sqref="B2:B4">');
    expect(parts['xl/tables/table1.xml']).toContain('ref="A1:E4"');
    expect(parts['xl/tables/table1.xml']).toContain('<autoFilter ref="A1:E4"/>');
    expect(parts['xl/tables/table1.xml']).toMatch(/<tableColumns count="5">.*name="Qty"\/><tableColumn id="5" name="Discount"\/><tableColumn id="3" name="Rate"\/>/);
    expect(parts['xl/workbook.xml']).toContain(`<definedName name="Rates">'Q1 Invoice'!$D$2:$D$4</definedName>`);
    expect(parts['xl/workbook.xml']).toContain(`localSheetId="0">'Q1 Invoice'!$A$1:$E$5</definedName>`);
    expect(parts['xl/workbook.xml']).toContain('<sheet name="Q1 Invoice"');
    expect(parts['xl/charts/chart1.xml']).toContain(`<c:f>'Q1 Invoice'!$E$2:$E$4</c:f>`);
    expect(parts['xl/comments1.xml']).toContain('<comment ref="B4"');
    expect(parts['xl/drawings/drawing1.xml']).toContain('<xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row>');
    expect(parts['xl/drawings/vmlDrawing1.vml']).toContain('<x:Row>3</x:Row>');
    // Parts that didn't change are copied as they were.
    expect(parts['xl/theme/theme1.xml']).toBe(originalParts['xl/theme/theme1.xml']);
    expect(parts['xl/worksheets/sheet2.xml']).toBe(originalParts['xl/worksheets/sheet2.xml']);

    const reopened = parseWorkbookBytes(bytes, 'rich.xlsx');
    expect(reopened.sheets[0].name).toBe('Q1 Invoice');
    expect(cellOf(reopened, 0, 'A4')?.value).toBe('Paper');
    expect(cellOf(reopened, 0, 'E5')).toMatchObject({ formula: '=SUM(E2:E4)', value: 250 });
    expect(reopened.sheets[0].tables?.[0].columns).toEqual(['Item', 'Qty', 'Discount', 'Rate', 'Amount']);
  });

  it('adds, deletes and reorders sheets in the package', () => {
    const original = richWorkbook();
    const session = new EditSession(parseWorkbookBytes(original, 'rich.xlsx'), original);
    session.addSheet(2, 'Extra');
    session.setText([{ sheet: 2, row: 0, col: 0, text: '=Notes!A1' }]);
    session.duplicateSheet(0);
    session.deleteSheet(2); // Notes
    session.moveSheet(2, 0);
    const bytes = session.buildFile('xlsx', 1);
    const parts = partsOf(bytes);
    expectWellFormed(parts);
    expect(Object.keys(parts)).not.toContain('xl/worksheets/sheet2.xml');
    expect(parts['[Content_Types].xml']).not.toContain('/xl/worksheets/sheet2.xml');
    // The copy doesn't share the original's table, drawing or comments.
    const copyPath = Object.keys(parts).find((p) => /worksheets\/sheet\d+\.xml$/.test(p) && parts[p].includes('Pens') && p !== 'xl/worksheets/sheet1.xml')!;
    expect(parts[copyPath]).not.toContain('tablePart');
    expect(parts[copyPath]).not.toContain('<drawing');

    const reopened = parseWorkbookBytes(bytes, 'rich.xlsx');
    expect(reopened.sheets.map((s) => s.name)).toEqual(['Extra', 'Invoice', 'Invoice (2)']);
    expect(reopened.activeSheetIndex).toBe(1);
    expect(cellOf(reopened, 0, 'A1')).toMatchObject({ formula: '=#REF!' });
    expect(cellOf(reopened, 2, 'D4')?.value).toBe(250);
    // Invoice's print area moves with it to position 1.
    expect(parts['xl/workbook.xml']).toContain('<definedName name="_xlnm.Print_Area" localSheetId="1">Invoice!$A$1:$D$4</definedName>');
  });

  it('writes new workbooks with their formatting', () => {
    const session = new EditSession(createBlankWorkbook(), null);
    session.setText([
      { sheet: 0, row: 0, col: 0, text: 'GST Invoice' },
      { sheet: 0, row: 1, col: 0, text: '1250.5' },
    ]);
    session.format(0, { r0: 0, c0: 0, r1: 0, c1: 3 }, { bold: true, fontScale: 16 / 11, fill: '#1a7a4c', color: '#ffffff' });
    session.merge(0, { r0: 0, c0: 0, r1: 0, c1: 3 }, 'center');
    session.format(0, { r0: 1, c0: 0, r1: 1, c1: 0 }, { numberFormat: '"₹"#,##0.00', fontName: 'Georgia' });
    session.setColumnWidths(0, [0], 150);
    session.setFrozen(0, 1, 0);
    const bytes = session.buildFile('xlsx', 0);
    const parts = partsOf(bytes);
    expectWellFormed(parts);
    const reopened = parseWorkbookBytes(bytes, 'new.xlsx');
    const title = reopened.styles[cellOf(reopened, 0, 'A1')!.styleId];
    expect(title).toMatchObject({ bold: true, fill: '#1a7a4c', color: '#ffffff', hAlign: 'center', fontScale: expect.closeTo(16 / 11, 3) });
    expect(reopened.styles[cellOf(reopened, 0, 'A2')!.styleId]).toMatchObject({ numberFormat: '"₹"#,##0.00', fontName: 'Georgia' });
    expect(cellOf(reopened, 0, 'A2')?.displayValue).toBe('₹1,250.50');
    expect(reopened.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 3 }]);
    expect(reopened.sheets[0].columnWidths.size[0]).toBe(150);
    expect(reopened.sheets[0].frozen).toEqual({ rows: 1, columns: 0 });
  });

  it('saves again after saving, and after undoing a structural change', () => {
    const original = richWorkbook();
    const session = new EditSession(parseWorkbookBytes(original, 'rich.xlsx'), original);
    session.insertRows(0, 0, 1);
    session.buildFile('xlsx', 0);
    session.markSaved();
    session.undo();
    const reopened = parseWorkbookBytes(session.buildFile('xlsx', 0), 'rich.xlsx');
    expect(cellOf(reopened, 0, 'A1')?.value).toBe('Item');
    expect(cellOf(reopened, 0, 'D4')?.formula).toBe('=SUM(D2:D3)');
  });
});
