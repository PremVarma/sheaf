import { stringAt } from '../../models/stringPool';
import type { WorkbookModel, WorksheetModel } from '../../models/workbook';
import { SheetBuilder } from '../workbook/sheetBuilder';

/** A new, empty workbook. */
export function createBlankWorkbook(name = 'Untitled.xlsx'): WorkbookModel {
  return {
    fileName: name,
    fileSize: 0,
    format: 'xlsx',
    properties: {},
    sheets: [new SheetBuilder({ name: 'Sheet1', index: 0 }).finish()],
    styles: [{}],
    activeSheetIndex: 0,
    date1904: false,
    warnings: [],
  };
}

function quoteField(value: string, delimiter: string): string {
  return value.includes(delimiter) || /["\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** A sheet as delimited text: the displayed values, like Excel's CSV export. */
export function serializeDelimited(sheet: WorksheetModel, delimiter = ',', bom = false): Uint8Array {
  const { rowStart, cols, texts } = sheet.cells;
  const lines: string[] = [];
  const fields: string[] = new Array(sheet.columnCount);
  for (let r = 0; r < sheet.rowCount; r++) {
    fields.fill('');
    let width = 0;
    for (let i = rowStart[r]; i < rowStart[r + 1]; i++) {
      const text = stringAt(sheet.strings, texts[i]);
      if (text === '') continue;
      fields[cols[i]] = quoteField(text, delimiter);
      width = Math.max(width, cols[i] + 1);
    }
    lines.push(fields.slice(0, width).join(delimiter));
  }
  const text = lines.join('\r\n') + (lines.length > 0 ? '\r\n' : '');
  const body = new TextEncoder().encode(text);
  if (!bom) return body;
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return out;
}
