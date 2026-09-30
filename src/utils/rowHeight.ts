import { stringAt } from '../models/stringPool';
import { CellKind, type CellStyle, type WorksheetModel } from '../models/workbook';
import { BASE_FONT_PX, measureText } from './textMeasure';

/**
 * Height (px at 100%) a row needs for its largest font and its wrapped text,
 * like Excel's row auto-fit. Returns the default height when nothing is taller.
 */
export function neededRowHeight(
  sheet: WorksheetModel,
  row: number,
  styles: readonly CellStyle[],
  columnWidth: (col: number) => number,
): number {
  if (row >= sheet.rowCount) return sheet.defaultRowHeight;
  const { rowStart, cols, kinds, texts } = sheet.cells;
  let needed = sheet.defaultRowHeight;
  for (let i = rowStart[row]; i < rowStart[row + 1]; i++) {
    const style = styles[sheet.cells.styles[i]] ?? {};
    const scale = style.fontScale ?? 1;
    const family = style.fontName ? `"${style.fontName}"` : undefined;
    const lineHeight = Math.ceil(BASE_FONT_PX * 1.25 * scale);
    let lines = 1;
    const text = kinds[i] === CellKind.Empty ? '' : stringAt(sheet.strings, texts[i]);
    if (style.wrap && text) {
      const width = Math.max(1, columnWidth(cols[i]) - 7 - (style.indent ?? 0) * 9);
      lines = 0;
      for (const line of text.split('\n')) {
        const measured = measureText(line, Boolean(style.bold), Boolean(style.italic), false, family) * scale;
        lines += Math.max(1, Math.ceil(measured / width));
      }
    }
    if (scale === 1 && lines === 1) continue;
    needed = Math.max(needed, lines * lineHeight + 5);
  }
  return Math.min(546, needed);
}
