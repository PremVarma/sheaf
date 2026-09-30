import { stringAt } from '../../models/stringPool';
import type { CellRange, WorksheetModel } from '../../models/workbook';
import { rangeCellCount } from '../../utils/cellAddress';
import { clampToUsedRange, lowerBoundColumn } from '../workbook/sheetService';

/** Larger selections are refused rather than freezing the app building a huge string. */
export const MAX_COPY_CELLS = 2_000_000;

export interface CopyPayload {
  text: string;
  rows: number;
  columns: number;
}

/** Quotes a field the way Excel does when it contains tabs, line breaks or quotes. */
function quoteField(value: string): string {
  return /[\t\n\r"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/**
 * Tab-separated display text for a range, which pastes cell-for-cell into
 * Excel, Numbers, Google Sheets and plain text editors.
 * Returns null when the selection is too large.
 */
export function buildCopyPayload(sheet: WorksheetModel, selection: CellRange): CopyPayload | null {
  const range = clampToUsedRange(selection, sheet);
  if (rangeCellCount(range) > MAX_COPY_CELLS) return null;

  const width = range.c1 - range.c0 + 1;
  const { rowStart, cols, texts } = sheet.cells;
  const lines: string[] = [];
  const fields = new Array<string>(width);
  for (let r = range.r0; r <= range.r1; r++) {
    fields.fill('');
    if (r < sheet.rowCount) {
      const end = rowStart[r + 1];
      for (let i = lowerBoundColumn(cols, rowStart[r], end, range.c0); i < end && cols[i] <= range.c1; i++) {
        fields[cols[i] - range.c0] = quoteField(stringAt(sheet.strings, texts[i]));
      }
    }
    lines.push(fields.join('\t'));
  }
  return { text: lines.join('\n'), rows: range.r1 - range.r0 + 1, columns: width };
}
