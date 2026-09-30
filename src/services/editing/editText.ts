import { CellKind, type WorkbookModel } from '../../models/workbook';
import { mergeAt } from '../workbook/sheetService';
import { readStoredCell } from './cells';

/**
 * Text shown when editing a cell: its formula, or a value that re-parses to
 * the same thing (raw numbers, percentages as "12%", dates as displayed).
 */
export function editTextAt(workbook: WorkbookModel, sheetIndex: number, row: number, col: number): string {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return '';
  const merge = mergeAt(sheet, row, col);
  const cell = merge ? readStoredCell(sheet, merge.r0, merge.c0) : readStoredCell(sheet, row, col);
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  if (cell.kind === CellKind.Number) {
    const format = workbook.styles[cell.style]?.numberFormat;
    if (format?.includes('%')) return `${Number((cell.num * 100).toPrecision(15))}%`;
    return String(Number(cell.num.toPrecision(15)));
  }
  return cell.text;
}
