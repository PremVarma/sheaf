import { stringAt } from '../../models/stringPool';
import {
  CELL_VALUE_TYPES,
  CellKind,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type CellModel,
  type CellRange,
  type CellRef,
  type WorksheetModel,
} from '../../models/workbook';

/** Read access to the compact sheet storage. */

/** First stored-cell index in [start, end) whose column is >= col. */
export function lowerBoundColumn(cols: Uint16Array, start: number, end: number, col: number): number {
  let lo = start;
  let hi = end;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (cols[mid] < col) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Stored-cell index range [start, end) of a row. */
export function rowSpan(sheet: WorksheetModel, row: number): [number, number] {
  if (row < 0 || row >= sheet.rowCount) return [0, 0];
  return [sheet.cells.rowStart[row], sheet.cells.rowStart[row + 1]];
}

/** Stored-cell index at (row, col), or -1. */
export function findCellIndex(sheet: WorksheetModel, row: number, col: number): number {
  const [start, end] = rowSpan(sheet, row);
  if (start === end) return -1;
  const i = lowerBoundColumn(sheet.cells.cols, start, end, col);
  return i < end && sheet.cells.cols[i] === col ? i : -1;
}

export function displayTextAt(sheet: WorksheetModel, index: number): string {
  return stringAt(sheet.strings, sheet.cells.texts[index]);
}

export function hasValueAt(sheet: WorksheetModel, row: number, col: number): boolean {
  const i = findCellIndex(sheet, row, col);
  return i >= 0 && sheet.cells.kinds[i] !== CellKind.Empty;
}

export function formulaAt(sheet: WorksheetModel, index: number): string | undefined {
  const { cells, texts } = sheet.formulas;
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (cells[mid] < index) lo = mid + 1;
    else hi = mid;
  }
  return lo < cells.length && cells[lo] === index ? texts[lo] : undefined;
}

/** Materializes one cell; null when nothing is stored there. */
export function getCell(sheet: WorksheetModel, row: number, col: number): CellModel | null {
  const i = findCellIndex(sheet, row, col);
  if (i < 0) return null;
  const kind = sheet.cells.kinds[i] as CellKind;
  const displayValue = displayTextAt(sheet, i);
  let value: CellModel['value'] = null;
  switch (kind) {
    case CellKind.String:
    case CellKind.Error:
      value = displayValue;
      break;
    case CellKind.Number:
    case CellKind.Date:
      value = sheet.cells.numbers[i];
      break;
    case CellKind.Boolean:
      value = sheet.cells.numbers[i] === 1;
      break;
    default:
      break;
  }
  const formula = formulaAt(sheet, i);
  return {
    row,
    column: col,
    value,
    displayValue,
    type: CELL_VALUE_TYPES[kind],
    ...(formula ? { formula } : {}),
    styleId: sheet.cells.styles[i],
  };
}

/** What the formula bar shows: the formula, the raw number, or the text. */
export function formulaBarText(sheet: WorksheetModel, row: number, col: number): string {
  const merge = mergeAt(sheet, row, col);
  const cell = merge ? getCell(sheet, merge.r0, merge.c0) : getCell(sheet, row, col);
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  if (cell.type === 'number' && typeof cell.value === 'number') return String(Number(cell.value.toPrecision(15)));
  return cell.displayValue;
}

/** Whole-row/whole-column selections are limited to the used range. */
export function clampToUsedRange(range: CellRange, sheet: WorksheetModel): CellRange {
  return {
    r0: range.r0,
    c0: range.c0,
    r1: range.r1 >= EXCEL_MAX_ROWS - 1 ? Math.max(range.r0, sheet.rowCount - 1) : range.r1,
    c1: range.c1 >= EXCEL_MAX_COLS - 1 ? Math.max(range.c0, sheet.columnCount - 1) : range.c1,
  };
}

export interface RangeStats {
  /** Non-empty cells. */
  count: number;
  numericCount: number;
  sum: number;
  average: number | null;
}

/** Excel-style status bar summary for a selection. O(stored cells in range). */
export function computeRangeStats(sheet: WorksheetModel, range: CellRange): RangeStats {
  const { rowStart, cols, kinds, numbers } = sheet.cells;
  let count = 0;
  let numericCount = 0;
  let sum = 0;
  const lastRow = Math.min(range.r1, sheet.rowCount - 1);
  for (let r = range.r0; r <= lastRow; r++) {
    const end = rowStart[r + 1];
    for (let i = lowerBoundColumn(cols, rowStart[r], end, range.c0); i < end && cols[i] <= range.c1; i++) {
      const kind = kinds[i];
      if (kind === CellKind.Empty) continue;
      count++;
      if (kind === CellKind.Number || kind === CellKind.Date) {
        numericCount++;
        sum += numbers[i];
      }
    }
  }
  return { count, numericCount, sum, average: numericCount > 0 ? sum / numericCount : null };
}

/**
 * Cmd/Ctrl+Arrow: jump to the edge of the current data block, or to the next
 * non-empty cell, or to the edge of the grid (`limits`), like Excel.
 */
export function findDataEdge(
  sheet: WorksheetModel,
  from: CellRef,
  dRow: -1 | 0 | 1,
  dCol: -1 | 0 | 1,
  limits: { rows: number; cols: number },
): CellRef {
  const inBounds = (r: number, c: number) => r >= 0 && c >= 0 && r < limits.rows && c < limits.cols;
  let r = from.row;
  let c = from.col;
  if (!inBounds(r + dRow, c + dCol)) return from;

  if (hasValueAt(sheet, r, c) && hasValueAt(sheet, r + dRow, c + dCol)) {
    while (inBounds(r + dRow, c + dCol) && hasValueAt(sheet, r + dRow, c + dCol)) {
      r += dRow;
      c += dCol;
    }
    return { row: r, col: c };
  }

  // Skip the empty run and stop on the next value, never past the used range.
  const lastRow = Math.min(limits.rows, sheet.rowCount) - 1;
  const lastCol = Math.min(limits.cols, sheet.columnCount) - 1;
  r += dRow;
  c += dCol;
  while (r >= 0 && c >= 0 && r <= lastRow && c <= lastCol) {
    if (hasValueAt(sheet, r, c)) return { row: r, col: c };
    r += dRow;
    c += dCol;
  }
  return {
    row: dRow > 0 ? limits.rows - 1 : dRow < 0 ? 0 : from.row,
    col: dCol > 0 ? limits.cols - 1 : dCol < 0 ? 0 : from.col,
  };
}

/** The merged range containing a cell, if any. */
export function mergeAt(sheet: WorksheetModel, row: number, col: number): CellRange | undefined {
  return sheet.merges.find((m) => row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1);
}

/** The used range's last cell (Cmd/Ctrl+End). */
export function lastUsedCell(sheet: WorksheetModel): CellRef {
  return { row: Math.max(0, sheet.rowCount - 1), col: Math.max(0, sheet.columnCount - 1) };
}

/**
 * The block of data around a cell, bounded by empty rows and columns
 * (Excel's "current region", used by sort and double-click fill).
 */
export function currentRegion(sheet: WorksheetModel, cell: CellRef): CellRange {
  const range = { r0: cell.row, c0: cell.col, r1: cell.row, c1: cell.col };
  const rowHasValue = (r: number, c0: number, c1: number) => {
    if (r < 0 || r >= sheet.rowCount) return false;
    const { rowStart, cols, kinds } = sheet.cells;
    const end = rowStart[r + 1];
    for (let i = lowerBoundColumn(cols, rowStart[r], end, Math.max(0, c0)); i < end && cols[i] <= c1; i++) {
      if (kinds[i] !== CellKind.Empty) return true;
    }
    return false;
  };
  const colHasValue = (c: number, r0: number, r1: number) => {
    if (c < 0) return false;
    for (let r = Math.max(0, r0); r <= Math.min(r1, sheet.rowCount - 1); r++) if (hasValueAt(sheet, r, c)) return true;
    return false;
  };
  for (let grew = true; grew; ) {
    grew = false;
    if (rowHasValue(range.r0 - 1, range.c0 - 1, range.c1 + 1)) {
      range.r0--;
      grew = true;
    }
    if (rowHasValue(range.r1 + 1, range.c0 - 1, range.c1 + 1)) {
      range.r1++;
      grew = true;
    }
    if (colHasValue(range.c0 - 1, range.r0 - 1, range.r1 + 1)) {
      range.c0--;
      grew = true;
    }
    if (colHasValue(range.c1 + 1, range.r0 - 1, range.r1 + 1)) {
      range.c1++;
      grew = true;
    }
  }
  return range;
}

/**
 * Whether the first row of a block looks like column headings: all text, and
 * the rows below hold numbers or dates, or the headings are formatted
 * differently (e.g. bold) — roughly Excel's guess when sorting.
 */
export function hasHeaderRow(sheet: WorksheetModel, range: CellRange, styles: readonly { bold?: boolean; fill?: string }[]): boolean {
  if (range.r1 <= range.r0) return false;
  let texts = 0;
  let numbersBelow = false;
  let styledDifferently = false;
  for (let c = range.c0; c <= range.c1; c++) {
    const head = getCell(sheet, range.r0, c);
    if (head && head.type !== 'string' && head.type !== 'empty') return false;
    if (head?.type === 'string') texts++;
    const below = getCell(sheet, range.r0 + 1, c);
    if (below && (below.type === 'number' || below.type === 'date')) numbersBelow = true;
    const headStyle = styles[head?.styleId ?? 0];
    const belowStyle = styles[below?.styleId ?? 0];
    if (Boolean(headStyle?.bold) !== Boolean(belowStyle?.bold) || headStyle?.fill !== belowStyle?.fill) styledDifferently = true;
  }
  return texts > 0 && (numbersBelow || styledDifferently);
}

/**
 * What to sort when a single cell is selected: the block of data around it,
 * without title rows above it (merged, or a single value), its heading row,
 * or a totals row at the bottom.
 */
export function sortRange(sheet: WorksheetModel, cell: CellRef, styles: readonly { bold?: boolean; fill?: string }[]): CellRange {
  const range = currentRegion(sheet, cell);
  const wide = range.c1 > range.c0;
  const valuesIn = (r: number) => {
    let count = 0;
    for (let c = range.c0; c <= range.c1; c++) if (hasValueAt(sheet, r, c)) count++;
    return count;
  };
  const isTitle = (r: number) => sheet.merges.some((m) => m.r0 <= r && m.r1 >= r && m.c0 <= range.c1 && m.c1 >= range.c0) || (wide && valuesIn(r) <= 1);
  while (range.r0 < cell.row && isTitle(range.r0)) range.r0++;
  if (hasHeaderRow(sheet, range, styles)) range.r0++;
  const isTotal = (r: number) => {
    for (let c = range.c0; c <= range.c1; c++) {
      const value = getCell(sheet, r, c);
      if (!value) continue;
      if (typeof value.value === 'string' && /^\s*(grand\s+)?(sub)?totals?\b/i.test(value.value)) return true;
      if (value.formula && /^=\s*(SUM|SUBTOTAL)\(/i.test(value.formula)) return true;
    }
    return false;
  };
  if (range.r1 > cell.row && range.r1 > range.r0 && isTotal(range.r1)) range.r1--;
  return range;
}
