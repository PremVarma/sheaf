import { stringAt, stringCount, type StringPool } from '../../models/stringPool';
import { CellKind, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, type CellRange, type WorksheetModel } from '../../models/workbook';
import { findCellIndex, formulaAt } from '../workbook/sheetService';

/** Everything stored for one cell. Writing `null` removes the cell. */
export interface StoredCell {
  kind: CellKind;
  num: number;
  text: string;
  style: number;
  formula?: string;
}

/** Cell key within a sheet. */
export const cellKey = (row: number, col: number): number => row * EXCEL_MAX_COLS + col;
export const keyRow = (key: number): number => Math.floor(key / EXCEL_MAX_COLS);
export const keyCol = (key: number): number => key % EXCEL_MAX_COLS;

/** Cell key across sheets. */
const SHEET_STRIDE = EXCEL_MAX_ROWS * EXCEL_MAX_COLS;
export const globalKey = (sheet: number, row: number, col: number): number => sheet * SHEET_STRIDE + cellKey(row, col);
export const globalSheet = (key: number): number => Math.floor(key / SHEET_STRIDE);
export const globalLocal = (key: number): number => key % SHEET_STRIDE;

export function readStoredCell(sheet: WorksheetModel, row: number, col: number): StoredCell | null {
  const i = findCellIndex(sheet, row, col);
  if (i < 0) return null;
  const { kinds, numbers, texts, styles } = sheet.cells;
  const formula = formulaAt(sheet, i);
  return {
    kind: kinds[i] as CellKind,
    num: numbers[i],
    text: stringAt(sheet.strings, texts[i]),
    style: styles[i],
    ...(formula ? { formula } : {}),
  };
}

/** Cells copied inside the app, pasted with their formulas, formats and merged cells. */
export interface CellBlock {
  sheetName: string;
  range: CellRange;
  /** Row-major, (range rows) × (range columns). */
  cells: (StoredCell | null)[];
  /** Merged ranges inside the block, relative to its top-left cell. */
  merges: CellRange[];
}

/** The cells of a range, for copying. */
export function readBlock(sheet: WorksheetModel, range: CellRange): CellBlock {
  const cells: (StoredCell | null)[] = [];
  for (let r = range.r0; r <= range.r1; r++) {
    for (let c = range.c0; c <= range.c1; c++) cells.push(readStoredCell(sheet, r, c));
  }
  const inside = (m: CellRange) => m.r0 >= range.r0 && m.r1 <= range.r1 && m.c0 >= range.c0 && m.c1 <= range.c1;
  const merges = sheet.merges.filter(inside).map((m) => ({ r0: m.r0 - range.r0, c0: m.c0 - range.c0, r1: m.r1 - range.r0, c1: m.c1 - range.c0 }));
  return { sheetName: sheet.name, range, cells, merges };
}

export function sameCell(a: StoredCell | null, b: StoredCell | null): boolean {
  if (a === null || b === null) return a === b;
  return a.kind === b.kind && a.num === b.num && a.text === b.text && a.style === b.style && a.formula === b.formula;
}

function extendPool(pool: StringPool, added: string[]): StringPool {
  if (added.length === 0) return pool;
  const count = stringCount(pool);
  const offsets = new Uint32Array(count + added.length + 1);
  offsets.set(pool.offsets);
  let end = pool.offsets[count];
  for (let k = 0; k < added.length; k++) {
    end += added[k].length;
    offsets[count + k + 1] = end;
  }
  return { data: pool.data + added.join(''), offsets };
}

/**
 * Returns a new sheet with the writes applied. One linear pass over the stored
 * cells, so every reader (grid, search, copy) keeps working on plain arrays.
 */
export function applyWrites(sheet: WorksheetModel, writes: ReadonlyMap<number, StoredCell | null>): WorksheetModel {
  if (writes.size === 0) return sheet;
  const keys = [...writes.keys()].sort((a, b) => a - b);
  const { rowStart, cols, kinds, numbers, texts, styles } = sheet.cells;
  const oldCount = sheet.cellCount;
  const capacity = oldCount + keys.length;
  const outCols = new Uint16Array(capacity);
  const outKinds = new Uint8Array(capacity);
  const outNumbers = new Float64Array(capacity);
  const outTexts = new Uint32Array(capacity);
  const outStyles = new Uint16Array(capacity);
  const formulaCells: number[] = [];
  const formulaTexts: string[] = [];
  const oldFormulaCells = sheet.formulas.cells;
  const oldFormulaTexts = sheet.formulas.texts;

  const added: string[] = [];
  const addedIds = new Map<string, number>();
  const firstNewId = stringCount(sheet.strings);
  const textId = (text: string): number => {
    if (text === '') return 0;
    let id = addedIds.get(text);
    if (id === undefined) {
      id = firstNewId + added.length;
      added.push(text);
      addedIds.set(text, id);
    }
    return id;
  };

  const lastWrite = keys[keys.length - 1];
  const rowCount = Math.max(sheet.rowCount, keyRow(lastWrite) + 1);
  let columnCount = sheet.columnCount;
  const outRowStart = new Uint32Array(rowCount + 1);
  let out = 0;
  let k = 0;
  let fp = 0;

  const copyOld = (i: number) => {
    outCols[out] = cols[i];
    outKinds[out] = kinds[i];
    outNumbers[out] = numbers[i];
    outTexts[out] = texts[i];
    outStyles[out] = styles[i];
    while (fp < oldFormulaCells.length && oldFormulaCells[fp] < i) fp++;
    if (fp < oldFormulaCells.length && oldFormulaCells[fp] === i) {
      formulaCells.push(out);
      formulaTexts.push(oldFormulaTexts[fp]);
    }
    out++;
  };
  const emit = (col: number, cell: StoredCell) => {
    outCols[out] = col;
    outKinds[out] = cell.kind;
    outNumbers[out] = cell.num;
    outTexts[out] = textId(cell.text);
    outStyles[out] = cell.style;
    if (cell.formula) {
      formulaCells.push(out);
      formulaTexts.push(cell.formula);
    }
    if (col + 1 > columnCount) columnCount = col + 1;
    out++;
  };

  for (let r = 0; r < rowCount; r++) {
    outRowStart[r] = out;
    const start = r < sheet.rowCount ? rowStart[r] : oldCount;
    const end = r < sheet.rowCount ? rowStart[r + 1] : oldCount;
    let i = start;
    for (;;) {
      const writeKey = k < keys.length && keyRow(keys[k]) === r ? keys[k] : -1;
      if (i >= end && writeKey < 0) break;
      const oldCol = i < end ? cols[i] : Infinity;
      const writeCol = writeKey >= 0 ? keyCol(writeKey) : Infinity;
      if (writeCol <= oldCol) {
        const cell = writes.get(writeKey);
        if (cell) emit(writeCol, cell);
        if (writeCol === oldCol) i++;
        k++;
      } else {
        copyOld(i++);
      }
    }
  }
  outRowStart[rowCount] = out;

  return {
    ...sheet,
    rowCount,
    columnCount,
    cellCount: out,
    cells: {
      rowStart: outRowStart,
      cols: outCols.slice(0, out),
      kinds: outKinds.slice(0, out),
      numbers: outNumbers.slice(0, out),
      texts: outTexts.slice(0, out),
      styles: outStyles.slice(0, out),
    },
    strings: extendPool(sheet.strings, added),
    formulas: { cells: Uint32Array.from(formulaCells), texts: formulaTexts },
  };
}
