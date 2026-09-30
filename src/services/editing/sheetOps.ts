import { stringCount, type StringPool } from '../../models/stringPool';
import {
  type CellKind,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type CellRange,
  type IndexMap,
  type SizeOverrides,
  type TableInfo,
  type WorksheetModel,
} from '../../models/workbook';
import { lowerBoundColumn } from '../workbook/sheetService';
import { toSizeOverrides } from '../workbook/sheetBuilder';
import { fromBase, insertAt, mapRange, removeAt, toBase } from './indexMap';

/** Pure transformations of a worksheet used by the edit session. */

export type Axis = 'rows' | 'cols';

/** Row of a stored cell index. */
export function rowOfCell(sheet: WorksheetModel, index: number): number {
  const { rowStart } = sheet.cells;
  let lo = 0;
  let hi = sheet.rowCount - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >>> 1;
    if (rowStart[mid] <= index) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Calls `visit` for each stored cell in a range (rows ascending, columns ascending). */
export function forEachStoredCell(sheet: WorksheetModel, range: CellRange, visit: (index: number, row: number, col: number) => void): void {
  const { rowStart, cols } = sheet.cells;
  const lastRow = Math.min(range.r1, sheet.rowCount - 1);
  for (let r = Math.max(0, range.r0); r <= lastRow; r++) {
    const end = rowStart[r + 1];
    for (let i = lowerBoundColumn(cols, rowStart[r], end, range.c0); i < end && cols[i] <= range.c1; i++) visit(i, r, cols[i]);
  }
}

/** Grows a range until no merged cell is cut by its edge (Excel selects whole merges). */
export function expandToMerges(sheet: WorksheetModel, range: CellRange): CellRange {
  const out = { ...range };
  for (let changed = true; changed; ) {
    changed = false;
    for (const m of sheet.merges) {
      if (m.r0 > out.r1 || m.r1 < out.r0 || m.c0 > out.c1 || m.c1 < out.c0) continue;
      if (m.r0 < out.r0 || m.r1 > out.r1 || m.c0 < out.c0 || m.c1 > out.c1) {
        out.r0 = Math.min(out.r0, m.r0);
        out.r1 = Math.max(out.r1, m.r1);
        out.c0 = Math.min(out.c0, m.c0);
        out.c1 = Math.max(out.c1, m.c1);
        changed = true;
      }
    }
  }
  return out;
}

export function extendPool(pool: StringPool, added: string[]): StringPool {
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

/** New display text and kind of a cell after its number format changed. */
export type Retext = (index: number, style: number) => { kind: CellKind; text: string } | null;

/**
 * Changes the style of stored cells in a range in one pass over the typed
 * arrays (no per-cell objects), so formatting a whole column stays fast.
 * Arrays that don't change are shared with the old sheet.
 */
export function restyleCells(
  sheet: WorksheetModel,
  range: CellRange,
  styleOf: (style: number, row: number, col: number) => number,
  retext?: Retext,
): WorksheetModel {
  const { kinds, texts, styles } = sheet.cells;
  let newStyles: Uint16Array | null = null;
  let newKinds: Uint8Array | null = null;
  let newTexts: Uint32Array | null = null;
  const added: string[] = [];
  const addedIds = new Map<string, number>();
  const firstNew = stringCount(sheet.strings);

  forEachStoredCell(sheet, range, (i, row, col) => {
    const style = styleOf(styles[i], row, col);
    if (style === styles[i]) return;
    newStyles ??= styles.slice();
    newStyles[i] = style;
    const change = retext?.(i, style);
    if (!change) return;
    if (change.kind !== kinds[i]) {
      newKinds ??= kinds.slice();
      newKinds[i] = change.kind;
    }
    let id = addedIds.get(change.text);
    if (id === undefined) {
      id = firstNew + added.length;
      added.push(change.text);
      addedIds.set(change.text, id);
    }
    newTexts ??= texts.slice();
    newTexts[i] = id;
  });
  if (!newStyles) return sheet;
  return {
    ...sheet,
    cells: { ...sheet.cells, styles: newStyles, kinds: newKinds ?? kinds, texts: newTexts ?? texts },
    strings: extendPool(sheet.strings, added),
  };
}

/** Sizes with some entries set (null = default size, 0 = hidden). */
export function withSizes(sizes: SizeOverrides, indexes: Iterable<number>, size: number | null): SizeOverrides {
  const map = new Map<number, number>();
  sizes.index.forEach((index, k) => map.set(index, sizes.size[k]));
  for (const index of indexes) {
    if (size === null) map.delete(index);
    else map.set(index, size);
  }
  return toSizeOverrides(map);
}

export function sizeAt(sizes: SizeOverrides, index: number): number | undefined {
  let lo = 0;
  let hi = sizes.index.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (sizes.index[mid] < index) lo = mid + 1;
    else hi = mid;
  }
  return lo < sizes.index.length && sizes.index[lo] === index ? sizes.size[lo] : undefined;
}

/** Whether rows/columns can be inserted at `at` without pushing stored cells off the sheet. */
export function canInsert(sheet: WorksheetModel, axis: Axis, at: number, count: number): boolean {
  const used = axis === 'rows' ? sheet.rowCount : sheet.columnCount;
  const limit = axis === 'rows' ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS;
  return at >= used || used + count <= limit;
}

/** A table column name not used in the table yet ("Column5"). */
function freshColumnName(taken: Set<string>): string {
  for (let n = 1; ; n++) {
    const name = `Column${n}`;
    if (!taken.has(name.toLowerCase())) {
      taken.add(name.toLowerCase());
      return name;
    }
  }
}

export interface ShiftResult {
  sheet: WorksheetModel;
  /** Header cells for table columns created by an insert: [row, col, text]. */
  headers: [number, number, string][];
}

/**
 * Inserts or deletes rows/columns: moves cells, merged ranges, sizes, frozen
 * panes and tables, and records the change in the sheet's origin. Formula text
 * is updated separately (references on every sheet may change).
 */
export function shiftSheet(sheet: WorksheetModel, axis: Axis, kind: 'insert' | 'delete', at: number, count: number, map: IndexMap): ShiftResult {
  const limit = axis === 'rows' ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS;
  const newIndex = (old: number) => {
    const index = fromBase(map, old);
    return index >= limit ? -1 : index;
  };
  const { rowStart, cols, kinds, numbers, texts, styles } = sheet.cells;
  const n = sheet.cellCount;
  const outCols = new Uint16Array(n);
  const outKinds = new Uint8Array(n);
  const outNumbers = new Float64Array(n);
  const outTexts = new Uint32Array(n);
  const outStyles = new Uint16Array(n);
  const formulaCells: number[] = [];
  const formulaTexts: string[] = [];
  const fc = sheet.formulas.cells;
  const ft = sheet.formulas.texts;
  let fp = 0;
  let out = 0;
  let maxCol = -1;

  const copy = (i: number, col: number) => {
    outCols[out] = col;
    outKinds[out] = kinds[i];
    outNumbers[out] = numbers[i];
    outTexts[out] = texts[i];
    outStyles[out] = styles[i];
    while (fp < fc.length && fc[fp] < i) fp++;
    if (fp < fc.length && fc[fp] === i) {
      formulaCells.push(out);
      formulaTexts.push(ft[fp]);
    }
    if (col > maxCol) maxCol = col;
    out++;
  };

  let rowCount: number;
  let outRowStart: Uint32Array;
  if (axis === 'rows') {
    const rowIndex = new Int32Array(sheet.rowCount);
    let last = -1;
    for (let r = 0; r < sheet.rowCount; r++) {
      rowIndex[r] = newIndex(r);
      if (rowIndex[r] >= 0 && rowStart[r] < rowStart[r + 1]) last = rowIndex[r];
    }
    rowCount = last + 1;
    outRowStart = new Uint32Array(rowCount + 1);
    let nextRow = 0;
    for (let r = 0; r < sheet.rowCount; r++) {
      const nr = rowIndex[r];
      if (nr < 0 || nr >= rowCount) continue;
      while (nextRow <= nr) outRowStart[nextRow++] = out;
      for (let i = rowStart[r]; i < rowStart[r + 1]; i++) copy(i, cols[i]);
    }
    while (nextRow <= rowCount) outRowStart[nextRow++] = out;
  } else {
    rowCount = sheet.rowCount;
    outRowStart = new Uint32Array(rowCount + 1);
    for (let r = 0; r < rowCount; r++) {
      outRowStart[r] = out;
      for (let i = rowStart[r]; i < rowStart[r + 1]; i++) {
        const col = newIndex(cols[i]);
        if (col >= 0) copy(i, col);
      }
    }
    outRowStart[rowCount] = out;
  }
  // Trailing empty rows keep the used range tidy.
  while (rowCount > 0 && outRowStart[rowCount - 1] === out) rowCount--;
  outRowStart = outRowStart.slice(0, rowCount + 1);

  const mapSpan = (a: number, b: number): [number, number] | null => mapRange(map, a, b, limit);
  const merges: CellRange[] = [];
  for (const m of sheet.merges) {
    const span = axis === 'rows' ? mapSpan(m.r0, m.r1) : mapSpan(m.c0, m.c1);
    if (!span) continue;
    const next = axis === 'rows' ? { ...m, r0: span[0], r1: span[1] } : { ...m, c0: span[0], c1: span[1] };
    if (next.r1 > next.r0 || next.c1 > next.c0) merges.push(next);
  }

  const sizes = axis === 'rows' ? sheet.rowHeights : sheet.columnWidths;
  const movedSizes = new Map<number, number>();
  sizes.index.forEach((index, k) => {
    const next = newIndex(index);
    if (next >= 0) movedSizes.set(next, sizes.size[k]);
  });

  const frozenCount = axis === 'rows' ? sheet.frozen.rows : sheet.frozen.columns;
  const frozen = frozenCount > 0 ? (mapSpan(0, frozenCount - 1)?.[1] ?? -1) + 1 : 0;

  const headers: [number, number, string][] = [];
  let tables: TableInfo[] | undefined;
  if (sheet.tables) {
    tables = [];
    for (const table of sheet.tables) {
      const span = axis === 'rows' ? mapSpan(table.range.r0, table.range.r1) : mapSpan(table.range.c0, table.range.c1);
      if (!span) continue;
      if (axis === 'rows') {
        tables.push({ ...table, range: { ...table.range, r0: span[0], r1: span[1] } });
        continue;
      }
      const taken = new Set(table.columns.map((name) => name.toLowerCase()));
      const columns: string[] = [];
      for (let col = span[0]; col <= span[1]; col++) {
        const old = toBase(map, col);
        if (old >= table.range.c0 && old <= table.range.c1) columns.push(table.columns[old - table.range.c0] ?? freshColumnName(taken));
        else {
          const name = freshColumnName(taken);
          columns.push(name);
          if (table.headerRow) headers.push([table.range.r0, col, name]);
        }
      }
      tables.push({ ...table, range: { ...table.range, c0: span[0], c1: span[1] }, columns });
    }
  }

  const origin = sheet.origin && {
    ...sheet.origin,
    [axis]: kind === 'insert' ? insertAt(sheet.origin[axis], at, count) : removeAt(sheet.origin[axis], at, count),
  };

  let columnCount = maxCol + 1;
  for (const m of merges) {
    rowCount = Math.max(rowCount, m.r1 + 1);
    columnCount = Math.max(columnCount, m.c1 + 1);
  }
  if (rowCount + 1 !== outRowStart.length) {
    const grown = new Uint32Array(rowCount + 1);
    grown.set(outRowStart);
    grown.fill(out, outRowStart.length);
    outRowStart = grown;
  }

  return {
    headers,
    sheet: {
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
      formulas: { cells: Uint32Array.from(formulaCells), texts: formulaTexts },
      merges,
      ...(axis === 'rows' ? { rowHeights: toSizeOverrides(movedSizes) } : { columnWidths: toSizeOverrides(movedSizes) }),
      frozen: axis === 'rows' ? { ...sheet.frozen, rows: frozen } : { ...sheet.frozen, columns: frozen },
      ...(tables ? { tables } : {}),
      ...(origin ? { origin } : {}),
    },
  };
}

/**
 * The sheet with every formula passed through `rewrite`. Returns the stored-cell
 * indexes whose formula changed.
 */
export function rewriteFormulas(sheet: WorksheetModel, rewrite: (formula: string) => string): { sheet: WorksheetModel; changed: number[] } {
  const { cells, texts } = sheet.formulas;
  let next: string[] | null = null;
  const changed: number[] = [];
  for (let k = 0; k < texts.length; k++) {
    const updated = rewrite(texts[k]);
    if (updated === texts[k]) continue;
    next ??= texts.slice();
    next[k] = updated;
    changed.push(cells[k]);
  }
  return next ? { sheet: { ...sheet, formulas: { cells, texts: next } }, changed } : { sheet, changed };
}
