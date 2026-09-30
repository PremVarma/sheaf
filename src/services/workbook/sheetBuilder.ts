import { createStringPool } from '../../models/stringPool';
import {
  CellKind,
  type CellRange,
  type SheetKind,
  type SheetVisibility,
  type SizeOverrides,
  type TableInfo,
  type WorksheetModel,
} from '../../models/workbook';

/** Excel's defaults at 100%: 8.43 characters of Calibri 11 and 15pt rows. */
export const DEFAULT_COLUMN_WIDTH = 64;
export const DEFAULT_ROW_HEIGHT = 20;

export interface SheetIdentity {
  name: string;
  index: number;
  /** Defaults to `index`. */
  id?: number;
  kind?: SheetKind;
  visibility?: SheetVisibility;
}

export interface SheetLayout {
  merges?: CellRange[];
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
  columnWidths?: ReadonlyMap<number, number>;
  rowHeights?: ReadonlyMap<number, number>;
  frozen?: { rows: number; columns: number };
  showGridlines?: boolean;
  tables?: TableInfo[];
}

export function toSizeOverrides(sizes: ReadonlyMap<number, number> | undefined): SizeOverrides {
  if (!sizes || sizes.size === 0) return { index: new Uint32Array(0), size: new Float32Array(0) };
  const keys = [...sizes.keys()].sort((a, b) => a - b);
  return { index: Uint32Array.from(keys), size: Float32Array.from(keys, (k) => sizes.get(k)!) };
}

/**
 * Accumulates cells into the compact CSR layout of `WorksheetModel`.
 * Cells must be added in row-major order (rows ascending, columns ascending).
 */
export class SheetBuilder {
  private capacity = 1024;
  private count = 0;
  private cols = new Uint16Array(this.capacity);
  private kinds = new Uint8Array(this.capacity);
  private numbers = new Float64Array(this.capacity);
  private texts = new Uint32Array(this.capacity);
  private styles = new Uint16Array(this.capacity);
  private rowStart = new Uint32Array(1024);
  private currentRow = -1;
  private lastCol = -1;
  private maxCol = -1;
  private readonly stringIds = new Map<string, number>([['', 0]]);
  private readonly strings: string[] = [''];
  private readonly formulaCells: number[] = [];
  private readonly formulaTexts: string[] = [];
  private readonly identity: SheetIdentity;

  constructor(identity: SheetIdentity) {
    this.identity = identity;
  }

  get cellCount(): number {
    return this.count;
  }

  get rowCount(): number {
    return this.currentRow + 1;
  }

  add(row: number, col: number, kind: CellKind, num: number, text: string, style = 0, formula?: string): void {
    if (row < this.currentRow || (row === this.currentRow && col <= this.lastCol)) {
      throw new Error(`Cell ${row},${col} added out of order`);
    }
    if (row > this.currentRow) {
      this.ensureRowCapacity(row + 2);
      for (let r = this.currentRow + 1; r <= row; r++) this.rowStart[r] = this.count;
      this.currentRow = row;
    }
    if (this.count === this.capacity) this.grow();
    const i = this.count++;
    this.cols[i] = col;
    this.kinds[i] = kind;
    this.numbers[i] = num;
    this.texts[i] = this.intern(text);
    this.styles[i] = style;
    if (formula) {
      this.formulaCells.push(i);
      this.formulaTexts.push(formula);
    }
    this.lastCol = col;
    if (col > this.maxCol) this.maxCol = col;
  }

  finish(layout: SheetLayout = {}): WorksheetModel {
    const merges = layout.merges ?? [];
    let rowCount = this.currentRow + 1;
    let columnCount = this.maxCol + 1;
    for (const m of merges) {
      rowCount = Math.max(rowCount, m.r1 + 1);
      columnCount = Math.max(columnCount, m.c1 + 1);
    }

    const n = this.count;
    const rowStart = new Uint32Array(rowCount + 1);
    rowStart.set(this.rowStart.subarray(0, this.currentRow + 1));
    for (let r = this.currentRow + 1; r <= rowCount; r++) rowStart[r] = n;

    return {
      id: this.identity.id ?? this.identity.index,
      name: this.identity.name,
      index: this.identity.index,
      kind: this.identity.kind ?? 'worksheet',
      visibility: this.identity.visibility ?? 'visible',
      rowCount,
      columnCount,
      cellCount: n,
      cells: {
        rowStart,
        cols: this.cols.slice(0, n),
        kinds: this.kinds.slice(0, n),
        numbers: this.numbers.slice(0, n),
        texts: this.texts.slice(0, n),
        styles: this.styles.slice(0, n),
      },
      strings: createStringPool(this.strings),
      formulas: { cells: Uint32Array.from(this.formulaCells), texts: this.formulaTexts },
      merges,
      defaultColumnWidth: layout.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH,
      defaultRowHeight: layout.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
      columnWidths: toSizeOverrides(layout.columnWidths),
      rowHeights: toSizeOverrides(layout.rowHeights),
      frozen: layout.frozen ?? { rows: 0, columns: 0 },
      showGridlines: layout.showGridlines ?? true,
      ...(layout.tables?.length ? { tables: layout.tables } : {}),
    };
  }

  private intern(text: string): number {
    let id = this.stringIds.get(text);
    if (id === undefined) {
      id = this.strings.length;
      this.strings.push(text);
      this.stringIds.set(text, id);
    }
    return id;
  }

  private grow(): void {
    const capacity = this.capacity * 2;
    const extend = <T extends Uint8Array | Uint16Array | Uint32Array | Float64Array>(src: T, make: (n: number) => T): T => {
      const next = make(capacity);
      next.set(src);
      return next;
    };
    this.cols = extend(this.cols, (n) => new Uint16Array(n));
    this.kinds = extend(this.kinds, (n) => new Uint8Array(n));
    this.numbers = extend(this.numbers, (n) => new Float64Array(n));
    this.texts = extend(this.texts, (n) => new Uint32Array(n));
    this.styles = extend(this.styles, (n) => new Uint16Array(n));
    this.capacity = capacity;
  }

  private ensureRowCapacity(size: number): void {
    if (this.rowStart.length >= size) return;
    const next = new Uint32Array(Math.max(size, this.rowStart.length * 2));
    next.set(this.rowStart);
    this.rowStart = next;
  }
}
