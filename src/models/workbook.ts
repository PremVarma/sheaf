import type { StringPool } from './stringPool';

/**
 * Normalized workbook model.
 *
 * UI code only ever sees these types - never SheetJS objects. Cell data is stored
 * in a compressed-sparse-row layout (typed arrays) so that a worksheet with millions
 * of cells stays compact, can be transferred from the parser worker without copying,
 * and gives O(log n) access to any cell and O(k) iteration over a visible row window.
 */

/** Excel's hard limits. Ranges that extend to these bounds mean "whole row/column". */
export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLS = 16_384;

/** Storage-level cell kind. Kept numeric so it fits in a Uint8Array. */
export const CellKind = {
  Empty: 0,
  String: 1,
  Number: 2,
  Date: 3,
  Boolean: 4,
  Error: 5,
} as const;
export type CellKind = (typeof CellKind)[keyof typeof CellKind];

export type CellValueType = 'empty' | 'string' | 'number' | 'date' | 'boolean' | 'error';

export const CELL_VALUE_TYPES: Record<CellKind, CellValueType> = {
  [CellKind.Empty]: 'empty',
  [CellKind.String]: 'string',
  [CellKind.Number]: 'number',
  [CellKind.Date]: 'date',
  [CellKind.Boolean]: 'boolean',
  [CellKind.Error]: 'error',
};

/**
 * A single materialized cell. Built on demand (formula bar, tests, copy) from the
 * compact storage; the grid reads the typed arrays directly instead.
 *
 * A formula is not a value type: `type` describes the cached result and `formula`
 * holds the expression when the cell has one.
 */
export interface CellModel {
  row: number;
  column: number;
  /** Raw value. Dates are Excel serial numbers; booleans are true/false. */
  value: string | number | boolean | null;
  /** Text exactly as the viewer shows it (number formats applied). */
  displayValue: string;
  type: CellValueType;
  formula?: string;
  styleId: number;
}

/**
 * Compressed sparse row storage. The stored cells of row `r` occupy indexes
 * `rowStart[r] .. rowStart[r + 1] - 1`, sorted by column.
 */
export interface SheetCells {
  rowStart: Uint32Array;
  cols: Uint16Array;
  kinds: Uint8Array;
  /** Numeric value for numbers, dates (serial) and booleans (0/1). */
  numbers: Float64Array;
  /** Index into `WorksheetModel.strings` (see stringAt) of the display text. */
  texts: Uint32Array;
  /** Index into `WorkbookModel.styles`. 0 is the default style. */
  styles: Uint16Array;
}

/** Sparse per-index sizes in CSS pixels at 100% zoom, sorted by index. Size 0 = hidden. */
export interface SizeOverrides {
  index: Uint32Array;
  size: Float32Array;
}

/** Inclusive cell range (zero-based). */
export interface CellRange {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export interface CellRef {
  row: number;
  col: number;
}

/**
 * A monotone map from current row (or column) indexes to indexes in the file as
 * it was opened, kept as segments so inserting or deleting rows stays cheap.
 * Segment k covers current indexes [starts[k], starts[k + 1]) (the last one is
 * open-ended) and starts at file index bases[k]; -1 marks inserted entries.
 */
export interface IndexMap {
  readonly starts: readonly number[];
  readonly bases: readonly number[];
}

/** Where an edited sheet came from in the opened file, so saving can keep what the model doesn't hold. */
export interface SheetOrigin {
  /** Index of the sheet in the opened file; null for a sheet added since. */
  base: number | null;
  /** A duplicate of `base`, written to a new part. */
  copy: boolean;
  rows: IndexMap;
  cols: IndexMap;
}

/** An Excel table (ListObject). Kept so row/column edits leave it consistent. */
export interface TableInfo {
  name: string;
  range: CellRange;
  /** Column names, left to right (the header cells' text). */
  columns: string[];
  headerRow: boolean;
}

export type SheetKind = 'worksheet' | 'chart' | 'macro' | 'dialog';
export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';

export interface WorksheetModel {
  /** Stable identity while the workbook is open (survives renaming and reordering). */
  id: number;
  name: string;
  /** Position in the workbook. */
  index: number;
  kind: SheetKind;
  visibility: SheetVisibility;
  /** Rows/columns in the used range (last used index + 1); 0 for an empty sheet. */
  rowCount: number;
  columnCount: number;
  /** Number of stored cells (non-empty or visibly styled). */
  cellCount: number;
  cells: SheetCells;
  /** Deduplicated display strings. Entry 0 is always the empty string. */
  strings: StringPool;
  /** Formulas, keyed by stored-cell index (sorted ascending). */
  formulas: { cells: Uint32Array; texts: string[] };
  merges: CellRange[];
  defaultColumnWidth: number;
  defaultRowHeight: number;
  columnWidths: SizeOverrides;
  rowHeights: SizeOverrides;
  frozen: { rows: number; columns: number };
  showGridlines: boolean;
  tables?: TableInfo[];
  /** Set once the workbook is being edited. */
  origin?: SheetOrigin;
}

export type HorizontalAlign = 'left' | 'center' | 'right' | 'justify' | 'fill' | 'centerContinuous' | 'distributed';
export type VerticalAlign = 'top' | 'center' | 'bottom';

export interface BorderEdge {
  /** Line width in CSS pixels. */
  width: 1 | 2 | 3;
  style: 'solid' | 'dashed' | 'dotted' | 'double';
  /** CSS color, or undefined for "automatic" (the default text color). */
  color?: string;
}

/** Visual cell formatting the viewer can reproduce. All fields optional. */
export interface CellStyle {
  /** Excel number format code (absent = General). */
  numberFormat?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font family, when it isn't the workbook's default font. */
  fontName?: string;
  /** Font color as #rrggbb. */
  color?: string;
  /** Solid fill as #rrggbb. */
  fill?: string;
  /** Font size relative to the workbook's default font (1 = normal). */
  fontScale?: number;
  hAlign?: HorizontalAlign;
  vAlign?: VerticalAlign;
  wrap?: boolean;
  indent?: number;
  borders?: { top?: BorderEdge; right?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge };
}

export type WorkbookFormat = 'xlsx' | 'xlsm' | 'xls' | 'csv' | 'tsv';

export interface WorkbookProperties {
  title?: string;
  subject?: string;
  author?: string;
  lastModifiedBy?: string;
  company?: string;
  application?: string;
  /** ISO 8601 timestamps. */
  created?: string;
  modified?: string;
}

export interface WorkbookModel {
  fileName: string;
  fileSize: number;
  format: WorkbookFormat;
  properties: WorkbookProperties;
  sheets: WorksheetModel[];
  /** Shared style table. Index 0 is the default (empty) style. */
  styles: CellStyle[];
  /** Sheet that was active when the file was saved. */
  activeSheetIndex: number;
  date1904: boolean;
  /** Non-fatal notices, e.g. rows that were not loaded. */
  warnings: string[];
  /** Delimited text details, kept so saving writes the same format back. */
  csv?: { delimiter: string; bom: boolean };
  /** The workbook's default font (font sizes in styles are relative to it). */
  defaultFont?: { name: string; size: number };
}

export const DEFAULT_FONT = { name: 'Calibri', size: 11 } as const;
