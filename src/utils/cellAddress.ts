import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, type CellRange, type CellRef } from '../models/workbook';

const labelCache: string[] = [];

/** 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD". */
export function columnLabel(index: number): string {
  const cached = labelCache[index];
  if (cached !== undefined) return cached;
  let n = index + 1;
  let label = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  if (index >= 0 && index < EXCEL_MAX_COLS) labelCache[index] = label;
  return label;
}

/** "A" → 0, "xfd" → 16383. Returns -1 for anything that isn't a valid column. */
export function columnIndex(label: string): number {
  if (label.length === 0 || label.length > 3) return -1;
  let n = 0;
  for (let i = 0; i < label.length; i++) {
    const code = label.charCodeAt(i) | 0x20; // fold to lower case
    if (code < 97 || code > 122) return -1;
    n = n * 26 + (code - 96);
  }
  return n <= EXCEL_MAX_COLS ? n - 1 : -1;
}

export function cellAddress(row: number, col: number): string {
  return columnLabel(col) + (row + 1);
}

export function normalizeRange(a: CellRef, b: CellRef): CellRange {
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  };
}

export function isWholeColumns(range: CellRange): boolean {
  return range.r0 === 0 && range.r1 >= EXCEL_MAX_ROWS - 1;
}

export function isWholeRows(range: CellRange): boolean {
  return range.c0 === 0 && range.c1 >= EXCEL_MAX_COLS - 1;
}

export function rangeContains(range: CellRange, row: number, col: number): boolean {
  return row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1;
}

export function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
}

export function rangeCellCount(range: CellRange): number {
  return (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
}

/** "A1", "A1:C5", "B:D" (whole columns), "3:7" (whole rows). */
export function rangeAddress(range: CellRange): string {
  const cols = isWholeColumns(range);
  const rows = isWholeRows(range);
  if (cols) return `${columnLabel(range.c0)}:${columnLabel(range.c1)}`;
  if (rows) return `${range.r0 + 1}:${range.r1 + 1}`;
  const start = cellAddress(range.r0, range.c0);
  if (range.r0 === range.r1 && range.c0 === range.c1) return start;
  return `${start}:${cellAddress(range.r1, range.c1)}`;
}

export interface ParsedReference {
  /** Sheet name when the reference was qualified, e.g. `Sheet2!A1`. */
  sheetName?: string;
  range: CellRange;
}

const CELL_PATTERN = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;
const COLUMN_PATTERN = /^\$?([A-Za-z]{1,3})$/;
const ROW_PATTERN = /^\$?([0-9]{1,7})$/;

function parseCell(text: string): CellRef | null {
  const match = CELL_PATTERN.exec(text.trim());
  if (!match) return null;
  const col = columnIndex(match[1]);
  const row = Number(match[2]) - 1;
  if (col < 0 || row < 0 || row >= EXCEL_MAX_ROWS) return null;
  return { row, col };
}

function parseColumn(text: string): number {
  const match = COLUMN_PATTERN.exec(text.trim());
  return match ? columnIndex(match[1]) : -1;
}

function parseRow(text: string): number {
  const match = ROW_PATTERN.exec(text.trim());
  if (!match) return -1;
  const row = Number(match[1]) - 1;
  return row >= 0 && row < EXCEL_MAX_ROWS ? row : -1;
}

/**
 * Parses Name Box input: `A152`, `$B$7`, `A1:C10`, `B:D`, `3:7`, `Sheet2!A1`,
 * `'My Sheet'!A1:B2`. Returns null when the text isn't a valid reference.
 */
export function parseReference(input: string): ParsedReference | null {
  let text = input.trim();
  if (!text) return null;

  let sheetName: string | undefined;
  const bang = text.lastIndexOf('!');
  if (bang >= 0) {
    let name = text.slice(0, bang).trim();
    if (name.length >= 2 && name.startsWith("'") && name.endsWith("'")) {
      name = name.slice(1, -1).replace(/''/g, "'");
    }
    if (!name) return null;
    sheetName = name;
    text = text.slice(bang + 1).trim();
  }

  const parts = text.split(':');
  if (parts.length === 1) {
    const cell = parseCell(parts[0]);
    return cell ? { sheetName, range: { r0: cell.row, c0: cell.col, r1: cell.row, c1: cell.col } } : null;
  }
  if (parts.length !== 2) return null;

  const [first, second] = parts;
  const a = parseCell(first);
  const b = parseCell(second);
  if (a && b) return { sheetName, range: normalizeRange(a, b) };

  const colA = parseColumn(first);
  const colB = parseColumn(second);
  if (colA >= 0 && colB >= 0) {
    return {
      sheetName,
      range: { r0: 0, r1: EXCEL_MAX_ROWS - 1, c0: Math.min(colA, colB), c1: Math.max(colA, colB) },
    };
  }

  const rowA = parseRow(first);
  const rowB = parseRow(second);
  if (rowA >= 0 && rowB >= 0) {
    return {
      sheetName,
      range: { r0: Math.min(rowA, rowB), r1: Math.max(rowA, rowB), c0: 0, c1: EXCEL_MAX_COLS - 1 },
    };
  }
  return null;
}
