import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, type CellRange, type IndexMap } from '../../models/workbook';
import { columnLabel } from '../../utils/cellAddress';
import { fromBase, mapRange } from '../editing/indexMap';
import type { RefEndpoint, ReferenceToken } from './ast';
import { scanReferences } from './parser';

/**
 * Rewriting the references inside formula text: when a formula is copied,
 * when rows/columns are inserted or deleted, and when sheets are renamed,
 * moved or deleted. Only the references change; the rest of the text is kept.
 */

type Reference = Pick<ReferenceToken, 'sheet' | 'kind' | 'from' | 'to'>;

const REF_ERROR = '#REF!';

/** Sheet name as written in a formula: quoted unless it's a plain identifier. */
export function quoteSheetName(name: string): string {
  const plain = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name);
  return plain && !/^(TRUE|FALSE)$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function formatEndpoint(e: RefEndpoint, part: 'cell' | 'row' | 'col'): string {
  const col = `${e.colAbs ? '$' : ''}${columnLabel(e.col)}`;
  const row = `${e.rowAbs ? '$' : ''}${e.row + 1}`;
  return part === 'col' ? col : part === 'row' ? row : col + row;
}

export function formatReference(ref: Reference): string {
  const prefix = ref.sheet !== undefined ? `${quoteSheetName(ref.sheet)}!` : '';
  switch (ref.kind) {
    case 'columns':
      return `${prefix}${formatEndpoint(ref.from, 'col')}:${formatEndpoint(ref.to, 'col')}`;
    case 'rows':
      return `${prefix}${formatEndpoint(ref.from, 'row')}:${formatEndpoint(ref.to, 'row')}`;
    case 'area':
      return `${prefix}${formatEndpoint(ref.from, 'cell')}:${formatEndpoint(ref.to, 'cell')}`;
    default:
      return prefix + formatEndpoint(ref.from, 'cell');
  }
}

/** Replaces references for which `rewrite` returns text; null keeps a reference as written. */
export function rewriteReferences(formula: string, rewrite: (ref: ReferenceToken) => string | null): string {
  if (!formula) return formula;
  let out = '';
  let last = 0;
  for (const ref of scanReferences(formula)) {
    const next = rewrite(ref);
    if (next === null) continue;
    out += formula.slice(last, ref.start) + next;
    last = ref.end;
  }
  return last === 0 ? formula : out + formula.slice(last);
}

const sameName = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

/** Whether a reference points at `sheet` (unqualified references point at the formula's own sheet). */
function refersTo(ref: ReferenceToken, homeSheet: string | null, sheet: string): boolean {
  const target = ref.sheet ?? homeSheet;
  return target !== null && sameName(target, sheet);
}

/** Orders a reference's ends so `from` is the top-left, keeping each part's `$`. */
function normalized(ref: Reference): Reference {
  const from = { ...ref.from };
  const to = { ...ref.to };
  if (from.row > to.row) {
    [from.row, to.row] = [to.row, from.row];
    [from.rowAbs, to.rowAbs] = [to.rowAbs, from.rowAbs];
  }
  if (from.col > to.col) {
    [from.col, to.col] = [to.col, from.col];
    [from.colAbs, to.colAbs] = [to.colAbs, from.colAbs];
  }
  return { ...ref, from, to };
}

const sameEnds = (a: Reference, b: Reference) =>
  a.from.row === b.from.row && a.from.col === b.from.col && a.to.row === b.to.row && a.to.col === b.to.col;

/**
 * Copying a formula by (dRow, dCol): relative parts move, `$` parts stay.
 * References pushed off the sheet become #REF!, as in Excel.
 */
export function translateFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  return rewriteReferences(formula, (ref) => {
    const move = (e: RefEndpoint): RefEndpoint => ({
      ...e,
      row: ref.kind === 'columns' || e.rowAbs ? e.row : e.row + dRow,
      col: ref.kind === 'rows' || e.colAbs ? e.col : e.col + dCol,
    });
    const next = { ...ref, from: move(ref.from), to: move(ref.to) };
    if (sameEnds(ref, next)) return null;
    const valid = (e: RefEndpoint) => e.row >= 0 && e.row < EXCEL_MAX_ROWS && e.col >= 0 && e.col < EXCEL_MAX_COLS;
    return valid(next.from) && valid(next.to) ? formatReference(next) : REF_ERROR;
  });
}

/** How rows/columns moved: maps from current indexes to earlier ones (see indexMap.ts). */
export interface AxisChange {
  rows?: IndexMap;
  cols?: IndexMap;
}

/** A reference after rows/columns moved, or null when everything it pointed at was deleted. */
function remapped(ref: Reference, change: AxisChange): Reference | null {
  const r = normalized(ref);
  const from = { ...r.from };
  const to = { ...r.to };
  if (change.rows && ref.kind !== 'columns') {
    if (ref.kind === 'cell') {
      const row = fromBase(change.rows, r.from.row);
      if (row < 0 || row >= EXCEL_MAX_ROWS) return null;
      from.row = to.row = row;
    } else {
      const span = mapRange(change.rows, r.from.row, r.to.row, EXCEL_MAX_ROWS);
      if (!span) return null;
      [from.row, to.row] = span;
    }
  }
  if (change.cols && ref.kind !== 'rows') {
    if (ref.kind === 'cell') {
      const col = fromBase(change.cols, r.from.col);
      if (col < 0 || col >= EXCEL_MAX_COLS) return null;
      from.col = to.col = col;
    } else {
      const span = mapRange(change.cols, r.from.col, r.to.col, EXCEL_MAX_COLS);
      if (!span) return null;
      [from.col, to.col] = span;
    }
  }
  return { ...r, from, to };
}

/**
 * Updates references to `target` after rows/columns were inserted or deleted
 * there. References to deleted cells become #REF!.
 */
export function remapReferences(formula: string, homeSheet: string, target: string, change: AxisChange): string {
  return rewriteReferences(formula, (ref) => {
    if (!refersTo(ref, homeSheet, target)) return null;
    const next = remapped(ref, change);
    if (!next) return REF_ERROR;
    return sameEnds(normalized(ref), next) ? null : formatReference(next);
  });
}

/** Sheet renamed: qualified references follow it. */
export function renameSheetReferences(formula: string, from: string, to: string): string {
  return rewriteReferences(formula, (ref) =>
    ref.sheet !== undefined && sameName(ref.sheet, from) ? `${quoteSheetName(to)}!${formula.slice(ref.cellStart, ref.end)}` : null,
  );
}

/** Sheet deleted: references to it become #REF!. */
export function invalidateSheetReferences(formula: string, name: string): string {
  return rewriteReferences(formula, (ref) => (ref.sheet !== undefined && sameName(ref.sheet, name) ? REF_ERROR : null));
}

/**
 * Cut and paste: references to cells inside the moved block follow it to
 * `dest` (ignoring `$`), everything else stays, as in Excel.
 */
export function moveReferences(
  formula: string,
  homeSheet: string,
  source: { sheet: string; range: CellRange },
  dest: { sheet: string; dRow: number; dCol: number },
): string {
  const { range } = source;
  return rewriteReferences(formula, (ref) => {
    if (ref.kind === 'columns' || ref.kind === 'rows' || !refersTo(ref, homeSheet, source.sheet)) return null;
    const r = normalized(ref);
    const inside = r.from.row >= range.r0 && r.to.row <= range.r1 && r.from.col >= range.c0 && r.to.col <= range.c1;
    if (!inside) return null;
    const move = (e: RefEndpoint): RefEndpoint => ({ ...e, row: e.row + dest.dRow, col: e.col + dest.dCol });
    // Unqualified references stay unqualified while they point at the formula's own sheet.
    const sheet = ref.sheet === undefined && sameName(dest.sheet, homeSheet) ? undefined : dest.sheet;
    return formatReference({ ...r, sheet, from: move(r.from), to: move(r.to) });
  });
}

/** How the sheets of the opened file map to the workbook being saved. */
export interface SheetChanges {
  /** The current name of a sheet from the file: null when it was deleted, undefined when unknown. */
  currentName(fileName: string): string | null | undefined;
  /** Row/column maps (current → file) of a sheet whose rows or columns moved. */
  axes(fileName: string): AxisChange | null;
}

/**
 * Formulas stored elsewhere in the file (defined names, conditional formats,
 * validations, charts), written against the file's sheet names and positions.
 * `homeSheet` is the file name of the sheet the formula belongs to, if any.
 */
export function remapFileFormula(formula: string, homeSheet: string | null, changes: SheetChanges): string {
  return rewriteReferences(formula, (ref) => {
    const target = ref.sheet ?? homeSheet;
    if (target === null) return null;
    const current = changes.currentName(target);
    if (current === null) return REF_ERROR;
    if (current === undefined) return null;
    const axes = changes.axes(target);
    let next: Reference | null = ref;
    if (axes) {
      next = remapped(ref, axes);
      if (!next) return REF_ERROR;
    }
    const renamed = ref.sheet !== undefined && ref.sheet !== current;
    if (!renamed && (!axes || sameEnds(normalized(ref), next))) return null;
    return formatReference({ ...next, sheet: ref.sheet !== undefined ? current : undefined });
  });
}
