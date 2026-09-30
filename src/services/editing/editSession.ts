import { stringAt } from '../../models/stringPool';
import {
  CellKind,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type BorderEdge,
  type CellRange,
  type CellStyle,
  type SheetVisibility,
  type WorkbookModel,
  type WorksheetModel,
} from '../../models/workbook';
import { isWholeColumns, isWholeRows } from '../../utils/cellAddress';
import { neededRowHeight } from '../../utils/rowHeight';
import {
  invalidateSheetReferences,
  moveReferences,
  remapReferences,
  renameSheetReferences,
  translateFormula,
  type AxisChange,
} from '../formula/references';
import { compareValues, isError, type Scalar } from '../formula/values';
import { formatNumber, isDateFormat } from '../workbook/numberFormat';
import { clampToUsedRange } from '../workbook/sheetService';
import { SheetBuilder } from '../workbook/sheetBuilder';
import { canonicalStyle, MAX_STYLES, styleKey } from '../workbook/styles';
import { extendRun } from './autofill';
import { Calculator } from './calculator';
import {
  applyWrites,
  cellKey,
  globalKey,
  globalLocal,
  globalSheet,
  keyCol,
  keyRow,
  readBlock,
  readStoredCell,
  sameCell,
  type CellBlock,
  type StoredCell,
} from './cells';
import { serializeDelimited } from './export';
import { changeMap, fromBase, IDENTITY } from './indexMap';
import { parseInput } from './inputParser';
import {
  canInsert,
  expandToMerges,
  forEachStoredCell,
  restyleCells,
  rewriteFormulas,
  rowOfCell,
  shiftSheet,
  sizeAt,
  withSizes,
  type Axis,
  type Retext,
} from './sheetOps';
import { writeXlsx } from './xlsxWriter';

export type SaveFormat = 'xlsx' | 'xlsm' | 'csv' | 'tsv';

/** A change that can't be made; the message is shown to the user. */
export class EditRefusedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'EditRefusedError';
  }
}

interface CellChange {
  sheet: number;
  row: number;
  col: number;
  before: StoredCell | null;
  after: StoredCell | null;
}

interface Snapshot {
  workbook: WorkbookModel;
  stale: number[];
}

type HistoryEntry =
  | { kind: 'cells'; id: number; label: string; changes: CellChange[]; weight: number }
  | { kind: 'snapshot'; id: number; label: string; before: Snapshot; after: Snapshot; weight: number };

export interface EditOutcome {
  workbook: WorkbookModel;
  /** Formulas whose value may be out of date because they couldn't be recalculated. */
  staleFormulas: number;
  circular: boolean;
}

export interface TextInput {
  sheet: number;
  row: number;
  col: number;
  text: string;
}

export type BorderSide = 'top' | 'right' | 'bottom' | 'left';

/** Style changes: a value sets a property, null removes it. */
export type StylePatch = { [K in Exclude<keyof CellStyle, 'borders'>]?: CellStyle[K] | null } & {
  borders?: { [E in BorderSide]?: BorderEdge | null } | null;
};

export type MergeMode = 'merge' | 'center' | 'across';
export type PasteMode = 'all' | 'values' | 'formats';

/** What a change needs recalculated. */
interface Recalc {
  /** Cells whose dependents must be recalculated. */
  changed?: number[];
  /** Formula cells to recalculate themselves. */
  formulas?: number[];
  /** Cell positions or sheets moved: rebuild the dependency index. */
  reset?: boolean;
}

const MAX_HISTORY = 200;
/** Memory the undo history may keep alive in old copies of sheets. */
const MAX_HISTORY_BYTES = 384 * 1024 * 1024;
/** Empty cells a formatting change may create inside a selection. */
const MAX_CREATED_CELLS = 250_000;
const INVALID_SHEET_CHARS = /[\\/?*[\]:]/;

/** Applies a style patch. */
export function patchStyle(style: CellStyle, patch: StylePatch): CellStyle {
  const out: Record<string, unknown> = { ...style };
  for (const [key, value] of Object.entries(patch)) {
    if (key === 'borders' || value === undefined) continue;
    if (value === null) delete out[key];
    else out[key] = value;
  }
  if (patch.borders !== undefined) {
    if (patch.borders === null) delete out.borders;
    else {
      const borders: Record<string, BorderEdge> = { ...style.borders };
      for (const [side, edge] of Object.entries(patch.borders)) {
        if (edge === null) delete borders[side];
        else if (edge) borders[side] = edge;
      }
      if (Object.keys(borders).length > 0) out.borders = borders;
      else delete out.borders;
    }
  }
  return out as CellStyle;
}

/** Why a sheet name can't be used, or null when it can. */
export function sheetNameError(name: string, sheets: readonly WorksheetModel[], except = -1): string | null {
  const trimmed = name.trim();
  if (!trimmed) return 'A sheet name can’t be empty.';
  if (trimmed.length > 31) return 'A sheet name can have at most 31 characters.';
  if (INVALID_SHEET_CHARS.test(trimmed)) return 'A sheet name can’t contain \\ / ? * [ ] or :';
  if (trimmed.startsWith("'") || trimmed.endsWith("'")) return 'A sheet name can’t start or end with an apostrophe.';
  if (trimmed.toLowerCase() === 'history') return '“History” is reserved by Excel.';
  const lower = trimmed.toLowerCase();
  if (sheets.some((s, i) => i !== except && s.name.toLowerCase() === lower)) return `There’s already a sheet named “${trimmed}”.`;
  return null;
}

function uniqueSheetName(sheets: readonly WorksheetModel[], stem = 'Sheet'): string {
  const taken = new Set(sheets.map((s) => s.name.toLowerCase()));
  for (let n = sheets.length + 1; ; n++) if (!taken.has(`${stem}${n}`.toLowerCase())) return `${stem}${n}`;
}

function copyName(name: string, sheets: readonly WorksheetModel[]): string {
  const taken = new Set(sheets.map((s) => s.name.toLowerCase()));
  const stem = name.replace(/ \(\d+\)$/, '');
  for (let n = 2; ; n++) {
    const suffix = ` (${n})`;
    const candidate = stem.slice(0, 31 - suffix.length) + suffix;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}

function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
}

function contains(outer: CellRange, inner: CellRange): boolean {
  return inner.r0 >= outer.r0 && inner.r1 <= outer.r1 && inner.c0 >= outer.c0 && inner.c1 <= outer.c1;
}

/** Bytes of sheet data that `before` holds and `after` no longer shares. */
function snapshotWeight(before: WorkbookModel, after: WorkbookModel): number {
  const shared = new Set<unknown>();
  for (const sheet of after.sheets) {
    const { cells } = sheet;
    for (const array of [cells.rowStart, cells.cols, cells.kinds, cells.numbers, cells.texts, cells.styles, sheet.strings.data]) shared.add(array);
  }
  let bytes = 1024;
  for (const sheet of before.sheets) {
    const { cells } = sheet;
    for (const array of [cells.rowStart, cells.cols, cells.kinds, cells.numbers, cells.texts, cells.styles]) {
      if (!shared.has(array)) bytes += array.byteLength;
    }
    if (!shared.has(sheet.strings.data)) bytes += sheet.strings.data.length * 2;
  }
  return bytes;
}

/** Excel's sort order: numbers, text, logical values, errors; blanks always last. */
function sortRank(cell: StoredCell | null): number {
  if (!cell || cell.kind === CellKind.Empty) return 4;
  if (cell.kind === CellKind.Number || cell.kind === CellKind.Date) return 0;
  if (cell.kind === CellKind.String) return 1;
  if (cell.kind === CellKind.Boolean) return 2;
  return 3;
}

const collator = new Intl.Collator(undefined, { sensitivity: 'base', numeric: false });

/**
 * Editing state for one open workbook: applies changes, recalculates
 * dependent formulas, keeps undo/redo history, and builds the file to save.
 */
export class EditSession {
  private workbook: WorkbookModel;
  /** The workbook as it was opened; saving compares against it. */
  private readonly baseline: WorkbookModel;
  /** The .xlsx package it was opened from (null for other formats and new files). */
  private readonly original: Uint8Array | null;
  private calculator: Calculator;
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private nextId = 1;
  private savedId = 0;
  private stale = new Set<number>();
  private styleIds: Map<string, number> | null = null;
  private nextSheetId: number;

  constructor(workbook: WorkbookModel, original: Uint8Array | null) {
    this.baseline = workbook;
    this.original = original;
    this.workbook = {
      ...workbook,
      sheets: workbook.sheets.map((sheet, index) =>
        sheet.origin ? sheet : { ...sheet, origin: { base: index, copy: false, rows: IDENTITY, cols: IDENTITY } },
      ),
    };
    this.nextSheetId = workbook.sheets.reduce((max, sheet) => Math.max(max, sheet.id), -1) + 1;
    this.calculator = new Calculator(() => this.workbook);
  }

  get current(): WorkbookModel {
    return this.workbook;
  }

  get dirty(): boolean {
    return (this.undoStack[this.undoStack.length - 1]?.id ?? 0) !== this.savedId;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get staleFormulas(): number {
    return this.stale.size;
  }

  get undoLabel(): string | undefined {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }

  get redoLabel(): string | undefined {
    return this.redoStack[this.redoStack.length - 1]?.label;
  }

  // Cell contents

  /** Sets cells from typed text (formulas, numbers, dates, text…). */
  setText(inputs: TextInput[], label = 'Typing'): EditOutcome {
    const changes: CellChange[] = [];
    for (const input of inputs) {
      const sheet = this.workbook.sheets[input.sheet];
      if (!sheet) continue;
      const before = readStoredCell(sheet, input.row, input.col);
      const styleId = before?.style ?? 0;
      const parsed = parseInput(input.text, this.workbook.styles[styleId], this.workbook.date1904);
      let after: StoredCell | null;
      if (parsed) {
        const style = parsed.suggestedFormat ? this.styleWith(styleId, { numberFormat: parsed.suggestedFormat }) : styleId;
        after = { kind: parsed.kind, num: parsed.num, text: parsed.text, style, ...(parsed.formula ? { formula: parsed.formula } : {}) };
      } else {
        // Clearing keeps the cell's formatting, like Excel.
        after = styleId !== 0 ? { kind: CellKind.Empty, num: 0, text: '', style: styleId } : null;
      }
      if (!sameCell(before, after)) changes.push({ sheet: input.sheet, row: input.row, col: input.col, before, after });
    }
    return this.commit(changes, label);
  }

  /** Clears the contents (not the formatting) of a range. */
  clear(sheetIndex: number, range: CellRange): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    const changes: CellChange[] = [];
    if (sheet) {
      forEachStoredCell(sheet, range, (_i, row, col) => {
        const before = readStoredCell(sheet, row, col);
        if (!before || (before.kind === CellKind.Empty && !before.formula)) return;
        const after = before.style !== 0 ? { kind: CellKind.Empty, num: 0, text: '', style: before.style } : null;
        changes.push({ sheet: sheetIndex, row, col, before, after });
      });
    }
    return this.commit(changes, 'Clear');
  }

  /** Copies a range for pasting inside the app. */
  readBlock(sheetIndex: number, range: CellRange): CellBlock | null {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return null;
    return readBlock(sheet, range);
  }

  /**
   * Pastes copied cells. A single cell fills the whole selection, and a block
   * repeats when the selection is a multiple of its size, as in Excel.
   */
  paste(sheetIndex: number, target: CellRange, block: CellBlock, mode: PasteMode = 'all'): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    const rows = block.range.r1 - block.range.r0 + 1;
    const cols = block.range.c1 - block.range.c0 + 1;
    const height = target.r1 - target.r0 + 1;
    const width = target.c1 - target.c0 + 1;
    const tiles = height % rows === 0 && width % cols === 0 && height * width <= 1_000_000;
    const area = tiles
      ? target
      : { r0: target.r0, c0: target.c0, r1: Math.min(EXCEL_MAX_ROWS - 1, target.r0 + rows - 1), c1: Math.min(EXCEL_MAX_COLS - 1, target.c0 + cols - 1) };

    const changes: CellChange[] = [];
    for (let r = area.r0; r <= area.r1; r++) {
      for (let c = area.c0; c <= area.c1; c++) {
        const sr = (r - area.r0) % rows;
        const sc = (c - area.c0) % cols;
        const source = block.cells[sr * cols + sc];
        const before = readStoredCell(sheet, r, c);
        let after: StoredCell | null;
        if (mode === 'values') {
          after = source && source.kind !== CellKind.Empty ? this.withStyle({ kind: source.kind, num: source.num, text: source.text, style: 0 }, before?.style ?? 0) : before && before.style !== 0 ? { kind: CellKind.Empty, num: 0, text: '', style: before.style } : null;
        } else if (mode === 'formats') {
          const style = source?.style ?? 0;
          after = before ? this.withStyle(before, style) : style !== 0 ? { kind: CellKind.Empty, num: 0, text: '', style } : null;
        } else if (source) {
          const formula = source.formula ? translateFormula(source.formula, r - (block.range.r0 + sr), c - (block.range.c0 + sc)) : undefined;
          after = { kind: source.kind, num: source.num, text: source.text, style: source.style, ...(formula ? { formula } : {}) };
        } else {
          after = null;
        }
        if (!sameCell(before, after)) changes.push({ sheet: sheetIndex, row: r, col: c, before, after });
      }
    }

    // Merged cells in the paste area are replaced by the copied ones.
    const clash = sheet.merges.find((m) => rangesIntersect(m, area) && !contains(area, m));
    if (clash) throw new EditRefusedError('You can’t paste over part of a merged cell.');
    const replaced = mode === 'values' ? [] : sheet.merges.filter((m) => contains(area, m));
    const added: CellRange[] = [];
    if (mode !== 'values') {
      for (let tr = 0; tr < (area.r1 - area.r0 + 1) / rows; tr++) {
        for (let tc = 0; tc < (area.c1 - area.c0 + 1) / cols; tc++) {
          for (const m of block.merges) {
            const r0 = area.r0 + tr * rows;
            const c0 = area.c0 + tc * cols;
            added.push({ r0: r0 + m.r0, c0: c0 + m.c0, r1: r0 + m.r1, c1: c0 + m.c1 });
          }
        }
      }
    }
    const label = mode === 'values' ? 'Paste Values' : mode === 'formats' ? 'Paste Formatting' : 'Paste';
    if (replaced.length === 0 && added.length === 0) return this.commit(changes, label);
    return this.snapshotEdit(label, () => {
      const keys = this.writeChanges(changes);
      const current = this.workbook.sheets[sheetIndex];
      this.setSheet(sheetIndex, { ...current, merges: [...current.merges.filter((m) => !replaced.includes(m)), ...added] });
      return { changed: keys };
    });
  }

  /**
   * Cut and paste: moves cells with their formatting. Formulas anywhere that
   * pointed at the moved cells follow them.
   */
  move(sourceSheet: number, source: CellRange, destSheet: number, destRow: number, destCol: number): EditOutcome {
    const from = this.workbook.sheets[sourceSheet];
    const to = this.workbook.sheets[destSheet];
    if (!from || !to) return this.outcome(false);
    const dRow = destRow - source.r0;
    const dCol = destCol - source.c0;
    if (dRow === 0 && dCol === 0 && sourceSheet === destSheet) return this.outcome(false);
    const dest = { r0: destRow, c0: destCol, r1: destRow + source.r1 - source.r0, c1: destCol + source.c1 - source.c0 };
    if (dest.r1 >= EXCEL_MAX_ROWS || dest.c1 >= EXCEL_MAX_COLS) throw new EditRefusedError('The cells don’t fit there.');
    const block = readBlock(from, source);
    const sourceMerges = from.merges.filter((m) => contains(source, m));
    if (from.merges.some((m) => rangesIntersect(m, source) && !contains(source, m)) || to.merges.some((m) => rangesIntersect(m, dest) && !contains(dest, m) && !(sourceSheet === destSheet && contains(source, m)))) {
      throw new EditRefusedError('You can’t move part of a merged cell.');
    }

    return this.snapshotEdit('Move', () => {
      const changes: CellChange[] = [];
      forEachStoredCell(from, source, (_i, row, col) => {
        if (sourceSheet === destSheet && row >= dest.r0 && row <= dest.r1 && col >= dest.c0 && col <= dest.c1) return;
        changes.push({ sheet: sourceSheet, row, col, before: readStoredCell(from, row, col), after: null });
      });
      this.writeChanges(changes);
      const target = this.workbook.sheets[destSheet];
      const writes: CellChange[] = [];
      for (let r = 0; r <= source.r1 - source.r0; r++) {
        for (let c = 0; c <= source.c1 - source.c0; c++) {
          writes.push({ sheet: destSheet, row: dest.r0 + r, col: dest.c0 + c, before: readStoredCell(target, dest.r0 + r, dest.c0 + c), after: block.cells[r * (source.c1 - source.c0 + 1) + c] });
        }
      }
      this.writeChanges(writes);

      // References to the moved cells follow them.
      const sourceName = from.name;
      const destName = to.name;
      const formulas: number[] = [];
      const sheets = this.workbook.sheets.map((sheet, i) => {
        const { sheet: next, changed } = rewriteFormulas(sheet, (f) =>
          moveReferences(f, sheet.name, { sheet: sourceName, range: source }, { sheet: destName, dRow, dCol }),
        );
        for (const cell of changed) formulas.push(globalKey(i, rowOfCell(next, cell), next.cells.cols[cell]));
        return next;
      });
      this.workbook = { ...this.workbook, sheets };
      const moved = sourceMerges.map((m) => ({ r0: m.r0 + dRow, c0: m.c0 + dCol, r1: m.r1 + dRow, c1: m.c1 + dCol }));
      const fromSheet = this.workbook.sheets[sourceSheet];
      this.setSheet(sourceSheet, { ...fromSheet, merges: fromSheet.merges.filter((m) => !sourceMerges.includes(m)) });
      const toSheet = this.workbook.sheets[destSheet];
      this.setSheet(destSheet, { ...toSheet, merges: [...toSheet.merges.filter((m) => !contains(dest, m)), ...moved] });
      this.remapStale((s, r, c) =>
        s === sourceSheet && r >= source.r0 && r <= source.r1 && c >= source.c0 && c <= source.c1 ? [destSheet, r + dRow, c + dCol] : [s, r, c],
      );
      return { reset: true, formulas, changed: [...changes, ...writes].map((c) => globalKey(c.sheet, c.row, c.col)) };
    });
  }

  /** Fill down/right (copy) or the fill handle (series), from `source` into the adjacent `target`. */
  fill(sheetIndex: number, source: CellRange, target: CellRange, series: boolean): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    const vertical = target.c0 === source.c0 && target.c1 === source.c1;
    const forward = vertical ? target.r0 > source.r1 : target.c0 > source.c1;
    const changes: CellChange[] = [];
    const lines = vertical ? [source.c0, source.c1] : [source.r0, source.r1];
    for (let line = lines[0]; line <= lines[1]; line++) {
      // The run in fill order (upwards/leftwards fills read the source backwards).
      const positions: number[] = [];
      const [a, b] = vertical ? [source.r0, source.r1] : [source.c0, source.c1];
      for (let p = a; p <= b; p++) positions.push(p);
      if (!forward) positions.reverse();
      const at = (p: number) => (vertical ? readStoredCell(sheet, p, line) : readStoredCell(sheet, line, p));
      const run = positions.map(at);
      const targets: number[] = [];
      const [t0, t1] = vertical ? [target.r0, target.r1] : [target.c0, target.c1];
      for (let p = t0; p <= t1; p++) targets.push(p);
      if (!forward) targets.reverse();
      const format = (num: number, from: number) => formatNumber(num, this.workbook.styles[run[from]?.style ?? 0]?.numberFormat, this.workbook.date1904);
      const filled = extendRun(run, targets.length, series, format, this.workbook.date1904);
      targets.forEach((p, k) => {
        const { from, value } = filled[k];
        const src = run[from];
        const row = vertical ? p : line;
        const col = vertical ? line : p;
        const before = at(p);
        let after: StoredCell | null = null;
        if (src) {
          if (value) after = { ...value, style: src.style };
          else {
            const dist = p - positions[from];
            const formula = src.formula ? translateFormula(src.formula, vertical ? dist : 0, vertical ? 0 : dist) : undefined;
            after = { kind: src.kind, num: src.num, text: src.text, style: src.style, ...(formula ? { formula } : {}) };
          }
        }
        if (!sameCell(before, after)) changes.push({ sheet: sheetIndex, row, col, before, after });
      });
    }
    return this.commit(changes, series ? 'AutoFill' : vertical ? 'Fill Down' : 'Fill Right');
  }

  /** Sorts the rows of a range by one column. */
  sort(sheetIndex: number, range: CellRange, keyCol: number, ascending: boolean): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet || range.r1 <= range.r0) return this.outcome(false);
    if (sheet.merges.some((m) => rangesIntersect(m, range))) throw new EditRefusedError('Ranges with merged cells can’t be sorted. Unmerge them first.');
    const width = range.c1 - range.c0 + 1;
    const rows: (StoredCell | null)[][] = [];
    for (let r = range.r0; r <= range.r1; r++) {
      const cells: (StoredCell | null)[] = [];
      for (let c = range.c0; c <= range.c1; c++) cells.push(readStoredCell(sheet, r, c));
      rows.push(cells);
    }
    const valueOf = (cell: StoredCell | null): Scalar => {
      if (!cell) return null;
      if (cell.kind === CellKind.Number || cell.kind === CellKind.Date) return cell.num;
      if (cell.kind === CellKind.Boolean) return cell.num === 1;
      return cell.text;
    };
    const key = keyCol - range.c0;
    const order = rows.map((_, i) => i);
    order.sort((a, b) => {
      const x = rows[a][key];
      const y = rows[b][key];
      const rx = sortRank(x);
      const ry = sortRank(y);
      if (rx === 4 || ry === 4) return rx === ry ? a - b : rx - ry; // blanks last either way
      let cmp = rx !== ry ? rx - ry : rx === 1 ? collator.compare(x!.text, y!.text) : rx === 3 ? 0 : compareValues(valueOf(x), valueOf(y));
      if (!ascending) cmp = -cmp;
      return cmp || a - b;
    });
    const changes: CellChange[] = [];
    order.forEach((from, to) => {
      for (let k = 0; k < width; k++) {
        const source = rows[from][k];
        const before = rows[to][k];
        const formula = source?.formula ? translateFormula(source.formula, to - from, 0) : undefined;
        const after = source ? { ...source, ...(formula ? { formula } : {}) } : null;
        if (!sameCell(before, after)) changes.push({ sheet: sheetIndex, row: range.r0 + to, col: range.c0 + k, before, after });
      }
    });
    return this.commit(changes, 'Sort');
  }

  // Formatting

  /** Changes the formatting of a range; `patch` may depend on the cell (borders). */
  format(sheetIndex: number, range: CellRange, patch: StylePatch | ((row: number, col: number) => StylePatch | null), label = 'Format Cells'): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    const whole = isWholeRows(range) || isWholeColumns(range);
    const target = whole ? clampToUsedRange(range, sheet) : expandToMerges(sheet, range);
    const patchAt = typeof patch === 'function' ? patch : () => patch;
    const cache = new Map<string, number>();
    const styleOf = (style: number, row: number, col: number): number => {
      const p = patchAt(row, col);
      if (!p) return style;
      const key = `${style}|${JSON.stringify(p)}`;
      let id = cache.get(key);
      if (id === undefined) {
        id = this.internStyle(patchStyle(this.workbook.styles[style] ?? {}, p));
        cache.set(key, id);
      }
      return id;
    };
    // Bigger text and wrapping make rows grow to fit, as in Excel.
    const affectsHeight = typeof patch === 'function' || 'fontScale' in patch || 'wrap' in patch || 'fontName' in patch;
    return this.snapshotEdit(label, () => {
      let next = restyleCells(sheet, target, styleOf, this.retexter(sheet));
      next = this.styleEmptyCells(next, target, styleOf);
      if (affectsHeight) next = this.growRows(next, target);
      this.setSheet(sheetIndex, next);
    });
  }

  /** Rows in a range made taller where their text no longer fits (never shorter). */
  private growRows(sheet: WorksheetModel, range: CellRange): WorksheetModel {
    const last = Math.min(range.r1, sheet.rowCount - 1, range.r0 + 5000);
    const width = (col: number) => sizeAt(sheet.columnWidths, col) ?? sheet.defaultColumnWidth;
    const grown = new Map<number, number>();
    for (let r = range.r0; r <= last; r++) {
      const current = sizeAt(sheet.rowHeights, r) ?? sheet.defaultRowHeight;
      if (current === 0) continue;
      const needed = neededRowHeight(sheet, r, this.workbook.styles, width);
      if (needed > current) grown.set(r, needed);
    }
    if (grown.size === 0) return sheet;
    let heights = sheet.rowHeights;
    grown.forEach((height, row) => {
      heights = withSizes(heights, [row], height);
    });
    return { ...sheet, rowHeights: heights };
  }

  /** Removes all formatting (and merged cells) from a range. */
  clearFormats(sheetIndex: number, range: CellRange): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    const target = isWholeRows(range) || isWholeColumns(range) ? clampToUsedRange(range, sheet) : expandToMerges(sheet, range);
    return this.snapshotEdit('Clear Formats', () => {
      const next = restyleCells(sheet, target, () => 0, this.retexter(sheet));
      // Formatting-only cells go away entirely.
      const writes = new Map<number, StoredCell | null>();
      forEachStoredCell(next, target, (i, row, col) => {
        if (next.cells.kinds[i] === CellKind.Empty && !readStoredCell(next, row, col)?.formula) writes.set(cellKey(row, col), null);
      });
      const merges = next.merges.filter((m) => !rangesIntersect(m, target));
      this.setSheet(sheetIndex, { ...applyWrites(next, writes), merges });
    });
  }

  /** Whether merging would discard values other than the top-left one (Excel warns). */
  mergeLosesData(sheetIndex: number, range: CellRange, mode: MergeMode): boolean {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return false;
    const target = expandToMerges(sheet, range);
    let lost = false;
    forEachStoredCell(sheet, target, (i, row, col) => {
      const anchor = mode === 'across' ? col === target.c0 : row === target.r0 && col === target.c0;
      if (!anchor && sheet.cells.kinds[i] !== CellKind.Empty) lost = true;
    });
    return lost;
  }

  merge(sheetIndex: number, range: CellRange, mode: MergeMode): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    const target = expandToMerges(sheet, range);
    const merges: CellRange[] = [];
    if (mode === 'across') {
      if (target.c1 > target.c0) for (let r = target.r0; r <= target.r1; r++) merges.push({ r0: r, c0: target.c0, r1: r, c1: target.c1 });
    } else if (target.r1 > target.r0 || target.c1 > target.c0) {
      merges.push(target);
    }
    if (merges.length === 0) return this.outcome(false);
    return this.snapshotEdit(mode === 'center' ? 'Merge & Center' : 'Merge Cells', () => {
      // Only the top-left value of each merged cell is kept; formatting stays.
      const changes: CellChange[] = [];
      for (const m of merges) {
        forEachStoredCell(sheet, m, (_i, row, col) => {
          if (row === m.r0 && col === m.c0) return;
          const before = readStoredCell(sheet, row, col);
          if (!before || (before.kind === CellKind.Empty && !before.formula)) return;
          changes.push({ sheet: sheetIndex, row, col, before, after: before.style ? { kind: CellKind.Empty, num: 0, text: '', style: before.style } : null });
        });
      }
      const keys = this.writeChanges(changes);
      const current = this.workbook.sheets[sheetIndex];
      let next: WorksheetModel = { ...current, merges: [...current.merges.filter((m) => !rangesIntersect(m, target)), ...merges] };
      if (mode === 'center') {
        const center = (style: number) => this.styleWith(style, { hAlign: 'center' });
        next = this.styleEmptyCells(restyleCells(next, target, center), { r0: target.r0, c0: target.c0, r1: target.r0, c1: target.c0 }, center);
      }
      this.setSheet(sheetIndex, next);
      return { changed: keys };
    });
  }

  unmerge(sheetIndex: number, range: CellRange): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet || !sheet.merges.some((m) => rangesIntersect(m, range))) return this.outcome(false);
    return this.snapshotEdit('Unmerge Cells', () => {
      this.setSheet(sheetIndex, { ...sheet, merges: sheet.merges.filter((m) => !rangesIntersect(m, range)) });
    });
  }

  // Rows and columns

  insertRows(sheetIndex: number, at: number, count: number): EditOutcome {
    return this.shift(sheetIndex, 'rows', 'insert', at, count);
  }

  deleteRows(sheetIndex: number, at: number, count: number): EditOutcome {
    return this.shift(sheetIndex, 'rows', 'delete', at, count);
  }

  insertColumns(sheetIndex: number, at: number, count: number): EditOutcome {
    return this.shift(sheetIndex, 'cols', 'insert', at, count);
  }

  deleteColumns(sheetIndex: number, at: number, count: number): EditOutcome {
    return this.shift(sheetIndex, 'cols', 'delete', at, count);
  }

  /** Column widths in px at 100% (null = default width, 0 = hidden). */
  setColumnWidths(sheetIndex: number, cols: Iterable<number>, width: number | null, label = 'Column Width'): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    return this.snapshotEdit(label, () => this.setSheet(sheetIndex, { ...sheet, columnWidths: withSizes(sheet.columnWidths, cols, width) }));
  }

  setRowHeights(sheetIndex: number, rows: Iterable<number>, height: number | null, label = 'Row Height'): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet) return this.outcome(false);
    return this.snapshotEdit(label, () => this.setSheet(sheetIndex, { ...sheet, rowHeights: withSizes(sheet.rowHeights, rows, height) }));
  }

  setFrozen(sheetIndex: number, rows: number, columns: number): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet || (sheet.frozen.rows === rows && sheet.frozen.columns === columns)) return this.outcome(false);
    return this.snapshotEdit(rows || columns ? 'Freeze Panes' : 'Unfreeze Panes', () => this.setSheet(sheetIndex, { ...sheet, frozen: { rows, columns } }));
  }

  setGridlines(sheetIndex: number, show: boolean): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    if (!sheet || sheet.showGridlines === show) return this.outcome(false);
    return this.snapshotEdit('Gridlines', () => this.setSheet(sheetIndex, { ...sheet, showGridlines: show }));
  }

  // Sheets

  addSheet(at: number, name?: string): EditOutcome {
    const sheets = this.workbook.sheets;
    const sheetName = name ?? uniqueSheetName(sheets);
    const error = sheetNameError(sheetName, sheets);
    if (error) throw new EditRefusedError(error);
    const blank = new SheetBuilder({ name: sheetName.trim(), index: at, id: this.nextSheetId++ }).finish();
    blank.origin = { base: null, copy: false, rows: IDENTITY, cols: IDENTITY };
    const index = Math.max(0, Math.min(at, sheets.length));
    return this.sheetsEdit('Insert Sheet', [...sheets.slice(0, index), blank, ...sheets.slice(index)]);
  }

  renameSheet(index: number, name: string): EditOutcome {
    const sheets = this.workbook.sheets;
    const sheet = sheets[index];
    const trimmed = name.trim();
    if (!sheet || sheet.name === trimmed) return this.outcome(false);
    const error = sheetNameError(trimmed, sheets, index);
    if (error) throw new EditRefusedError(error);
    const renamed = sheets.map((s) => rewriteFormulas(s, (f) => renameSheetReferences(f, sheet.name, trimmed)).sheet);
    renamed[index] = { ...renamed[index], name: trimmed };
    return this.sheetsEdit('Rename Sheet', renamed);
  }

  deleteSheet(index: number): EditOutcome {
    const sheets = this.workbook.sheets;
    const sheet = sheets[index];
    if (!sheet) return this.outcome(false);
    if (sheet.visibility === 'visible' && sheets.filter((s) => s.visibility === 'visible').length === 1) {
      throw new EditRefusedError('A workbook must contain at least one visible sheet.');
    }
    return this.sheetsEdit('Delete Sheet', sheets.filter((_, i) => i !== index), () => {
      const formulas: number[] = [];
      const updated = this.workbook.sheets.map((s, i) => {
        const { sheet: next, changed } = rewriteFormulas(s, (f) => invalidateSheetReferences(f, sheet.name));
        for (const cell of changed) formulas.push(globalKey(i, rowOfCell(next, cell), next.cells.cols[cell]));
        return next;
      });
      this.workbook = { ...this.workbook, sheets: updated };
      return { formulas, changed: formulas };
    });
  }

  moveSheet(from: number, to: number): EditOutcome {
    const sheets = this.workbook.sheets.slice();
    if (from === to || !sheets[from] || to < 0 || to >= sheets.length) return this.outcome(false);
    const [sheet] = sheets.splice(from, 1);
    sheets.splice(to, 0, sheet);
    return this.sheetsEdit('Move Sheet', sheets);
  }

  duplicateSheet(index: number): EditOutcome {
    const sheets = this.workbook.sheets;
    const source = sheets[index];
    if (!source) return this.outcome(false);
    const copy: WorksheetModel = {
      ...source,
      id: this.nextSheetId++,
      name: copyName(source.name, sheets),
      visibility: 'visible',
      tables: undefined,
      origin: source.origin && source.origin.base !== null ? { ...source.origin, copy: true } : { base: null, copy: false, rows: IDENTITY, cols: IDENTITY },
    };
    return this.sheetsEdit('Duplicate Sheet', [...sheets.slice(0, index + 1), copy, ...sheets.slice(index + 1)]);
  }

  setSheetVisibility(index: number, visibility: SheetVisibility): EditOutcome {
    const sheets = this.workbook.sheets;
    const sheet = sheets[index];
    if (!sheet || sheet.visibility === visibility) return this.outcome(false);
    if (visibility !== 'visible' && sheets.filter((s) => s.visibility === 'visible').length === 1) {
      throw new EditRefusedError('A workbook must contain at least one visible sheet.');
    }
    const next = sheets.slice();
    next[index] = { ...sheet, visibility };
    return this.sheetsEdit(visibility === 'visible' ? 'Unhide Sheet' : 'Hide Sheet', next);
  }

  // History

  undo(): EditOutcome | null {
    const entry = this.undoStack.pop();
    if (!entry) return null;
    this.redoStack.push(entry);
    if (entry.kind === 'snapshot') return this.restore(entry.before);
    return this.apply([...entry.changes].reverse().map((c) => ({ ...c, after: c.before, before: c.after })));
  }

  redo(): EditOutcome | null {
    const entry = this.redoStack.pop();
    if (!entry) return null;
    this.undoStack.push(entry);
    if (entry.kind === 'snapshot') return this.restore(entry.after);
    return this.apply(entry.changes);
  }

  // Saving

  /** Bytes of the file to save. `activeSheet` is the sheet written for CSV/TSV and shown first in Excel. */
  buildFile(format: SaveFormat, activeSheet: number): Uint8Array {
    if (format === 'csv' || format === 'tsv') {
      const sheet = this.workbook.sheets[activeSheet];
      const delimiter = format === 'tsv' ? '\t' : (this.workbook.csv?.delimiter ?? ',');
      const bom = this.workbook.csv?.bom ?? /[^\x00-\x7f]/.test(sheet.strings.data);
      return serializeDelimited(sheet, delimiter === '\t' && format === 'csv' ? ',' : delimiter, bom);
    }
    return writeXlsx({
      original: this.original,
      baseline: this.original ? this.baseline : null,
      workbook: { ...this.workbook, activeSheetIndex: activeSheet },
    });
  }

  /** Call after the file was written. */
  markSaved(workbook: Partial<WorkbookModel> = {}): WorkbookModel {
    this.savedId = this.undoStack[this.undoStack.length - 1]?.id ?? 0;
    this.workbook = { ...this.workbook, ...workbook };
    return this.workbook;
  }

  // Internals

  private outcome(circular: boolean): EditOutcome {
    return { workbook: this.workbook, staleFormulas: this.stale.size, circular };
  }

  private push(entry: HistoryEntry): void {
    this.undoStack.push(entry);
    this.redoStack = [];
    let total = 0;
    for (const e of this.undoStack) total += e.weight;
    while (this.undoStack.length > 1 && (this.undoStack.length > MAX_HISTORY || total > MAX_HISTORY_BYTES)) {
      total -= this.undoStack.shift()!.weight;
    }
  }

  private commit(changes: CellChange[], label: string): EditOutcome {
    if (changes.length === 0) return this.outcome(false);
    this.push({ kind: 'cells', id: this.nextId++, label, changes, weight: changes.length * 160 });
    return this.apply(changes);
  }

  /**
   * An edit recorded as before/after copies of the workbook (cheap: unchanged
   * sheets and arrays are shared). Used for structural and formatting changes.
   */
  private snapshotEdit(label: string, mutate: () => Recalc | void): EditOutcome {
    const before: Snapshot = { workbook: this.workbook, stale: [...this.stale] };
    let request: Recalc;
    try {
      request = mutate() ?? {};
    } catch (error) {
      this.workbook = before.workbook;
      this.stale = new Set(before.stale);
      throw error;
    }
    if (this.workbook === before.workbook) return this.outcome(false);
    if (request.reset) this.calculator = new Calculator(() => this.workbook);
    const changed = request.changed ?? [];
    const formulas = new Set(request.formulas ?? []);
    for (const key of changed) {
      const local = globalLocal(key);
      const sheet = this.workbook.sheets[globalSheet(key)];
      const formula = sheet ? readStoredCell(sheet, keyRow(local), keyCol(local))?.formula : undefined;
      if (!request.reset) this.calculator.formulaChanged(globalSheet(key), keyRow(local), keyCol(local), formula);
      if (formula) formulas.add(key);
      else this.stale.delete(key);
    }
    const circular = formulas.size > 0 || changed.length > 0 ? this.recalc(changed, [...formulas]) : false;
    const after: Snapshot = { workbook: this.workbook, stale: [...this.stale] };
    this.push({ kind: 'snapshot', id: this.nextId++, label, before, after, weight: snapshotWeight(before.workbook, after.workbook) });
    return this.outcome(circular);
  }

  private restore(snapshot: Snapshot): EditOutcome {
    const current = this.workbook;
    // Styles are only ever appended, so the longest list serves every state.
    const styles = current.styles.length >= snapshot.workbook.styles.length ? current.styles : snapshot.workbook.styles;
    if (styles !== current.styles) this.styleIds = null;
    this.workbook = {
      ...snapshot.workbook,
      styles,
      fileName: current.fileName,
      fileSize: current.fileSize,
      format: current.format,
      csv: current.csv,
    };
    this.stale = new Set(snapshot.stale);
    this.calculator = new Calculator(() => this.workbook);
    return this.outcome(false);
  }

  private setSheet(index: number, sheet: WorksheetModel): void {
    if (this.workbook.sheets[index] === sheet) return;
    const sheets = this.workbook.sheets.slice();
    sheets[index] = sheet;
    this.workbook = { ...this.workbook, sheets };
  }

  private internStyle(style: CellStyle): number {
    if (!this.styleIds) {
      this.styleIds = new Map();
      this.workbook.styles.forEach((s, i) => {
        const key = styleKey(s);
        if (!this.styleIds!.has(key)) this.styleIds!.set(key, i);
      });
    }
    const canonical = canonicalStyle(style);
    const key = JSON.stringify(canonical);
    let id = this.styleIds.get(key);
    if (id === undefined) {
      const styles = this.workbook.styles;
      if (styles.length > MAX_STYLES) throw new EditRefusedError('This workbook has too many different cell formats to add another.');
      id = styles.length;
      this.workbook = { ...this.workbook, styles: [...styles, canonical] };
      this.styleIds.set(key, id);
    }
    return id;
  }

  private styleWith(styleId: number, patch: StylePatch): number {
    return this.internStyle(patchStyle(this.workbook.styles[styleId] ?? {}, patch));
  }

  /** A cell given another style, its text reformatted when the number format changes. */
  private withStyle(cell: StoredCell, style: number): StoredCell {
    if (cell.style === style) return cell;
    const next = { ...cell, style };
    if (cell.kind === CellKind.Number || cell.kind === CellKind.Date) {
      const format = this.workbook.styles[style]?.numberFormat;
      next.kind = isDateFormat(format) ? CellKind.Date : CellKind.Number;
      next.text = formatNumber(cell.num, format, this.workbook.date1904);
    }
    return next;
  }

  /** Reformats numbers whose number format changes with their style. */
  private retexter(sheet: WorksheetModel): Retext {
    const { kinds, numbers, styles } = sheet.cells;
    return (i, style) => {
      const kind = kinds[i];
      if (kind !== CellKind.Number && kind !== CellKind.Date) return null;
      const format = this.workbook.styles[style]?.numberFormat;
      if (format === this.workbook.styles[styles[i]]?.numberFormat) return null;
      return { kind: isDateFormat(format) ? CellKind.Date : CellKind.Number, text: formatNumber(numbers[i], format, this.workbook.date1904) };
    };
  }

  /** Creates formatted empty cells where a selection has no cells yet (so borders and fills show). */
  private styleEmptyCells(sheet: WorksheetModel, range: CellRange, styleOf: (style: number, row: number, col: number) => number): WorksheetModel {
    const r1 = Math.min(range.r1, EXCEL_MAX_ROWS - 1);
    const c1 = Math.min(range.c1, EXCEL_MAX_COLS - 1);
    if ((r1 - range.r0 + 1) * (c1 - range.c0 + 1) > MAX_CREATED_CELLS) return sheet;
    const writes = new Map<number, StoredCell | null>();
    const { rowStart, cols } = sheet.cells;
    for (let r = range.r0; r <= r1; r++) {
      let i = r < sheet.rowCount ? rowStart[r] : 0;
      const end = r < sheet.rowCount ? rowStart[r + 1] : 0;
      for (let c = range.c0; c <= c1; c++) {
        while (i < end && cols[i] < c) i++;
        if (i < end && cols[i] === c) continue;
        const style = styleOf(0, r, c);
        if (style !== 0) writes.set(cellKey(r, c), { kind: CellKind.Empty, num: 0, text: '', style });
      }
    }
    return applyWrites(sheet, writes);
  }

  /** Writes cell changes and keeps the dependency index in step. Returns their keys. */
  private writeChanges(changes: CellChange[]): number[] {
    const writes = new Map<number, Map<number, StoredCell | null>>();
    for (const change of changes) {
      let cells = writes.get(change.sheet);
      if (!cells) writes.set(change.sheet, (cells = new Map()));
      cells.set(cellKey(change.row, change.col), change.after);
    }
    this.writeCells(writes);
    return changes.map((c) => globalKey(c.sheet, c.row, c.col));
  }

  private writeCells(writes: Map<number, Map<number, StoredCell | null>>): void {
    if (writes.size === 0) return;
    const sheets = this.workbook.sheets.slice();
    writes.forEach((cells, index) => {
      sheets[index] = applyWrites(sheets[index], cells);
    });
    this.workbook = { ...this.workbook, sheets };
  }

  /** Writes cell states, then recalculates everything that depends on them. */
  private apply(changes: CellChange[]): EditOutcome {
    this.writeChanges(changes);
    for (const change of changes) {
      this.calculator.formulaChanged(change.sheet, change.row, change.col, change.after?.formula);
    }
    const changedKeys = changes.map((c) => globalKey(c.sheet, c.row, c.col));
    const formulaKeys = changes.filter((c) => c.after?.formula).map((c) => globalKey(c.sheet, c.row, c.col));
    // Cells that no longer hold a formula can't be stale.
    for (const change of changes) if (!change.after?.formula) this.stale.delete(globalKey(change.sheet, change.row, change.col));
    const circular = this.recalc(changedKeys, formulaKeys);
    return this.outcome(circular);
  }

  private recalc(changedKeys: number[], formulaKeys: number[]): boolean {
    const result = this.calculator.recalculate(changedKeys, formulaKeys);
    const results = new Map<number, Map<number, StoredCell | null>>();
    result.values.forEach((value, key) => {
      this.stale.delete(key);
      const sheetIndex = globalSheet(key);
      const local = globalLocal(key);
      const sheet = this.workbook.sheets[sheetIndex];
      const existing = sheet && readStoredCell(sheet, keyRow(local), keyCol(local));
      if (!existing?.formula) return;
      const next = this.formulaResult(existing, value);
      if (sameCell(existing, next)) return;
      let cells = results.get(sheetIndex);
      if (!cells) results.set(sheetIndex, (cells = new Map()));
      cells.set(local, next);
    });
    for (const key of result.unsupported) this.stale.add(key);
    this.writeCells(results);
    return result.circular;
  }

  private formulaResult(cell: StoredCell, value: Scalar): StoredCell {
    const format = this.workbook.styles[cell.style]?.numberFormat;
    const base = { style: cell.style, formula: cell.formula };
    if (value === null) value = 0; // a formula pointing at an empty cell shows 0
    if (typeof value === 'number') {
      return { ...base, kind: isDateFormat(format) ? CellKind.Date : CellKind.Number, num: value, text: formatNumber(value, format, this.workbook.date1904) };
    }
    if (typeof value === 'boolean') return { ...base, kind: CellKind.Boolean, num: value ? 1 : 0, text: value ? 'TRUE' : 'FALSE' };
    if (isError(value)) return { ...base, kind: CellKind.Error, num: 0, text: value.message };
    return { ...base, kind: value === '' ? CellKind.Empty : CellKind.String, num: 0, text: value };
  }

  private remapStale(map: (sheet: number, row: number, col: number) => [number, number, number] | null): void {
    const next = new Set<number>();
    for (const key of this.stale) {
      const local = globalLocal(key);
      const moved = map(globalSheet(key), keyRow(local), keyCol(local));
      if (moved && moved[1] >= 0 && moved[2] >= 0) next.add(globalKey(...moved));
    }
    this.stale = next;
  }

  /** Replaces the sheet list (add, delete, move, rename…), renumbering sheets. */
  private sheetsEdit(label: string, sheets: WorksheetModel[], after?: () => Recalc): EditOutcome {
    return this.snapshotEdit(label, () => {
      const renumbered = sheets.map((sheet, i) => (sheet.index === i ? sheet : { ...sheet, index: i }));
      const newIndex = new Map(renumbered.map((sheet, i) => [sheet.id, i]));
      const oldSheets = this.workbook.sheets;
      this.remapStale((s, r, c) => {
        const index = newIndex.get(oldSheets[s]?.id ?? -1);
        return index === undefined ? null : [index, r, c];
      });
      this.workbook = { ...this.workbook, sheets: renumbered };
      return { ...(after?.() ?? {}), reset: true };
    });
  }

  private shift(sheetIndex: number, axis: Axis, kind: 'insert' | 'delete', at: number, count: number): EditOutcome {
    const sheet = this.workbook.sheets[sheetIndex];
    const limit = axis === 'rows' ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS;
    if (!sheet || count <= 0 || at < 0 || at >= limit) return this.outcome(false);
    count = Math.min(count, limit - at);
    const noun = axis === 'rows' ? 'rows' : 'columns';
    if (kind === 'insert' && !canInsert(sheet, axis, at, count)) {
      throw new EditRefusedError(`To avoid losing data, ${noun} can’t be inserted here: cells with data would be pushed off the end of the sheet.`);
    }
    if (kind === 'delete' && axis === 'rows') {
      for (const table of sheet.tables ?? []) {
        const header = table.headerRow && table.range.r0 >= at && table.range.r0 < at + count;
        const whole = table.range.r0 >= at && table.range.r1 < at + count;
        if (header && !whole) throw new EditRefusedError(`The header row of table “${table.name}” can’t be deleted on its own.`);
      }
    }
    const map = changeMap(kind, at, count);
    const change: AxisChange = axis === 'rows' ? { rows: map } : { cols: map };
    const label = `${kind === 'insert' ? 'Insert' : 'Delete'} ${axis === 'rows' ? 'Rows' : 'Columns'}`;
    return this.snapshotEdit(label, () => {
      const shifted = shiftSheet(sheet, axis, kind, at, count, map);
      const formulas: number[] = [];
      const sheets = this.workbook.sheets.map((s, i) => {
        const base = i === sheetIndex ? shifted.sheet : s;
        const { sheet: next, changed } = rewriteFormulas(base, (f) => remapReferences(f, s.name, sheet.name, change));
        for (const cell of changed) formulas.push(globalKey(i, rowOfCell(next, cell), next.cells.cols[cell]));
        return next;
      });
      this.workbook = { ...this.workbook, sheets };
      if (kind === 'insert' && at > 0) this.copyFormatting(sheetIndex, axis, at, count);
      if (shifted.headers.length > 0) {
        const target = this.workbook.sheets[sheetIndex];
        const writes = new Map<number, StoredCell | null>();
        for (const [row, col, text] of shifted.headers) {
          writes.set(cellKey(row, col), { kind: CellKind.String, num: 0, text, style: readStoredCell(target, row, col)?.style ?? 0 });
        }
        this.setSheet(sheetIndex, applyWrites(target, writes));
      }
      this.remapStale((s, r, c) => {
        if (s !== sheetIndex) return [s, r, c];
        return axis === 'rows' ? [s, fromBase(map, r), c] : [s, r, fromBase(map, c)];
      });
      return { reset: true, formulas, changed: formulas };
    });
  }

  /** Inserted rows/columns take the formatting of the row above / column to the left, like Excel. */
  private copyFormatting(sheetIndex: number, axis: Axis, at: number, count: number): void {
    let sheet = this.workbook.sheets[sheetIndex];
    const source = at - 1;
    const writes = new Map<number, StoredCell | null>();
    const range = axis === 'rows' ? { r0: source, r1: source, c0: 0, c1: EXCEL_MAX_COLS - 1 } : { r0: 0, r1: EXCEL_MAX_ROWS - 1, c0: source, c1: source };
    forEachStoredCell(sheet, range, (i, row, col) => {
      const style = sheet.cells.styles[i];
      if (style === 0 || writes.size > 200_000) return;
      for (let k = 0; k < count; k++) {
        const key = axis === 'rows' ? cellKey(at + k, col) : cellKey(row, at + k);
        writes.set(key, { kind: CellKind.Empty, num: 0, text: '', style });
      }
    });
    sheet = applyWrites(sheet, writes);
    const sizes = axis === 'rows' ? sheet.rowHeights : sheet.columnWidths;
    const size = sizeAt(sizes, source);
    if (size !== undefined && size > 0) {
      const indexes = Array.from({ length: count }, (_, k) => at + k);
      sheet = axis === 'rows' ? { ...sheet, rowHeights: withSizes(sizes, indexes, size) } : { ...sheet, columnWidths: withSizes(sizes, indexes, size) };
    }
    this.setSheet(sheetIndex, sheet);
  }
}

/** Display text of a stored cell (for tests and callers without a sheet). */
export function cellText(sheet: WorksheetModel, index: number): string {
  return stringAt(sheet.strings, sheet.cells.texts[index]);
}
