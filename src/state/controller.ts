import {
  DEFAULT_FONT,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type BorderEdge,
  type CellRange,
  type CellRef,
  type CellStyle,
  type HorizontalAlign,
  type VerticalAlign,
  type WorkbookModel,
  type WorksheetModel,
} from '../models/workbook';
import { buildCopyPayload } from '../services/clipboard/clipboardService';
import type { BorderSide, EditOutcome, EditSession, MergeMode, PasteMode, SaveFormat, StylePatch } from '../services/editing';
import { readBlock, type CellBlock } from '../services/editing/cells';
import { editTextAt } from '../services/editing/editText';
import { changeDecimals, currencyFormat, formatPresets, localCurrency } from '../services/editing/numberFormats';
import type { Platform, SaveFileType } from '../services/platform/types';
import { addRecentFile, clearRecentFiles, removeRecentFile } from '../services/recentFiles';
import { firstMatchFrom, searchWorkbook, stepMatch } from '../services/search/searchService';
import { isCancelled, toSaveError, toUserFacingError, WorkbookError } from '../services/workbook/errors';
import { getExtension } from '../services/workbook/formatDetection';
import { clampToUsedRange, findCellIndex, getCell, hasValueAt, mergeAt, sortRange } from '../services/workbook/sheetService';
import { loadWorkbook, type ParseFunction, type WorkbookSource } from '../services/workbook/workbookService';
import { isWholeColumns, isWholeRows, normalizeRange, parseReference } from '../utils/cellAddress';
import { baseName, formatCount } from '../utils/format';
import { isCommandId, type CommandId } from './commands';
import type { Store } from './store';
import {
  MAX_ZOOM,
  MIN_ZOOM,
  ZOOM_LEVELS,
  type CellEditState,
  type DialogState,
  type SearchState,
  type Selection,
  type SheetViewState,
  type ViewerAction,
  type ViewerState,
} from './viewerState';

export const APP_NAME = 'Sheaf';

type EditingModule = typeof import('../services/editing');
export type CommitMove = 'down' | 'up' | 'right' | 'left' | 'none';

export type BorderPreset =
  | 'bottom'
  | 'top'
  | 'left'
  | 'right'
  | 'none'
  | 'all'
  | 'outside'
  | 'inside'
  | 'thickOutside'
  | 'thickBottom'
  | 'doubleBottom'
  | 'topBottom';

/** Font sizes offered in the toolbar, in points. */
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** Copied cells larger than this are pasted as text only. */
const MAX_BLOCK_CELLS = 250_000;

const EMPTY_STYLE: CellStyle = Object.freeze({});

/** Splits clipboard text (tab-separated, Excel-style quoting) into rows of cells. */
export function parseClipboardTable(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let atFieldStart = true;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (ch === '"') quoted = false;
      else field += ch;
      continue;
    }
    if (ch === '"' && atFieldStart) {
      quoted = true;
      atFieldStart = false;
    } else if (ch === '\t') {
      row.push(field);
      field = '';
      atFieldStart = true;
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      atFieldStart = true;
    } else {
      field += ch;
      atFieldStart = false;
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** The style shown for the active cell (merged cells use their top-left cell). */
export function activeCellStyle(state: ViewerState): CellStyle {
  const { workbook, activeSheet } = state;
  const view = state.sheetViews[activeSheet];
  const sheet = workbook?.sheets[activeSheet];
  if (!workbook || !sheet || !view) return EMPTY_STYLE;
  let { row, col } = view.selection.anchor;
  const merge = mergeAt(sheet, row, col);
  if (merge) {
    row = merge.r0;
    col = merge.c0;
  }
  const i = findCellIndex(sheet, row, col);
  return (i >= 0 ? workbook.styles[sheet.cells.styles[i]] : undefined) ?? EMPTY_STYLE;
}

function borderPatch(range: CellRange, preset: BorderPreset, color?: string): (row: number, col: number) => StylePatch | null {
  const edge = (width: BorderEdge['width'], style: BorderEdge['style'] = 'solid'): BorderEdge => (color ? { width, style, color } : { width, style });
  const thin = edge(1);
  return (row, col) => {
    const top = row === range.r0;
    const bottom = row === range.r1;
    const left = col === range.c0;
    const right = col === range.c1;
    const set = (sides: Partial<Record<BorderSide, BorderEdge | null>>): StylePatch | null => (Object.keys(sides).length > 0 ? { borders: sides } : null);
    switch (preset) {
      case 'none':
        return { borders: null };
      case 'all':
        return set({ top: thin, right: thin, bottom: thin, left: thin });
      case 'outside':
      case 'thickOutside': {
        const e = preset === 'outside' ? thin : edge(2);
        return set({ ...(top ? { top: e } : {}), ...(bottom ? { bottom: e } : {}), ...(left ? { left: e } : {}), ...(right ? { right: e } : {}) });
      }
      case 'inside':
        return set({ ...(!top ? { top: thin } : {}), ...(!bottom ? { bottom: thin } : {}), ...(!left ? { left: thin } : {}), ...(!right ? { right: thin } : {}) });
      case 'bottom':
        return bottom ? set({ bottom: thin }) : null;
      case 'top':
        return top ? set({ top: thin }) : null;
      case 'left':
        return left ? set({ left: thin }) : null;
      case 'right':
        return right ? set({ right: thin }) : null;
      case 'thickBottom':
        return bottom ? set({ bottom: edge(2) }) : null;
      case 'doubleBottom':
        return bottom ? set({ bottom: edge(3, 'double') }) : null;
      case 'topBottom':
        return set({ ...(top ? { top: thin } : {}), ...(bottom ? { bottom: thin } : {}) });
    }
  };
}

interface CopiedCells {
  text: string;
  block: CellBlock | null;
  sheetId: number;
  range: CellRange;
  cut: boolean;
}

/**
 * Application logic that isn't rendering: opening, editing and saving files,
 * commands, search navigation, clipboard and desktop integration. Components
 * call into this; it only talks to the store and the platform.
 */
export class ViewerController {
  readonly store: Store<ViewerState, ViewerAction>;
  readonly platform: Platform;
  private readonly parse?: ParseFunction;
  private loadAbort: AbortController | null = null;
  private searchTimer: ReturnType<typeof setTimeout> | undefined;
  private lastKeyboardCommand = { id: '', time: 0 };
  private editingModule: Promise<EditingModule> | null = null;
  /** Editing state for the open workbook, created on the first edit. */
  private session: EditSession | null = null;
  /** Bytes of the open file, kept so .xlsx edits can be saved into the original package. */
  private fileBytes: Uint8Array | null = null;
  private untitledCount = 0;
  private copied: CopiedCells | null = null;
  private dialogs = new Map<number, (ok: boolean) => void>();
  private nextDialogId = 1;
  /** Saving as CSV was confirmed for this document (it drops formatting and other sheets). */
  private textSaveConfirmed = false;
  /** Set by the sheet tabs so "Rename Sheet" can start editing the tab. */
  requestRename: ((index: number) => void) | null = null;

  constructor(store: Store<ViewerState, ViewerAction>, platform: Platform, parse?: ParseFunction) {
    this.store = store;
    this.platform = platform;
    this.parse = parse;
  }

  private get state(): ViewerState {
    return this.store.getState();
  }

  private dispatch(action: ViewerAction): void {
    this.store.dispatch(action);
  }

  activeView(): SheetViewState | null {
    return this.state.sheetViews[this.state.activeSheet] ?? null;
  }

  private status(text: string): void {
    this.dispatch({ type: 'status/show', text });
  }

  // Dialogs

  /** Asks a question in a modal dialog; resolves true when confirmed. */
  confirm(dialog: Omit<DialogState, 'id' | 'kind'>): Promise<boolean> {
    return this.showDialog({ ...dialog, kind: 'confirm' });
  }

  async alert(title: string, message: string): Promise<void> {
    await this.showDialog({ title, message, confirmLabel: 'OK', kind: 'alert' });
  }

  private showDialog(dialog: Omit<DialogState, 'id'>): Promise<boolean> {
    const previous = this.state.dialog;
    if (previous) this.resolveDialog(previous.id, false);
    const id = this.nextDialogId++;
    this.dispatch({ type: 'dialog/open', dialog: { ...dialog, id } });
    return new Promise((resolve) => this.dialogs.set(id, resolve));
  }

  resolveDialog(id: number, ok: boolean): void {
    const resolve = this.dialogs.get(id);
    this.dialogs.delete(id);
    this.dispatch({ type: 'dialog/close', id });
    resolve?.(ok);
    if (!this.state.dialog) this.focusGrid();
  }

  // Files

  async openDialog(): Promise<void> {
    let source: WorkbookSource | null;
    try {
      source = await this.platform.pickWorkbookFile();
    } catch (error) {
      this.dispatch({ type: 'error/show', error: toUserFacingError(new WorkbookError('READ_FAILED', String(error))) });
      return;
    }
    if (source) await this.open(source);
  }

  openPath(path: string): Promise<void> {
    return this.open(this.platform.sourceFromPath(path));
  }

  async open(source: WorkbookSource): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    this.loadAbort?.abort();
    const abort = new AbortController();
    this.loadAbort = abort;
    this.dispatch({ type: 'load/start', fileName: source.name, size: source.size });

    try {
      const { workbook, bytes } = await loadWorkbook(source, {
        signal: abort.signal,
        parse: this.parse,
        onStage: (stage) => {
          if (this.loadAbort === abort) this.dispatch({ type: 'load/stage', stage });
        },
      });
      if (this.loadAbort !== abort) return;
      this.loadAbort = null;
      this.resetDocument(bytes);
      this.dispatch({ type: 'load/success', workbook, path: source.path ?? null });
      this.documentChanged();
      if (source.path) this.setRecentFiles(addRecentFile(this.state.recentFiles, source.path));
      if (workbook.warnings.length > 0) this.status(workbook.warnings.join(' '));
      if (this.state.search.query) this.runSearch();
    } catch (error) {
      if (this.loadAbort !== abort) return; // superseded or cancelled
      this.loadAbort = null;
      if (isCancelled(error)) {
        this.dispatch({ type: 'load/cancel' });
        return;
      }
      // Technical detail goes to the console only; the user sees a friendly message.
      console.warn(`Could not open ${source.name}:`, error instanceof WorkbookError ? `[${error.code}] ${error.message}` : error);
      this.dispatch({ type: 'load/failure', error: toUserFacingError(error, source.name) });
      if (source.path && error instanceof WorkbookError && error.code === 'NOT_FOUND') {
        this.setRecentFiles(removeRecentFile(this.state.recentFiles, source.path));
      }
    }
  }

  private resetDocument(bytes: Uint8Array | null, session: EditSession | null = null): void {
    this.session = session;
    this.fileBytes = bytes;
    this.copied = null;
    this.textSaveConfirmed = false;
  }

  cancelLoad(): void {
    if (!this.loadAbort) return;
    this.loadAbort.abort();
    this.loadAbort = null;
    this.dispatch({ type: 'load/cancel' });
  }

  async closeWorkbook(): Promise<void> {
    if (!this.state.workbook || !(await this.confirmDiscard())) return;
    this.resetDocument(null);
    this.dispatch({ type: 'workbook/close' });
    this.documentChanged();
  }

  async newWorkbook(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    const { createBlankWorkbook, EditSession } = await this.loadEditing();
    this.untitledCount++;
    const workbook = createBlankWorkbook(this.untitledCount === 1 ? 'Untitled.xlsx' : `Untitled ${this.untitledCount}.xlsx`);
    const session = new EditSession(workbook, null);
    this.resetDocument(null, session);
    this.dispatch({ type: 'load/success', workbook: session.current, path: null });
    this.documentChanged();
  }

  /** The app is closing (window close button / quit): save or discard changes first. */
  async quit(): Promise<void> {
    if (await this.confirmDiscard()) this.platform.quit();
  }

  dismissError(): void {
    this.dispatch({ type: 'error/dismiss' });
  }

  private setRecentFiles(paths: string[]): void {
    this.dispatch({ type: 'recent/set', paths });
    this.platform.setRecentFiles(paths);
  }

  // Commands

  runCommand(id: CommandId, source: 'keyboard' | 'menu' | 'ui' = 'ui'): void {
    // On some platforms a shortcut reaches both the page and the native menu.
    const now = Date.now();
    if (source === 'menu' && this.lastKeyboardCommand.id === id && now - this.lastKeyboardCommand.time < 300) return;
    if (source === 'keyboard') this.lastKeyboardCommand = { id, time: now };

    const view = this.activeView();
    const textField = typeof document !== 'undefined' ? document.activeElement : null;
    const inTextField = textField instanceof HTMLInputElement || textField instanceof HTMLTextAreaElement;
    switch (id) {
      case 'app.quit':
        void this.quit();
        break;
      case 'file.new':
        void this.newWorkbook();
        break;
      case 'file.open':
        void this.openDialog();
        break;
      case 'file.close':
        void this.closeWorkbook();
        break;
      case 'file.save':
        void this.save();
        break;
      case 'file.saveAs':
        void this.saveAs();
        break;
      case 'edit.undo':
        if (inTextField && source === 'menu') document.execCommand('undo');
        else void this.undo();
        break;
      case 'edit.redo':
        if (inTextField && source === 'menu') document.execCommand('redo');
        else void this.redo();
        break;
      case 'edit.cut':
        void this.cutSelection();
        break;
      case 'edit.paste':
        void this.pasteFromClipboard();
        break;
      case 'edit.pasteValues':
        void this.pasteFromClipboard('values');
        break;
      case 'edit.clear':
        void this.clearSelection();
        break;
      case 'edit.copy':
        void this.copySelection();
        break;
      case 'edit.selectAll':
        this.selectAll();
        break;
      case 'edit.fillDown':
        void this.fill('down');
        break;
      case 'edit.fillRight':
        void this.fill('right');
        break;
      case 'edit.find':
        this.dispatch({ type: 'replace/open', open: false });
        this.dispatch({ type: 'focus/request', target: 'search' });
        break;
      case 'edit.replace':
        if (!this.state.workbook) break;
        this.dispatch({ type: 'replace/open', open: true });
        this.dispatch({ type: 'focus/request', target: this.state.search.query ? 'replace' : 'search' });
        break;
      case 'edit.findNext':
        this.findNext(1);
        break;
      case 'edit.findPrevious':
        this.findNext(-1);
        break;
      case 'edit.goTo':
        if (this.state.workbook) this.dispatch({ type: 'focus/request', target: 'nameBox' });
        break;
      case 'insert.cells':
        void this.insertCells();
        break;
      case 'insert.rowsAbove':
        void this.insertRows('above');
        break;
      case 'insert.rowsBelow':
        void this.insertRows('below');
        break;
      case 'insert.columnsLeft':
        void this.insertColumns('left');
        break;
      case 'insert.columnsRight':
        void this.insertColumns('right');
        break;
      case 'insert.sheet':
        void this.addSheet();
        break;
      case 'insert.date':
        this.insertNow('date');
        break;
      case 'insert.time':
        this.insertNow('time');
        break;
      case 'delete.cells':
        void this.deleteCells();
        break;
      case 'delete.rows':
        void this.deleteRows();
        break;
      case 'delete.columns':
        void this.deleteColumns();
        break;
      case 'delete.sheet':
        void this.deleteSheet(this.state.activeSheet);
        break;
      case 'format.bold':
        void this.toggleStyle('bold');
        break;
      case 'format.italic':
        void this.toggleStyle('italic');
        break;
      case 'format.underline':
        void this.toggleStyle('underline');
        break;
      case 'format.strikethrough':
        void this.toggleStyle('strike');
        break;
      case 'format.increaseFont':
        void this.stepFontSize(1);
        break;
      case 'format.decreaseFont':
        void this.stepFontSize(-1);
        break;
      case 'format.alignLeft':
        void this.setAlignment('left');
        break;
      case 'format.alignCenter':
        void this.setAlignment('center');
        break;
      case 'format.alignRight':
        void this.setAlignment('right');
        break;
      case 'format.wrap':
        void this.toggleWrap();
        break;
      case 'format.mergeCenter':
        void this.toggleMergeCenter();
        break;
      case 'format.unmerge':
        void this.unmerge();
        break;
      case 'format.numberGeneral':
        void this.setNumberFormat(undefined);
        break;
      case 'format.numberNumber':
        void this.setNumberFormat('#,##0.00');
        break;
      case 'format.numberCurrency':
        void this.setNumberFormat(currencyFormat(localCurrency()));
        break;
      case 'format.numberPercent':
        void this.setNumberFormat('0%');
        break;
      case 'format.numberDate':
        void this.setNumberFormat(formatPresets().find((p) => p.id === 'shortDate')?.format);
        break;
      case 'format.increaseDecimals':
        void this.stepDecimals(1);
        break;
      case 'format.decreaseDecimals':
        void this.stepDecimals(-1);
        break;
      case 'format.clear':
        void this.clearFormats();
        break;
      case 'format.hideRows':
        void this.setRowsHidden(true);
        break;
      case 'format.unhideRows':
        void this.setRowsHidden(false);
        break;
      case 'format.hideColumns':
        void this.setColumnsHidden(true);
        break;
      case 'format.unhideColumns':
        void this.setColumnsHidden(false);
        break;
      case 'format.renameSheet':
        this.requestRename?.(this.state.activeSheet);
        break;
      case 'data.sortAscending':
        void this.sort(true);
        break;
      case 'data.sortDescending':
        void this.sort(false);
        break;
      case 'view.zoomIn':
        this.stepZoom(1);
        break;
      case 'view.zoomOut':
        this.stepZoom(-1);
        break;
      case 'view.zoomReset':
        this.setZoom(1);
        break;
      case 'view.freezeTopRow':
        void this.freeze(1, 0);
        break;
      case 'view.freezeFirstColumn':
        void this.freeze(0, 1);
        break;
      case 'view.freezeAtSelection':
        if (view) void this.freeze(view.selection.anchor.row, view.selection.anchor.col);
        break;
      case 'view.unfreeze':
        void this.freeze(0, 0);
        break;
      case 'view.toggleGridlines':
        void this.toggleGridlines();
        break;
      case 'view.nextSheet':
        this.stepSheet(1);
        break;
      case 'view.previousSheet':
        this.stepSheet(-1);
        break;
      case 'view.toggleInfo':
        this.dispatch({ type: 'info/set', open: !this.state.infoOpen });
        break;
      case 'recent.clear':
        this.setRecentFiles(clearRecentFiles());
        break;
    }
  }

  // Sheets

  activateSheet(index: number): void {
    if (index === this.state.activeSheet) return;
    if (this.state.editing) void this.commitEdit('none');
    this.dispatch({ type: 'sheet/activate', index });
    if (this.state.search.query && this.state.search.scope === 'sheet') this.runSearch();
  }

  /** Moves to the next/previous visible sheet (Excel doesn't wrap around). */
  stepSheet(direction: 1 | -1): void {
    const { workbook, activeSheet } = this.state;
    if (!workbook) return;
    const order = workbook.sheets.filter((s) => s.visibility === 'visible' || s.index === activeSheet).map((s) => s.index);
    const next = order[order.indexOf(activeSheet) + direction];
    if (next !== undefined) this.activateSheet(next);
  }

  /** Adds a blank sheet after the active one and shows it. */
  async addSheet(): Promise<void> {
    if (!this.state.workbook) return;
    await this.commitEdit('none');
    const at = this.state.activeSheet + 1;
    if (await this.edit((s) => s.addSheet(at))) this.activateSheet(at);
  }

  /** Renames a sheet. Returns a message when the name can't be used. */
  async renameSheet(index: number, name: string): Promise<string | null> {
    const workbook = this.state.workbook;
    if (!workbook?.sheets[index] || workbook.sheets[index].name === name.trim()) return null;
    const { sheetNameError } = await this.loadEditing();
    const error = sheetNameError(name, workbook.sheets, index);
    if (error) return error;
    await this.edit((s) => s.renameSheet(index, name));
    return null;
  }

  async deleteSheet(index: number): Promise<void> {
    const sheet = this.state.workbook?.sheets[index];
    if (!sheet) return;
    await this.commitEdit('none');
    if (sheet.cellCount > 0) {
      const ok = await this.confirm({
        title: `Delete “${sheet.name}”?`,
        message: 'The sheet and everything on it will be deleted. Formulas that use it will show #REF!.',
        confirmLabel: 'Delete',
        cancelLabel: 'Cancel',
        destructive: true,
      });
      if (!ok) return;
    }
    await this.edit((s) => s.deleteSheet(index));
  }

  async duplicateSheet(index: number): Promise<void> {
    const sheet = this.state.workbook?.sheets[index];
    if (!sheet) return;
    if (sheet.kind !== 'worksheet') {
      void this.alert('Can’t duplicate this sheet', `“${sheet.name}” isn’t a worksheet, so it can’t be copied here.`);
      return;
    }
    await this.commitEdit('none');
    if (await this.edit((s) => s.duplicateSheet(index))) this.activateSheet(index + 1);
  }

  async moveSheet(from: number, to: number): Promise<void> {
    await this.edit((s) => s.moveSheet(from, to));
  }

  async setSheetHidden(index: number, hidden: boolean): Promise<void> {
    await this.commitEdit('none');
    if ((await this.edit((s) => s.setSheetVisibility(index, hidden ? 'hidden' : 'visible'))) && !hidden) this.activateSheet(index);
  }

  // Selection and navigation

  setSelection(selection: Selection, reveal: boolean | 'focus' = false): void {
    this.dispatch({ type: 'selection/set', selection, reveal });
  }

  selectAll(): void {
    if (!this.state.workbook) return;
    this.setSelection({ anchor: { row: 0, col: 0 }, focus: { row: EXCEL_MAX_ROWS - 1, col: EXCEL_MAX_COLS - 1 } });
  }

  /** Name Box / Go To. Returns an error message, or null on success. */
  goToReference(text: string): string | null {
    const { workbook } = this.state;
    if (!workbook) return null;
    const reference = parseReference(text);
    if (!reference) return 'Enter a cell reference such as A152 or B2:D10.';
    if (reference.sheetName !== undefined) {
      const wanted = reference.sheetName.toLowerCase();
      const index = workbook.sheets.findIndex((s) => s.name.toLowerCase() === wanted);
      if (index < 0) return `There’s no sheet named “${reference.sheetName}”.`;
      this.activateSheet(index);
    }
    const { r0, c0, r1, c1 } = reference.range;
    this.setSelection({ anchor: { row: r0, col: c0 }, focus: { row: r1, col: c1 } }, true);
    this.dispatch({ type: 'focus/request', target: 'grid' });
    return null;
  }

  focusGrid(): void {
    this.dispatch({ type: 'focus/request', target: 'grid' });
  }

  // Zoom

  setZoom(zoom: number): void {
    this.dispatch({ type: 'zoom/set', zoom });
  }

  stepZoom(direction: 1 | -1): void {
    const zoom = this.state.zoom;
    const next =
      direction > 0
        ? (ZOOM_LEVELS.find((level) => level > zoom + 0.001) ?? MAX_ZOOM)
        : ([...ZOOM_LEVELS].reverse().find((level) => level < zoom - 0.001) ?? MIN_ZOOM);
    this.setZoom(next);
  }

  // View

  async freeze(rows: number, columns: number): Promise<void> {
    const sheet = this.state.activeSheet;
    await this.edit((s) => s.setFrozen(sheet, rows, columns), { keepClipboard: true });
  }

  async toggleGridlines(): Promise<void> {
    const sheet = this.state.workbook?.sheets[this.state.activeSheet];
    if (!sheet) return;
    const index = this.state.activeSheet;
    await this.edit((s) => s.setGridlines(index, !sheet.showGridlines), { keepClipboard: true });
  }

  /** Live size while a column divider is dragged (committed with resizeColumns). */
  setColumnSize(col: number, size: number | null): void {
    this.dispatch({ type: 'view/columnSize', col, size });
  }

  setRowSize(row: number, size: number | null): void {
    this.dispatch({ type: 'view/rowSize', row, size });
  }

  /** Sets column widths; resizing one of several selected columns resizes them all, like Excel. */
  async resizeColumns(sizes: ReadonlyMap<number, number | null>): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    let entries = [...sizes];
    if (entries.length === 1 && isWholeColumns(target.range)) {
      const [col, size] = entries[0];
      if (col >= target.range.c0 && col <= target.range.c1 && target.range.c1 - target.range.c0 < 1000) {
        entries = [];
        for (let c = target.range.c0; c <= target.range.c1; c++) entries.push([c, size]);
      }
    }
    const groups = new Map<number | null, number[]>();
    for (const [col, size] of entries) groups.set(size, [...(groups.get(size) ?? []), col]);
    await this.edit((s) => {
      let outcome: EditOutcome | null = null;
      groups.forEach((cols, size) => {
        outcome = s.setColumnWidths(target.sheet, cols, size);
      });
      return outcome;
    }, { keepClipboard: true });
    this.dispatch({ type: 'view/clearSizes' });
  }

  async resizeRows(row: number, size: number | null): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const rows: number[] = [];
    if (isWholeRows(target.range) && row >= target.range.r0 && row <= target.range.r1 && target.range.r1 - target.range.r0 < 10_000) {
      for (let r = target.range.r0; r <= target.range.r1; r++) rows.push(r);
    } else rows.push(row);
    await this.edit((s) => s.setRowHeights(target.sheet, rows, size), { keepClipboard: true });
    this.dispatch({ type: 'view/clearSizes' });
  }

  saveScroll(sheetId: number, scroll: { left: number; top: number }, revealHandled: number): void {
    this.dispatch({ type: 'view/saveScroll', sheetId, scroll, revealHandled });
  }

  setInfoOpen(open: boolean): void {
    this.dispatch({ type: 'info/set', open });
  }

  // Editing

  private loadEditing(): Promise<EditingModule> {
    return (this.editingModule ??= import('../services/editing'));
  }

  /** Loads the editing code in the background so the first edit is instant. */
  preloadEditing(): void {
    void this.loadEditing().catch(() => {
      this.editingModule = null;
    });
  }

  private async getSession(): Promise<EditSession | null> {
    const workbook = this.state.workbook;
    if (!workbook) return null;
    if (this.session) return this.session;
    const { EditSession } = await this.loadEditing();
    if (this.state.workbook !== workbook) return null; // another file was opened meanwhile
    const packaged = workbook.format === 'xlsx' || workbook.format === 'xlsm';
    this.session ??= new EditSession(workbook, packaged ? this.fileBytes : null);
    return this.session;
  }

  /** The active worksheet and the selected range. */
  private selectionTarget(): { sheet: number; model: WorksheetModel; range: CellRange; anchor: CellRef } | null {
    const { workbook, activeSheet } = this.state;
    const view = this.activeView();
    const model = workbook?.sheets[activeSheet];
    if (!model || !view || model.kind !== 'worksheet') return null;
    return { sheet: activeSheet, model, range: normalizeRange(view.selection.anchor, view.selection.focus), anchor: view.selection.anchor };
  }

  /** Starts editing the active cell. `initial` replaces the content (typing); otherwise it's edited in place. */
  startEdit(mode: CellEditState['mode'], initial?: string, source: CellEditState['source'] = 'cell'): void {
    const { workbook, activeSheet, editing } = this.state;
    const view = this.activeView();
    const sheet = workbook?.sheets[activeSheet];
    if (!workbook || !view || sheet?.kind !== 'worksheet') return;
    if (editing) {
      if (source !== editing.source) this.dispatch({ type: 'edit/draft', draft: editing.draft, source });
      return;
    }
    let { row, col } = view.selection.anchor;
    // A merged cell's value lives in its top-left cell.
    const merge = mergeAt(sheet, row, col);
    if (merge) {
      row = merge.r0;
      col = merge.c0;
    }
    const draft = initial ?? editTextAt(workbook, activeSheet, row, col);
    this.dispatch({ type: 'edit/start', edit: { sheet: activeSheet, row, col, draft, mode, source } });
  }

  setDraft(draft: string, source?: CellEditState['source']): void {
    this.dispatch({ type: 'edit/draft', draft, source });
  }

  cancelEdit(): void {
    if (!this.state.editing) return;
    this.dispatch({ type: 'edit/end' });
    this.focusGrid();
  }

  /** Finishes the edit, stores the value and moves the selection like Excel's Enter/Tab. */
  async commitEdit(move: CommitMove = 'down'): Promise<void> {
    const edit = this.state.editing;
    if (!edit || !this.state.workbook) return;
    this.dispatch({ type: 'edit/end' });
    if (move !== 'none') {
      const target = {
        row: Math.max(0, edit.row + (move === 'down' ? 1 : move === 'up' ? -1 : 0)),
        col: Math.max(0, edit.col + (move === 'right' ? 1 : move === 'left' ? -1 : 0)),
      };
      if (this.state.activeSheet === edit.sheet) this.setSelection({ anchor: target, focus: target }, true);
    }
    this.focusGrid();
    if (edit.draft === editTextAt(this.state.workbook, edit.sheet, edit.row, edit.col)) return;
    await this.edit((session) => session.setText([{ sheet: edit.sheet, row: edit.row, col: edit.col, text: edit.draft }], 'Typing'));
  }

  /** ⌘; / ⌘⇧; : starts typing today's date or the current time into the active cell. */
  private insertNow(kind: 'date' | 'time'): void {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    const text =
      kind === 'date' ? `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}` : `${now.getHours()}:${pad(now.getMinutes())}`;
    if (this.state.editing) {
      this.setDraft(this.state.editing.draft + text);
      return;
    }
    this.startEdit('enter', text);
  }

  /** Delete key: clears the selected cells' contents. */
  async clearSelection(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const range = clampToUsedRange(target.range, target.model);
    await this.edit((session) => session.clear(target.sheet, range));
  }

  async undo(): Promise<void> {
    if (!this.state.document.canUndo) return;
    const label = this.session?.undoLabel;
    if ((await this.edit((session) => session.undo())) && label && label !== 'Typing') this.status(`Undid ${label}`);
  }

  async redo(): Promise<void> {
    if (!this.state.document.canRedo) return;
    const label = this.session?.redoLabel;
    if ((await this.edit((session) => session.redo())) && label && label !== 'Typing') this.status(`Redid ${label}`);
  }

  /**
   * Runs an edit and publishes the result. Changes that can't be made are
   * explained in a dialog; unexpected failures leave the workbook as it was.
   */
  private async edit(action: (session: EditSession) => EditOutcome | null, options: { keepClipboard?: boolean } = {}): Promise<boolean> {
    const session = await this.getSession();
    if (!session) return false;
    let outcome: EditOutcome | null;
    try {
      outcome = action(session);
    } catch (error) {
      if (error instanceof Error && error.name === 'EditRefusedError') {
        void this.alert('That can’t be done', error.message);
        return false;
      }
      console.error('Edit failed', error);
      this.status('That change couldn’t be applied.');
      return false;
    }
    if (!outcome) return false;
    this.dispatch({ type: 'workbook/update', workbook: outcome.workbook });
    if (!options.keepClipboard && this.state.clipboard) this.clearClipboardMark();
    this.documentChanged();
    if (outcome.circular) this.status('Circular reference: a formula refers to its own result.');
    if (this.state.search.query) this.runSearch();
    return true;
  }

  // Formatting

  /** The active cell's style, which the toolbar shows and toggles against. */
  currentStyle(): CellStyle {
    return activeCellStyle(this.state);
  }

  async format(patch: StylePatch | ((row: number, col: number) => StylePatch | null), label?: string): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.edit((s) => s.format(target.sheet, target.range, patch, label), { keepClipboard: true });
  }

  /** Bold, italic…: on for the whole selection unless the active cell already has it (Excel). */
  async toggleStyle(key: 'bold' | 'italic' | 'underline' | 'strike'): Promise<void> {
    const on = !this.currentStyle()[key];
    await this.format({ [key]: on ? true : null } as StylePatch, 'Font Style');
  }

  private defaultFont(): { name: string; size: number } {
    return this.state.workbook?.defaultFont ?? DEFAULT_FONT;
  }

  /** Font size of the active cell in points. */
  currentFontSize(): number {
    return Math.round(this.defaultFont().size * (this.currentStyle().fontScale ?? 1) * 2) / 2;
  }

  async setFontSize(points: number): Promise<void> {
    if (!Number.isFinite(points) || points < 1 || points > 409) return;
    const scale = Math.round((points / this.defaultFont().size) * 10_000) / 10_000;
    await this.format({ fontScale: scale === 1 ? null : scale }, 'Font Size');
  }

  async stepFontSize(direction: 1 | -1): Promise<void> {
    const size = this.currentFontSize();
    const next = direction > 0 ? FONT_SIZES.find((s) => s > size + 0.01) : [...FONT_SIZES].reverse().find((s) => s < size - 0.01);
    if (next !== undefined) await this.setFontSize(next);
  }

  async setFontName(name: string): Promise<void> {
    const isDefault = name.toLowerCase() === this.defaultFont().name.toLowerCase();
    await this.format({ fontName: isDefault ? null : name }, 'Font');
  }

  async setTextColor(color: string | null): Promise<void> {
    await this.format({ color }, 'Font Color');
  }

  async setFillColor(color: string | null): Promise<void> {
    await this.format({ fill: color }, 'Fill Color');
  }

  /** Horizontal alignment; choosing the current one again returns to General. */
  async setAlignment(align: HorizontalAlign): Promise<void> {
    const same = this.currentStyle().hAlign === align;
    await this.format({ hAlign: same ? null : align }, 'Alignment');
  }

  async setVerticalAlignment(align: VerticalAlign): Promise<void> {
    await this.format({ vAlign: align === 'bottom' ? null : align }, 'Alignment');
  }

  async toggleWrap(): Promise<void> {
    await this.format({ wrap: this.currentStyle().wrap ? null : true }, 'Wrap Text');
  }

  async stepIndent(direction: 1 | -1): Promise<void> {
    const style = this.currentStyle();
    const indent = Math.max(0, Math.min(15, (style.indent ?? 0) + direction));
    await this.format({ indent: indent || null, ...(indent && !style.hAlign ? { hAlign: 'left' as const } : {}) }, 'Indent');
  }

  async setNumberFormat(format: string | undefined): Promise<void> {
    await this.format({ numberFormat: format ?? null }, 'Number Format');
  }

  async stepDecimals(delta: 1 | -1): Promise<void> {
    const { workbook, activeSheet } = this.state;
    const view = this.activeView();
    const sheet = workbook?.sheets[activeSheet];
    if (!sheet || !view) return;
    const cell = getCell(sheet, view.selection.anchor.row, view.selection.anchor.col);
    const sample = typeof cell?.value === 'number' ? cell.value : null;
    const next = changeDecimals(this.currentStyle().numberFormat, sample, delta);
    if (next) await this.format({ numberFormat: next }, delta > 0 ? 'Increase Decimal' : 'Decrease Decimal');
  }

  async applyBorders(preset: BorderPreset, color?: string): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const range = isWholeRows(target.range) || isWholeColumns(target.range) ? clampToUsedRange(target.range, target.model) : target.range;
    await this.format(borderPatch(range, preset, color), 'Borders');
  }

  async clearFormats(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.edit((s) => s.clearFormats(target.sheet, target.range));
  }

  /** Merges the selection. Excel asks first when values other than the top-left one would be lost. */
  async merge(mode: MergeMode): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const session = await this.getSession();
    if (!session) return;
    if (session.mergeLosesData(target.sheet, target.range, mode)) {
      const ok = await this.confirm({
        title: 'Merge cells?',
        message: 'Merging keeps only the upper-left value and discards the others.',
        confirmLabel: 'Merge',
        cancelLabel: 'Cancel',
      });
      if (!ok) return;
    }
    await this.edit((s) => s.merge(target.sheet, target.range, mode));
  }

  /** The Merge & Center button: merges, or unmerges when the selection is merged. */
  async toggleMergeCenter(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const { range } = target;
    const merged = target.model.merges.some((m) => m.r0 <= range.r1 && m.r1 >= range.r0 && m.c0 <= range.c1 && m.c1 >= range.c0);
    if (merged) await this.unmerge();
    else await this.merge('center');
  }

  async unmerge(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.edit((s) => s.unmerge(target.sheet, target.range));
  }

  // Rows and columns

  async insertRows(where: 'above' | 'below'): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const { r0, r1 } = isWholeColumns(target.range) ? { r0: target.anchor.row, r1: target.anchor.row } : target.range;
    const count = r1 - r0 + 1;
    const at = where === 'above' ? r0 : r1 + 1;
    if ((await this.edit((s) => s.insertRows(target.sheet, at, count))) && where === 'below') {
      this.setSelection({ anchor: { row: at, col: 0 }, focus: { row: at + count - 1, col: EXCEL_MAX_COLS - 1 } }, true);
    }
  }

  async insertColumns(where: 'left' | 'right'): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const { c0, c1 } = isWholeRows(target.range) ? { c0: target.anchor.col, c1: target.anchor.col } : target.range;
    const count = c1 - c0 + 1;
    const at = where === 'left' ? c0 : c1 + 1;
    if ((await this.edit((s) => s.insertColumns(target.sheet, at, count))) && where === 'right') {
      this.setSelection({ anchor: { row: 0, col: at }, focus: { row: EXCEL_MAX_ROWS - 1, col: at + count - 1 } }, true);
    }
  }

  async deleteRows(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const range = isWholeColumns(target.range) ? clampToUsedRange(target.range, target.model) : target.range;
    await this.edit((s) => s.deleteRows(target.sheet, range.r0, range.r1 - range.r0 + 1));
  }

  async deleteColumns(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const range = isWholeRows(target.range) ? clampToUsedRange(target.range, target.model) : target.range;
    await this.edit((s) => s.deleteColumns(target.sheet, range.c0, range.c1 - range.c0 + 1));
  }

  /** ⌘⌥= : whole columns selected → insert columns, otherwise insert rows. */
  async insertCells(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    if (isWholeColumns(target.range) && !isWholeRows(target.range)) await this.insertColumns('left');
    else await this.insertRows('above');
  }

  async deleteCells(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    if (isWholeColumns(target.range) && !isWholeRows(target.range)) await this.deleteColumns();
    else await this.deleteRows();
  }

  async setRowsHidden(hidden: boolean): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const range = isWholeColumns(target.range) ? clampToUsedRange(target.range, target.model) : target.range;
    if (!hidden) {
      // Unhide hidden rows in the selection (or right next to it).
      const heights = target.model.rowHeights;
      const rows: number[] = [];
      heights.index.forEach((r, k) => {
        if (heights.size[k] === 0 && r >= range.r0 - 1 && r <= range.r1 + 1) rows.push(r);
      });
      if (rows.length > 0) await this.edit((s) => s.setRowHeights(target.sheet, rows, null, 'Unhide Rows'));
      return;
    }
    const rows: number[] = [];
    for (let r = range.r0; r <= range.r1 && rows.length < 100_000; r++) rows.push(r);
    await this.edit((s) => s.setRowHeights(target.sheet, rows, 0, 'Hide Rows'));
  }

  async setColumnsHidden(hidden: boolean): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const range = isWholeRows(target.range) ? clampToUsedRange(target.range, target.model) : target.range;
    if (!hidden) {
      const widths = target.model.columnWidths;
      const cols: number[] = [];
      widths.index.forEach((c, k) => {
        if (widths.size[k] === 0 && c >= range.c0 - 1 && c <= range.c1 + 1) cols.push(c);
      });
      if (cols.length > 0) await this.edit((s) => s.setColumnWidths(target.sheet, cols, null, 'Unhide Columns'));
      return;
    }
    const cols: number[] = [];
    for (let c = range.c0; c <= range.c1; c++) cols.push(c);
    await this.edit((s) => s.setColumnWidths(target.sheet, cols, 0, 'Hide Columns'));
  }

  // Fill and sort

  /** Fill Down (⌘D) / Fill Right (⌘R): copies the first row/column of the selection (or the cell before it). */
  async fill(direction: 'down' | 'right'): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const range = clampToUsedRange(target.range, target.model);
    let source: CellRange;
    let dest: CellRange;
    if (direction === 'down') {
      if (range.r0 === range.r1) {
        if (range.r0 === 0) return;
        source = { ...range, r0: range.r0 - 1, r1: range.r0 - 1 };
        dest = range;
      } else {
        source = { ...range, r1: range.r0 };
        dest = { ...range, r0: range.r0 + 1 };
      }
    } else if (range.c0 === range.c1) {
      if (range.c0 === 0) return;
      source = { ...range, c0: range.c0 - 1, c1: range.c0 - 1 };
      dest = range;
    } else {
      source = { ...range, c1: range.c0 };
      dest = { ...range, c0: range.c0 + 1 };
    }
    await this.edit((s) => s.fill(target.sheet, source, dest, false));
  }

  /** The fill handle was dragged over `dest`, next to the selection. */
  async fillTo(dest: CellRange): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    const source = target.range;
    if (await this.edit((s) => s.fill(target.sheet, source, dest, true))) {
      const all = { r0: Math.min(source.r0, dest.r0), c0: Math.min(source.c0, dest.c0), r1: Math.max(source.r1, dest.r1), c1: Math.max(source.c1, dest.c1) };
      this.setSelection({ anchor: { row: all.r0, col: all.c0 }, focus: { row: all.r1, col: all.c1 } });
    }
  }

  /** Double-clicking the fill handle fills down as far as the data next to it goes. */
  async autoFill(): Promise<void> {
    const target = this.selectionTarget();
    if (!target) return;
    const { model, range } = target;
    const beside = [range.c0 - 1, range.c1 + 1].find((c) => c >= 0 && hasValueAt(model, range.r1 + 1, c));
    if (beside === undefined) return;
    let last = range.r1 + 1;
    while (last + 1 < model.rowCount && hasValueAt(model, last + 1, beside)) last++;
    await this.fillTo({ r0: range.r1 + 1, c0: range.c0, r1: last, c1: range.c1 });
  }

  /**
   * Sorts by the active cell's column. With a single cell selected, the block of
   * data around it is sorted, keeping titles, the heading row and a totals row in place.
   */
  async sort(ascending: boolean): Promise<void> {
    const target = this.selectionTarget();
    const workbook = this.state.workbook;
    if (!target || !workbook) return;
    await this.commitEdit('none');
    let range: CellRange;
    const single = target.range.r0 === target.range.r1 && target.range.c0 === target.range.c1;
    if (single) {
      range = sortRange(target.model, target.anchor, workbook.styles);
    } else {
      range = clampToUsedRange(target.range, target.model);
    }
    if (range.r1 <= range.r0) {
      this.status('Select at least two rows to sort.');
      return;
    }
    const key = Math.min(Math.max(target.anchor.col, range.c0), range.c1);
    await this.edit((s) => s.sort(target.sheet, range, key, ascending));
  }

  // Find and replace

  /** Replaces every match of the search in its scope. Formulas are searched as written. */
  async replaceAll(replacement: string): Promise<void> {
    const { workbook, search, activeSheet } = this.state;
    if (!workbook || !search.query) return;
    await this.commitEdit('none');
    const inputs = this.replacements(replacement, search.scope === 'sheet' ? [activeSheet] : workbook.sheets.map((_, i) => i));
    if (inputs.length === 0) {
      this.status('Nothing to replace.');
      return;
    }
    if (await this.edit((s) => s.setText(inputs, 'Replace'))) {
      this.status(inputs.length === 1 ? 'Replaced 1 cell' : `Replaced ${formatCount(inputs.length)} cells`);
    }
  }

  /** Replaces the current match, then moves to the next one. */
  async replaceCurrent(replacement: string): Promise<void> {
    const { search } = this.state;
    const match = search.current >= 0 ? search.results?.matches[search.current] : undefined;
    if (!match) {
      this.findNext(1);
      return;
    }
    const inputs = this.replacements(replacement, [match.sheet], match);
    if (inputs.length > 0) await this.edit((s) => s.setText(inputs, 'Replace'));
    this.findNext(1);
  }

  private replacements(replacement: string, sheets: number[], only?: { sheet: number; row: number; col: number }) {
    const { workbook, search } = this.state;
    if (!workbook) return [];
    const escaped = search.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(search.wholeCell ? `^${escaped}$` : escaped, search.matchCase ? 'g' : 'gi');
    const inputs: { sheet: number; row: number; col: number; text: string }[] = [];
    for (const index of sheets) {
      const sheet = workbook.sheets[index];
      if (!sheet || sheet.kind !== 'worksheet') continue;
      const { rowStart, cols } = sheet.cells;
      const rows = only ? [only.row] : Array.from({ length: sheet.rowCount }, (_, r) => r);
      for (const r of rows) {
        if (r >= sheet.rowCount) continue;
        for (let i = rowStart[r]; i < rowStart[r + 1]; i++) {
          if (only && cols[i] !== only.col) continue;
          const text = editTextAt(workbook, index, r, cols[i]);
          if (!text) continue;
          pattern.lastIndex = 0;
          const next = text.replace(pattern, () => replacement);
          if (next !== text) inputs.push({ sheet: index, row: r, col: cols[i], text: next });
        }
      }
    }
    return inputs;
  }

  /** Syncs document flags, the window title and the host's "unsaved changes" state. */
  private documentChanged(): void {
    const session = this.session;
    const dirty = session?.dirty ?? false;
    this.dispatch({
      type: 'document/set',
      document: {
        dirty,
        canUndo: session?.canUndo ?? false,
        canRedo: session?.canRedo ?? false,
        staleFormulas: session?.staleFormulas ?? 0,
      },
    });
    const name = this.state.workbook?.fileName;
    this.platform.setWindowTitle(name ? `${name}${dirty ? ' — Edited' : ''} — ${APP_NAME}` : APP_NAME);
    this.platform.setDocumentEdited(dirty);
  }

  // Saving

  /** Formats a workbook can be saved as, the current one first. */
  private saveTypes(workbook: WorkbookModel): SaveFileType[] {
    switch (workbook.format) {
      case 'xlsm':
        return ['xlsm', 'csv'];
      case 'csv':
        return ['csv', 'xlsx', 'tsv'];
      case 'tsv':
        return ['tsv', 'xlsx', 'csv'];
      default:
        return ['xlsx', 'csv'];
    }
  }

  /** Save: writes to the open file, or asks where when there isn't one (new file, .xls). */
  async save(): Promise<boolean> {
    const { workbook, filePath } = this.state;
    if (!workbook) return false;
    if (this.state.editing) await this.commitEdit('none');
    const format = workbook.format === 'xls' ? null : (workbook.format as SaveFormat);
    if (!filePath || !format || this.platform.kind === 'browser') return this.saveAs();
    return this.writeTo(filePath, format);
  }

  async saveAs(): Promise<boolean> {
    const workbook = this.state.workbook;
    if (!workbook) return false;
    if (this.state.editing) await this.commitEdit('none');
    const types = this.saveTypes(workbook);
    const stem = workbook.fileName.replace(/\.[^.]+$/, '') || 'Untitled';
    const path = await this.platform.saveDialog({ defaultName: `${stem}.${types[0]}`, types });
    if (!path) return false;
    const ext = getExtension(path);
    const format = (types as string[]).includes(ext) ? (ext as SaveFormat) : types[0];
    return this.writeTo((types as string[]).includes(ext) ? path : `${path}.${format}`, format);
  }

  /** CSV keeps only the active sheet's values; say so once before losing formatting or other sheets. */
  private async confirmTextFormat(workbook: WorkbookModel, format: SaveFormat): Promise<boolean> {
    if (this.textSaveConfirmed || (format !== 'csv' && format !== 'tsv')) return true;
    const sheet = workbook.sheets[this.state.activeSheet];
    let styled = false;
    if (sheet) for (let i = 0; i < sheet.cellCount && !styled; i++) styled = sheet.cells.styles[i] !== 0;
    const loses = workbook.sheets.length > 1 || styled || (sheet?.formulas.cells.length ?? 0) > 0 || (sheet?.merges.length ?? 0) > 0;
    if (!loses) return true;
    const kind = format.toUpperCase();
    const ok = await this.confirm({
      title: `Save as ${kind}?`,
      message: `${kind} files keep only the values of the current sheet. Formatting, formulas, merged cells and other sheets won’t be saved. To keep them, choose Save As and pick Excel Workbook.`,
      confirmLabel: `Save as ${kind}`,
      cancelLabel: 'Cancel',
    });
    if (ok) this.textSaveConfirmed = true;
    return ok;
  }

  private async writeTo(path: string, format: SaveFormat): Promise<boolean> {
    const session = await this.getSession();
    const workbook = this.state.workbook;
    if (!session || !workbook) return false;
    if (!(await this.confirmTextFormat(workbook, format))) return false;
    const name = baseName(path);
    this.dispatch({ type: 'document/set', document: { saving: true } });
    try {
      const bytes = session.buildFile(format, this.state.activeSheet);
      await this.platform.writeFile(path, bytes);
      const text = format === 'csv' || format === 'tsv';
      const saved = session.markSaved({
        fileName: name,
        fileSize: bytes.length,
        format,
        csv: text ? { delimiter: format === 'tsv' ? '\t' : (workbook.csv?.delimiter ?? ','), bom: workbook.csv?.bom ?? false } : undefined,
      });
      this.fileBytes = text ? null : bytes;
      const path_ = this.platform.kind === 'desktop' ? path : name;
      this.dispatch({ type: 'file/saved', workbook: saved, path: path_ });
      this.documentChanged();
      if (this.platform.kind === 'desktop') this.setRecentFiles(addRecentFile(this.state.recentFiles, path));
      this.status(text && workbook.sheets.length > 1 ? `Saved “${name}” (only the active sheet is kept in ${format.toUpperCase()})` : `Saved “${name}”`);
      return true;
    } catch (error) {
      console.warn(`Could not save ${name}:`, error);
      this.dispatch({ type: 'error/show', error: toSaveError(error, name) });
      return false;
    } finally {
      this.dispatch({ type: 'document/set', document: { saving: false } });
    }
  }

  /** Asks to save unsaved changes. Resolves false when the user cancels. */
  async confirmDiscard(): Promise<boolean> {
    if (this.state.editing) await this.commitEdit('none');
    if (!this.state.document.dirty) return true;
    const choice = await this.platform.confirmUnsavedChanges(this.state.workbook?.fileName ?? 'Untitled');
    if (choice === 'cancel') return false;
    if (choice === 'save') return this.save();
    return true;
  }

  // Clipboard

  /** TSV for the current selection, or null when there's nothing (or too much) to copy. */
  selectionText(): { text: string; cells: number } | null {
    const { workbook, activeSheet } = this.state;
    const view = this.activeView();
    const sheet = workbook?.sheets[activeSheet];
    if (!sheet || !view) return null;
    const payload = buildCopyPayload(sheet, normalizeRange(view.selection.anchor, view.selection.focus));
    return payload ? { text: payload.text, cells: payload.rows * payload.columns } : null;
  }

  /**
   * Remembers the selection for pasting inside the app (with formulas and
   * formats) and returns its text for the system clipboard.
   */
  prepareCopy(cut: boolean): { text: string; cells: number } | null {
    const target = this.selectionTarget();
    const payload = this.selectionText();
    if (!target || !payload) return null;
    const range = clampToUsedRange(target.range, target.model);
    const cells = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    const block = cells <= MAX_BLOCK_CELLS ? readBlock(target.model, range) : null;
    this.copied = { text: payload.text, block, sheetId: target.model.id, range, cut: cut && block !== null };
    this.dispatch({ type: 'clipboard/set', clipboard: { sheetId: target.model.id, range, cut: this.copied.cut } });
    return payload;
  }

  async copySelection(cut = false): Promise<void> {
    if (!this.state.workbook) return;
    const payload = this.prepareCopy(cut);
    if (!payload) {
      this.status('The selection is too large to copy.');
      return;
    }
    try {
      await this.platform.writeClipboardText(payload.text);
      if (cut) this.status('Select where to move the cells, then paste.');
      else this.announceCopy(payload.cells);
    } catch {
      this.status('Couldn’t copy to the clipboard.');
    }
  }

  async cutSelection(): Promise<void> {
    await this.copySelection(true);
  }

  announceCopy(cells: number): void {
    this.status(cells === 1 ? 'Copied 1 cell' : `Copied ${formatCount(cells)} cells`);
  }

  clearClipboardMark(): void {
    this.dispatch({ type: 'clipboard/set', clipboard: null });
    if (this.copied?.cut) this.copied = null;
  }

  async pasteFromClipboard(mode: PasteMode = 'all'): Promise<void> {
    let text: string;
    try {
      text = await this.platform.readClipboardText();
    } catch {
      this.status('Couldn’t read the clipboard.');
      return;
    }
    await this.pasteData(text, mode);
  }

  /** Pastes clipboard text: cells copied in the app keep their formulas and formats. */
  async pasteData(text: string, mode: PasteMode = 'all'): Promise<void> {
    const copied = this.copied;
    const target = this.selectionTarget();
    if (!target) return;
    await this.commitEdit('none');
    if (copied?.block && text === copied.text) {
      const sourceSheet = this.state.workbook?.sheets.findIndex((s) => s.id === copied.sheetId) ?? -1;
      const rows = copied.range.r1 - copied.range.r0 + 1;
      const cols = copied.range.c1 - copied.range.c0 + 1;
      const { r0, c0 } = target.range;
      if (copied.cut && mode === 'all' && sourceSheet >= 0) {
        if (await this.edit((s) => s.move(sourceSheet, copied.range, target.sheet, r0, c0))) {
          this.copied = null;
          this.dispatch({ type: 'clipboard/set', clipboard: null });
          this.setSelection({ anchor: { row: r0, col: c0 }, focus: { row: r0 + rows - 1, col: c0 + cols - 1 } });
        }
        return;
      }
      const block = copied.block;
      const height = target.range.r1 - target.range.r0 + 1;
      const width = target.range.c1 - target.range.c0 + 1;
      const tiles = height % rows === 0 && width % cols === 0 && height * width <= 1_000_000;
      if ((await this.edit((s) => s.paste(target.sheet, target.range, block, mode), { keepClipboard: true })) && !tiles) {
        this.setSelection({ anchor: { row: r0, col: c0 }, focus: { row: r0 + rows - 1, col: c0 + cols - 1 } });
      }
      return;
    }
    if (mode !== 'formats') await this.pasteText(text);
  }

  /** Pastes tab-separated text at the active cell (or fills the selection with a single value). */
  async pasteText(text: string): Promise<void> {
    const view = this.activeView();
    if (!view || !this.state.workbook || text === '') return;
    const rows = parseClipboardTable(text);
    if (rows.length === 0) return;
    const sheet = this.state.activeSheet;
    const { anchor } = view.selection;
    const range = normalizeRange(view.selection.anchor, view.selection.focus);
    const inputs: { sheet: number; row: number; col: number; text: string }[] = [];
    const single = rows.length === 1 && rows[0].length === 1;
    const selectionCells = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    if (single && selectionCells > 1 && selectionCells <= 100_000) {
      for (let r = range.r0; r <= range.r1; r++) {
        for (let c = range.c0; c <= range.c1; c++) inputs.push({ sheet, row: r, col: c, text: rows[0][0] });
      }
    } else {
      rows.forEach((cells, r) => cells.forEach((value, c) => inputs.push({ sheet, row: anchor.row + r, col: anchor.col + c, text: value })));
      const width = Math.max(...rows.map((cells) => cells.length));
      this.setSelection({ anchor, focus: { row: anchor.row + rows.length - 1, col: anchor.col + width - 1 } });
    }
    await this.edit((session) => session.setText(inputs, 'Paste'));
  }

  // Search

  setSearchQuery(query: string): void {
    this.dispatch({ type: 'search/query', query });
    clearTimeout(this.searchTimer);
    if (query) this.searchTimer = setTimeout(() => this.runSearch(), 150);
  }

  setSearchOptions(options: Partial<Pick<SearchState, 'scope' | 'matchCase' | 'wholeCell'>>): void {
    this.dispatch({ type: 'search/options', options });
    this.runSearch();
  }

  setReplaceOpen(open: boolean): void {
    this.dispatch({ type: 'replace/open', open });
  }

  clearSearch(): void {
    clearTimeout(this.searchTimer);
    this.dispatch({ type: 'search/query', query: '' });
    this.dispatch({ type: 'replace/open', open: false });
  }

  /** Esc outside a dialog: drops the copy outline, then closes an active search. Returns whether it did something. */
  handleEscape(): boolean {
    if (this.state.clipboard) {
      this.clearClipboardMark();
      return true;
    }
    if (!this.state.search.query && !this.state.replaceOpen) return false;
    this.clearSearch();
    this.focusGrid();
    return true;
  }

  runSearch(): void {
    clearTimeout(this.searchTimer);
    const { workbook, search, activeSheet } = this.state;
    if (!workbook || !search.query) {
      if (search.results) this.dispatch({ type: 'search/results', results: null });
      return;
    }
    const results = searchWorkbook(workbook, search.query, {
      scope: search.scope,
      sheetIndex: activeSheet,
      matchCase: search.matchCase,
      wholeCell: search.wholeCell,
    });
    this.dispatch({ type: 'search/results', results });
  }

  findNext(direction: 1 | -1): void {
    if (!this.state.workbook) return;
    if (!this.state.search.query) {
      this.dispatch({ type: 'focus/request', target: 'search' });
      return;
    }
    if (!this.state.search.results) this.runSearch();
    const { results, current } = this.state.search;
    const count = results?.matches.length ?? 0;
    if (!results || count === 0) return;

    let index: number;
    if (current < 0) {
      // Start from the active cell, like Excel's Find Next.
      const anchor = this.activeView()?.selection.anchor ?? { row: 0, col: 0 };
      const from = firstMatchFrom(results.matches, this.state.activeSheet, anchor.row, anchor.col);
      index = direction === 1 ? from : (from - 1 + count) % count;
    } else {
      index = stepMatch(current, count, direction);
    }
    const match = results.matches[index];
    if (match.sheet !== this.state.activeSheet) this.dispatch({ type: 'sheet/activate', index: match.sheet });
    this.dispatch({ type: 'search/current', index });
    const cell = { row: match.row, col: match.col };
    this.setSelection({ anchor: cell, focus: cell }, true);
  }

  // Desktop integration

  /** Subscribes to drag and drop, OS file-open requests, menu commands, close requests and native clipboard events. */
  attach(): () => void {
    const unsubscribe = this.platform.subscribe({
      onDragStateChange: (active) => this.dispatch({ type: 'drag/set', active }),
      onFilesDropped: (files) => {
        if (files.length > 1) this.status('Only the first file was opened.');
        if (files[0]) void this.open(files[0]);
      },
      onOpenPaths: (paths) => {
        if (paths[0]) void this.openPath(paths[0]);
      },
      onMenuCommand: (command) => {
        if (isCommandId(command)) this.runCommand(command, 'menu');
      },
      onCloseRequested: () => void this.quit(),
    });

    // Native Edit ▸ Copy / Cut / Paste while the grid has focus (keyboard shortcuts are handled by the grid).
    const gridFocused = () => document.activeElement instanceof HTMLElement && document.activeElement.closest('[data-grid-root]') !== null;
    const enable = (event: Event) => {
      if (gridFocused()) event.preventDefault(); // enables the command in WebKit without a text selection
    };
    const onCopy = (event: ClipboardEvent) => {
      if (!gridFocused() || !event.clipboardData) return;
      const payload = this.prepareCopy(event.type === 'cut');
      if (!payload) return;
      event.clipboardData.setData('text/plain', payload.text);
      event.preventDefault();
      if (event.type !== 'cut') this.announceCopy(payload.cells);
    };
    const onPaste = (event: ClipboardEvent) => {
      if (!gridFocused() || !event.clipboardData) return;
      event.preventDefault();
      void this.pasteData(event.clipboardData.getData('text/plain'));
    };
    const listeners: [string, EventListener][] = [
      ['beforecopy', enable],
      ['beforecut', enable],
      ['beforepaste', enable],
      ['copy', onCopy as EventListener],
      ['cut', onCopy as EventListener],
      ['paste', onPaste as EventListener],
    ];
    for (const [type, listener] of listeners) document.addEventListener(type, listener);

    this.platform.setRecentFiles(this.state.recentFiles);
    void this.platform
      .takeStartupFiles()
      .then((paths) => {
        if (paths[0]) void this.openPath(paths[0]);
      })
      .catch(() => undefined);
    const preload = setTimeout(() => this.preloadEditing(), 1500);

    return () => {
      unsubscribe();
      for (const [type, listener] of listeners) document.removeEventListener(type, listener);
      clearTimeout(this.searchTimer);
      clearTimeout(preload);
    };
  }
}
