import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, type CellRange, type CellRef, type WorkbookModel, type WorksheetModel } from '../models/workbook';
import type { SearchResults, SearchScope } from '../services/search/searchService';
import type { UserFacingError } from '../services/workbook/errors';
import type { LoadStage } from '../services/workbook/workbookService';

export const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2] as const;
export const MIN_ZOOM = ZOOM_LEVELS[0];
export const MAX_ZOOM = ZOOM_LEVELS[ZOOM_LEVELS.length - 1];

/** The active cell is the anchor; the range spans anchor → focus. */
export interface Selection {
  anchor: CellRef;
  focus: CellRef;
}

export interface SheetViewState {
  /** The sheet this view belongs to (views follow sheets that are added, moved or deleted). */
  sheetId: number;
  selection: Selection;
  /** Sizes being dragged (px at 100%), shown over the sheet's sizes until the drag ends. */
  columnSizes: ReadonlyMap<number, number>;
  rowSizes: ReadonlyMap<number, number>;
  /** Scroll position saved when leaving the sheet. */
  scroll: { left: number; top: number };
  /** The grid scrolls `cell` into view whenever `token` changes. */
  reveal: { cell: CellRef; token: number } | null;
  /** Last reveal token the grid acted on (kept across sheet switches). */
  revealHandled: number;
  /** Minimum rows/columns shown; grows when navigating past the data. */
  extent: { rows: number; cols: number };
}

export interface SearchState {
  query: string;
  scope: SearchScope;
  matchCase: boolean;
  wholeCell: boolean;
  /** Results for `query`; null while typing or when empty. */
  results: SearchResults | null;
  /** Index into results.matches of the match the user navigated to, or -1. */
  current: number;
}

export interface DocumentState {
  /** Unsaved changes. */
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  saving: boolean;
  /** Formulas that couldn't be recalculated after an edit (they keep their saved value). */
  staleFormulas: number;
}

/** An in-progress cell edit. The in-cell editor and the formula bar share the draft. */
export interface CellEditState {
  sheet: number;
  row: number;
  col: number;
  draft: string;
  /** 'enter': typing replaced the content (arrow keys commit); 'edit': caret editing (F2, double-click). */
  mode: 'enter' | 'edit';
  source: 'cell' | 'formulaBar';
}

/** A question or notice shown in a modal dialog. */
export interface DialogState {
  id: number;
  kind: 'confirm' | 'alert' | 'prompt';
  title: string;
  message: string;
  /** Initial text of a prompt's field. */
  value?: string;
  confirmLabel: string;
  cancelLabel?: string;
  /** The confirm button does something destructive (shown in red). */
  destructive?: boolean;
}

/** Cells copied or cut in the app, outlined until pasted (cut) or dismissed. */
export interface ClipboardMark {
  sheetId: number;
  range: CellRange;
  cut: boolean;
}

export const cleanDocument: DocumentState = { dirty: false, canUndo: false, canRedo: false, saving: false, staleFormulas: 0 };

export interface ViewerState {
  workbook: WorkbookModel | null;
  document: DocumentState;
  editing: CellEditState | null;
  filePath: string | null;
  activeSheet: number;
  sheetViews: SheetViewState[];
  zoom: number;
  loading: { fileName: string; stage: LoadStage | 'starting'; size?: number } | null;
  error: UserFacingError | null;
  search: SearchState;
  infoOpen: boolean;
  dragActive: boolean;
  recentFiles: string[];
  statusMessage: { text: string; token: number } | null;
  focusRequest: { target: 'grid' | 'nameBox' | 'search' | 'replace'; token: number } | null;
  dialog: DialogState | null;
  clipboard: ClipboardMark | null;
  /** Find & Replace shows the replace field. */
  replaceOpen: boolean;
}

export type ViewerAction =
  | { type: 'load/start'; fileName: string; size?: number }
  | { type: 'load/stage'; stage: LoadStage }
  | { type: 'load/success'; workbook: WorkbookModel; path: string | null }
  | { type: 'load/failure'; error: UserFacingError }
  | { type: 'load/cancel' }
  | { type: 'workbook/close' }
  | { type: 'workbook/update'; workbook: WorkbookModel; activeSheetId?: number }
  | { type: 'file/saved'; workbook: WorkbookModel; path: string | null }
  | { type: 'document/set'; document: Partial<DocumentState> }
  | { type: 'edit/start'; edit: CellEditState }
  | { type: 'edit/draft'; draft: string; source?: CellEditState['source'] }
  | { type: 'edit/end' }
  | { type: 'error/show'; error: UserFacingError }
  | { type: 'error/dismiss' }
  | { type: 'sheet/activate'; index: number }
  | { type: 'selection/set'; selection: Selection; reveal?: boolean | 'focus' }
  | { type: 'view/saveScroll'; sheetId: number; scroll: { left: number; top: number }; revealHandled: number }
  | { type: 'view/columnSize'; col: number; size: number | null }
  | { type: 'view/rowSize'; row: number; size: number | null }
  | { type: 'view/clearSizes' }
  | { type: 'dialog/open'; dialog: DialogState }
  | { type: 'dialog/close'; id: number }
  | { type: 'clipboard/set'; clipboard: ClipboardMark | null }
  | { type: 'replace/open'; open: boolean }
  | { type: 'zoom/set'; zoom: number }
  | { type: 'info/set'; open: boolean }
  | { type: 'drag/set'; active: boolean }
  | { type: 'recent/set'; paths: string[] }
  | { type: 'status/show'; text: string }
  | { type: 'focus/request'; target: 'grid' | 'nameBox' | 'search' | 'replace' }
  | { type: 'search/query'; query: string }
  | { type: 'search/options'; options: Partial<Pick<SearchState, 'scope' | 'matchCase' | 'wholeCell'>> }
  | { type: 'search/results'; results: SearchResults | null }
  | { type: 'search/current'; index: number };

export const initialSearchState: SearchState = {
  query: '',
  scope: 'sheet',
  matchCase: false,
  wholeCell: false,
  results: null,
  current: -1,
};

export function createInitialState(recentFiles: string[] = []): ViewerState {
  return {
    workbook: null,
    document: cleanDocument,
    editing: null,
    filePath: null,
    activeSheet: 0,
    sheetViews: [],
    zoom: 1,
    loading: null,
    error: null,
    search: initialSearchState,
    infoOpen: false,
    dragActive: false,
    recentFiles,
    statusMessage: null,
    focusRequest: null,
    dialog: null,
    clipboard: null,
    replaceOpen: false,
  };
}

export function initialSheetView(sheet: WorksheetModel): SheetViewState {
  const origin = { row: 0, col: 0 };
  return {
    sheetId: sheet.id,
    selection: { anchor: origin, focus: origin },
    columnSizes: new Map(),
    rowSizes: new Map(),
    scroll: { left: 0, top: 0 },
    reveal: null,
    revealHandled: 0,
    extent: { rows: 0, cols: 0 },
  };
}

let token = 0;
const nextToken = () => ++token;

function updateActiveView(state: ViewerState, update: (view: SheetViewState) => SheetViewState): ViewerState {
  const view = state.sheetViews[state.activeSheet];
  if (!view) return state;
  const next = update(view);
  if (next === view) return state;
  const sheetViews = state.sheetViews.slice();
  sheetViews[state.activeSheet] = next;
  return { ...state, sheetViews };
}

/** Views and the active sheet after the workbook's sheets changed (added, deleted, moved). */
function syncViews(state: ViewerState, workbook: WorkbookModel, activeSheetId?: number): Pick<ViewerState, 'sheetViews' | 'activeSheet'> {
  const views = state.sheetViews;
  const sheets = workbook.sheets;
  const same = views.length === sheets.length && views.every((view, i) => view.sheetId === sheets[i].id);
  const currentId = activeSheetId ?? views[state.activeSheet]?.sheetId;
  const sheetViews = same ? views : sheets.map((sheet) => views.find((view) => view.sheetId === sheet.id) ?? initialSheetView(sheet));
  let activeSheet = sheets.findIndex((sheet) => sheet.id === currentId);
  if (activeSheet < 0) {
    // The active sheet was deleted: show the one that took its place, like Excel.
    const start = Math.min(state.activeSheet, sheets.length - 1);
    activeSheet = sheets.findIndex((sheet, i) => i >= start && sheet.visibility === 'visible');
    if (activeSheet < 0) activeSheet = Math.max(0, sheets.findIndex((sheet) => sheet.visibility === 'visible'));
  }
  return { sheetViews, activeSheet };
}

function withSize(sizes: ReadonlyMap<number, number>, index: number, size: number | null): ReadonlyMap<number, number> {
  const next = new Map(sizes);
  if (size === null) next.delete(index);
  else next.set(index, size);
  return next;
}

/** Rows/columns needed so a selected cell exists in the grid (whole rows/columns excluded). */
function growExtent(extent: SheetViewState['extent'], selection: Selection): SheetViewState['extent'] {
  const cells = [selection.anchor, selection.focus];
  let rows = extent.rows;
  let cols = extent.cols;
  for (const cell of cells) {
    if (cell.row < EXCEL_MAX_ROWS - 1) rows = Math.max(rows, cell.row + 1);
    if (cell.col < EXCEL_MAX_COLS - 1) cols = Math.max(cols, cell.col + 1);
  }
  return rows === extent.rows && cols === extent.cols ? extent : { rows, cols };
}

export function viewerReducer(state: ViewerState, action: ViewerAction): ViewerState {
  switch (action.type) {
    case 'load/start':
      return { ...state, loading: { fileName: action.fileName, stage: 'starting', size: action.size }, error: null };
    case 'load/stage':
      return state.loading ? { ...state, loading: { ...state.loading, stage: action.stage } } : state;
    case 'load/success': {
      const { workbook } = action;
      return {
        ...state,
        workbook,
        filePath: action.path,
        activeSheet: workbook.activeSheetIndex,
        sheetViews: workbook.sheets.map(initialSheetView),
        document: cleanDocument,
        editing: null,
        loading: null,
        error: null,
        clipboard: null,
        search: { ...state.search, results: null, current: -1 },
        focusRequest: { target: 'grid', token: nextToken() },
      };
    }
    case 'load/failure':
      return { ...state, loading: null, error: action.error };
    case 'load/cancel':
      return { ...state, loading: null };
    case 'workbook/close':
      return {
        ...state,
        workbook: null,
        document: cleanDocument,
        editing: null,
        filePath: null,
        activeSheet: 0,
        sheetViews: [],
        clipboard: null,
        search: { ...state.search, results: null, current: -1 },
      };
    case 'workbook/update': {
      if (state.workbook === action.workbook && action.activeSheetId === undefined) return state;
      // Views follow their sheets; selections and scroll positions stay.
      const views = syncViews(state, action.workbook, action.activeSheetId);
      const sheetChanged = views.activeSheet !== state.activeSheet || views.sheetViews !== state.sheetViews;
      return {
        ...state,
        ...views,
        workbook: action.workbook,
        editing: sheetChanged ? null : state.editing,
        search: { ...state.search, results: null, current: -1 },
      };
    }
    case 'file/saved':
      return { ...state, workbook: action.workbook, filePath: action.path };
    case 'document/set': {
      const document = { ...state.document, ...action.document };
      const same = (Object.keys(document) as (keyof DocumentState)[]).every((k) => document[k] === state.document[k]);
      return same ? state : { ...state, document };
    }
    case 'edit/start':
      return { ...state, editing: action.edit };
    case 'edit/draft':
      return state.editing
        ? { ...state, editing: { ...state.editing, draft: action.draft, source: action.source ?? state.editing.source } }
        : state;
    case 'edit/end':
      return state.editing ? { ...state, editing: null } : state;
    case 'error/show':
      return { ...state, error: action.error };
    case 'error/dismiss':
      return state.error ? { ...state, error: null } : state;
    case 'sheet/activate': {
      if (!state.workbook || action.index === state.activeSheet) return state;
      if (action.index < 0 || action.index >= state.workbook.sheets.length) return state;
      return { ...state, activeSheet: action.index, editing: null };
    }
    case 'selection/set':
      return updateActiveView(state, (view) => ({
        ...view,
        selection: action.selection,
        extent: growExtent(view.extent, action.selection),
        reveal: action.reveal
          ? { cell: action.reveal === 'focus' ? action.selection.focus : action.selection.anchor, token: nextToken() }
          : view.reveal,
      }));
    case 'view/saveScroll': {
      const index = state.sheetViews.findIndex((view) => view.sheetId === action.sheetId);
      const view = state.sheetViews[index];
      if (!view) return state;
      const sheetViews = state.sheetViews.slice();
      sheetViews[index] = { ...view, scroll: action.scroll, revealHandled: action.revealHandled };
      return { ...state, sheetViews };
    }
    case 'view/columnSize':
      return updateActiveView(state, (view) => ({ ...view, columnSizes: withSize(view.columnSizes, action.col, action.size) }));
    case 'view/rowSize':
      return updateActiveView(state, (view) => ({ ...view, rowSizes: withSize(view.rowSizes, action.row, action.size) }));
    case 'view/clearSizes':
      return updateActiveView(state, (view) =>
        view.columnSizes.size === 0 && view.rowSizes.size === 0 ? view : { ...view, columnSizes: new Map(), rowSizes: new Map() },
      );
    case 'dialog/open':
      return { ...state, dialog: action.dialog };
    case 'dialog/close':
      return state.dialog?.id === action.id ? { ...state, dialog: null } : state;
    case 'clipboard/set':
      return state.clipboard === action.clipboard ? state : { ...state, clipboard: action.clipboard };
    case 'replace/open':
      return state.replaceOpen === action.open ? state : { ...state, replaceOpen: action.open };
    case 'zoom/set': {
      const zoom = Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, action.zoom)) * 100) / 100;
      return zoom === state.zoom ? state : { ...state, zoom };
    }
    case 'info/set':
      return state.infoOpen === action.open ? state : { ...state, infoOpen: action.open };
    case 'drag/set':
      return state.dragActive === action.active ? state : { ...state, dragActive: action.active };
    case 'recent/set':
      return { ...state, recentFiles: action.paths };
    case 'status/show':
      return { ...state, statusMessage: { text: action.text, token: nextToken() } };
    case 'focus/request':
      return { ...state, focusRequest: { target: action.target, token: nextToken() } };
    case 'search/query':
      return action.query === state.search.query
        ? state
        : { ...state, search: { ...state.search, query: action.query, results: null, current: -1 } };
    case 'search/options':
      return { ...state, search: { ...state.search, ...action.options, results: null, current: -1 } };
    case 'search/results':
      return { ...state, search: { ...state.search, results: action.results, current: -1 } };
    case 'search/current':
      return { ...state, search: { ...state.search, current: action.index } };
    default:
      return state;
  }
}
