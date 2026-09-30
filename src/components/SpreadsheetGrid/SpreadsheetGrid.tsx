import {
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type Ref,
} from 'react';
import { useDarkMode } from '../../hooks/useColorScheme';
import { useElementSize } from '../../hooks/useElementSize';
import { stringAt } from '../../models/stringPool';
import {
  CellKind,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type CellRange,
  type CellRef,
  type CellStyle,
  type WorksheetModel,
} from '../../models/workbook';
import { findCellIndex, findDataEdge, lastUsedCell } from '../../services/workbook/sheetService';
import type { CommitMove } from '../../state/controller';
import type { CellEditState, Selection, SheetViewState } from '../../state/viewerState';
import { Axis } from '../../utils/axis';
import { normalizeRange } from '../../utils/cellAddress';
import { neededRowHeight } from '../../utils/rowHeight';
import { BASE_FONT_PX, measureText } from '../../utils/textMeasure';
import { buildCellCss } from './cellCss';
import { CellEditor } from './CellEditor';
import { ColumnHeaders, RowHeaders } from './GridHeaders';
import { GridPane } from './GridPane';
import {
  clampFrozen,
  displayCounts,
  expandWindow,
  GridLayout,
  visibleWindow,
  windowContains,
  type RenderWindow,
} from './gridLayout';
import { MergeIndex } from './mergeIndex';
import './grid.css';

const MIN_COLUMN_WIDTH = 12;
const MIN_ROW_HEIGHT = 8;

export type GridHitKind = 'corner' | 'column' | 'row' | 'cell';

export interface SpreadsheetGridProps {
  sheet: WorksheetModel;
  styles: readonly CellStyle[];
  view: SheetViewState;
  zoom: number;
  matchFlags: Uint8Array | null;
  currentMatch: CellRef | null;
  isMac: boolean;
  onSelectionChange(selection: Selection, reveal?: boolean | 'focus'): void;
  /** Live size while a divider is dragged. */
  onColumnResize(col: number, size: number | null): void;
  onRowResize(row: number, size: number | null): void;
  /** A resize finished (drag released or double-click to fit); null = default size. */
  onColumnSizes?(sizes: ReadonlyMap<number, number | null>): void;
  onRowSize?(row: number, size: number | null): void;
  /** Right-click on a cell or header (the selection has already been updated). */
  onContextMenu?(kind: GridHitKind, x: number, y: number): void;
  /** The fill handle was dragged over `range` (next to the selection). */
  onFill?(range: CellRange): void;
  /** The fill handle was double-clicked. */
  onAutoFill?(): void;
  /** Cells copied or cut on this sheet, outlined with a moving border. */
  clipboard?: { range: CellRange; cut: boolean } | null;
  onZoomChange(zoom: number): void;
  onCopy(): void;
  onSelectAll(): void;
  /** Called on unmount with the final scroll position. */
  onScrollSave(scroll: { left: number; top: number }, revealHandled: number): void;
  /** The cell being edited on this sheet, if any. */
  editing?: CellEditState | null;
  onStartEdit?(mode: CellEditState['mode'], initial?: string): void;
  onDraftChange?(text: string): void;
  onCommitEdit?(move: CommitMove): void;
  onCancelEdit?(): void;
  /** Delete key: clear the selection's contents. */
  onClear?(): void;
  onPaste?(): void;
  onCut?(): void;
  ref?: Ref<HTMLDivElement>;
}

type Hit =
  | { kind: 'corner' }
  | { kind: 'column'; col: number }
  | { kind: 'row'; row: number }
  | { kind: 'cell'; cell: CellRef };

type Drag =
  | { kind: 'cells'; anchor: CellRef }
  | { kind: 'columns'; anchor: CellRef }
  | { kind: 'rows'; anchor: CellRef }
  | { kind: 'fill'; source: CellRange }
  | { kind: 'resize-col'; col: number; startX: number; startSize: number; size?: number }
  | { kind: 'resize-row'; row: number; startY: number; startSize: number; size?: number };

/** The range the fill handle extends into when dragged to `cell`: the axis the pointer moved furthest along wins. */
function fillTarget(source: CellRange, cell: CellRef): CellRange | null {
  const down = cell.row - source.r1;
  const up = source.r0 - cell.row;
  const right = cell.col - source.c1;
  const left = source.c0 - cell.col;
  const vertical = Math.max(down, up);
  const horizontal = Math.max(right, left);
  if (vertical <= 0 && horizontal <= 0) return null;
  if (vertical >= horizontal) {
    return down > 0 ? { ...source, r0: source.r1 + 1, r1: cell.row } : { ...source, r0: cell.row, r1: source.r0 - 1 };
  }
  return right > 0 ? { ...source, c0: source.c1 + 1, c1: cell.col } : { ...source, c0: cell.col, c1: source.c0 - 1 };
}

const nextFrame: (callback: () => void) => number =
  typeof requestAnimationFrame === 'function' ? requestAnimationFrame : (callback) => window.setTimeout(callback, 16);
const cancelFrame: (handle: number) => void =
  typeof cancelAnimationFrame === 'function' ? cancelAnimationFrame : (handle) => window.clearTimeout(handle);

/**
 * Virtualized spreadsheet grid.
 *
 * Layout: one native scroll container whose content is the full sheet size.
 * Headers and frozen panes are `position: sticky` bands, so all scrolling is
 * done by the browser (smooth, momentum-friendly); React only renders the
 * cells inside the current window plus overscan.
 */
export function SpreadsheetGrid(props: SpreadsheetGridProps) {
  const { sheet, view, zoom, matchFlags, currentMatch } = props;
  const scrollerRef = useRef<HTMLDivElement>(null);
  useImperativeHandle(props.ref, () => scrollerRef.current as HTMLDivElement, []);
  const viewport = useElementSize(scrollerRef);
  const dark = useDarkMode();

  const css = useMemo(() => buildCellCss(props.styles, dark), [props.styles, dark]);
  const merges = useMemo(() => new MergeIndex(sheet.merges), [sheet]);

  const counts = displayCounts(sheet, view.extent, zoom, viewport);
  const [fillPreview, setFillPreview] = useState<CellRange | null>(null);
  const fillPreviewRef = useRef<CellRange | null>(null);
  const showFillPreview = useCallback((range: CellRange | null) => {
    fillPreviewRef.current = range;
    setFillPreview(range);
  }, []);
  const colAxis = useMemo(
    () => Axis.create(counts.cols, sheet.defaultColumnWidth, sheet.columnWidths, view.columnSizes),
    [counts.cols, sheet, view.columnSizes],
  );
  const rowAxis = useMemo(
    () => Axis.create(counts.rows, sheet.defaultRowHeight, sheet.rowHeights, view.rowSizes),
    [counts.rows, sheet, view.rowSizes],
  );
  const frozen = clampFrozen(sheet.frozen, rowAxis, colAxis, zoom, viewport);
  const layout = useMemo(
    () => new GridLayout(rowAxis, colAxis, zoom, frozen.rows, frozen.cols),
    [rowAxis, colAxis, zoom, frozen.rows, frozen.cols],
  );

  // Latest values for event handlers without re-binding them.
  const layoutRef = useRef(layout);
  layoutRef.current = layout;
  const propsRef = useRef(props);
  propsRef.current = props;

  // The selection we last emitted, so rapid events (key repeat, drags) never act on a
  // selection that hasn't been rendered yet.
  const selectionRef = useRef(view.selection);
  const renderedSelection = useRef(view.selection);
  if (renderedSelection.current !== view.selection) {
    renderedSelection.current = view.selection;
    selectionRef.current = view.selection;
  }
  const emitSelection = useCallback((selection: Selection, reveal?: boolean | 'focus') => {
    selectionRef.current = selection;
    propsRef.current.onSelectionChange(selection, reveal);
  }, []);

  const scrollPos = useRef({ left: view.scroll.left, top: view.scroll.top });
  const [win, setWin] = useState<RenderWindow>(() => ({ r0: 0, r1: -1, c0: 0, c1: -1 }));
  // Re-render on scroll while a cell editor is open, so it follows its cell.
  const [, setScrollTick] = useState(0);
  const editingRef = useRef(props.editing);
  editingRef.current = props.editing;

  const updateWindow = useCallback(() => {
    const el = scrollerRef.current;
    if (!el) return;
    scrollPos.current = { left: el.scrollLeft, top: el.scrollTop };
    const current = layoutRef.current;
    const visible = visibleWindow(current, el.scrollLeft, el.scrollTop, { width: el.clientWidth, height: el.clientHeight });
    setWin((prev) =>
      windowContains(prev, visible) && prev.r0 >= current.frozenRows && prev.c0 >= current.frozenCols
        ? prev
        : expandWindow(visible, current),
    );
  }, []);

  // Restore the sheet's scroll position on mount; save it on unmount.
  const revealHandled = useRef(view.revealHandled);
  useLayoutEffect(() => {
    const el = scrollerRef.current;
    if (el) {
      el.scrollLeft = view.scroll.left;
      el.scrollTop = view.scroll.top;
    }
    return () => propsRef.current.onScrollSave(scrollPos.current, revealHandled.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount/unmount only
  }, []);

  // Keep scroll position proportional when zooming.
  const previousZoom = useRef(layout.zoom);
  useLayoutEffect(() => {
    const el = scrollerRef.current;
    const old = previousZoom.current;
    previousZoom.current = layout.zoom;
    if (el && old !== layout.zoom) {
      el.scrollLeft = (el.scrollLeft * layout.zoom) / old;
      el.scrollTop = (el.scrollTop * layout.zoom) / old;
    }
  }, [layout.zoom]);

  useLayoutEffect(() => {
    updateWindow();
  }, [layout, viewport.width, viewport.height, updateWindow]);

  useEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    let frame = 0;
    const onScroll = () => {
      if (!frame) {
        frame = nextFrame(() => {
          frame = 0;
          updateWindow();
          if (editingRef.current) setScrollTick((tick) => tick + 1);
        });
      }
    };
    el.addEventListener('scroll', onScroll, { passive: true });
    return () => {
      el.removeEventListener('scroll', onScroll);
      if (frame) cancelFrame(frame);
    };
  }, [updateWindow]);

  /** Scrolls the minimum needed to show a cell outside the frozen panes. */
  const ensureVisible = useCallback(
    (cell: CellRef) => {
      const el = scrollerRef.current;
      if (!el) return;
      const current = layoutRef.current;
      const bodyWidth = el.clientWidth - current.rowHeaderWidth;
      const bodyHeight = el.clientHeight - current.colHeaderHeight;
      if (cell.col >= current.frozenCols && cell.col < current.cols.count) {
        const left = current.x(cell.col);
        const right = current.x(cell.col + 1);
        if (left < el.scrollLeft + current.frozenWidth) el.scrollLeft = left - current.frozenWidth;
        else if (right > el.scrollLeft + bodyWidth) el.scrollLeft = Math.min(left - current.frozenWidth, right - bodyWidth);
      }
      if (cell.row >= current.frozenRows && cell.row < current.rows.count) {
        const top = current.y(cell.row);
        const bottom = current.y(cell.row + 1);
        if (top < el.scrollTop + current.frozenHeight) el.scrollTop = top - current.frozenHeight;
        else if (bottom > el.scrollTop + bodyHeight) el.scrollTop = Math.min(top - current.frozenHeight, bottom - bodyHeight);
      }
      updateWindow();
    },
    [updateWindow],
  );

  useLayoutEffect(() => {
    const reveal = view.reveal;
    if (!reveal || reveal.token <= revealHandled.current) return;
    revealHandled.current = reveal.token;
    ensureVisible(reveal.cell);
  }, [view.reveal, layout, ensureVisible]);

  // Pinch / ctrl+wheel zoom. Needs a non-passive native listener to prevent page zoom.
  useEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    let pending: number | null = null;
    let gestureStart = 1;
    const apply = (next: number) => {
      pending = next;
      propsRef.current.onZoomChange(next);
    };
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      const base = pending ?? propsRef.current.zoom;
      apply(base * Math.exp(-event.deltaY * 0.01));
    };
    const onGestureStart = (event: Event) => {
      event.preventDefault();
      gestureStart = propsRef.current.zoom;
    };
    const onGestureChange = (event: Event) => {
      event.preventDefault();
      apply(gestureStart * ((event as Event & { scale?: number }).scale ?? 1));
    };
    const onGestureEnd = () => {
      pending = null;
    };
    // Two-finger pinch on touch screens without WebKit gesture events (Android).
    let pinch: { distance: number; zoom: number } | null = null;
    const spread = (touches: TouchList) => Math.hypot(touches[0].clientX - touches[1].clientX, touches[0].clientY - touches[1].clientY);
    const onTouchStart = (event: TouchEvent) => {
      if (event.touches.length === 2) pinch = { distance: spread(event.touches), zoom: propsRef.current.zoom };
    };
    const onTouchMove = (event: TouchEvent) => {
      if (!pinch || event.touches.length !== 2) return;
      event.preventDefault();
      apply(pinch.zoom * (spread(event.touches) / pinch.distance));
    };
    const onTouchEnd = (event: TouchEvent) => {
      if (event.touches.length < 2) pinch = pending = null;
    };
    const touchPinch = !('ongesturestart' in window);
    el.addEventListener('wheel', onWheel, { passive: false });
    el.addEventListener('gesturestart', onGestureStart);
    el.addEventListener('gesturechange', onGestureChange);
    el.addEventListener('gestureend', onGestureEnd);
    if (touchPinch) {
      el.addEventListener('touchstart', onTouchStart, { passive: true });
      el.addEventListener('touchmove', onTouchMove, { passive: false });
      el.addEventListener('touchend', onTouchEnd);
      el.addEventListener('touchcancel', onTouchEnd);
    }
    return () => {
      el.removeEventListener('wheel', onWheel);
      el.removeEventListener('gesturestart', onGestureStart);
      el.removeEventListener('gesturechange', onGestureChange);
      el.removeEventListener('gestureend', onGestureEnd);
      el.removeEventListener('touchstart', onTouchStart);
      el.removeEventListener('touchmove', onTouchMove);
      el.removeEventListener('touchend', onTouchEnd);
      el.removeEventListener('touchcancel', onTouchEnd);
    };
  }, []);

  // Hit testing (client coordinates → header/cell)

  const hitTest = useCallback((clientX: number, clientY: number, clamp = false): Hit | null => {
    const el = scrollerRef.current;
    if (!el) return null;
    const current = layoutRef.current;
    const rect = el.getBoundingClientRect();
    let x = clientX - rect.left;
    let y = clientY - rect.top;
    if (clamp) {
      x = Math.min(Math.max(x, current.rowHeaderWidth), el.clientWidth - 1);
      y = Math.min(Math.max(y, current.colHeaderHeight), el.clientHeight - 1);
    } else if (x < 0 || y < 0 || x >= el.clientWidth || y >= el.clientHeight) {
      return null; // scrollbar
    }
    const inColumnHeader = y < current.colHeaderHeight;
    const inRowHeader = x < current.rowHeaderWidth;
    if (inColumnHeader && inRowHeader) return { kind: 'corner' };
    const bx = x - current.rowHeaderWidth;
    const by = y - current.colHeaderHeight;
    const col = bx < current.frozenWidth ? current.colAt(bx) : current.colAt(bx + el.scrollLeft);
    const row = by < current.frozenHeight ? current.rowAt(by) : current.rowAt(by + el.scrollTop);
    if (inColumnHeader) return { kind: 'column', col };
    if (inRowHeader) return { kind: 'row', row };
    return { kind: 'cell', cell: { row, col } };
  }, []);

  // Pointer interaction: selection by drag (with auto-scroll) and resizing.

  const drag = useRef<Drag | null>(null);
  const pointer = useRef<{ x: number; y: number } | null>(null);
  const autoScrollFrame = useRef(0);
  const resizeFrame = useRef(0);

  const extendDrag = useCallback(() => {
    const d = drag.current;
    const p = pointer.current;
    if (!d || !p) return;
    const hit = hitTest(p.x, p.y, true);
    if (!hit) return;
    const onSelectionChange = emitSelection;
    if (d.kind === 'cells' && hit.kind === 'cell') {
      onSelectionChange({ anchor: d.anchor, focus: hit.cell });
    } else if (d.kind === 'columns') {
      const col = hit.kind === 'cell' ? hit.cell.col : hit.kind === 'column' ? hit.col : d.anchor.col;
      onSelectionChange({ anchor: d.anchor, focus: { row: EXCEL_MAX_ROWS - 1, col } });
    } else if (d.kind === 'rows') {
      const row = hit.kind === 'cell' ? hit.cell.row : hit.kind === 'row' ? hit.row : d.anchor.row;
      onSelectionChange({ anchor: d.anchor, focus: { row, col: EXCEL_MAX_COLS - 1 } });
    } else if (d.kind === 'fill' && hit.kind === 'cell') {
      showFillPreview(fillTarget(d.source, hit.cell));
    }
  }, [hitTest, emitSelection, showFillPreview]);

  const autoScroll = useCallback(() => {
    autoScrollFrame.current = 0;
    const el = scrollerRef.current;
    const p = pointer.current;
    const d = drag.current;
    if (!el || !p || !d || d.kind === 'resize-col' || d.kind === 'resize-row') return;
    const current = layoutRef.current;
    const rect = el.getBoundingClientRect();
    const left = rect.left + current.rowHeaderWidth + current.frozenWidth;
    const top = rect.top + current.colHeaderHeight + current.frozenHeight;
    const right = rect.left + el.clientWidth;
    const bottom = rect.top + el.clientHeight;
    const speed = (distance: number) => Math.sign(distance) * Math.min(40, Math.abs(distance) / 2 + 4);
    const dx = d.kind === 'rows' ? 0 : p.x < left ? speed(p.x - left) : p.x > right ? speed(p.x - right) : 0;
    const dy = d.kind === 'columns' ? 0 : p.y < top ? speed(p.y - top) : p.y > bottom ? speed(p.y - bottom) : 0;
    if (dx || dy) {
      el.scrollLeft += dx;
      el.scrollTop += dy;
      extendDrag();
      autoScrollFrame.current = nextFrame(autoScroll);
    }
  }, [extendDrag]);

  const endDrag = useCallback(() => {
    const d = drag.current;
    drag.current = null;
    pointer.current = null;
    // Resizes and fills take effect when the pointer is released.
    if (d?.kind === 'resize-col' && d.size !== undefined) propsRef.current.onColumnSizes?.(new Map([[d.col, d.size]]));
    if (d?.kind === 'resize-row' && d.size !== undefined) propsRef.current.onRowSize?.(d.row, d.size);
    if (d?.kind === 'fill') {
      const target = fillPreviewRef.current;
      showFillPreview(null);
      if (target) propsRef.current.onFill?.(target);
    }
    if (autoScrollFrame.current) cancelFrame(autoScrollFrame.current);
    autoScrollFrame.current = 0;
    document.documentElement.classList.remove('xv-resizing-col', 'xv-resizing-row');
  }, [showFillPreview]);

  const onWindowPointerMove = useCallback(
    (event: PointerEvent) => {
      const d = drag.current;
      if (!d) return;
      pointer.current = { x: event.clientX, y: event.clientY };
      if (d.kind === 'resize-col' || d.kind === 'resize-row') {
        if (resizeFrame.current) return;
        resizeFrame.current = nextFrame(() => {
          resizeFrame.current = 0;
          const active = drag.current;
          const p = pointer.current;
          if (!active || !p) return;
          const zoomLevel = layoutRef.current.zoom;
          if (active.kind === 'resize-col') {
            active.size = Math.round(Math.max(MIN_COLUMN_WIDTH, active.startSize + (p.x - active.startX) / zoomLevel));
            propsRef.current.onColumnResize(active.col, active.size);
          } else if (active.kind === 'resize-row') {
            active.size = Math.round(Math.max(MIN_ROW_HEIGHT, active.startSize + (p.y - active.startY) / zoomLevel));
            propsRef.current.onRowResize(active.row, active.size);
          }
        });
        return;
      }
      extendDrag();
      if (!autoScrollFrame.current) autoScrollFrame.current = nextFrame(autoScroll);
    },
    [autoScroll, extendDrag],
  );

  const onWindowPointerUp = useCallback(() => {
    endDrag();
    window.removeEventListener('pointermove', onWindowPointerMove);
    window.removeEventListener('pointerup', onWindowPointerUp);
    window.removeEventListener('pointercancel', onWindowPointerUp);
  }, [endDrag, onWindowPointerMove]);

  const startDrag = useCallback(
    (next: Drag, event: { clientX: number; clientY: number }) => {
      drag.current = next;
      pointer.current = { x: event.clientX, y: event.clientY };
      window.addEventListener('pointermove', onWindowPointerMove);
      window.addEventListener('pointerup', onWindowPointerUp);
      window.addEventListener('pointercancel', onWindowPointerUp);
    },
    [onWindowPointerMove, onWindowPointerUp],
  );

  useEffect(() => onWindowPointerUp, [onWindowPointerUp]);

  const onPointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    const el = scrollerRef.current;
    if (!el) return;
    const target = event.target as HTMLElement;
    const current = layoutRef.current;

    const colHandle = target.closest<HTMLElement>('[data-resize-col]');
    if (colHandle) {
      event.preventDefault();
      const col = Number(colHandle.dataset.resizeCol);
      document.documentElement.classList.add('xv-resizing-col');
      startDrag({ kind: 'resize-col', col, startX: event.clientX, startSize: current.cols.sizeOf(col) }, event);
      return;
    }
    const rowHandle = target.closest<HTMLElement>('[data-resize-row]');
    if (rowHandle) {
      event.preventDefault();
      const row = Number(rowHandle.dataset.resizeRow);
      document.documentElement.classList.add('xv-resizing-row');
      startDrag({ kind: 'resize-row', row, startY: event.clientY, startSize: current.rows.sizeOf(row) }, event);
      return;
    }
    if (target.closest('[data-fill-handle]')) {
      event.preventDefault();
      el.focus({ preventScroll: true });
      const { anchor: a, focus: f } = selectionRef.current;
      startDrag({ kind: 'fill', source: normalizeRange(a, f) }, event);
      return;
    }

    const hit = hitTest(event.clientX, event.clientY);
    if (!hit) return;
    event.preventDefault();
    el.focus({ preventScroll: true });
    const { anchor } = selectionRef.current;
    const onSelectionChange = emitSelection;
    const { onSelectAll } = propsRef.current;

    switch (hit.kind) {
      case 'corner':
        onSelectAll();
        break;
      case 'column': {
        const start = event.shiftKey ? { row: 0, col: anchor.col } : { row: 0, col: hit.col };
        onSelectionChange({ anchor: start, focus: { row: EXCEL_MAX_ROWS - 1, col: hit.col } });
        startDrag({ kind: 'columns', anchor: start }, event);
        break;
      }
      case 'row': {
        const start = event.shiftKey ? { row: anchor.row, col: 0 } : { row: hit.row, col: 0 };
        onSelectionChange({ anchor: start, focus: { row: hit.row, col: EXCEL_MAX_COLS - 1 } });
        startDrag({ kind: 'rows', anchor: start }, event);
        break;
      }
      case 'cell': {
        const start = event.shiftKey ? anchor : hit.cell;
        onSelectionChange({ anchor: start, focus: hit.cell });
        startDrag({ kind: 'cells', anchor: start }, event);
        break;
      }
    }
  };

  /** Double-click a column divider to fit its content (all selected columns, like Excel); a row divider fits the row. */
  const onDoubleClick = (event: React.MouseEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement;
    const colHandle = target.closest<HTMLElement>('[data-resize-col]');
    if (colHandle) {
      const col = Number(colHandle.dataset.resizeCol);
      const range = normalizeRange(selectionRef.current.anchor, selectionRef.current.focus);
      const columns = range.r0 === 0 && range.r1 >= EXCEL_MAX_ROWS - 1 && col >= range.c0 && col <= range.c1 && range.c1 - range.c0 < 200
        ? Array.from({ length: range.c1 - range.c0 + 1 }, (_, k) => range.c0 + k)
        : [col];
      const sizes = new Map(columns.map((c) => [c, fitColumnWidth(c)] as const));
      propsRef.current.onColumnSizes?.(sizes);
      if (!propsRef.current.onColumnSizes) propsRef.current.onColumnResize(col, sizes.get(col) ?? null);
      return;
    }
    const rowHandle = target.closest<HTMLElement>('[data-resize-row]');
    if (rowHandle) {
      const row = Number(rowHandle.dataset.resizeRow);
      if (propsRef.current.onRowSize) propsRef.current.onRowSize(row, fitRowHeight(row));
      else propsRef.current.onRowResize(row, null);
      return;
    }
    if (target.closest('[data-fill-handle]')) {
      propsRef.current.onAutoFill?.();
      return;
    }
    if (hitTest(event.clientX, event.clientY)?.kind === 'cell') propsRef.current.onStartEdit?.('edit');
  };

  /** Right-click: select what was clicked unless it's already selected, then show the menu. */
  const onContextMenu = (event: React.MouseEvent<HTMLDivElement>) => {
    const hit = hitTest(event.clientX, event.clientY);
    if (!hit || !propsRef.current.onContextMenu) return;
    event.preventDefault();
    scrollerRef.current?.focus({ preventScroll: true });
    const range = normalizeRange(selectionRef.current.anchor, selectionRef.current.focus);
    if (hit.kind === 'cell') {
      const { row, col } = hit.cell;
      if (row < range.r0 || row > range.r1 || col < range.c0 || col > range.c1) emitSelection({ anchor: hit.cell, focus: hit.cell });
    } else if (hit.kind === 'column') {
      const whole = range.r0 === 0 && range.r1 >= EXCEL_MAX_ROWS - 1 && hit.col >= range.c0 && hit.col <= range.c1;
      if (!whole) emitSelection({ anchor: { row: 0, col: hit.col }, focus: { row: EXCEL_MAX_ROWS - 1, col: hit.col } });
    } else if (hit.kind === 'row') {
      const whole = range.c0 === 0 && range.c1 >= EXCEL_MAX_COLS - 1 && hit.row >= range.r0 && hit.row <= range.r1;
      if (!whole) emitSelection({ anchor: { row: hit.row, col: 0 }, focus: { row: hit.row, col: EXCEL_MAX_COLS - 1 } });
    }
    propsRef.current.onContextMenu(hit.kind, event.clientX, event.clientY);
  };

  const fitColumnWidth = (col: number): number | null => {
    const { rowStart, kinds, texts, styles } = sheet.cells;
    const rowCount = sheet.rowCount;
    // Measure everything for normal sheets; sample around the viewport for huge ones.
    const ranges: [number, number][] =
      rowCount <= 50_000
        ? [[0, rowCount - 1]]
        : [
            [0, 999],
            [Math.max(0, win.r0 - 5000), Math.min(rowCount - 1, win.r1 + 5000)],
          ];
    let widest = 0;
    for (const [from, to] of ranges) {
      for (let r = from; r <= to; r++) {
        if (rowStart[r] === rowStart[r + 1]) continue;
        const i = findCellIndex(sheet, r, col);
        if (i < 0 || kinds[i] === CellKind.Empty || merges.at(r, col)) continue;
        const cellCss = css[styles[i]] ?? css[0];
        const lines = stringAt(sheet.strings, texts[i]).split('\n');
        const tabular = kinds[i] === CellKind.Number || kinds[i] === CellKind.Date;
        for (const line of cellCss.wrap ? lines : lines.slice(0, 1)) {
          widest = Math.max(widest, measureText(line, cellCss.bold, cellCss.italic, tabular, cellCss.fontFamily) * cellCss.fontScale + cellCss.indentPx);
        }
      }
    }
    return widest > 0 ? Math.min(800, Math.ceil(widest + 10)) : null;
  };

  /** Height a row needs for its largest font and its wrapped text; null when the default fits. */
  const fitRowHeight = (row: number): number | null => {
    const needed = neededRowHeight(sheet, row, propsRef.current.styles, (col) => layoutRef.current.cols.sizeOf(col));
    return needed > sheet.defaultRowHeight ? needed : null;
  };

  // Keyboard navigation

  const onKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (event.nativeEvent.isComposing) return;
    const primary = props.isMac ? event.metaKey : event.ctrlKey;
    /** A printable key starts editing with that character (Option on macOS still types symbols). */
    const startTyping = (): boolean => {
      const printable = event.key.length === 1 && !event.metaKey && !event.ctrlKey && (props.isMac || !event.altKey);
      if (printable && propsRef.current.onStartEdit) propsRef.current.onStartEdit('enter', event.key);
      return printable && Boolean(propsRef.current.onStartEdit);
    };
    const current = layoutRef.current;
    const selection = selectionRef.current;
    const onSelectionChange = emitSelection;
    const edgeLimits = { rows: Math.max(current.rows.count, sheet.rowCount), cols: Math.max(current.cols.count, sheet.columnCount) };

    const step = (from: CellRef, dRow: number, dCol: number): CellRef => {
      let { row, col } = from;
      if (dRow) {
        let r = row + dRow;
        while (r >= 0 && r < current.rows.count && current.rows.isHidden(r)) r += dRow;
        if (r >= 0 && r < EXCEL_MAX_ROWS) row = r;
      }
      if (dCol) {
        let c = col + dCol;
        while (c >= 0 && c < current.cols.count && current.cols.isHidden(c)) c += dCol;
        if (c >= 0 && c < EXCEL_MAX_COLS) col = c;
      }
      return { row, col };
    };
    const move = (dRow: -1 | 0 | 1, dCol: -1 | 0 | 1, extend: boolean, jump: boolean) => {
      const from = extend ? selection.focus : selection.anchor;
      const target = jump ? findDataEdge(sheet, from, dRow, dCol, edgeLimits) : step(from, dRow, dCol);
      onSelectionChange(extend ? { anchor: selection.anchor, focus: target } : { anchor: target, focus: target }, extend ? 'focus' : true);
    };
    const pageRows = Math.max(1, Math.floor((viewport.height - current.colHeaderHeight - current.frozenHeight) / (sheet.defaultRowHeight * current.zoom)) - 1);
    const pageCols = Math.max(1, Math.floor((viewport.width - current.rowHeaderWidth - current.frozenWidth) / (sheet.defaultColumnWidth * current.zoom)) - 1);
    const jumpTo = (cell: CellRef, extend: boolean) =>
      onSelectionChange(extend ? { anchor: selection.anchor, focus: cell } : { anchor: cell, focus: cell }, extend ? 'focus' : true);

    switch (event.key) {
      case ' ':
        // Shift+Space selects the rows, Ctrl+Space the columns of the selection (Excel).
        if (event.shiftKey && !event.ctrlKey && !event.metaKey) {
          const range = normalizeRange(selection.anchor, selection.focus);
          onSelectionChange({ anchor: { row: range.r0, col: 0 }, focus: { row: range.r1, col: EXCEL_MAX_COLS - 1 } });
          break;
        }
        if (event.ctrlKey && !event.metaKey && !event.altKey) {
          const range = normalizeRange(selection.anchor, selection.focus);
          onSelectionChange({ anchor: { row: 0, col: range.c0 }, focus: { row: EXCEL_MAX_ROWS - 1, col: range.c1 } });
          break;
        }
        if (!startTyping()) return;
        break;
      case 'ArrowUp':
        move(-1, 0, event.shiftKey, primary);
        break;
      case 'ArrowDown':
        move(1, 0, event.shiftKey, primary);
        break;
      case 'ArrowLeft':
        move(0, -1, event.shiftKey, primary);
        break;
      case 'ArrowRight':
        move(0, 1, event.shiftKey, primary);
        break;
      case 'PageDown':
      case 'PageUp': {
        if (event.ctrlKey) return; // sheet switching is a global shortcut
        const sign = event.key === 'PageDown' ? 1 : -1;
        const from = event.shiftKey ? selection.focus : selection.anchor;
        const target = event.altKey
          ? { row: from.row, col: Math.min(EXCEL_MAX_COLS - 1, Math.max(0, from.col + sign * pageCols)) }
          : { row: Math.min(EXCEL_MAX_ROWS - 1, Math.max(0, from.row + sign * pageRows)), col: from.col };
        jumpTo(target, event.shiftKey);
        break;
      }
      case 'Home':
        jumpTo(primary ? { row: 0, col: 0 } : { row: (event.shiftKey ? selection.focus : selection.anchor).row, col: 0 }, event.shiftKey);
        break;
      case 'End':
        if (!primary) return;
        jumpTo(lastUsedCell(sheet), event.shiftKey);
        break;
      case 'Tab':
        onSelectionChange(
          { anchor: step(selection.anchor, 0, event.shiftKey ? -1 : 1), focus: step(selection.anchor, 0, event.shiftKey ? -1 : 1) },
          true,
        );
        break;
      case 'Enter':
        onSelectionChange(
          { anchor: step(selection.anchor, event.shiftKey ? -1 : 1, 0), focus: step(selection.anchor, event.shiftKey ? -1 : 1, 0) },
          true,
        );
        break;
      case 'a':
      case 'A':
        if (!primary || event.altKey) return;
        propsRef.current.onSelectAll();
        break;
      case 'c':
      case 'C':
        if (!primary || event.altKey || event.shiftKey) {
          if (!startTyping()) return;
          break;
        }
        propsRef.current.onCopy();
        break;
      case 'v':
      case 'V':
        if (!primary || event.altKey || event.shiftKey) {
          if (!startTyping()) return;
          break;
        }
        propsRef.current.onPaste?.();
        break;
      case 'x':
      case 'X':
        if (!primary || event.altKey || event.shiftKey) {
          if (!startTyping()) return;
          break;
        }
        propsRef.current.onCut?.();
        break;
      case 'F2':
        propsRef.current.onStartEdit?.('edit');
        break;
      case 'Backspace':
        // Like Excel: clear the cell and start typing.
        propsRef.current.onStartEdit?.('enter', '');
        break;
      case 'Delete':
        propsRef.current.onClear?.();
        break;
      default:
        if (!startTyping()) return;
    }
    event.preventDefault();
  };

  // Rendering

  const selection = normalizeRange(view.selection.anchor, view.selection.focus);
  const active: CellRange = useMemo(() => {
    const { row, col } = view.selection.anchor;
    const merge = merges.at(row, col);
    return merge ?? { r0: row, r1: row, c0: col, c1: col };
  }, [view.selection.anchor, merges]);

  const {
    rowHeaderWidth: RH,
    colHeaderHeight: CH,
    frozenRows: FR,
    frozenCols: FC,
    frozenWidth: FW,
    frozenHeight: FH,
  } = layout;
  const lastRow = layout.rows.count - 1;
  const lastCol = layout.cols.count - 1;
  const w = {
    r0: Math.max(win.r0, FR),
    r1: Math.min(win.r1, lastRow),
    c0: Math.max(win.c0, FC),
    c1: Math.min(win.c1, lastCol),
  };
  const shared = {
    sheet,
    layout,
    css,
    merges,
    matchFlags,
    showGridlines: sheet.showGridlines,
    selection,
    active,
    currentMatch,
    fillPreview,
    fillHandle: Boolean(props.onFill) && !props.editing,
    clipboard: props.clipboard ?? null,
  };

  const rootStyle = {
    '--xv-zoom': layout.zoom,
    '--xv-font': `${BASE_FONT_PX * layout.zoom}px`,
    '--xv-hdr-font': `${Math.max(8, 11 * layout.zoom)}px`,
    '--xv-hdr-h': `${layout.colHeaderHeight}px`,
  } as CSSProperties;

  // In-cell editor position, in the container's coordinates.
  const editing = props.editing;
  let editor: React.ReactNode = null;
  if (editing && scrollerRef.current) {
    const el = scrollerRef.current;
    const x0 = layout.x(editing.col);
    const y0 = layout.y(editing.row);
    const cellIndex = findCellIndex(sheet, editing.row, editing.col);
    editor = (
      <CellEditor
        rect={{
          left: RH + x0 - (editing.col >= FC ? el.scrollLeft : 0),
          top: CH + y0 - (editing.row >= FR ? el.scrollTop : 0),
          width: layout.x(editing.col + 1) - x0,
          height: layout.y(editing.row + 1) - y0,
        }}
        maxRight={el.clientWidth - 4}
        draft={editing.draft}
        mode={editing.mode}
        zoom={layout.zoom}
        css={css[cellIndex >= 0 ? sheet.cells.styles[cellIndex] : 0] ?? css[0]}
        focused={editing.source === 'cell'}
        isMac={props.isMac}
        onChange={(text) => props.onDraftChange?.(text)}
        onCommit={(move) => props.onCommitEdit?.(move)}
        onCancel={() => props.onCancelEdit?.()}
      />
    );
  }

  return (
    <div className="xv-grid" style={rootStyle}>
    <div
      ref={scrollerRef}
      className="xv-scroller"
      tabIndex={0}
      role="grid"
      aria-label={`Sheet ${sheet.name}`}
      aria-rowcount={sheet.rowCount}
      aria-colcount={sheet.columnCount}
      data-grid-root=""
      onPointerDown={onPointerDown}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
      onKeyDown={onKeyDown}
    >
      <div className="xv-content" style={{ width: layout.contentWidth, height: layout.contentHeight }}>
        {/* Sticky top band: column headers + frozen rows */}
        <div className="xv-top" style={{ height: CH + FH, width: layout.contentWidth }}>
          <ColumnHeaders layout={layout} c0={w.c0} c1={w.c1} xOrigin={RH} selection={selection} />
          <GridPane {...shared} r0={0} r1={FR - 1} c0={w.c0} c1={w.c1} colMax={lastCol} xOrigin={RH} yOrigin={CH} />
          {FR > 0 && <div className="xv-freeze-h" style={{ top: CH + FH - 1, width: layout.contentWidth }} />}
          {/* Sticky corner: select-all button + frozen corner cells */}
          <div className="xv-corner" style={{ width: RH + FW, height: CH + FH }}>
            <div className="xv-select-all" style={{ width: RH, height: CH }} title="Select all" />
            <ColumnHeaders layout={layout} c0={0} c1={FC - 1} xOrigin={RH} selection={selection} />
            <RowHeaders layout={layout} r0={0} r1={FR - 1} yOrigin={CH} selection={selection} />
            <GridPane {...shared} r0={0} r1={FR - 1} c0={0} c1={FC - 1} colMax={FC - 1} xOrigin={RH} yOrigin={CH} />
            {FR > 0 && <div className="xv-freeze-h" style={{ top: CH + FH - 1, width: RH + FW }} />}
            {FC > 0 && <div className="xv-freeze-v" style={{ left: RH + FW - 1, top: 0, height: CH + FH }} />}
          </div>
        </div>
        {/* Sticky left band: row headers + frozen columns */}
        <div className="xv-left" style={{ width: RH + FW, height: layout.totalHeight - FH }}>
          <RowHeaders layout={layout} r0={w.r0} r1={w.r1} yOrigin={-FH} selection={selection} />
          <GridPane {...shared} r0={w.r0} r1={w.r1} c0={0} c1={FC - 1} colMax={FC - 1} xOrigin={RH} yOrigin={-FH} />
          {FC > 0 && <div className="xv-freeze-v" style={{ left: RH + FW - 1, top: 0, height: layout.totalHeight - FH }} />}
        </div>
        {/* Scrolling cells */}
        <GridPane {...shared} r0={w.r0} r1={w.r1} c0={w.c0} c1={w.c1} colMax={lastCol} xOrigin={RH} yOrigin={CH} />
      </div>
    </div>
    {editor}
    </div>
  );
}
