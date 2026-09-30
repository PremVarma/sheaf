import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, type WorksheetModel } from '../../models/workbook';
import { Axis, scaled } from '../../utils/axis';

/** Column header height at 100%. */
export const HEADER_HEIGHT = 22;
/** Empty rows/columns shown past the data so the sheet reads as a spreadsheet, not a table. */
const ROW_PADDING = 20;
const COL_PADDING = 3;
/** Browsers stop laying out elements beyond ~33.5M px; stay well below. */
const MAX_CONTENT_PX = 30_000_000;
/** Frozen panes may take at most this share of the viewport. */
const MAX_FROZEN_SHARE = 0.75;

export interface Viewport {
  width: number;
  height: number;
}

/** Rows/columns in the grid: data + padding, enough to fill the window, rounded to limit re-layouts. */
export function displayCounts(
  sheet: WorksheetModel,
  extent: { rows: number; cols: number },
  zoom: number,
  viewport: Viewport,
): { rows: number; cols: number } {
  const fillRows = Math.ceil(viewport.height / Math.max(1, sheet.defaultRowHeight * zoom)) + 1;
  const fillCols = Math.ceil(viewport.width / Math.max(1, sheet.defaultColumnWidth * zoom)) + 1;
  const rows = Math.max(sheet.rowCount + ROW_PADDING, extent.rows + ROW_PADDING, fillRows);
  const cols = Math.max(sheet.columnCount + COL_PADDING, extent.cols + COL_PADDING, fillCols);
  return {
    rows: Math.min(EXCEL_MAX_ROWS, Math.ceil(rows / 50) * 50),
    cols: Math.min(EXCEL_MAX_COLS, Math.ceil(cols / 8) * 8),
  };
}

/** Frozen rows/columns, reduced if they would leave no room to scroll. */
export function clampFrozen(
  frozen: { rows: number; columns: number },
  rows: Axis,
  cols: Axis,
  zoom: number,
  viewport: Viewport,
): { rows: number; cols: number } {
  let r = Math.max(0, Math.min(frozen.rows, rows.count - 1));
  let c = Math.max(0, Math.min(frozen.columns, cols.count - 1));
  if (viewport.height > 0) {
    while (r > 0 && rows.offsetOf(r) * zoom > viewport.height * MAX_FROZEN_SHARE) r--;
  }
  if (viewport.width > 0) {
    while (c > 0 && cols.offsetOf(c) * zoom > viewport.width * MAX_FROZEN_SHARE) c--;
  }
  return { rows: r, cols: c };
}

/**
 * Pixel geometry of a sheet at a zoom level. Column/row positions are relative
 * to the top-left of the cell area (headers excluded).
 */
export class GridLayout {
  readonly rows: Axis;
  readonly cols: Axis;
  readonly zoom: number;
  readonly frozenRows: number;
  readonly frozenCols: number;
  readonly rowHeaderWidth: number;
  readonly colHeaderHeight: number;
  readonly frozenWidth: number;
  readonly frozenHeight: number;
  readonly totalWidth: number;
  readonly totalHeight: number;
  readonly contentWidth: number;
  readonly contentHeight: number;

  constructor(rows: Axis, cols: Axis, zoom: number, frozenRows: number, frozenCols: number) {
    this.rows = rows;
    this.cols = cols;
    // Keep the scrollable content within what the browser can lay out.
    this.zoom = Math.min(zoom, MAX_CONTENT_PX / Math.max(1, rows.total, cols.total));
    this.frozenRows = frozenRows;
    this.frozenCols = frozenCols;
    const digits = String(rows.count).length;
    this.rowHeaderWidth = Math.max(24, Math.round((digits * 7.5 + 16) * this.zoom));
    this.colHeaderHeight = Math.max(14, Math.round(HEADER_HEIGHT * this.zoom));
    this.frozenWidth = this.x(frozenCols);
    this.frozenHeight = this.y(frozenRows);
    this.totalWidth = this.x(cols.count);
    this.totalHeight = this.y(rows.count);
    this.contentWidth = this.rowHeaderWidth + this.totalWidth;
    this.contentHeight = this.colHeaderHeight + this.totalHeight;
  }

  x(col: number): number {
    return scaled(this.cols.offsetOf(col), this.zoom);
  }

  y(row: number): number {
    return scaled(this.rows.offsetOf(row), this.zoom);
  }

  colAt(px: number): number {
    return this.cols.indexAt(px / this.zoom);
  }

  rowAt(px: number): number {
    return this.rows.indexAt(px / this.zoom);
  }
}

/** Inclusive row/column range of the scrollable pane. Empty when r0 > r1 or c0 > c1. */
export interface RenderWindow {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

export function visibleWindow(layout: GridLayout, scrollLeft: number, scrollTop: number, viewport: Viewport): RenderWindow {
  const bodyWidth = Math.max(0, viewport.width - layout.rowHeaderWidth);
  const bodyHeight = Math.max(0, viewport.height - layout.colHeaderHeight);
  return {
    r0: Math.max(layout.frozenRows, layout.rowAt(scrollTop + layout.frozenHeight)),
    r1: Math.max(layout.frozenRows, layout.rowAt(scrollTop + bodyHeight)),
    c0: Math.max(layout.frozenCols, layout.colAt(scrollLeft + layout.frozenWidth)),
    c1: Math.max(layout.frozenCols, layout.colAt(scrollLeft + bodyWidth)),
  };
}

/** Adds overscan so fast scrolling doesn't reveal unrendered areas. */
export function expandWindow(window: RenderWindow, layout: GridLayout, rows = 12, cols = 4): RenderWindow {
  return {
    r0: Math.max(layout.frozenRows, window.r0 - rows),
    r1: Math.min(layout.rows.count - 1, window.r1 + rows),
    c0: Math.max(layout.frozenCols, window.c0 - cols),
    c1: Math.min(layout.cols.count - 1, window.c1 + cols),
  };
}

export function windowContains(outer: RenderWindow, inner: RenderWindow): boolean {
  return outer.r0 <= inner.r0 && outer.r1 >= inner.r1 && outer.c0 <= inner.c0 && outer.c1 >= inner.c1;
}
