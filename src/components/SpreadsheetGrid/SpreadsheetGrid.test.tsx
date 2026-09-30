import { act, fireEvent, render, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import type { WorksheetModel } from '../../models/workbook';
import { initialSheetView, type SheetViewState } from '../../state/viewerState';
import { createSheet } from '../../test/fixtures';
import { SpreadsheetGrid, type SpreadsheetGridProps } from './SpreadsheetGrid';

function renderGrid(sheet: WorksheetModel, overrides: Omit<Partial<SpreadsheetGridProps>, 'view'> & { view?: Partial<SheetViewState> } = {}) {
  const { view, ...rest } = overrides;
  const props: SpreadsheetGridProps = {
    sheet,
    styles: [{}],
    view: { ...initialSheetView(sheet), ...view },
    zoom: 1,
    matchFlags: null,
    currentMatch: null,
    isMac: true,
    onSelectionChange: vi.fn(),
    onColumnResize: vi.fn(),
    onRowResize: vi.fn(),
    onZoomChange: vi.fn(),
    onCopy: vi.fn(),
    onSelectAll: vi.fn(),
    onScrollSave: vi.fn(),
    ...rest,
  };
  const result = render(<SpreadsheetGrid {...props} />);
  const scroller = result.container.querySelector('.xv-scroller') as HTMLDivElement;
  return { ...result, props, scroller };
}

const cellTexts = (container: HTMLElement) => Array.from(container.querySelectorAll('.xv-cell'), (el) => el.textContent);

describe('SpreadsheetGrid', () => {
  it('renders only the visible part of a large worksheet', async () => {
    // 200,000 rows × 10 columns = 2 million cells.
    const sheet = createSheet(200_000, 10, (r, c) => (r * 7 + c) % 1000);
    const { container, scroller } = renderGrid(sheet);

    const initial = container.querySelectorAll('.xv-cell').length;
    expect(initial).toBeGreaterThan(100);
    expect(initial).toBeLessThan(1500);
    expect(container.querySelector('.xv-content')).toHaveStyle({ height: `${22 + 200_050 * 20}px` });

    await act(async () => {
      scroller.scrollTop = 199_950 * 20;
      fireEvent.scroll(scroller);
      await new Promise((resolve) => setTimeout(resolve, 50));
    });
    const rowHeaders = Array.from(container.querySelectorAll('.xv-rowhdr'), (el) => el.textContent);
    expect(rowHeaders).toContain('199960');
    expect(container.querySelectorAll('.xv-cell').length).toBeLessThan(1500);
  });

  it('moves the selection with the keyboard', () => {
    const sheet = createSheet(20, 5);
    const { scroller, props } = renderGrid(sheet);
    fireEvent.keyDown(scroller, { key: 'ArrowDown' });
    expect(props.onSelectionChange).toHaveBeenLastCalledWith({ anchor: { row: 1, col: 0 }, focus: { row: 1, col: 0 } }, true);
    fireEvent.keyDown(scroller, { key: 'ArrowRight', shiftKey: true });
    expect(props.onSelectionChange).toHaveBeenLastCalledWith({ anchor: { row: 1, col: 0 }, focus: { row: 1, col: 1 } }, 'focus');
    fireEvent.keyDown(scroller, { key: 'ArrowDown', metaKey: true });
    expect(props.onSelectionChange).toHaveBeenLastCalledWith({ anchor: { row: 19, col: 0 }, focus: { row: 19, col: 0 } }, true);
  });

  it('copies and selects all with Cmd+C / Cmd+A', () => {
    const { scroller, props } = renderGrid(createSheet(3, 3));
    fireEvent.keyDown(scroller, { key: 'c', metaKey: true });
    fireEvent.keyDown(scroller, { key: 'a', metaKey: true });
    expect(props.onCopy).toHaveBeenCalledTimes(1);
    expect(props.onSelectAll).toHaveBeenCalledTimes(1);
  });

  it('selects cells by clicking and extends with Shift+click', () => {
    const sheet = createSheet(10, 5);
    const { scroller, props } = renderGrid(sheet);
    // Row header is 31px wide and column headers are 22px tall at 100%.
    fireEvent.pointerDown(scroller, { button: 0, clientX: 31 + 64 + 10, clientY: 22 + 20 + 5 });
    fireEvent.pointerUp(window);
    expect(props.onSelectionChange).toHaveBeenLastCalledWith({ anchor: { row: 1, col: 1 }, focus: { row: 1, col: 1 } }, undefined);
    fireEvent.pointerDown(scroller, { button: 0, shiftKey: true, clientX: 31 + 64 * 3 + 10, clientY: 22 + 20 * 4 + 5 });
    fireEvent.pointerUp(window);
    expect(props.onSelectionChange).toHaveBeenLastCalledWith({ anchor: { row: 1, col: 1 }, focus: { row: 4, col: 3 } }, undefined);
  });

  it('resizes a column by dragging its header divider', async () => {
    const { container, props } = renderGrid(createSheet(5, 5));
    const handle = container.querySelector('[data-resize-col="1"]')!;
    fireEvent.pointerDown(handle, { button: 0, clientX: 160, clientY: 10 });
    fireEvent.pointerMove(window, { clientX: 220, clientY: 10 });
    await waitFor(() => expect(props.onColumnResize).toHaveBeenLastCalledWith(1, 124));
    fireEvent.pointerUp(window);
  });

  it('auto-fits a column when its divider is double-clicked', () => {
    const sheet = createSheet(3, 2, (_, c) => (c === 0 ? 'A fairly long piece of text' : 'x'));
    const { container, props } = renderGrid(sheet);
    fireEvent.doubleClick(container.querySelector('[data-resize-col="0"]')!);
    const [col, width] = (props.onColumnResize as ReturnType<typeof vi.fn>).mock.calls.at(-1)!;
    expect(col).toBe(0);
    expect(width).toBeGreaterThan(64);
  });

  it('applies column widths from the view', () => {
    const { container } = renderGrid(createSheet(2, 3), { view: { columnSizes: new Map([[0, 150]]) } });
    const headerB = Array.from(container.querySelectorAll('.xv-colhdr')).find((el) => el.textContent === 'B') as HTMLElement;
    expect(headerB.style.left).toBe(`${31 + 150}px`);
  });

  it('keeps frozen rows in the sticky header band', () => {
    const { container } = renderGrid(createSheet(50, 3, undefined, { frozen: { rows: 1, columns: 0 } }));
    const top = container.querySelector('.xv-top')!;
    expect(cellTexts(top as HTMLElement)).toContain('R1C1');
    const main = Array.from(container.querySelectorAll('.xv-content > .xv-row'));
    expect(main.some((row) => row.textContent?.includes('R1C1'))).toBe(false);
    expect(container.querySelector('.xv-freeze-h')).not.toBeNull();
  });

  it('draws a merged range once, with the top-left value', () => {
    const sheet = createSheet(3, 3, (r, c) => (r === 0 && c === 0 ? 'Merged title' : r === 0 ? 'hidden' : 'v'), {
      merges: [{ r0: 0, c0: 0, r1: 0, c1: 2 }],
    });
    const { container } = renderGrid(sheet);
    const merged = container.querySelectorAll('.xv-merge');
    expect(merged).toHaveLength(1);
    expect(merged[0].textContent).toBe('Merged title');
    expect(cellTexts(container)).not.toContain('hidden');
  });

  it('scales with the zoom level', () => {
    const { scroller } = renderGrid(createSheet(2, 2), { zoom: 1.5 });
    // 12px base font × 150%, set on the grid container so the cell editor inherits it too.
    expect((scroller.parentElement as HTMLElement).style.getPropertyValue('--xv-font')).toBe('18px');
    const headerB = Array.from(scroller.querySelectorAll('.xv-colhdr')).find((el) => el.textContent === 'B') as HTMLElement;
    expect(headerB.style.width).toBe('96px');
  });

  it('shows ### for numbers that do not fit, like Excel', () => {
    const sheet = createSheet(1, 2, (_, c) => (c === 0 ? 123456789012 : 'text'));
    const { container } = renderGrid(sheet);
    expect(cellTexts(container)[0]).toMatch(/^#+$/);
  });
});
