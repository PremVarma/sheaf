import { memo, type ReactNode } from 'react';
import type { CellRange } from '../../models/workbook';
import { columnLabel, isWholeColumns, isWholeRows } from '../../utils/cellAddress';
import type { GridLayout } from './gridLayout';

interface ColumnHeadersProps {
  layout: GridLayout;
  c0: number;
  c1: number;
  xOrigin: number;
  selection: CellRange;
}

export const ColumnHeaders = memo(function ColumnHeaders({ layout, c0, c1, xOrigin, selection }: ColumnHeadersProps) {
  const whole = isWholeColumns(selection);
  const headers: ReactNode[] = [];
  for (let c = c0; c <= c1; c++) {
    const x0 = layout.x(c);
    const x1 = layout.x(c + 1);
    if (x1 <= x0) continue;
    const selected = c >= selection.c0 && c <= selection.c1;
    headers.push(
      <div
        key={c}
        className={`xv-colhdr${selected ? (whole ? ' is-whole' : ' is-sel') : ''}`}
        style={{ left: xOrigin + x0, width: x1 - x0 }}
        role="columnheader"
        aria-label={`Column ${columnLabel(c)}`}
      >
        {columnLabel(c)}
        <div className="xv-resize-col" data-resize-col={c} title="Drag to resize · double-click to fit" />
      </div>,
    );
  }
  return <>{headers}</>;
});

interface RowHeadersProps {
  layout: GridLayout;
  r0: number;
  r1: number;
  yOrigin: number;
  selection: CellRange;
}

export const RowHeaders = memo(function RowHeaders({ layout, r0, r1, yOrigin, selection }: RowHeadersProps) {
  const whole = isWholeRows(selection);
  const headers: ReactNode[] = [];
  for (let r = r0; r <= r1; r++) {
    const y0 = layout.y(r);
    const y1 = layout.y(r + 1);
    if (y1 <= y0) continue;
    const selected = r >= selection.r0 && r <= selection.r1;
    headers.push(
      <div
        key={r}
        className={`xv-rowhdr${selected ? (whole ? ' is-whole' : ' is-sel') : ''}`}
        style={{ top: yOrigin + y0, height: y1 - y0, width: layout.rowHeaderWidth }}
        role="rowheader"
      >
        {r + 1}
        <div className="xv-resize-row" data-resize-row={r} title="Drag to resize · double-click to reset" />
      </div>,
    );
  }
  return <>{headers}</>;
});
