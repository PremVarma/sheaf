import { memo, type CSSProperties, type ReactNode } from 'react';
import { stringAt } from '../../models/stringPool';
import { CellKind, type CellRange, type CellRef, type WorksheetModel } from '../../models/workbook';
import { findCellIndex, lowerBoundColumn } from '../../services/workbook/sheetService';
import { measureText } from '../../utils/textMeasure';
import { KIND_ALIGN_CLASS, type CellCss } from './cellCss';
import type { GridLayout } from './gridLayout';
import { isCovered, type ColumnSpan, type MergeIndex } from './mergeIndex';

/** Horizontal cell padding at 100%. */
const PAD = 3;

export interface PaneProps {
  sheet: WorksheetModel;
  layout: GridLayout;
  css: readonly CellCss[];
  merges: MergeIndex;
  matchFlags: Uint8Array | null;
  /** Rows/columns to render (inclusive). */
  r0: number;
  r1: number;
  c0: number;
  c1: number;
  /** Last column of this pane; text never overflows past it. */
  colMax: number;
  /** Position of column 0 / row 0 inside the pane's container. */
  xOrigin: number;
  yOrigin: number;
  showGridlines: boolean;
  selection: CellRange;
  /** Active cell, expanded to its merged range. */
  active: CellRange;
  currentMatch: CellRef | null;
  /** Range the fill handle is being dragged over. */
  fillPreview: CellRange | null;
  /** Show the fill handle at the selection's bottom-right corner. */
  fillHandle: boolean;
  /** Copied or cut cells, outlined with a moving border. */
  clipboard: { range: CellRange; cut: boolean } | null;
}

/** One pane (main, frozen top, frozen left or frozen corner) of the grid. */
export function GridPane(props: PaneProps) {
  const { sheet, layout, r0, r1, c0, c1 } = props;
  if (r0 > r1 || c0 > c1) return null;

  const rows: ReactNode[] = [];
  for (let r = r0; r <= r1; r++) {
    const top = layout.y(r);
    const height = layout.y(r + 1) - top;
    if (height <= 0) continue;
    rows.push(
      <PaneRow
        key={r}
        sheet={sheet}
        row={r}
        top={props.yOrigin + top}
        height={height}
        c0={c0}
        c1={c1}
        colMax={props.colMax}
        layout={layout}
        css={props.css}
        xOrigin={props.xOrigin}
        matchFlags={props.matchFlags}
        covered={props.merges.covered(r)}
      />,
    );
  }

  return (
    <>
      {props.showGridlines && (
        <GridLines layout={layout} r0={r0} r1={r1} c0={c0} c1={c1} xOrigin={props.xOrigin} yOrigin={props.yOrigin} />
      )}
      {rows}
      <MergedCells {...props} />
      <SelectionOverlay {...props} />
    </>
  );
}

interface PaneRowProps {
  sheet: WorksheetModel;
  row: number;
  top: number;
  height: number;
  c0: number;
  c1: number;
  colMax: number;
  layout: GridLayout;
  css: readonly CellCss[];
  xOrigin: number;
  matchFlags: Uint8Array | null;
  covered: readonly ColumnSpan[];
}

/**
 * The cells of one row within the render window. Memoized: while scrolling
 * vertically, rows that stay in view are not re-rendered.
 */
const PaneRow = memo(function PaneRow(props: PaneRowProps) {
  const { sheet, row, c0, c1, colMax, layout, css, xOrigin, matchFlags, covered } = props;
  const children: ReactNode[] = [];

  if (row < sheet.rowCount) {
    const { rowStart, cols, kinds, texts, styles } = sheet.cells;
    const start = rowStart[row];
    const end = rowStart[row + 1];
    const zoom = layout.zoom;
    const pad = PAD * zoom * 2;

    const render = (i: number) => {
      const col = cols[i];
      if (isCovered(covered, col)) return;
      const x0 = layout.x(col);
      const x1 = layout.x(col + 1);
      if (x1 <= x0) return;
      const kind = kinds[i];
      const cellCss = css[styles[i]] ?? css[0];
      let text = stringAt(sheet.strings, texts[i]);
      let spanEnd = col + 1;

      if (text) {
        if (kind === CellKind.String) {
          if (!cellCss.wrap && text.includes('\n')) text = text.replace(/\r?\n/g, ' ');
          if (cellCss.canOverflow) {
            const needed = (measureText(text, cellCss.bold, cellCss.italic, false, cellCss.fontFamily) * cellCss.fontScale + cellCss.indentPx) * zoom + pad;
            if (needed > x1 - x0) {
              // Spill into following empty cells, like Excel.
              let limit = colMax + 1;
              for (let k = i + 1; k < end; k++) {
                if (kinds[k] !== CellKind.Empty) {
                  limit = cols[k];
                  break;
                }
              }
              for (const span of covered) if (span[0] > col && span[0] < limit) limit = span[0];
              while (spanEnd < limit && layout.x(spanEnd) - x0 < needed) spanEnd++;
            }
          }
        } else if ((kind === CellKind.Number || kind === CellKind.Date) && !cellCss.wrap) {
          // Numbers that don't fit show ### rather than misleading truncated digits.
          // Text may use the cell padding, so allow 1px either side.
          const scale = cellCss.fontScale * zoom;
          if (measureText(text, cellCss.bold, cellCss.italic, true, cellCss.fontFamily) * scale + 2 > x1 - x0 - 1) {
            const hash = measureText('#', cellCss.bold, cellCss.italic, false, cellCss.fontFamily) * scale;
            text = '#'.repeat(Math.max(1, Math.floor((x1 - x0 - pad) / hash)));
          }
        }
      }
      if (col < c0 && spanEnd <= c0) return; // a left neighbor that doesn't reach the window

      const width = layout.x(spanEnd) - x0;
      const overflow = spanEnd > col + 1;
      let className = `xv-cell ${cellCss.alignClass ?? KIND_ALIGN_CLASS[kind]}${cellCss.extraClass}`;
      if (kind === CellKind.Number || kind === CellKind.Date) className += ' xv-num';
      if (overflow) className += ' xv-ovf';
      if (matchFlags && matchFlags[i]) className += ' xv-match';
      const style: CSSProperties = { left: xOrigin + x0, width: cellCss.opaque && !overflow ? width : width - 1, ...cellCss.style };
      children.push(
        <div key={col} className={className} style={style}>
          {text}
        </div>,
      );
      if (cellCss.borderStyle) {
        children.push(
          <div key={`b${col}`} className="xv-border" style={{ left: xOrigin + x0 - 1, width: x1 - x0 + 1, ...cellCss.borderStyle }} />,
        );
      }
    };

    let i = lowerBoundColumn(cols, start, end, c0);
    // Text in a cell left of the window can spill into it.
    let j = i - 1;
    while (j >= start && kinds[j] === CellKind.Empty && !css[styles[j]]?.opaque) j--;
    if (j >= start && kinds[j] === CellKind.String) render(j);
    for (; i < end && cols[i] <= c1; i++) render(i);
  }

  return (
    <div className="xv-row" style={{ top: props.top, height: props.height }}>
      {children}
    </div>
  );
});

interface GridLinesProps {
  layout: GridLayout;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
  xOrigin: number;
  yOrigin: number;
}

const GridLines = memo(function GridLines({ layout, r0, r1, c0, c1, xOrigin, yOrigin }: GridLinesProps) {
  const top = yOrigin + layout.y(r0);
  const height = layout.y(r1 + 1) - layout.y(r0);
  const left = xOrigin + layout.x(c0);
  const width = layout.x(c1 + 1) - layout.x(c0);
  const lines: ReactNode[] = [];
  for (let c = c0; c <= c1; c++) {
    const x = layout.x(c + 1);
    if (x === layout.x(c)) continue;
    lines.push(<div key={`v${c}`} className="xv-vline" style={{ left: xOrigin + x - 1, top, height }} />);
  }
  for (let r = r0; r <= r1; r++) {
    const y = layout.y(r + 1);
    if (y === layout.y(r)) continue;
    lines.push(<div key={`h${r}`} className="xv-hline" style={{ top: yOrigin + y - 1, left, width }} />);
  }
  return <>{lines}</>;
});

function MergedCells({ sheet, layout, css, merges, matchFlags, r0, r1, c0, c1, xOrigin, yOrigin }: PaneProps) {
  const list = merges.intersecting(r0, r1, c0, c1);
  if (list.length === 0) return null;
  const { kinds, texts, styles } = sheet.cells;

  return (
    <>
      {list.map((m) => {
        const x0 = layout.x(m.c0);
        const x1 = layout.x(m.c1 + 1);
        const y0 = layout.y(m.r0);
        const y1 = layout.y(m.r1 + 1);
        if (x1 <= x0 || y1 <= y0) return null;
        const i = findCellIndex(sheet, m.r0, m.c0);
        const kind = i >= 0 ? kinds[i] : CellKind.Empty;
        const cellCss = (i >= 0 ? css[styles[i]] : undefined) ?? css[0];
        let text = i >= 0 ? stringAt(sheet.strings, texts[i]) : '';
        if (!cellCss.wrap && text.includes('\n')) text = text.replace(/\r?\n/g, ' ');
        const className = `xv-cell xv-merge ${cellCss.alignClass ?? KIND_ALIGN_CLASS[kind]}${cellCss.extraClass}${
          i >= 0 && matchFlags?.[i] ? ' xv-match' : ''
        }`;
        const key = `m${m.r0}:${m.c0}`;
        const border = mergeBorders(sheet, css, m);
        return [
          <div
            key={key}
            className={className}
            style={{ left: xOrigin + x0, top: yOrigin + y0, width: x1 - x0 - 1, height: y1 - y0 - 1, ...cellCss.style }}
          >
            {text}
          </div>,
          border && (
            <div
              key={`${key}b`}
              className="xv-border xv-merge-border"
              style={{ left: xOrigin + x0 - 1, top: yOrigin + y0 - 1, width: x1 - x0 + 1, height: y1 - y0 + 1, ...border }}
            />
          ),
        ];
      })}
    </>
  );
}

/** A merged range's outline comes from the borders of the cells on its edges. */
function mergeBorders(sheet: WorksheetModel, css: readonly CellCss[], m: CellRange): CSSProperties | null {
  const at = (row: number, col: number) => {
    const i = findCellIndex(sheet, row, col);
    return i >= 0 ? css[sheet.cells.styles[i]]?.borderStyle : null;
  };
  const topLeft = at(m.r0, m.c0);
  const topRight = m.c1 !== m.c0 ? at(m.r0, m.c1) : topLeft;
  const bottomLeft = m.r1 !== m.r0 ? at(m.r1, m.c0) : topLeft;
  const style: CSSProperties = {
    borderTop: topLeft?.borderTop,
    borderLeft: topLeft?.borderLeft,
    borderRight: topRight?.borderRight,
    borderBottom: bottomLeft?.borderBottom,
  };
  return style.borderTop || style.borderLeft || style.borderRight || style.borderBottom ? style : null;
}

function SelectionOverlay({ layout, r0, r1, c0, c1, xOrigin, yOrigin, selection, active, currentMatch, fillPreview, fillHandle, clipboard }: PaneProps) {
  const rect = (range: CellRange) => {
    const x0 = layout.x(range.c0);
    const y0 = layout.y(range.r0);
    return {
      left: xOrigin + x0 - 1,
      top: yOrigin + y0 - 1,
      width: layout.x(range.c1 + 1) - x0 + 1,
      height: layout.y(range.r1 + 1) - y0 + 1,
    };
  };
  const clip = (range: CellRange): CellRange | null => {
    const clipped = {
      r0: Math.max(range.r0, r0),
      r1: Math.min(range.r1, r1),
      c0: Math.max(range.c0, c0),
      c1: Math.min(range.c1, c1),
    };
    return clipped.r0 <= clipped.r1 && clipped.c0 <= clipped.c1 ? clipped : null;
  };
  /** Edges cut off by the pane boundary have no border (the neighboring pane draws them). */
  const edges = (range: CellRange) =>
    `${range.r0 < r0 ? ' no-t' : ''}${range.r1 > r1 ? ' no-b' : ''}${range.c0 < c0 ? ' no-l' : ''}${range.c1 > c1 ? ' no-r' : ''}`;

  const out: ReactNode[] = [];
  const multiple = selection.r0 !== active.r0 || selection.r1 !== active.r1 || selection.c0 !== active.c0 || selection.c1 !== active.c1;
  const selected = multiple ? clip(selection) : null;
  if (selected) out.push(<div key="sel" className={`xv-sel${edges(selection)}`} style={rect(selected)} />);
  if (currentMatch) {
    const match = clip({ r0: currentMatch.row, r1: currentMatch.row, c0: currentMatch.col, c1: currentMatch.col });
    if (match) out.push(<div key="match" className="xv-current-match" style={rect(match)} />);
  }
  const activeClipped = clip(active);
  if (activeClipped) out.push(<div key="active" className="xv-active" style={rect(activeClipped)} />);
  if (clipboard) {
    const copied = clip(clipboard.range);
    if (copied) out.push(<div key="clip" className={`xv-clip${edges(clipboard.range)}`} style={rect(copied)} />);
  }
  if (fillPreview) {
    const whole = {
      r0: Math.min(fillPreview.r0, selection.r0),
      c0: Math.min(fillPreview.c0, selection.c0),
      r1: Math.max(fillPreview.r1, selection.r1),
      c1: Math.max(fillPreview.c1, selection.c1),
    };
    const preview = clip(whole);
    if (preview) out.push(<div key="fill" className={`xv-fill-preview${edges(whole)}`} style={rect(preview)} />);
  }
  // The fill handle sits on the selection's bottom-right corner, in the pane that shows that corner.
  const corner = multiple ? selection : active;
  if (fillHandle && corner.r1 >= r0 && corner.r1 <= r1 && corner.c1 >= c0 && corner.c1 <= c1 && corner.r1 < layout.rows.count) {
    const x = xOrigin + layout.x(corner.c1 + 1);
    const y = yOrigin + layout.y(corner.r1 + 1);
    out.push(<div key="handle" className="xv-fill-handle" data-fill-handle="" style={{ left: x - 4, top: y - 4 }} title="Drag to fill · double-click to fill down" />);
  }
  return <>{out}</>;
}
