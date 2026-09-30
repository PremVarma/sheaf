import type { CellRange } from '../../models/workbook';

export type ColumnSpan = readonly [number, number];

const NONE: readonly ColumnSpan[] = Object.freeze([]);
const MAX_CACHED_ROWS = 20_000;

/** Fast "which columns of this row are covered by merged cells" lookups for rendering. */
export class MergeIndex {
  readonly merges: readonly CellRange[];
  private readonly byRow = new Map<number, readonly ColumnSpan[]>();

  constructor(merges: readonly CellRange[]) {
    this.merges = merges;
  }

  /** Column spans covered by merges in a row. Returns the same array for the same row. */
  covered(row: number): readonly ColumnSpan[] {
    if (this.merges.length === 0) return NONE;
    let spans = this.byRow.get(row);
    if (!spans) {
      const list: ColumnSpan[] = [];
      for (const m of this.merges) if (row >= m.r0 && row <= m.r1) list.push([m.c0, m.c1]);
      spans = list.length > 0 ? list : NONE;
      if (this.byRow.size >= MAX_CACHED_ROWS) this.byRow.clear();
      this.byRow.set(row, spans);
    }
    return spans;
  }

  intersecting(r0: number, r1: number, c0: number, c1: number): CellRange[] {
    if (this.merges.length === 0 || r0 > r1 || c0 > c1) return [];
    return this.merges.filter((m) => m.r0 <= r1 && m.r1 >= r0 && m.c0 <= c1 && m.c1 >= c0);
  }

  /** The merge containing a cell, if any. */
  at(row: number, col: number): CellRange | undefined {
    for (const m of this.covered(row)) {
      if (col >= m[0] && col <= m[1]) {
        return this.merges.find((range) => row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1);
      }
    }
    return undefined;
  }
}

export function isCovered(spans: readonly ColumnSpan[], col: number): boolean {
  for (let k = 0; k < spans.length; k++) if (col >= spans[k][0] && col <= spans[k][1]) return true;
  return false;
}
