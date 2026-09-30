import { stringAt, stringCount } from '../../models/stringPool';
import { CellKind, type WorkbookModel, type WorksheetModel } from '../../models/workbook';

export type SearchScope = 'sheet' | 'workbook';

export interface SearchOptions {
  scope: SearchScope;
  /** Sheet searched when scope is 'sheet'. */
  sheetIndex: number;
  matchCase: boolean;
  wholeCell: boolean;
}

export interface SearchMatch {
  sheet: number;
  row: number;
  col: number;
}

export interface SearchResults {
  query: string;
  /** Matches in workbook order (sheet, row, column); capped at MAX_SEARCH_MATCHES. */
  matches: SearchMatch[];
  /** Total number of matches, which can exceed `matches.length`. */
  total: number;
  /** Per sheet: 1 for every stored cell that matches, for highlighting. */
  flags: Map<number, Uint8Array>;
}

export const MAX_SEARCH_MATCHES = 100_000;

function hiddenIndexes(sizes: WorksheetModel['rowHeights']): Set<number> {
  const hidden = new Set<number>();
  for (let k = 0; k < sizes.index.length; k++) if (sizes.size[k] === 0) hidden.add(sizes.index[k]);
  return hidden;
}

/**
 * Searches displayed cell text. Each distinct string is tested once, then the
 * cells are scanned with a lookup table, so large sheets search in milliseconds.
 */
export function searchWorkbook(workbook: WorkbookModel, query: string, options: SearchOptions): SearchResults {
  const results: SearchResults = { query, matches: [], total: 0, flags: new Map() };
  if (!query) return results;
  const needle = options.matchCase ? query : query.toLocaleLowerCase();
  const sheets =
    options.scope === 'sheet' ? [workbook.sheets[options.sheetIndex]].filter(Boolean) : workbook.sheets;

  for (const sheet of sheets) {
    const { strings } = sheet;
    const count = stringCount(strings);
    const matching = new Uint8Array(count);
    let any = false;
    for (let s = 1; s < count; s++) {
      const value = stringAt(strings, s);
      const text = options.matchCase ? value : value.toLocaleLowerCase();
      if (options.wholeCell ? text === needle : text.includes(needle)) {
        matching[s] = 1;
        any = true;
      }
    }
    if (!any) continue;

    const { rowStart, cols, kinds, texts } = sheet.cells;
    const hiddenRows = hiddenIndexes(sheet.rowHeights);
    const hiddenCols = hiddenIndexes(sheet.columnWidths);
    const flags = new Uint8Array(sheet.cellCount);
    for (let r = 0; r < sheet.rowCount; r++) {
      const end = rowStart[r + 1];
      if (rowStart[r] === end || hiddenRows.has(r)) continue;
      for (let i = rowStart[r]; i < end; i++) {
        if (!matching[texts[i]] || kinds[i] === CellKind.Empty || hiddenCols.has(cols[i])) continue;
        flags[i] = 1;
        results.total++;
        if (results.matches.length < MAX_SEARCH_MATCHES) results.matches.push({ sheet: sheet.index, row: r, col: cols[i] });
      }
    }
    results.flags.set(sheet.index, flags);
  }
  return results;
}

function compare(a: SearchMatch, sheet: number, row: number, col: number): number {
  return a.sheet - sheet || a.row - row || a.col - col;
}

/** Index of the first match at or after a position (wrapping to the start). */
export function firstMatchFrom(matches: SearchMatch[], sheet: number, row: number, col: number): number {
  if (matches.length === 0) return -1;
  let lo = 0;
  let hi = matches.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (compare(matches[mid], sheet, row, col) < 0) lo = mid + 1;
    else hi = mid;
  }
  return lo < matches.length ? lo : 0;
}

/** Next/previous match index with wrap-around. */
export function stepMatch(current: number, count: number, direction: 1 | -1): number {
  if (count === 0) return -1;
  if (current < 0) return direction === 1 ? 0 : count - 1;
  return (current + direction + count) % count;
}
