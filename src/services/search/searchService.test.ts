import { describe, expect, it } from 'vitest';
import { createSheet, workbookOf } from '../../test/fixtures';
import { firstMatchFrom, searchWorkbook, stepMatch } from './searchService';

const cities = [
  ['City', 'Region'],
  ['Mumbai', 'West'],
  ['Delhi', 'North'],
  ['Navi Mumbai', 'West'],
];
const first = createSheet(4, 2, (r, c) => cities[r][c], {}, 'Cities');
const second = createSheet(3, 1, (r) => ['Offices', 'mumbai', 'Pune'][r], {}, 'Offices');
const workbook = workbookOf(first, second);
const options = { scope: 'sheet' as const, sheetIndex: 0, matchCase: false, wholeCell: false };

describe('searchWorkbook', () => {
  it('searches the current sheet, case-insensitively', () => {
    const results = searchWorkbook(workbook, 'mumbai', options);
    expect(results.total).toBe(2);
    expect(results.matches).toEqual([
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 3, col: 0 },
    ]);
  });

  it('searches the whole workbook in sheet order', () => {
    const results = searchWorkbook(workbook, 'Mumbai', { ...options, scope: 'workbook' });
    expect(results.matches.map((m) => m.sheet)).toEqual([0, 0, 1]);
    expect(results.flags.get(1)?.some(Boolean)).toBe(true);
  });

  it('supports match case and whole-cell matching', () => {
    expect(searchWorkbook(workbook, 'mumbai', { ...options, scope: 'workbook', matchCase: true }).total).toBe(1);
    expect(searchWorkbook(workbook, 'mumbai', { ...options, wholeCell: true }).matches).toEqual([{ sheet: 0, row: 1, col: 0 }]);
  });

  it('flags matching cells for highlighting and ignores empty queries', () => {
    const results = searchWorkbook(workbook, 'west', options);
    expect(Array.from(results.flags.get(0)!).filter(Boolean)).toHaveLength(2);
    expect(searchWorkbook(workbook, '', options).total).toBe(0);
    expect(searchWorkbook(workbook, 'Chennai', options).total).toBe(0);
  });

  it('skips cells in hidden rows', () => {
    const hidden = createSheet(4, 2, (r, c) => cities[r][c], { rowHeights: new Map([[1, 0]]) });
    expect(searchWorkbook(workbookOf(hidden), 'mumbai', options).matches).toEqual([{ sheet: 0, row: 3, col: 0 }]);
  });
});

describe('match navigation', () => {
  const matches = [
    { sheet: 0, row: 1, col: 0 },
    { sheet: 0, row: 3, col: 0 },
    { sheet: 1, row: 1, col: 0 },
  ];

  it('starts from the active cell and wraps around', () => {
    expect(firstMatchFrom(matches, 0, 0, 0)).toBe(0);
    expect(firstMatchFrom(matches, 0, 2, 0)).toBe(1);
    expect(firstMatchFrom(matches, 0, 3, 0)).toBe(1);
    expect(firstMatchFrom(matches, 1, 5, 0)).toBe(0);
  });

  it('steps forward and backward with wrap-around', () => {
    expect(stepMatch(-1, 3, 1)).toBe(0);
    expect(stepMatch(-1, 3, -1)).toBe(2);
    expect(stepMatch(2, 3, 1)).toBe(0);
    expect(stepMatch(0, 3, -1)).toBe(2);
    expect(stepMatch(0, 0, 1)).toBe(-1);
  });
});
