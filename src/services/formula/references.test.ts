import { describe, expect, it } from 'vitest';
import { changeMap, fromBase, IDENTITY, insertAt, mapRange, removeAt, toBase } from '../editing/indexMap';
import {
  invalidateSheetReferences,
  moveReferences,
  quoteSheetName,
  remapFileFormula,
  remapReferences,
  renameSheetReferences,
  translateFormula,
} from './references';

const rows = (kind: 'insert' | 'delete', at: number, count: number) => ({ rows: changeMap(kind, at, count) });
const cols = (kind: 'insert' | 'delete', at: number, count: number) => ({ cols: changeMap(kind, at, count) });

describe('index maps', () => {
  it('tracks inserted and deleted entries', () => {
    let map = insertAt(IDENTITY, 2, 3); // rows 2..4 are new
    expect([0, 1, 2, 4, 5, 9].map((i) => toBase(map, i))).toEqual([0, 1, -1, -1, 2, 6]);
    map = removeAt(map, 1, 2); // delete one old row and one new row
    expect([0, 1, 2, 3].map((i) => toBase(map, i))).toEqual([0, -1, -1, 2]);
    expect(fromBase(map, 1)).toBe(-1);
    expect(fromBase(map, 5)).toBe(6);
    expect(removeAt(insertAt(IDENTITY, 4, 2), 4, 2)).toEqual(IDENTITY);
  });

  it('maps ranges like Excel: inserts inside grow them, deletes shrink them', () => {
    expect(mapRange(changeMap('insert', 5, 2), 1, 9, 100)).toEqual([1, 11]);
    expect(mapRange(changeMap('insert', 1, 2), 1, 9, 100)).toEqual([3, 11]);
    expect(mapRange(changeMap('insert', 10, 2), 1, 9, 100)).toEqual([1, 9]);
    expect(mapRange(changeMap('delete', 1, 2), 1, 9, 100)).toEqual([1, 7]);
    expect(mapRange(changeMap('delete', 1, 20), 1, 9, 100)).toBeNull();
  });
});

describe('formula references', () => {
  it('moves relative references when a formula is copied', () => {
    expect(translateFormula('=A1+$B$2+C$3+$D4', 2, 1)).toBe('=B3+$B$2+D$3+$D6');
    expect(translateFormula('=SUM(A1:A10)*A:A', 1, 1)).toBe('=SUM(B2:B11)*B:B');
    expect(translateFormula('=Sheet2!A1&"A1"', 1, 0)).toBe('=Sheet2!A2&"A1"');
    expect(translateFormula('=A1', -1, 0)).toBe('=#REF!');
  });

  it('follows inserted and deleted rows and columns', () => {
    expect(remapReferences('=SUM(B2:B10)+B12', 'S', 'S', rows('insert', 4, 2))).toBe('=SUM(B2:B12)+B14');
    expect(remapReferences('=SUM(B2:B10)+B1', 'S', 'S', rows('delete', 2, 3))).toBe('=SUM(B2:B7)+B1');
    expect(remapReferences('=B3*2', 'S', 'S', rows('delete', 2, 1))).toBe('=#REF!*2');
    expect(remapReferences('=C1+$D$1+A:A+D:E', 'S', 'S', cols('insert', 1, 1))).toBe('=D1+$E$1+A:A+E:F');
    expect(remapReferences("='My S'!B2+B2", 'Other', 'My S', rows('insert', 0, 1))).toBe("='My S'!B3+B2");
    expect(remapReferences('=A1', 'Other', 'S', rows('insert', 0, 1))).toBe('=A1');
  });

  it('renames, invalidates and moves references', () => {
    expect(renameSheetReferences("=Data!A1+'data'!B2+A1", 'data', 'Q1 Sales')).toBe("='Q1 Sales'!A1+'Q1 Sales'!B2+A1");
    expect(invalidateSheetReferences('=Data!A1+1', 'Data')).toBe('=#REF!+1');
    const moved = moveReferences('=A1+B5+SUM(A1:A2)', 'S', { sheet: 'S', range: { r0: 0, c0: 0, r1: 1, c1: 0 } }, { sheet: 'S', dRow: 5, dCol: 2 });
    expect(moved).toBe('=C6+B5+SUM(C6:C7)');
    const across = moveReferences('=A1', 'S', { sheet: 'S', range: { r0: 0, c0: 0, r1: 0, c1: 0 } }, { sheet: 'T', dRow: 0, dCol: 0 });
    expect(across).toBe('=T!A1');
  });

  it('remaps formulas stored in the file (names, charts)', () => {
    const changes = {
      currentName: (name: string) => (name === 'Gone' ? null : name === 'Old' ? 'New Name' : name === 'Sales' ? 'Sales' : undefined),
      axes: (name: string) => (name === 'Sales' ? rows('insert', 0, 1) : null),
    };
    expect(remapFileFormula('Sales!$A$1:$B$5', null, changes)).toBe('Sales!$A$2:$B$6');
    expect(remapFileFormula('Old!A1', null, changes)).toBe("'New Name'!A1");
    expect(remapFileFormula('Gone!A1+1', null, changes)).toBe('#REF!+1');
    expect(remapFileFormula('$A$1', 'Sales', changes)).toBe('$A$2');
  });

  it('quotes sheet names only when needed', () => {
    expect(quoteSheetName('Sheet1')).toBe('Sheet1');
    expect(quoteSheetName('Q1 2024')).toBe("'Q1 2024'");
    expect(quoteSheetName('AB12')).toBe("'AB12'");
    expect(quoteSheetName("Bob's")).toBe("'Bob''s'");
  });
});
