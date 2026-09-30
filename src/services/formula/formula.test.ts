import { describe, expect, it } from 'vitest';
import { EXCEL_MAX_ROWS } from '../../models/workbook';
import { UnsupportedFormulaError } from './ast';
import { evaluateFormula, type EvaluationContext } from './evaluator';
import { parseFormula } from './parser';
import { isError, type Scalar } from './values';

/** A tiny in-memory sheet: A1-style address → value. */
function context(cells: Record<string, Scalar>, other: Record<string, Scalar> = {}): EvaluationContext {
  const sheets = [cells, other];
  const addr = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`;
  return {
    sheetIndex: (name) => (name.toLowerCase() === 'data' ? 1 : name.toLowerCase() === 'main' ? 0 : -1),
    cell: (sheet, row, col) => sheets[sheet][addr(row, col)] ?? null,
    range: (sheet, ref) => {
      const rows: Scalar[][] = [];
      const lastRow = Math.min(ref.r1, 20);
      for (let r = ref.r0; r <= lastRow; r++) {
        const row: Scalar[] = [];
        for (let c = ref.c0; c <= Math.min(ref.c1, 10); c++) row.push(sheets[sheet][addr(r, c)] ?? null);
        rows.push(row);
      }
      return rows;
    },
  };
}

const run = (formula: string, cells: Record<string, Scalar> = {}, other: Record<string, Scalar> = {}) =>
  evaluateFormula(parseFormula(formula), { sheet: 0, row: 50, col: 0 }, context(cells, other));

describe('formula parser', () => {
  it('parses references, ranges and sheet names', () => {
    expect(parseFormula('=A1')).toEqual({ type: 'ref', sheet: undefined, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(parseFormula('$B$2:C10')).toMatchObject({ type: 'ref', r0: 1, c0: 1, r1: 9, c1: 2 });
    expect(parseFormula("'My Sheet'!A1")).toMatchObject({ type: 'ref', sheet: 'My Sheet' });
    expect(parseFormula('Data!C:C')).toMatchObject({ type: 'ref', sheet: 'Data', r0: 0, r1: EXCEL_MAX_ROWS - 1, c0: 2, c1: 2 });
    expect(parseFormula('SUM(1:3)')).toMatchObject({ type: 'call', name: 'SUM', args: [{ type: 'ref', r0: 0, r1: 2 }] });
  });

  it('distinguishes functions from cell-like names', () => {
    expect(parseFormula('LOG10(100)')).toMatchObject({ type: 'call', name: 'LOG10' });
    expect(parseFormula('_xlfn.CONCAT("a","b")')).toMatchObject({ type: 'call', name: 'CONCAT' });
  });

  it('rejects array formulas and structured references as unsupported', () => {
    expect(() => parseFormula('={1,2,3}')).toThrow(UnsupportedFormulaError);
    expect(() => parseFormula('=Table1[Amount]')).toThrow(UnsupportedFormulaError);
  });
});

describe('formula evaluation', () => {
  it('follows Excel precedence and coercion', () => {
    expect(run('=1+2*3')).toBe(7);
    expect(run('=-2^2')).toBe(4);
    expect(run('=2^3^2')).toBe(64);
    expect(run('=10%')).toBe(0.1);
    expect(run('="5"+1')).toBe(6);
    expect(run('=TRUE+1')).toBe(2);
    expect(run('=A1+1')).toBe(1); // empty cells are 0
    expect(run('="Total: "&A1', { A1: 42 })).toBe('Total: 42');
    expect(run('="abc"="ABC"')).toBe(true);
    expect(run('=2<"1"')).toBe(true); // numbers sort before text
  });

  it('returns Excel errors', () => {
    expect((run('=1/0') as Error).message).toBe('#DIV/0!');
    expect((run('="x"*2') as Error).message).toBe('#VALUE!');
    expect((run('=Nope!A1') as Error).message).toBe('#REF!');
    expect(isError(run('=A1+1', { A1: new Error('#N/A') }))).toBe(true);
  });

  it('evaluates common functions over ranges', () => {
    const cells = { A1: 10, A2: 20, A3: 'x', A4: null, B1: 'Mumbai', B2: 'Delhi', B3: 'Mumbai' };
    expect(run('=SUM(A1:A4)', cells)).toBe(30);
    expect(run('=AVERAGE(A1:A2)', cells)).toBe(15);
    expect(run('=COUNT(A1:A4)', cells)).toBe(2);
    expect(run('=COUNTIF(B1:B3,"Mumbai")', cells)).toBe(2);
    expect(run('=SUMIF(B1:B3,"Mumbai",A1:A3)', cells)).toBe(10);
    expect(run('=ROUND(2.345,2)')).toBe(2.35);
    expect(run('=VLOOKUP("Delhi",B1:B3,1,FALSE)', cells)).toBe('Delhi');
    expect(run('=INDEX(A1:A2,2)', cells)).toBe(20);
    expect(run('=MATCH("Delhi",B1:B3,0)', cells)).toBe(2);
    expect(run('=UPPER(B2)&LEN(B1)', cells)).toBe('DELHI6');
    expect(run('=MAX(A1:A2)-MIN(A1:A2)', cells)).toBe(10);
  });

  it('evaluates IF, IFERROR, AND and OR lazily', () => {
    expect(run('=IF(1>0,"yes",1/0)')).toBe('yes');
    expect(run('=IFERROR(1/0,"n/a")')).toBe('n/a');
    expect(run('=AND(TRUE,1,OR(FALSE,0))')).toBe(false);
    expect(run('=IF(A1,,5)')).toBe(5);
  });

  it('works with dates as serial numbers', () => {
    expect(run('=DATE(2024,3,15)')).toBe(45366);
    expect(run('=YEAR(45366)&"-"&MONTH(45366)&"-"&DAY(45366)')).toBe('2024-3-15');
    expect(run('=EOMONTH(45366,0)')).toBe(45382);
    expect(run('=EDATE(DATE(2024,1,31),1)')).toBe(45351); // 29 Feb 2024
    expect(run('=TEXT(45366,"dd-mmm-yyyy")')).toBe('15-Mar-2024');
    expect(run('=TEXT(1234.5,"#,##0.00")')).toBe('1,234.50');
  });

  it('reads other sheets', () => {
    expect(run('=Data!A1*2', {}, { A1: 21 })).toBe(42);
  });

  it('reports functions it cannot evaluate', () => {
    expect(() => run('=WEBSERVICE("http://x")')).toThrow(UnsupportedFormulaError);
    expect(() => run('=TaxRate*2')).toThrow(UnsupportedFormulaError);
  });
});
