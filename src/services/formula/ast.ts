/** Excel formula syntax tree. */

export type ErrorCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA';

export const ERROR_CODES: readonly ErrorCode[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];

/** A cell or range reference. Whole columns/rows extend to Excel's limits. */
export interface RefNode {
  type: 'ref';
  /** Sheet name when qualified (Sheet2!A1). */
  sheet?: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** One end of a reference as written: `$` makes a part absolute. */
export interface RefEndpoint {
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

/** A reference found in formula text, with its position (for rewriting it). */
export interface ReferenceToken {
  /** Sheet name when qualified (unquoted). */
  sheet?: string;
  /** A1, A1:B2, A:C (whole columns) or 3:5 (whole rows). */
  kind: 'cell' | 'area' | 'columns' | 'rows';
  from: RefEndpoint;
  to: RefEndpoint;
  /** Offsets in the formula text: the whole reference, and where the part after "Sheet!" starts. */
  start: number;
  cellStart: number;
  end: number;
}

export type FormulaNode =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'error'; code: ErrorCode }
  | RefNode
  | { type: 'name'; name: string }
  | { type: 'unary'; op: '-' | '+' | '%'; arg: FormulaNode }
  | { type: 'binary'; op: BinaryOperator; left: FormulaNode; right: FormulaNode }
  | { type: 'call'; name: string; args: FormulaNode[] };

export type BinaryOperator = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=';

/** Syntax we recognise but don't evaluate (array constants, structured refs, 3-D refs…). */
export class UnsupportedFormulaError extends Error {
  constructor(detail: string) {
    super(detail);
    this.name = 'UnsupportedFormulaError';
  }
}

export class FormulaSyntaxError extends Error {
  constructor(detail: string) {
    super(detail);
    this.name = 'FormulaSyntaxError';
  }
}

/** Every reference in a formula, for dependency tracking. */
export function collectRefs(node: FormulaNode, out: RefNode[] = []): RefNode[] {
  switch (node.type) {
    case 'ref':
      out.push(node);
      break;
    case 'unary':
      collectRefs(node.arg, out);
      break;
    case 'binary':
      collectRefs(node.left, out);
      collectRefs(node.right, out);
      break;
    case 'call':
      for (const arg of node.args) collectRefs(arg, out);
      break;
    default:
      break;
  }
  return out;
}
