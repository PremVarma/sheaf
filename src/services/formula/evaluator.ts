import { UnsupportedFormulaError, type FormulaNode, type RefNode } from './ast';
import { callFunction } from './functions';
import {
  compareValues,
  formulaError,
  isError,
  isMatrix,
  toNumber,
  toScalar,
  toText,
  type Matrix,
  type Scalar,
  type Value,
} from './values';

/** How the evaluator reads the workbook. */
export interface EvaluationContext {
  /** Sheet index for a name, or -1 if there's no such sheet. */
  sheetIndex(name: string): number;
  cell(sheet: number, row: number, col: number): Scalar;
  /** Values of a range, limited to the sheet's used range for whole rows/columns. */
  range(sheet: number, ref: RefNode): Matrix;
}

export interface FormulaPosition {
  sheet: number;
  row: number;
  col: number;
}

function arithmetic(op: string, a: number, b: number): Scalar {
  switch (op) {
    case '+':
      return a + b;
    case '-':
      return a - b;
    case '*':
      return a * b;
    case '/':
      return b === 0 ? formulaError('#DIV/0!') : a / b;
    case '^': {
      const result = a ** b;
      return Number.isFinite(result) ? result : formulaError('#NUM!');
    }
    default:
      return formulaError('#VALUE!');
  }
}

/** Evaluates a parsed formula at a position. Throws UnsupportedFormulaError for things we can't compute. */
export function evaluateFormula(node: FormulaNode, position: FormulaPosition, context: EvaluationContext): Scalar {
  const evaluate = (n: FormulaNode): Value => {
    switch (n.type) {
      case 'number':
        return n.value;
      case 'string':
        return n.value;
      case 'boolean':
        return n.value;
      case 'error':
        return formulaError(n.code);
      case 'name':
        throw new UnsupportedFormulaError(`Name ${n.name}`);
      case 'ref': {
        const sheet = n.sheet === undefined ? position.sheet : context.sheetIndex(n.sheet);
        if (sheet < 0) return formulaError('#REF!');
        if (n.r0 === n.r1 && n.c0 === n.c1) return context.cell(sheet, n.r0, n.c0);
        return context.range(sheet, n);
      }
      case 'unary': {
        const value = toScalar(evaluate(n.arg));
        const number = toNumber(value);
        if (isError(number)) return number;
        return n.op === '-' ? -number : n.op === '%' ? number / 100 : number;
      }
      case 'binary': {
        const left = toScalar(evaluate(n.left));
        const right = toScalar(evaluate(n.right));
        if (isError(left)) return left;
        if (isError(right)) return right;
        if (n.op === '&') {
          const a = toText(left);
          const b = toText(right);
          if (isError(a)) return a;
          if (isError(b)) return b;
          return a + b;
        }
        if (n.op === '=' || n.op === '<>' || n.op === '<' || n.op === '>' || n.op === '<=' || n.op === '>=') {
          const c = compareValues(left, right);
          switch (n.op) {
            case '=':
              return c === 0;
            case '<>':
              return c !== 0;
            case '<':
              return c < 0;
            case '>':
              return c > 0;
            case '<=':
              return c <= 0;
            default:
              return c >= 0;
          }
        }
        const a = toNumber(left);
        const b = toNumber(right);
        if (isError(a)) return a;
        if (isError(b)) return b;
        return arithmetic(n.op, a, b);
      }
      case 'call':
        return callFunction(n.name, n.args, { evaluate, row: position.row, col: position.col });
    }
  };

  const result = evaluate(node);
  // A formula whose result is a range shows its first value (no dynamic-array spill).
  const scalar = isMatrix(result) ? toScalar(result) : result;
  if (typeof scalar === 'number' && !Number.isFinite(scalar)) return formulaError('#NUM!');
  return scalar;
}
