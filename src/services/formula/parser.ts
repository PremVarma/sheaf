import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS } from '../../models/workbook';
import { columnIndex } from '../../utils/cellAddress';
import {
  ERROR_CODES,
  FormulaSyntaxError,
  UnsupportedFormulaError,
  type BinaryOperator,
  type ErrorCode,
  type FormulaNode,
  type RefEndpoint,
  type ReferenceToken,
  type RefNode,
} from './ast';

/**
 * Parser for Excel formulas (without the leading "="). Precedence, lowest first:
 * comparison, &, + -, * /, ^, %, unary - +, then references, literals, calls.
 */

type Token =
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'error'; value: ErrorCode }
  | { kind: 'ref'; node: RefNode; ref: ReferenceToken }
  | { kind: 'ident'; value: string }
  | { kind: 'op'; value: string }
  | { kind: 'lparen' }
  | { kind: 'rparen' }
  | { kind: 'comma' };

const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
const COLUMN_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/;
const ROW_RANGE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\w.])/;
const NUMBER = /^(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/;
const IDENT = /^[A-Za-z_\\][\w.]*/;

function endpointFrom(match: RegExpExecArray): RefEndpoint | null {
  const col = columnIndex(match[2]);
  const row = Number(match[4]) - 1;
  return col >= 0 && row >= 0 && row < EXCEL_MAX_ROWS ? { row, col, rowAbs: match[3] === '$', colAbs: match[1] === '$' } : null;
}

interface ScannedReference {
  kind: ReferenceToken['kind'];
  from: RefEndpoint;
  to: RefEndpoint;
  length: number;
}

/** Tries to read a reference (A1, $A$1:B2, A:C, 3:5) at the start of `text`. */
function readReference(text: string): ScannedReference | null {
  const cols = COLUMN_RANGE.exec(text);
  if (cols) {
    const a = columnIndex(cols[2]);
    const b = columnIndex(cols[4]);
    if (a >= 0 && b >= 0) {
      return {
        kind: 'columns',
        from: { row: 0, col: a, rowAbs: false, colAbs: cols[1] === '$' },
        to: { row: EXCEL_MAX_ROWS - 1, col: b, rowAbs: false, colAbs: cols[3] === '$' },
        length: cols[0].length,
      };
    }
  }
  const rows = ROW_RANGE.exec(text);
  if (rows) {
    const a = Number(rows[2]) - 1;
    const b = Number(rows[4]) - 1;
    if (a >= 0 && b >= 0 && a < EXCEL_MAX_ROWS && b < EXCEL_MAX_ROWS) {
      return {
        kind: 'rows',
        from: { row: a, col: 0, rowAbs: rows[1] === '$', colAbs: false },
        to: { row: b, col: EXCEL_MAX_COLS - 1, rowAbs: rows[3] === '$', colAbs: false },
        length: rows[0].length,
      };
    }
  }
  const first = CELL.exec(text);
  if (!first) return null;
  // Not a reference if it continues as a name or call, e.g. LOG10( or ABC1D.
  if (/[\w(]/.test(text.charAt(first[0].length))) return null;
  const start = endpointFrom(first);
  if (!start) return null;
  let length = first[0].length;
  if (text.charAt(length) === ':') {
    const second = CELL.exec(text.slice(length + 1));
    if (second && !/[\w(]/.test(text.charAt(length + 1 + second[0].length))) {
      const end = endpointFrom(second);
      if (end) return { kind: 'area', from: start, to: end, length: length + 1 + second[0].length };
    }
  }
  return { kind: 'cell', from: start, to: start, length };
}

function refToken(scanned: ScannedReference, sheet: string | undefined, start: number, cellStart: number): Extract<Token, { kind: 'ref' }> {
  const { from, to } = scanned;
  return {
    kind: 'ref',
    node: {
      type: 'ref',
      sheet,
      r0: Math.min(from.row, to.row),
      r1: Math.max(from.row, to.row),
      c0: Math.min(from.col, to.col),
      c1: Math.max(from.col, to.col),
    },
    ref: { sheet, kind: scanned.kind, from, to, start, cellStart, end: cellStart + scanned.length },
  };
}

/**
 * Splits a formula into tokens. `lenient` (used to find references for rewriting)
 * skips syntax the evaluator doesn't support instead of failing.
 */
function tokenize(formula: string, lenient = false): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = formula.length;
  const unsupported = (detail: string) => {
    if (!lenient) throw new UnsupportedFormulaError(detail);
  };
  while (i < n) {
    const ch = formula[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      i++;
      continue;
    }
    const rest = formula.slice(i);

    if (ch === '"') {
      let value = '';
      let j = i + 1;
      for (;;) {
        if (j >= n) {
          if (lenient) return tokens;
          throw new FormulaSyntaxError('Unterminated string');
        }
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') {
            value += '"';
            j += 2;
            continue;
          }
          break;
        }
        value += formula[j++];
      }
      tokens.push({ kind: 'string', value });
      i = j + 1;
      continue;
    }

    if (ch === '#') {
      const code = ERROR_CODES.find((c) => rest.toUpperCase().startsWith(c));
      if (!code) {
        if (!lenient) throw new FormulaSyntaxError(`Unknown error literal at ${i}`);
        i++;
        continue;
      }
      tokens.push({ kind: 'error', value: code });
      i += code.length;
      continue;
    }

    // Sheet-qualified reference: 'My Sheet'!A1 or Sheet1!A1:B2
    if (ch === "'") {
      let j = i + 1;
      let name = '';
      for (;;) {
        if (j >= n) {
          if (lenient) return tokens;
          throw new FormulaSyntaxError('Unterminated sheet name');
        }
        if (formula[j] === "'") {
          if (formula[j + 1] === "'") {
            name += "'";
            j += 2;
            continue;
          }
          break;
        }
        name += formula[j++];
      }
      const scanned = formula[j + 1] === '!' ? readReference(formula.slice(j + 2)) : null;
      if (!scanned) {
        if (!lenient) throw new FormulaSyntaxError(formula[j + 1] === '!' ? 'Expected a reference after sheet name' : 'Expected ! after sheet name');
        i = j + 1;
        continue;
      }
      if (name.includes(':')) unsupported('3-D references');
      else tokens.push(refToken(scanned, name, i, j + 2));
      i = j + 2 + scanned.length;
      continue;
    }
    const qualified = /^([A-Za-z_À-￿][\w.À-￿]*)!/.exec(rest);
    if (qualified) {
      const scanned = readReference(rest.slice(qualified[0].length));
      if (!scanned) {
        if (!lenient) throw new FormulaSyntaxError('Expected a reference after sheet name');
        i += qualified[0].length;
        continue;
      }
      tokens.push(refToken(scanned, qualified[1], i, i + qualified[0].length));
      i += qualified[0].length + scanned.length;
      continue;
    }

    if (/[A-Za-z$0-9]/.test(ch)) {
      const scanned = readReference(rest);
      if (scanned) {
        tokens.push(refToken(scanned, undefined, i, i));
        i += scanned.length;
        continue;
      }
    }

    const number = NUMBER.exec(rest);
    if (number && /[\d.]/.test(ch)) {
      tokens.push({ kind: 'number', value: Number(number[0]) });
      i += number[0].length;
      continue;
    }

    const ident = IDENT.exec(rest);
    if (ident) {
      tokens.push({ kind: 'ident', value: ident[0] });
      i += ident[0].length;
      continue;
    }

    const two = rest.slice(0, 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      tokens.push({ kind: 'op', value: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%'.includes(ch)) {
      tokens.push({ kind: 'op', value: ch });
      i++;
      continue;
    }
    if (ch === '(') tokens.push({ kind: 'lparen' });
    else if (ch === ')') tokens.push({ kind: 'rparen' });
    else if (ch === ',') tokens.push({ kind: 'comma' });
    else if (lenient && (ch === '[' || ch === '{')) {
      // Structured references and array constants hold no cell references to rewrite.
      const close = ch === '[' ? ']' : '}';
      let depth = 0;
      for (; i < n; i++) {
        if (formula[i] === ch) depth++;
        else if (formula[i] === close && --depth === 0) break;
      }
    } else if (ch === '{' || ch === '}' || ch === ';' || ch === '[' || ch === '@') {
      unsupported(`Syntax "${ch}" is not supported`);
    } else if (!lenient) throw new FormulaSyntaxError(`Unexpected "${ch}" at ${i}`);
    i++;
  }
  return tokens;
}

/** Where the expression starts: after "=" or "{=" (array formulas). */
function bodyStart(formula: string): number {
  return formula.startsWith('{=') ? 2 : formula.startsWith('=') ? 1 : 0;
}

/** Every cell reference in a formula (with or without "="), with offsets into `formula`. */
export function scanReferences(formula: string): ReferenceToken[] {
  const offset = bodyStart(formula);
  const body = formula.endsWith('}') && offset === 2 ? formula.slice(offset, -1) : formula.slice(offset);
  const refs: ReferenceToken[] = [];
  for (const token of tokenize(body, true)) {
    if (token.kind !== 'ref') continue;
    const { ref } = token;
    refs.push({ ...ref, start: ref.start + offset, cellStart: ref.cellStart + offset, end: ref.end + offset });
  }
  return refs;
}

const BINARY_PRECEDENCE: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

class Parser {
  private pos = 0;
  private readonly tokens: Token[];

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  parse(): FormulaNode {
    const node = this.expression(0);
    if (this.pos < this.tokens.length) throw new FormulaSyntaxError('Unexpected input after formula');
    return node;
  }

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private expression(minPrecedence: number): FormulaNode {
    let left = this.postfix();
    for (;;) {
      const token = this.peek();
      if (!token || token.kind !== 'op') break;
      const precedence = BINARY_PRECEDENCE[token.value];
      if (precedence === undefined || precedence <= minPrecedence - 1 || precedence < minPrecedence) break;
      this.pos++;
      // All Excel binary operators, including ^, are left-associative.
      const right = this.expression(precedence + 1);
      left = { type: 'binary', op: token.value as BinaryOperator, left, right };
    }
    return left;
  }

  private postfix(): FormulaNode {
    let node = this.unary();
    while (this.peek()?.kind === 'op' && (this.peek() as { value: string }).value === '%') {
      this.pos++;
      node = { type: 'unary', op: '%', arg: node };
    }
    return node;
  }

  private unary(): FormulaNode {
    const token = this.peek();
    if (token?.kind === 'op' && (token.value === '-' || token.value === '+')) {
      this.pos++;
      // In Excel, negation binds tighter than ^ (-2^2 = 4).
      return { type: 'unary', op: token.value, arg: this.unary() };
    }
    return this.primary();
  }

  private primary(): FormulaNode {
    const token = this.tokens[this.pos++];
    if (!token) throw new FormulaSyntaxError('Unexpected end of formula');
    switch (token.kind) {
      case 'number':
        return { type: 'number', value: token.value };
      case 'string':
        return { type: 'string', value: token.value };
      case 'error':
        return { type: 'error', code: token.value };
      case 'ref':
        return token.node;
      case 'lparen': {
        const inner = this.expression(0);
        if (this.tokens[this.pos++]?.kind !== 'rparen') throw new FormulaSyntaxError('Missing )');
        return inner;
      }
      case 'ident': {
        const name = token.value.toUpperCase().replace(/^_XLFN\.|^_XLWS\./, '');
        if (this.peek()?.kind === 'lparen') {
          this.pos++;
          const args: FormulaNode[] = [];
          if (this.peek()?.kind === 'rparen') {
            this.pos++;
            return { type: 'call', name, args };
          }
          for (;;) {
            // Omitted arguments (e.g. IF(A1,,1)) evaluate as empty.
            const next = this.peek();
            if (next?.kind === 'comma' || next?.kind === 'rparen') args.push({ type: 'string', value: '' });
            else args.push(this.expression(0));
            const separator = this.tokens[this.pos++];
            if (separator?.kind === 'rparen') break;
            if (separator?.kind !== 'comma') throw new FormulaSyntaxError('Expected , or ) in function call');
          }
          return { type: 'call', name, args };
        }
        if (name === 'TRUE' || name === 'FALSE') return { type: 'boolean', value: name === 'TRUE' };
        return { type: 'name', name: token.value };
      }
      default:
        throw new FormulaSyntaxError('Unexpected token');
    }
  }
}

/** Parses a formula; accepts it with or without the leading "=". */
export function parseFormula(formula: string): FormulaNode {
  const text = formula.startsWith('=') ? formula.slice(1) : formula;
  if (text.startsWith('{')) throw new UnsupportedFormulaError('Array formulas');
  if (!text.trim()) throw new FormulaSyntaxError('Empty formula');
  return new Parser(tokenize(text)).parse();
}
