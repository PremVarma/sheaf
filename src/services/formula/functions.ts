import * as formulajs from '@formulajs/formulajs';
import { formatNumber } from '../workbook/numberFormat';
import { UnsupportedFormulaError, type FormulaNode } from './ast';
import {
  compareValues,
  dateToSerial,
  formulaError,
  isError,
  isMatrix,
  serialToDate,
  toBoolean,
  toNumber,
  toScalar,
  toText,
  type Matrix,
  type Scalar,
  type Value,
} from './values';

/** Evaluates an argument node. Functions with lazy semantics call this themselves. */
export type Evaluate = (node: FormulaNode) => Value;

interface CallContext {
  evaluate: Evaluate;
  /** Row/column of the cell being calculated (for ROW()/COLUMN() without arguments). */
  row: number;
  col: number;
}

type NativeFunction = (args: FormulaNode[], ctx: CallContext) => Value;

const num = (value: Value): number | Error => toNumber(toScalar(value));

function flatten(value: Value): Scalar[] {
  return isMatrix(value) ? value.flat() : [value];
}

function dateParts(value: Value): Date | Error {
  const serial = num(value);
  return isError(serial) ? serial : serialToDate(serial);
}

/** Functions implemented here: lazy logic, references, and dates as Excel serial numbers. */
const NATIVE: Record<string, NativeFunction> = {
  IF: (args, { evaluate }) => {
    const condition = toBoolean(toScalar(evaluate(args[0])));
    if (isError(condition)) return condition;
    if (condition) return args.length > 1 ? evaluate(args[1]) : true;
    return args.length > 2 ? evaluate(args[2]) : false;
  },
  IFERROR: (args, { evaluate }) => {
    const value = evaluate(args[0]);
    return isError(toScalar(value)) ? evaluate(args[1]) : value;
  },
  IFNA: (args, { evaluate }) => {
    const value = evaluate(args[0]);
    const scalar = toScalar(value);
    return isError(scalar) && scalar.message === '#N/A' ? evaluate(args[1]) : value;
  },
  AND: (args, { evaluate }) => {
    for (const arg of args) {
      for (const item of flatten(evaluate(arg))) {
        if (item === null || typeof item === 'string') continue;
        const bool = toBoolean(item);
        if (isError(bool)) return bool;
        if (!bool) return false;
      }
    }
    return true;
  },
  OR: (args, { evaluate }) => {
    let result = false;
    for (const arg of args) {
      for (const item of flatten(evaluate(arg))) {
        if (item === null || typeof item === 'string') continue;
        const bool = toBoolean(item);
        if (isError(bool)) return bool;
        result ||= bool;
      }
    }
    return result;
  },
  NOT: (args, { evaluate }) => {
    const bool = toBoolean(toScalar(evaluate(args[0])));
    return isError(bool) ? bool : !bool;
  },
  CHOOSE: (args, { evaluate }) => {
    const index = num(evaluate(args[0]));
    if (isError(index)) return index;
    const pick = Math.trunc(index);
    return pick >= 1 && pick < args.length ? evaluate(args[pick]) : formulaError('#VALUE!');
  },
  ISBLANK: (args, { evaluate }) => toScalar(evaluate(args[0])) === null,
  ISERROR: (args, { evaluate }) => isError(toScalar(evaluate(args[0]))),
  ISERR: (args, { evaluate }) => {
    const value = toScalar(evaluate(args[0]));
    return isError(value) && value.message !== '#N/A';
  },
  ISNA: (args, { evaluate }) => {
    const value = toScalar(evaluate(args[0]));
    return isError(value) && value.message === '#N/A';
  },
  ISNUMBER: (args, { evaluate }) => typeof toScalar(evaluate(args[0])) === 'number',
  ISTEXT: (args, { evaluate }) => typeof toScalar(evaluate(args[0])) === 'string',
  ISLOGICAL: (args, { evaluate }) => typeof toScalar(evaluate(args[0])) === 'boolean',
  ROW: (args, { row }) => (args.length === 0 ? row + 1 : args[0].type === 'ref' ? args[0].r0 + 1 : formulaError('#VALUE!')),
  COLUMN: (args, { col }) => (args.length === 0 ? col + 1 : args[0].type === 'ref' ? args[0].c0 + 1 : formulaError('#VALUE!')),
  ROWS: (args, { evaluate }) => {
    const value = evaluate(args[0]);
    return isMatrix(value) ? value.length : 1;
  },
  COLUMNS: (args, { evaluate }) => {
    const value = evaluate(args[0]);
    return isMatrix(value) ? (value[0]?.length ?? 0) : 1;
  },
  TEXT: (args, { evaluate }) => {
    const value = toScalar(evaluate(args[0]));
    const format = toText(toScalar(evaluate(args[1])));
    if (isError(value)) return value;
    if (isError(format)) return format;
    const n = typeof value === 'string' ? toNumber(value) : toNumber(value);
    if (isError(n)) return typeof value === 'string' ? value : n;
    return formatNumber(n, format);
  },
  TODAY: () => Math.floor(dateToSerial(new Date()) + new Date().getTimezoneOffset() / -1440),
  NOW: () => dateToSerial(new Date()) - new Date().getTimezoneOffset() / 1440,
  DATE: (args, { evaluate }) => {
    const [y, m, d] = args.map((a) => num(evaluate(a)));
    for (const part of [y, m, d]) if (isError(part)) return part;
    let year = Math.trunc(y as number);
    if (year < 1900) year += 1900;
    const serial = dateToSerial(new Date(Date.UTC(year, Math.trunc(m as number) - 1, Math.trunc(d as number))));
    return serial < 0 ? formulaError('#NUM!') : serial;
  },
  TIME: (args, { evaluate }) => {
    const [h, m, s] = args.map((a) => num(evaluate(a)));
    for (const part of [h, m, s]) if (isError(part)) return part;
    return (((h as number) * 3600 + (m as number) * 60 + (s as number)) / 86_400) % 1;
  },
  YEAR: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    return isError(date) ? date : date.getUTCFullYear();
  },
  MONTH: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    return isError(date) ? date : date.getUTCMonth() + 1;
  },
  DAY: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    return isError(date) ? date : date.getUTCDate();
  },
  HOUR: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    return isError(date) ? date : date.getUTCHours();
  },
  MINUTE: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    return isError(date) ? date : date.getUTCMinutes();
  },
  SECOND: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    return isError(date) ? date : date.getUTCSeconds();
  },
  WEEKDAY: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    if (isError(date)) return date;
    const type = args.length > 1 ? num(evaluate(args[1])) : 1;
    if (isError(type)) return type;
    const day = date.getUTCDay(); // 0 = Sunday
    if (type === 2) return day === 0 ? 7 : day;
    if (type === 3) return day === 0 ? 6 : day - 1;
    return day + 1;
  },
  EDATE: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    const months = num(evaluate(args[1]));
    if (isError(date)) return date;
    if (isError(months)) return months;
    const result = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + Math.trunc(months), date.getUTCDate()));
    if (result.getUTCDate() !== date.getUTCDate()) result.setUTCDate(0); // clamp to month end
    return Math.floor(dateToSerial(result));
  },
  EOMONTH: (args, { evaluate }) => {
    const date = dateParts(evaluate(args[0]));
    const months = num(evaluate(args[1]));
    if (isError(date)) return date;
    if (isError(months)) return months;
    return Math.floor(dateToSerial(new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + Math.trunc(months) + 1, 0))));
  },
  DAYS: (args, { evaluate }) => {
    const end = num(evaluate(args[0]));
    const start = num(evaluate(args[1]));
    if (isError(end)) return end;
    if (isError(start)) return start;
    return Math.floor(end) - Math.floor(start);
  },
  N: (args, { evaluate }) => {
    const value = toScalar(evaluate(args[0]));
    return typeof value === 'number' ? value : typeof value === 'boolean' ? (value ? 1 : 0) : isError(value) ? value : 0;
  },
  T: (args, { evaluate }) => {
    const value = toScalar(evaluate(args[0]));
    return typeof value === 'string' ? value : isError(value) ? value : '';
  },
};

/** Functions that return or accept JS Dates in formulajs; they are replaced above. */
const DATE_FUNCTIONS = new Set(['DATEVALUE', 'DATEDIF', 'NETWORKDAYS', 'WORKDAY', 'YEARFRAC', 'WEEKNUM', 'ISOWEEKNUM', 'DAYS360', 'TIMEVALUE']);

/** Arguments that must stay one-dimensional for formulajs lookups. */
const FLATTEN_ARGS: Record<string, number[]> = { MATCH: [1], LOOKUP: [1, 2] };

/** Ranges that are summed/averaged: Excel ignores text there, formulajs would concatenate it. */
const NUMERIC_ARGS: Record<string, number[]> = { SUMIF: [2], SUMIFS: [0], AVERAGEIF: [2], AVERAGEIFS: [0] };

const numbersOnly = (matrix: Matrix): Matrix => matrix.map((row) => row.map((v) => (typeof v === 'number' ? v : null)));

type FormulaJsFunction = (...args: unknown[]) => unknown;

function formulajsFunction(name: string): FormulaJsFunction | undefined {
  const library = formulajs as unknown as Record<string, unknown>;
  const direct = library[name];
  if (typeof direct === 'function') return direct as FormulaJsFunction;
  // Dotted names: STDEV.S → STDEV.S or nested { STDEV: { S } }.
  const dot = name.indexOf('.');
  if (dot > 0) {
    const parent = library[name.slice(0, dot)] as Record<string, unknown> | undefined;
    const child = parent?.[name.slice(dot + 1)];
    if (typeof child === 'function') return child as FormulaJsFunction;
  }
  return undefined;
}

function fromLibraryResult(result: unknown): Value {
  if (result === undefined) return null;
  if (result instanceof Date) return dateToSerial(result);
  if (result instanceof Error) {
    const code = result.message as Parameters<typeof formulaError>[0];
    return formulaError(/^#/.test(code) ? code : '#VALUE!');
  }
  if (Array.isArray(result)) {
    return (Array.isArray(result[0]) ? result : [result]).map((row: unknown[]) => row.map((v) => toScalar(fromLibraryResult(v))));
  }
  if (typeof result === 'number') return Number.isFinite(result) ? result : formulaError('#NUM!');
  if (typeof result === 'string' || typeof result === 'boolean' || result === null) return result;
  return formulaError('#VALUE!');
}

export function isSupportedFunction(name: string): boolean {
  return name in NATIVE || (!DATE_FUNCTIONS.has(name) && formulajsFunction(name) !== undefined);
}

export function callFunction(name: string, args: FormulaNode[], ctx: CallContext): Value {
  const native = NATIVE[name];
  if (native) return native(args, ctx);
  const fn = DATE_FUNCTIONS.has(name) ? undefined : formulajsFunction(name);
  if (!fn) throw new UnsupportedFormulaError(`Function ${name}`);

  const flattenAt = FLATTEN_ARGS[name] ?? [];
  const numericAt = NUMERIC_ARGS[name] ?? [];
  const values = args.map((arg, index) => {
    const value = ctx.evaluate(arg);
    if (!isMatrix(value)) return value;
    if (numericAt.includes(index)) return numbersOnly(value);
    return flattenAt.includes(index) ? value.flat() : value;
  });
  try {
    return fromLibraryResult(fn(...values));
  } catch {
    return formulaError('#VALUE!');
  }
}

export { compareValues };
export type { Matrix };
