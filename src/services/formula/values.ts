import type { ErrorCode } from './ast';

/**
 * Runtime values. Errors are Error instances whose message is the Excel code,
 * which is also how @formulajs/formulajs represents them.
 */
export type Scalar = number | string | boolean | null | Error;
export type Matrix = Scalar[][];
export type Value = Scalar | Matrix;

const errorCache = new Map<string, Error>();

export function formulaError(code: ErrorCode): Error {
  let error = errorCache.get(code);
  if (!error) {
    error = new Error(code);
    errorCache.set(code, error);
  }
  return error;
}

export const isError = (value: unknown): value is Error => value instanceof Error;
export const isMatrix = (value: Value): value is Matrix => Array.isArray(value);

/** Top-left value of a range used where a single value is expected (implicit intersection light). */
export function toScalar(value: Value): Scalar {
  if (!isMatrix(value)) return value;
  return value[0]?.[0] ?? null;
}

export function toNumber(value: Scalar): number | Error {
  if (typeof value === 'number') return value;
  if (value === null) return 0;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (isError(value)) return value;
  const text = value.trim();
  if (text === '') return 0;
  const percent = text.endsWith('%');
  const numeric = Number((percent ? text.slice(0, -1) : text).replace(/,/g, ''));
  if (Number.isNaN(numeric)) return formulaError('#VALUE!');
  return percent ? numeric / 100 : numeric;
}

/** Number → text the way Excel's General format does (up to 15 significant digits). */
export function numberToText(value: number): string {
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
  const precise = Number(value.toPrecision(15));
  const text = String(precise);
  return text.includes('e') ? precise.toExponential().replace('e+', 'E+').replace('e-', 'E-') : text;
}

export function toText(value: Scalar): string | Error {
  if (typeof value === 'string') return value;
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (isError(value)) return value;
  return numberToText(value);
}

export function toBoolean(value: Scalar): boolean | Error {
  if (typeof value === 'boolean') return value;
  if (value === null) return false;
  if (typeof value === 'number') return value !== 0;
  if (isError(value)) return value;
  const upper = value.toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  return formulaError('#VALUE!');
}

/** Excel comparison: numbers < text < booleans; text compares case-insensitively. */
export function compareValues(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar) => (typeof v === 'number' || v === null ? 0 : typeof v === 'string' ? 1 : 2);
  // Empty compares as 0 against numbers and "" against text.
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : a < b ? -1 : 1;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x === y ? 0 : x < y ? -1 : 1;
  }
  return Number(a) - Number(b);
}

/** Excel serial date ⇄ JS Date (UTC), 1900 date system including the 1900 leap-year quirk. */
const EPOCH = Date.UTC(1899, 11, 30);
export function serialToDate(serial: number): Date {
  return new Date(EPOCH + Math.round(serial * 86_400_000));
}
export function dateToSerial(date: Date): number {
  return (date.getTime() - EPOCH) / 86_400_000;
}
