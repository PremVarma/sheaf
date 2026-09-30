import { CellKind } from '../../models/workbook';
import { dateToSerial, serialToDate } from '../formula/values';
import type { StoredCell } from './cells';

/**
 * Excel's fill handle: extending a run of cells continues a series when it can
 * recognise one (numbers, dates, "Item 1", month and day names) and repeats
 * the cells otherwise. Formulas are always copied (their references move).
 */

export interface FilledCell {
  /** Index in the run of the cell this one continues (its style and formula are copied). */
  from: number;
  /** Value when a series was continued; undefined to copy the source cell. */
  value?: { kind: CellKind; num: number; text: string };
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const LISTS = [MONTHS, MONTHS.map((m) => m.slice(0, 3)), DAYS, DAYS.map((d) => d.slice(0, 3))];

function matchCase(word: string, like: string): string {
  if (like === like.toUpperCase()) return word.toUpperCase();
  if (like === like.toLowerCase()) return word.toLowerCase();
  return word;
}

/** Least-squares line through (0, v0), (1, v1), … — Excel's linear trend. */
function linearTrend(values: number[]): (position: number) => number {
  const n = values.length;
  const meanX = (n - 1) / 2;
  const meanY = values.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  values.forEach((v, x) => {
    num += (x - meanX) * (v - meanY);
    den += (x - meanX) ** 2;
  });
  const slope = den === 0 ? 0 : num / den;
  return (position) => Number((meanY + slope * (position - meanX)).toPrecision(15));
}

function addMonths(serial: number, months: number, date1904: boolean): number {
  const date = serialToDate(date1904 ? serial + 1462 : serial);
  const day = date.getUTCDate();
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
  const lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
  target.setUTCDate(Math.min(day, lastDay));
  const result = dateToSerial(target) + (serial - Math.floor(serial));
  return date1904 ? result - 1462 : result;
}

/** Month step between consecutive dates when they fall on the same day of the month, else null. */
function monthStep(serials: number[], date1904: boolean): number | null {
  const dates = serials.map((s) => serialToDate(date1904 ? s + 1462 : s));
  const months = dates.map((d) => d.getUTCFullYear() * 12 + d.getUTCMonth());
  const step = months[1] - months[0];
  if (step === 0) return null;
  for (let k = 1; k < dates.length; k++) {
    if (dates[k].getUTCDate() !== dates[0].getUTCDate() || months[k] - months[k - 1] !== step) return null;
  }
  return step;
}

/**
 * Continues `run` (cells in fill order, nulls for empty cells) by `count` cells.
 * `format(num)` formats a number with the source cell's number format.
 */
export function extendRun(
  run: (StoredCell | null)[],
  count: number,
  series: boolean,
  format: (num: number, from: number) => string,
  date1904 = false,
): FilledCell[] {
  const length = run.length;
  const copy = (): FilledCell[] => Array.from({ length: count }, (_, k) => ({ from: k % length }));
  if (!series || length === 0 || run.some((cell) => !cell || cell.formula)) return copy();
  const cells = run as StoredCell[];
  const last = length - 1;

  // Dates: one date counts up by days; several follow their month or day step.
  if (cells.every((c) => c.kind === CellKind.Date)) {
    const serials = cells.map((c) => c.num);
    const months = length >= 2 ? monthStep(serials, date1904) : null;
    const trend = length >= 2 ? linearTrend(serials) : (position: number) => serials[0] + position;
    return Array.from({ length: count }, (_, k) => {
      const num = months !== null ? addMonths(serials[last], months * (k + 1), date1904) : trend(last + k + 1);
      return { from: last, value: { kind: CellKind.Date, num, text: format(num, last) } };
    });
  }

  // Numbers: a single number is repeated (as in Excel); two or more follow their trend.
  if (cells.every((c) => c.kind === CellKind.Number)) {
    if (length === 1) return copy();
    const trend = linearTrend(cells.map((c) => c.num));
    return Array.from({ length: count }, (_, k) => {
      const num = trend(last + k + 1);
      return { from: last, value: { kind: CellKind.Number, num, text: format(num, last) } };
    });
  }

  if (cells.every((c) => c.kind === CellKind.String)) {
    // Month and day names.
    for (const list of LISTS) {
      const positions = cells.map((c) => list.findIndex((word) => word.toLowerCase() === c.text.toLowerCase()));
      if (positions.some((p) => p < 0)) continue;
      const step = length >= 2 ? (positions[last] - positions[last - 1] + list.length) % list.length || list.length : 1;
      return Array.from({ length: count }, (_, k) => {
        const word = list[(positions[last] + step * (k + 1)) % list.length];
        return { from: last, value: { kind: CellKind.String, num: 0, text: matchCase(word, cells[last].text) } };
      });
    }
    // Text ending in a number: Item 1, Item 2, …
    const parts = cells.map((c) => /^(.*?)(\d+)(\D*)$/.exec(c.text));
    if (parts.every((p) => p && p[1] === parts[0]![1] && p[3] === parts[0]![3])) {
      const numbers = parts.map((p) => Number(p![2]));
      const step = length >= 2 ? (numbers[last] - numbers[0]) / last : 1;
      if (Number.isInteger(step)) {
        const width = parts[last]![2].startsWith('0') ? parts[last]![2].length : 0;
        return Array.from({ length: count }, (_, k) => {
          const n = Math.max(0, numbers[last] + step * (k + 1));
          const digits = String(n).padStart(width, '0');
          return { from: last, value: { kind: CellKind.String, num: 0, text: `${parts[0]![1]}${digits}${parts[0]![3]}` } };
        });
      }
    }
  }
  return copy();
}
