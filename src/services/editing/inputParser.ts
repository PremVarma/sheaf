import { CellKind, type CellStyle } from '../../models/workbook';
import { dateToSerial } from '../formula/values';
import { formatNumber, isDateFormat } from '../workbook/numberFormat';

/** How typed text should be stored, before formulas are calculated. */
export interface ParsedInput {
  kind: CellKind;
  num: number;
  /** Display text (formatted); empty for formulas until they're calculated. */
  text: string;
  formula?: string;
  /** Number format to apply when the cell has none (e.g. "12%" → 0%). */
  suggestedFormat?: string;
}

const NUMBER = /^[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:[eE][-+]?\d+)?$/;
const CURRENCY = /^([-+]?)[₹$€£¥]\s?(.+)$/;
const ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const SLASH_DATE = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/;
const TIME = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/;

/** Day-first unless the locale is US-style (month first), like Excel follows the OS locale. */
function monthFirst(): boolean {
  const locale = typeof navigator !== 'undefined' ? navigator.language : 'en-US';
  return /^en-(US|PH)$|^en$/i.test(locale);
}

function parseNumber(text: string): number | null {
  if (!text || !/\d/.test(text) || !NUMBER.test(text)) return null;
  const value = Number(text.replace(/,/g, ''));
  return Number.isFinite(value) ? value : null;
}

function makeDate(year: number, month: number, day: number): number | null {
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return dateToSerial(date);
}

/**
 * Interprets text typed into a cell the way Excel does: formulas start with
 * "=", a leading apostrophe forces text, and numbers, percentages, currency
 * amounts, dates and TRUE/FALSE are recognised.
 */
export function parseInput(input: string, style: CellStyle | undefined, date1904 = false): ParsedInput | null {
  if (input === '') return null;
  const format = style?.numberFormat;
  const display = (num: number, fmt = format) => formatNumber(num, fmt, date1904);

  if (input.startsWith("'")) return { kind: CellKind.String, num: 0, text: input.slice(1) };
  if (format === '@') return { kind: CellKind.String, num: 0, text: input };
  if (input.startsWith('=') && input.length > 1) {
    return { kind: CellKind.Empty, num: 0, text: '', formula: input };
  }

  const text = input.trim();
  const upper = text.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    return { kind: CellKind.Boolean, num: upper === 'TRUE' ? 1 : 0, text: upper };
  }

  const number = parseNumber(text);
  if (number !== null) {
    const kind = isDateFormat(format) ? CellKind.Date : CellKind.Number;
    return { kind, num: number, text: display(number) };
  }

  if (text.endsWith('%')) {
    const value = parseNumber(text.slice(0, -1).trim());
    if (value !== null) {
      const num = value / 100;
      const suggestedFormat = format ? undefined : text.includes('.') ? '0.00%' : '0%';
      return { kind: CellKind.Number, num, text: display(num, format ?? suggestedFormat), suggestedFormat };
    }
  }

  const currency = CURRENCY.exec(text);
  if (currency) {
    const value = parseNumber(currency[2]);
    if (value !== null) {
      const num = currency[1] === '-' ? -value : value;
      const symbol = text.replace(/^[-+]/, '').charAt(0);
      const suggestedFormat = format ? undefined : `"${symbol}"#,##0.00`;
      return { kind: CellKind.Number, num, text: display(num, format ?? suggestedFormat), suggestedFormat };
    }
  }

  const time = TIME.exec(text);
  if (time) {
    let hours = Number(time[1]);
    const minutes = Number(time[2]);
    const seconds = Number(time[3] ?? 0);
    const meridiem = time[4]?.toUpperCase();
    if (meridiem && hours >= 1 && hours <= 12) hours = (hours % 12) + (meridiem === 'PM' ? 12 : 0);
    if (hours < 24 && minutes < 60 && seconds < 60 && (!meridiem || Number(time[1]) <= 12)) {
      const num = (hours * 3600 + minutes * 60 + seconds) / 86_400;
      const suggestedFormat = isDateFormat(format) ? undefined : `h:mm${time[3] ? ':ss' : ''}${meridiem ? ' AM/PM' : ''}`;
      return { kind: CellKind.Date, num, text: display(num, suggestedFormat ?? format), suggestedFormat };
    }
  }

  let serial: number | null = null;
  const iso = ISO_DATE.exec(text);
  if (iso) serial = makeDate(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  const slash = iso ? null : SLASH_DATE.exec(text);
  if (slash) {
    const [a, b, year] = [Number(slash[1]), Number(slash[2]), Number(slash[3])];
    serial = monthFirst() ? makeDate(year, a, b) : makeDate(year, b, a);
  }
  if (serial !== null) {
    if (date1904) serial -= 1462;
    const suggestedFormat = isDateFormat(format) ? undefined : iso ? 'yyyy-mm-dd' : monthFirst() ? 'm/d/yyyy' : 'dd/mm/yyyy';
    return { kind: CellKind.Date, num: serial, text: display(serial, suggestedFormat ?? format), suggestedFormat };
  }

  return { kind: CellKind.String, num: 0, text: input };
}
