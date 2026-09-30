import * as XLSX from 'xlsx';

/** Excel number formatting (via SheetJS SSF), shared by the parser and the editor. */

const dateFormats = new Map<string, boolean>();

export function isDateFormat(format: string | undefined): boolean {
  if (!format || format === 'General') return false;
  let result = dateFormats.get(format);
  if (result === undefined) {
    try {
      result = Boolean(XLSX.SSF.is_date(format));
    } catch {
      result = false;
    }
    dateFormats.set(format, result);
  }
  return result;
}

/** Formats a number with an Excel number format (General when omitted). */
export function formatNumber(value: number, format?: string, date1904 = false): string {
  if (!format || format === 'General') {
    if (Number.isInteger(value) && Math.abs(value) < 1e11) return String(value);
    try {
      return XLSX.SSF.format('General', value);
    } catch {
      return String(value);
    }
  }
  try {
    return XLSX.SSF.format(format, value, { date1904 });
  } catch {
    return String(value);
  }
}
