import { CellKind, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, type WorksheetModel } from '../../models/workbook';
import { measureText } from '../../utils/textMeasure';
import { SheetBuilder } from './sheetBuilder';

/**
 * Streaming RFC 4180 parser for CSV/TSV that writes straight into the compact
 * sheet model. It avoids building one object per cell, so multi-million-cell
 * files stay within a few hundred MB.
 *
 * Cell text is kept exactly as written (no "00123" → 123 or date conversions).
 * Values that look numeric or boolean are typed only for alignment and sums.
 */

const QUOTE = 34;
const CR = 13;
const LF = 10;

const CANDIDATE_DELIMITERS = [',', ';', '\t', '|'];

/** Columns are auto-sized from this many leading rows. */
const MEASURE_ROWS = 500;
const MEASURE_COLS = 256;
const MIN_AUTO_WIDTH = 48;
const MAX_AUTO_WIDTH = 360;

const NUMBER = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;
const GROUPED_NUMBER = /^[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/;
const LEADING_ZERO = /^[-+]?0\d/;

/** Picks the delimiter that splits the first lines most consistently. */
export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 64 * 1024);
  const perLine: number[][] = CANDIDATE_DELIMITERS.map(() => []);
  let current = CANDIDATE_DELIMITERS.map(() => 0);
  let inQuotes = false;
  let lines = 0;

  const endLine = () => {
    if (current.some((count) => count > 0)) {
      current.forEach((count, k) => perLine[k].push(count));
      lines++;
    }
    current = CANDIDATE_DELIMITERS.map(() => 0);
  };

  for (let i = 0; i < sample.length && lines < 20; i++) {
    const ch = sample[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (!inQuotes) {
      if (ch === '\n' || ch === '\r') endLine();
      else {
        const k = CANDIDATE_DELIMITERS.indexOf(ch);
        if (k >= 0) current[k]++;
      }
    }
  }
  endLine();

  let best = ',';
  let bestScore = 0;
  CANDIDATE_DELIMITERS.forEach((delimiter, k) => {
    const counts = perLine[k];
    if (counts.length === 0 || counts[0] === 0) return;
    const present = counts.filter((c) => c > 0).length / counts.length;
    const consistent = counts.filter((c) => c === counts[0]).length / counts.length;
    const score = present * 2 + consistent;
    if (score > bestScore) {
      bestScore = score;
      best = delimiter;
    }
  });
  return best;
}

function inferKind(value: string): [CellKind, number] {
  const first = value.charCodeAt(0);
  const numericStart = (first >= 48 && first <= 57) || first === 45 || first === 43 || first === 46;
  if (numericStart && value.length <= 24) {
    if (NUMBER.test(value)) {
      // Identifiers such as zip codes and long account numbers stay text.
      const mantissaDigits = value.split(/[eE]/)[0].replace(/[^0-9]/g, '').length;
      if (!LEADING_ZERO.test(value) && mantissaDigits <= 15) return [CellKind.Number, Number(value)];
    } else if (GROUPED_NUMBER.test(value)) {
      return [CellKind.Number, Number(value.replace(/,/g, ''))];
    }
  } else if (value.length === 4 || value.length === 5) {
    const upper = value.toUpperCase();
    if (upper === 'TRUE') return [CellKind.Boolean, 1];
    if (upper === 'FALSE') return [CellKind.Boolean, 0];
  }
  return [CellKind.String, 0];
}

export interface DelimitedParseOptions {
  name: string;
  index?: number;
  /** Forced delimiter; detected from the content when omitted. */
  delimiter?: string;
}

export interface DelimitedParseResult {
  sheet: WorksheetModel;
  delimiter: string;
  truncatedRows: boolean;
  truncatedColumns: boolean;
}

export function parseDelimited(text: string, options: DelimitedParseOptions): DelimitedParseResult {
  let i = 0;
  let delimiter = options.delimiter;
  // Excel honours a leading "sep=;" line.
  const sep = /^sep=(.)\r?\n/.exec(text.slice(0, 8));
  if (sep) {
    delimiter = sep[1];
    i = sep[0].length;
  }
  delimiter ??= detectDelimiter(text);
  const delim = delimiter.charCodeAt(0);

  const builder = new SheetBuilder({ name: options.name, index: options.index ?? 0 });
  const widths: number[] = [];
  const n = text.length;
  let row = 0;
  let col = 0;
  let truncatedRows = false;
  let truncatedColumns = false;

  while (i < n) {
    if (row >= EXCEL_MAX_ROWS) {
      truncatedRows = true;
      break;
    }

    let ch = text.charCodeAt(i);
    let value: string;
    if (ch === QUOTE) {
      i++;
      let segment = i;
      let parts: string[] | null = null;
      while (i < n) {
        ch = text.charCodeAt(i);
        if (ch === QUOTE) {
          if (i + 1 < n && text.charCodeAt(i + 1) === QUOTE) {
            (parts ??= []).push(text.slice(segment, i + 1));
            i += 2;
            segment = i;
            continue;
          }
          break;
        }
        i++;
      }
      value = parts ? parts.join('') + text.slice(segment, i) : text.slice(segment, i);
      if (i < n) i++; // closing quote
      // Like Excel, keep stray characters between the closing quote and the delimiter.
      const tail = i;
      while (i < n) {
        ch = text.charCodeAt(i);
        if (ch === delim || ch === LF || ch === CR) break;
        i++;
      }
      if (i > tail) value += text.slice(tail, i);
    } else {
      const start = i;
      while (i < n) {
        ch = text.charCodeAt(i);
        if (ch === delim || ch === LF || ch === CR) break;
        i++;
      }
      value = text.slice(start, i);
    }

    if (value.length > 0) {
      if (col < EXCEL_MAX_COLS) {
        const [kind, num] = inferKind(value);
        builder.add(row, col, kind, num, value);
        if (row < MEASURE_ROWS && col < MEASURE_COLS) {
          const sample = value.length > 200 ? value.slice(0, 200) : value;
          const width = measureText(sample, false, false, kind === CellKind.Number);
          if (!(widths[col] >= width)) widths[col] = width;
        }
      } else {
        truncatedColumns = true;
      }
    }

    if (i >= n) break;
    ch = text.charCodeAt(i);
    i++;
    if (ch === delim) {
      col++;
    } else {
      if (ch === CR && i < n && text.charCodeAt(i) === LF) i++;
      row++;
      col = 0;
    }
  }

  const columnWidths = new Map<number, number>();
  widths.forEach((width, c) => {
    columnWidths.set(c, Math.min(MAX_AUTO_WIDTH, Math.max(MIN_AUTO_WIDTH, Math.ceil(width + 12))));
  });

  return { sheet: builder.finish({ columnWidths }), delimiter, truncatedRows, truncatedColumns };
}
