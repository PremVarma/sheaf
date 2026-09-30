/** Number format presets for the toolbar, and changing the decimals of a format. */

export interface FormatPreset {
  id: 'general' | 'number' | 'currency' | 'accounting' | 'shortDate' | 'longDate' | 'time' | 'percent' | 'fraction' | 'scientific' | 'text';
  label: string;
  /** Format code; undefined for General. */
  format?: string;
}

const REGION_CURRENCY: Record<string, string> = { IN: '₹', US: '$', GB: '£', JP: '¥', CN: '¥', CA: '$', AU: '$', NZ: '$', SG: '$', HK: '$' };
const EURO_REGIONS = /^(AT|BE|CY|DE|EE|ES|FI|FR|GR|HR|IE|IT|LT|LU|LV|MT|NL|PT|SI|SK)$/;

function locale(): string {
  return typeof navigator !== 'undefined' ? navigator.language : 'en-US';
}

/** The currency symbol for the system's region (₹ in India, $ in the US, € in the euro area…). */
export function localCurrency(language = locale()): string {
  const region = /[-_]([A-Za-z]{2})\b/.exec(language)?.[1]?.toUpperCase() ?? '';
  if (EURO_REGIONS.test(region)) return '€';
  return REGION_CURRENCY[region] ?? '$';
}

function monthFirst(language = locale()): boolean {
  return /^en-(US|PH)$|^en$/i.test(language);
}

export function formatPresets(language = locale()): FormatPreset[] {
  const symbol = localCurrency(language);
  return [
    { id: 'general', label: 'General' },
    { id: 'number', label: 'Number', format: '#,##0.00' },
    { id: 'currency', label: 'Currency', format: `"${symbol}"#,##0.00` },
    { id: 'accounting', label: 'Accounting', format: `_("${symbol}"* #,##0.00_);_("${symbol}"* (#,##0.00);_("${symbol}"* "-"??_);_(@_)` },
    { id: 'shortDate', label: 'Short Date', format: monthFirst(language) ? 'm/d/yyyy' : 'dd/mm/yyyy' },
    { id: 'longDate', label: 'Long Date', format: 'dddd, d mmmm yyyy' },
    { id: 'time', label: 'Time', format: 'h:mm:ss AM/PM' },
    { id: 'percent', label: 'Percentage', format: '0.00%' },
    { id: 'fraction', label: 'Fraction', format: '# ?/?' },
    { id: 'scientific', label: 'Scientific', format: '0.00E+00' },
    { id: 'text', label: 'Text', format: '@' },
  ];
}

/** Other currencies offered next to the local one. */
export const CURRENCY_SYMBOLS = ['₹', '$', '€', '£', '¥'];

export function currencyFormat(symbol: string): string {
  return `"${symbol}"#,##0.00`;
}

/** Which preset a format code is, if any (for the toolbar label). */
export function presetFor(format: string | undefined, presets: FormatPreset[]): FormatPreset | undefined {
  if (!format || format === 'General') return presets[0];
  const exact = presets.find((p) => p.format === format);
  if (exact) return exact;
  if (/^"[^"]+"#,##0(\.0+)?$/.test(format)) return presets.find((p) => p.id === 'currency');
  if (/^0(\.0+)?%$/.test(format)) return presets.find((p) => p.id === 'percent');
  if (/^#,##0(\.0+)?$/.test(format)) return presets.find((p) => p.id === 'number');
  return undefined;
}

/**
 * Adds or removes one decimal place in every section of a format.
 * General takes the decimals of `sample` as shown, like Excel.
 */
export function changeDecimals(format: string | undefined, sample: number | null, delta: 1 | -1): string | null {
  if (!format || format === 'General') {
    const text = sample === null ? '0' : String(Number(sample.toPrecision(10)));
    const decimals = /\.(\d+)$/.exec(text)?.[1].length ?? 0;
    const next = Math.max(0, decimals + delta);
    return next > 0 ? `0.${'0'.repeat(next)}` : '0';
  }
  let changed = false;
  const sections = splitSections(format).map((section) => {
    const next = adjustSection(section, delta);
    if (next !== section) changed = true;
    return next;
  });
  return changed ? sections.join(';') : null;
}

function splitSections(format: string): string[] {
  const out: string[] = [];
  let current = '';
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '\\' && !quoted) {
      current += ch + (format[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"' && !bracket) quoted = !quoted;
    else if (ch === '[' && !quoted) bracket = true;
    else if (ch === ']' && !quoted) bracket = false;
    if (ch === ';' && !quoted && !bracket) {
      out.push(current);
      current = '';
    } else current += ch;
  }
  out.push(current);
  return out;
}

/** Changes the decimals of the first number placeholder in a section (outside quotes and brackets). */
function adjustSection(section: string, delta: 1 | -1): string {
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '\\' && !quoted) {
      i++;
      continue;
    }
    if (ch === '"' && !bracket) {
      quoted = !quoted;
      continue;
    }
    if (quoted) continue;
    if (ch === '[') bracket = true;
    else if (ch === ']') bracket = false;
    if (bracket || !/[0#?]/.test(ch)) continue;
    // The integer part, then an optional "." with its decimal placeholders.
    let end = i;
    while (end < section.length && /[0#?,]/.test(section[end])) end++;
    if (/^\?+$/.test(section.slice(i, end)) && section[end] !== '.') {
      // Accounting's zero section ("-"??) pads with one "?" per decimal.
      if (delta > 0) return `${section.slice(0, end)}?${section.slice(end)}`;
      return end - i > 2 ? section.slice(0, end - 1) + section.slice(end) : section;
    }
    let decimals = 0;
    if (section[end] === '.') {
      let k = end + 1;
      while (k < section.length && /[0#?]/.test(section[k])) k++;
      decimals = k - end - 1;
      if (delta > 0) return `${section.slice(0, k)}0${section.slice(k)}`;
      if (decimals <= 1) return section.slice(0, end) + section.slice(k);
      return section.slice(0, k - 1) + section.slice(k);
    }
    return delta > 0 ? `${section.slice(0, end)}.0${section.slice(end)}` : section;
  }
  return section;
}
