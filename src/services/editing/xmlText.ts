/** Text helpers for writing SpreadsheetML. */

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // Characters not allowed in XML 1.0.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

export function unescapeXml(text: string): string {
  return text.replace(/&(?:#x([0-9a-f]+)|#(\d+)|(amp|lt|gt|quot|apos));/gi, (_, hex: string, dec: string, name: string) => {
    if (hex) return String.fromCodePoint(parseInt(hex, 16));
    if (dec) return String.fromCodePoint(Number(dec));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[name.toLowerCase()] ?? '';
  });
}

/** Functions newer than Excel 2007 must be stored with the _xlfn. prefix. */
const FUTURE_FUNCTIONS = [
  'CONCAT', 'TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'XLOOKUP', 'XMATCH', 'IFNA', 'DAYS',
  'ISOWEEKNUM', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'PERCENTILE.INC', 'PERCENTILE.EXC',
  'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'CEILING.MATH', 'FLOOR.MATH',
  'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'LET',
];
const FUTURE_PATTERN = new RegExp(`(?<![\\w.])(${FUTURE_FUNCTIONS.map((f) => f.replace('.', '\\.')).join('|')})\\(`, 'gi');

/** Formula text as stored in the file: no "=", string literals untouched, _xlfn. prefixes added. */
export function formulaForFile(formula: string): string {
  const body = formula.startsWith('{=') && formula.endsWith('}') ? formula.slice(2, -1) : formula.replace(/^=/, '');
  return body
    .split(/("(?:[^"]|"")*")/)
    .map((part, index) => (index % 2 === 1 ? part : part.replace(FUTURE_PATTERN, (_, name: string) => `_xlfn.${name.toUpperCase()}(`)))
    .join('');
}
