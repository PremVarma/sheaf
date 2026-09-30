import type { WorkbookFormat } from '../../models/workbook';

export const SUPPORTED_EXTENSIONS: readonly WorkbookFormat[] = ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'];

/** Largest files we attempt to open. Workbooks expand ~10x in memory when parsed. */
export const SIZE_LIMITS = {
  workbook: 200 * 1024 * 1024,
  text: 500 * 1024 * 1024,
} as const;

/** Files above this size show a "large file" hint while loading. */
export const LARGE_FILE_BYTES = 20 * 1024 * 1024;

export function getExtension(fileName: string): string {
  const base = fileName.split(/[\\/]/).pop() ?? fileName;
  const dot = base.lastIndexOf('.');
  return dot > 0 ? base.slice(dot + 1).toLowerCase() : '';
}

export function isSupportedExtension(ext: string): ext is WorkbookFormat {
  return (SUPPORTED_EXTENSIONS as readonly string[]).includes(ext);
}

/** Files without an extension are sniffed by content; anything else must be supported. */
export function isOpenableFileName(fileName: string): boolean {
  const ext = getExtension(fileName);
  return ext === '' || isSupportedExtension(ext);
}

export function sizeLimitFor(fileName: string): number {
  const ext = getExtension(fileName);
  return ext === 'csv' || ext === 'tsv' ? SIZE_LIMITS.text : SIZE_LIMITS.workbook;
}

export const FORMAT_LABELS: Record<WorkbookFormat, string> = {
  xlsx: 'Excel Workbook (.xlsx)',
  xlsm: 'Excel Macro-Enabled Workbook (.xlsm)',
  xls: 'Excel 97–2003 Workbook (.xls)',
  csv: 'Comma-Separated Values (.csv)',
  tsv: 'Tab-Separated Values (.tsv)',
};

export type Container = 'zip' | 'cfb' | 'biff' | 'text' | 'binary' | 'empty';

/** Identifies the physical container from the first bytes of a file. */
export function sniffContainer(bytes: Uint8Array): Container {
  if (bytes.length === 0) return 'empty';
  const b = bytes;
  if (b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && (b[2] === 0x03 || b[2] === 0x05) && (b[3] === 0x04 || b[3] === 0x06)) {
    return 'zip';
  }
  if (
    b.length >= 8 &&
    b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0 &&
    b[4] === 0xa1 && b[5] === 0xb1 && b[6] === 0x1a && b[7] === 0xe1
  ) {
    return 'cfb';
  }
  // Pre-BIFF8 worksheets stored as a bare stream start with a BOF record.
  if (b.length >= 4 && b[0] === 0x09 && (b[1] === 0x00 || b[1] === 0x02 || b[1] === 0x04 || b[1] === 0x08) && b[3] === 0x00) {
    return 'biff';
  }
  return looksLikeText(bytes) ? 'text' : 'binary';
}

function looksLikeText(bytes: Uint8Array): boolean {
  // UTF-16 with a byte order mark is text even though it contains NUL bytes.
  if (bytes.length >= 2 && ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))) {
    return true;
  }
  const sample = Math.min(bytes.length, 64 * 1024);
  let control = 0;
  for (let i = 0; i < sample; i++) {
    const c = bytes[i];
    if (c === 0) return false;
    // Tab, LF, VT, FF, CR are fine; other C0 control characters are suspicious.
    if (c < 0x20 && (c < 0x09 || c > 0x0d) && c !== 0x1b) control++;
  }
  return control <= sample * 0.01;
}

/** True when a ZIP file still has its end-of-central-directory record (not truncated). */
export function hasZipEndRecord(bytes: Uint8Array): boolean {
  const stop = Math.max(0, bytes.length - 65_557);
  for (let i = bytes.length - 22; i >= stop; i--) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) return true;
  }
  return false;
}

/**
 * Decodes delimited text. UTF-8 (with or without BOM) is tried first; UTF-16 is
 * detected by BOM; anything that isn't valid UTF-8 falls back to Windows-1252,
 * which is what Excel writes on Western Windows systems.
 */
export function decodeText(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/** HTML or XML spreadsheets that are often saved with an .xls extension. */
export function looksLikeMarkup(text: string): boolean {
  const head = text.slice(0, 2048).trimStart().toLowerCase();
  return (
    head.startsWith('<?xml') ||
    head.startsWith('<!doctype html') ||
    head.startsWith('<html') ||
    head.startsWith('<table') ||
    head.includes('urn:schemas-microsoft-com:office:spreadsheet')
  );
}
