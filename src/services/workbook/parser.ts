import * as XLSX from 'xlsx';
import * as cptable from 'xlsx/dist/cpexcel.full.mjs';
import {
  CellKind,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type CellRange,
  type SheetKind,
  type SheetVisibility,
  type TableInfo,
  type WorkbookFormat,
  type WorkbookModel,
  type WorkbookProperties,
  type WorksheetModel,
} from '../../models/workbook';
import { formatCount } from '../../utils/format';
import { parseDelimited } from './csvParser';
import { isOutOfMemory, WorkbookError } from './errors';
import { formatNumber, isDateFormat } from './numberFormat';
import {
  decodeText,
  getExtension,
  hasZipEndRecord,
  isSupportedExtension,
  looksLikeMarkup,
  sniffContainer,
  type Container,
} from './formatDetection';
import { SheetBuilder } from './sheetBuilder';
import { createFillResolver, createXfResolver, StyleTable, type SheetJsStyles } from './styles';
import {
  columnWidthToPx,
  pointsToPx,
  readBorders,
  readSheetTables,
  readSheetViewInfo,
  readWorkbookParts,
  readZipText,
  scanCellStyleRefs,
  zipEntrySize,
  type CellStyleRefs,
  type SheetViewInfo,
  type ZipFiles,
} from './xlsxParts';

// Legacy .xls files may store text in Windows code pages.
XLSX.set_cptable(cptable);

/** Sheet XML larger than this skips the per-cell style scan (fills still come from SheetJS). */
const STYLE_SCAN_LIMIT = 96 * 1024 * 1024;

const ERROR_TEXT: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
  0x2b: '#GETTING_DATA',
};

/**
 * Parses file bytes into the normalized model. Pure and synchronous, so it can
 * run in a Web Worker (the normal path) or directly in tests.
 */
export function parseWorkbookBytes(bytes: Uint8Array, fileName: string): WorkbookModel {
  const ext = getExtension(fileName);
  if (ext && !isSupportedExtension(ext)) throw new WorkbookError('UNSUPPORTED_FORMAT', `Extension .${ext}`);

  const container = sniffContainer(bytes);
  const file = { fileName, fileSize: bytes.length };

  switch (container) {
    case 'empty':
      throw new WorkbookError('EMPTY_FILE');
    case 'binary':
      throw new WorkbookError('INVALID_FILE', 'Unrecognized binary content');
    case 'text': {
      if (ext === 'xlsx' || ext === 'xlsm') throw new WorkbookError('INVALID_FILE', 'Plain text in an Office Open XML file');
      const text = decodeText(bytes);
      if (ext === 'xls' && looksLikeMarkup(text)) {
        return fromSheetJs(read(text, 'string', container, bytes), 'xls', file, null);
      }
      const format: WorkbookFormat = ext === 'tsv' ? 'tsv' : ext === 'xls' ? 'xls' : 'csv';
      const bom = (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) || bytes[0] === 0xff || bytes[0] === 0xfe;
      return fromDelimitedText(text, format, file, ext === 'tsv' ? '\t' : undefined, bom);
    }
    default: {
      const workbook = read(bytes, 'array', container, bytes);
      const format: WorkbookFormat = container === 'zip' ? (ext === 'xlsm' ? 'xlsm' : 'xlsx') : 'xls';
      const files = container === 'zip' ? ((workbook as unknown as { files?: ZipFiles }).files ?? null) : null;
      return fromSheetJs(workbook, format, file, files);
    }
  }
}

function read(input: Uint8Array | string, type: 'array' | 'string', container: Container, bytes: Uint8Array): XLSX.WorkBook {
  try {
    return XLSX.read(input, {
      type,
      dense: true,
      cellStyles: true,
      cellNF: true,
      cellText: false,
      cellHTML: false,
      cellFormula: true,
      cellDates: false,
      bookFiles: container === 'zip',
      bookVBA: false,
    });
  } catch (error) {
    throw classifyReadError(error, container, bytes);
  }
}

function classifyReadError(error: unknown, container: Container, bytes: Uint8Array): WorkbookError {
  const message = error instanceof Error ? error.message : String(error);
  if (isOutOfMemory(error)) return new WorkbookError('OUT_OF_MEMORY', message);
  if (/password|encrypt/i.test(message)) return new WorkbookError('PASSWORD_PROTECTED', message);
  if (container === 'zip') {
    if (!hasZipEndRecord(bytes)) return new WorkbookError('CORRUPTED', `Truncated ZIP archive: ${message}`);
    if (/unsupported zip|cannot find|missing|could not find/i.test(message)) return new WorkbookError('NOT_A_WORKBOOK', message);
  }
  if (container === 'cfb' && /cannot find workbook|workbook stream/i.test(message)) {
    return new WorkbookError('NOT_A_WORKBOOK', message);
  }
  if (/is not a spreadsheet/i.test(message)) return new WorkbookError('INVALID_FILE', message);
  return new WorkbookError('CORRUPTED', message);
}

function sheetNameFromFile(fileName: string): string {
  const base = fileName.split(/[\\/]/).pop() ?? fileName;
  const dot = base.lastIndexOf('.');
  return (dot > 0 ? base.slice(0, dot) : base).slice(0, 31) || 'Sheet1';
}

function fromDelimitedText(
  text: string,
  format: WorkbookFormat,
  file: { fileName: string; fileSize: number },
  forcedDelimiter: string | undefined,
  bom = false,
): WorkbookModel {
  const { sheet, delimiter, truncatedRows, truncatedColumns } = parseDelimited(text, {
    name: sheetNameFromFile(file.fileName),
    delimiter: forcedDelimiter,
  });
  const warnings: string[] = [];
  if (truncatedRows) warnings.push(`Only the first ${formatCount(EXCEL_MAX_ROWS)} rows were loaded.`);
  if (truncatedColumns) warnings.push(`Only the first ${formatCount(EXCEL_MAX_COLS)} columns were loaded.`);
  return {
    ...file,
    format,
    properties: {},
    sheets: [sheet],
    styles: [{}],
    activeSheetIndex: 0,
    date1904: false,
    warnings,
    csv: { delimiter, bom },
  };
}

interface SheetContext {
  date1904: boolean;
  table: StyleTable;
  resolveXf: ((xf: number) => number) | null;
  resolveFill: (fill: object) => number;
  styleRefs: CellStyleRefs | null;
  view: SheetViewInfo | null;
  tables?: TableInfo[];
}

function fromSheetJs(
  workbook: XLSX.WorkBook,
  format: WorkbookFormat,
  file: { fileName: string; fileSize: number },
  files: ZipFiles | null,
): WorkbookModel {
  const names = workbook.SheetNames ?? [];
  if (names.length === 0) throw new WorkbookError('NO_SHEETS');

  const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904);
  const table = new StyleTable();
  const parts = files ? readWorkbookParts(files) : null;
  const sheetJsStyles = (workbook as unknown as { Styles?: SheetJsStyles }).Styles;

  let resolveXf: SheetContext['resolveXf'] = null;
  if (files && parts && sheetJsStyles?.CellXf) {
    const themeColors = (
      (workbook as unknown as { Themes?: { themeElements?: { clrScheme?: { rgb?: string }[] } } }).Themes?.themeElements
        ?.clrScheme ?? []
    ).map((c) => (c?.rgb ? `#${c.rgb.slice(-6).toLowerCase()}` : '#000000'));
    const stylesPath = Object.keys(files).find((k) => /(^|\/)styles\.xml$/i.test(k));
    const stylesXml = stylesPath ? readZipText(files, stylesPath) : null;
    resolveXf = createXfResolver(sheetJsStyles, stylesXml ? readBorders(stylesXml, themeColors) : [], table);
  }
  const resolveFill = createFillResolver(table);

  const sheets: WorksheetModel[] = names.map((name, index) => {
    const worksheet = workbook.Sheets[name];
    const hidden = workbook.Workbook?.Sheets?.[index]?.Hidden;
    const visibility: SheetVisibility = hidden === 2 ? 'veryHidden' : hidden === 1 ? 'hidden' : 'visible';

    let styleRefs: CellStyleRefs | null = null;
    let view: SheetViewInfo | null = null;
    let tables: TableInfo[] | undefined;
    const path = files && parts ? (parts.sheets.find((s) => s.name === name)?.path ?? parts.sheets[index]?.path) : null;
    if (files && path) {
      tables = readSheetTables(files, path);
      const size = zipEntrySize(files, path);
      if (resolveXf && size > 0 && size <= STYLE_SCAN_LIMIT) {
        const xml = readZipText(files, path) ?? '';
        view = readSheetViewInfo(xml.slice(0, 256 * 1024));
        styleRefs = scanCellStyleRefs(xml);
      } else {
        const head = readZipText(files, path, 256 * 1024);
        if (head) view = readSheetViewInfo(head);
      }
    }

    return buildSheet(worksheet, { name, index, visibility, kind: sheetKind(worksheet) }, {
      date1904,
      table,
      resolveXf,
      resolveFill,
      styleRefs,
      view,
      tables,
    });
  });

  let activeSheetIndex = Math.min(parts?.activeTab ?? 0, sheets.length - 1);
  if (sheets[activeSheetIndex]?.visibility !== 'visible') {
    activeSheetIndex = Math.max(0, sheets.findIndex((s) => s.visibility === 'visible'));
  }

  const defaultFont = sheetJsStyles?.Fonts?.[0];
  return {
    ...file,
    format,
    properties: readProperties(workbook.Props),
    sheets,
    styles: table.styles,
    activeSheetIndex,
    date1904,
    warnings: [],
    ...(files && defaultFont ? { defaultFont: { name: defaultFont.name ?? 'Calibri', size: defaultFont.sz ?? 11 } } : {}),
  };
}

function sheetKind(worksheet: XLSX.WorkSheet | undefined): SheetKind {
  const type = (worksheet as { '!type'?: string } | undefined)?.['!type'];
  return type === 'chart' || type === 'macro' || type === 'dialog' ? type : 'worksheet';
}

type SheetJsCell = XLSX.CellObject & { F?: string; s?: object };

function buildSheet(
  worksheet: XLSX.WorkSheet | undefined,
  identity: { name: string; index: number; visibility: SheetVisibility; kind: SheetKind },
  ctx: SheetContext,
): WorksheetModel {
  const builder = new SheetBuilder(identity);
  const data = ((worksheet as { '!data'?: (SheetJsCell | undefined)[][] } | undefined)?.['!data'] ?? []) as (
    | (SheetJsCell | undefined)[]
    | undefined
  )[];
  const refs = ctx.styleRefs;
  let ref = 0;

  const rowLimit = Math.min(data.length, EXCEL_MAX_ROWS);
  for (let r = 0; r < rowLimit; r++) {
    const row = data[r];
    if (!row) continue;
    const colLimit = Math.min(row.length, EXCEL_MAX_COLS);
    for (let c = 0; c < colLimit; c++) {
      const cell = row[c];
      if (!cell) continue;

      let styleId = 0;
      if (refs && ctx.resolveXf) {
        while (ref < refs.count && (refs.rows[ref] < r || (refs.rows[ref] === r && refs.cols[ref] < c))) ref++;
        if (ref < refs.count && refs.rows[ref] === r && refs.cols[ref] === c) styleId = ctx.resolveXf(refs.xf[ref]);
      } else if (cell.s) {
        styleId = ctx.resolveFill(cell.s);
      }
      if (typeof cell.z === 'string' && cell.z !== 'General') styleId = ctx.table.withNumberFormat(styleId, cell.z);
      addCell(builder, r, c, cell, styleId, ctx);
    }
  }

  return builder.finish({ ...sheetLayout(worksheet, ctx.view), tables: ctx.tables });
}

/** JS Date → Excel serial date (only for sources that produce Date cells, e.g. HTML). */
function dateToSerial(date: Date, date1904: boolean): number {
  const serial = (date.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
  return date1904 ? serial - 1462 : serial;
}

function addCell(builder: SheetBuilder, r: number, c: number, cell: SheetJsCell, styleId: number, ctx: SheetContext): void {
  let kind: CellKind = CellKind.Empty;
  let num = 0;
  let text = '';
  const format = typeof cell.z === 'string' ? cell.z : undefined;

  switch (cell.t) {
    case 's': {
      text = cell.v == null ? '' : String(cell.v);
      if (text) kind = CellKind.String;
      break;
    }
    case 'n': {
      if (typeof cell.v === 'number' && Number.isFinite(cell.v)) {
        num = cell.v;
        kind = isDateFormat(format) ? CellKind.Date : CellKind.Number;
        text = formatNumber(num, format, ctx.date1904);
      }
      break;
    }
    case 'b':
      kind = CellKind.Boolean;
      num = cell.v ? 1 : 0;
      text = cell.v ? 'TRUE' : 'FALSE';
      break;
    case 'e':
      kind = CellKind.Error;
      num = typeof cell.v === 'number' ? cell.v : 0;
      text = (typeof cell.w === 'string' && cell.w) || ERROR_TEXT[num] || '#ERROR!';
      break;
    case 'd': {
      const date = cell.v instanceof Date ? cell.v : new Date(String(cell.v));
      if (!Number.isNaN(date.getTime())) {
        num = dateToSerial(date, ctx.date1904);
        kind = CellKind.Date;
        text = formatNumber(num, format ?? 'yyyy-mm-dd', ctx.date1904);
      }
      break;
    }
    default:
      break;
  }

  const formula = typeof cell.f === 'string' && cell.f ? (cell.F ? `{=${cell.f}}` : `=${cell.f}`) : undefined;
  if (kind === CellKind.Empty && !formula && !ctx.table.isVisibleWhenEmpty(styleId)) return;
  builder.add(r, c, kind, num, text, styleId, formula);
}

function sheetLayout(worksheet: XLSX.WorkSheet | undefined, view: SheetViewInfo | null) {
  const columnWidths = new Map<number, number>();
  const cols = worksheet?.['!cols'];
  cols?.forEach((col, c) => {
    if (!col || c >= EXCEL_MAX_COLS) return;
    if (col.hidden) {
      columnWidths.set(c, 0);
      return;
    }
    const info = col as XLSX.ColInfo & { width?: number };
    const px =
      typeof info.width === 'number'
        ? columnWidthToPx(info.width)
        : typeof info.wpx === 'number'
          ? info.wpx
          : typeof info.wch === 'number'
            ? Math.round(info.wch * 7 + 5)
            : undefined;
    if (px !== undefined && px > 0) columnWidths.set(c, Math.min(px, 2000));
  });

  const rowHeights = new Map<number, number>();
  const rows = worksheet?.['!rows'];
  rows?.forEach((row, r) => {
    if (!row || r >= EXCEL_MAX_ROWS) return;
    if (row.hidden) {
      rowHeights.set(r, 0);
      return;
    }
    const px = typeof row.hpt === 'number' ? pointsToPx(row.hpt) : typeof row.hpx === 'number' ? row.hpx : undefined;
    if (px !== undefined && px > 0) rowHeights.set(r, Math.min(px, 1000));
  });

  const merges: CellRange[] = [];
  for (const merge of worksheet?.['!merges'] ?? []) {
    const range = {
      r0: merge.s.r,
      c0: merge.s.c,
      r1: Math.min(merge.e.r, EXCEL_MAX_ROWS - 1),
      c1: Math.min(merge.e.c, EXCEL_MAX_COLS - 1),
    };
    if (range.r1 > range.r0 || range.c1 > range.c0) merges.push(range);
  }

  return {
    merges,
    columnWidths,
    rowHeights,
    defaultColumnWidth: view?.defaultColumnWidth,
    defaultRowHeight: view?.defaultRowHeight,
    frozen: view ? { rows: view.frozenRows, columns: view.frozenColumns } : undefined,
    showGridlines: view?.showGridlines,
  };
}

function readProperties(props: XLSX.FullProperties | undefined): WorkbookProperties {
  if (!props) return {};
  const text = (value: unknown) => (typeof value === 'string' && value.trim() ? value.trim() : undefined);
  const date = (value: unknown) => {
    const parsed = value instanceof Date ? value : typeof value === 'string' ? new Date(value) : null;
    return parsed && !Number.isNaN(parsed.getTime()) ? parsed.toISOString() : undefined;
  };
  const extended = props as XLSX.FullProperties & { Application?: string };
  const out: WorkbookProperties = {
    title: text(props.Title),
    subject: text(props.Subject),
    author: text(props.Author),
    lastModifiedBy: text(props.LastAuthor),
    company: text(props.Company),
    application: text(extended.Application),
    created: date(props.CreatedDate),
    modified: date(props.ModifiedDate),
  };
  for (const key of Object.keys(out) as (keyof WorkbookProperties)[]) {
    if (out[key] === undefined) delete out[key];
  }
  return out;
}

/** Typed-array buffers in a workbook, for zero-copy transfer out of the worker. */
export function collectTransferables(workbook: WorkbookModel): ArrayBuffer[] {
  const buffers = new Set<ArrayBuffer>();
  for (const sheet of workbook.sheets) {
    const { cells } = sheet;
    for (const array of [
      cells.rowStart,
      cells.cols,
      cells.kinds,
      cells.numbers,
      cells.texts,
      cells.styles,
      sheet.formulas.cells,
      sheet.strings.offsets,
      sheet.columnWidths.index,
      sheet.columnWidths.size,
      sheet.rowHeights.index,
      sheet.rowHeights.size,
    ]) {
      buffers.add(array.buffer as ArrayBuffer);
    }
  }
  return [...buffers];
}
