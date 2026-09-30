import * as XLSX from 'xlsx';
import { stringAt } from '../../models/stringPool';
import {
  CellKind,
  DEFAULT_FONT,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  type CellRange,
  type CellStyle,
  type IndexMap,
  type TableInfo,
  type WorkbookModel,
  type WorksheetModel,
} from '../../models/workbook';
import { cellAddress } from '../../utils/cellAddress';
import { remapFileFormula, type AxisChange, type SheetChanges } from '../formula/references';
import { formulaAt, rowSpan } from '../workbook/sheetService';
import { decodeRange, readRelationships, readWorkbookParts, relsPathOf, type ZipFiles } from '../workbook/xlsxParts';
import { applyWrites, cellKey, readStoredCell, sameCell, type StoredCell } from './cells';
import { fromBase, IDENTITY, isIdentity, mapRange, toBase } from './indexMap';
import { sizeAt } from './sheetOps';
import { getAttr, setAttrs, StyleWriter } from './styleWriter';
import { escapeXml, formulaForFile, unescapeXml } from './xmlText';

/**
 * Writes a workbook as .xlsx. When it was opened from an .xlsx, the original
 * package is the starting point: parts are rewritten only where the workbook
 * changed, and everything the viewer doesn't model (themes, charts, images,
 * comments, conditional formats, macros…) is carried over, following inserted
 * or deleted rows and columns and renamed sheets. New workbooks start from a
 * minimal package.
 */

type CfbEntry = { name: string; content: Uint8Array | number[]; size?: number };
type CfbContainer = { FullPaths: string[]; FileIndex: CfbEntry[] };

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORKSHEET_REL = `${REL_NS}/worksheet`;
const WORKSHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Top-level worksheet elements in schema order. */
const WORKSHEET_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
  'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
  'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];

const utf8 = new TextDecoder();
const encoder = new TextEncoder();

/** An .xlsx (OPC) package: named parts in a ZIP container. */
class Package {
  private readonly container: CfbContainer;

  private constructor(container: CfbContainer) {
    this.container = container;
  }

  static open(bytes: Uint8Array): Package {
    return new Package(XLSX.CFB.read(bytes, { type: 'array' }) as unknown as CfbContainer);
  }

  static create(parts: Record<string, string>): Package {
    const pkg = new Package(XLSX.CFB.utils.cfb_new() as unknown as CfbContainer);
    for (const [path, text] of Object.entries(parts)) pkg.write(path, text);
    return pkg;
  }

  paths(): string[] {
    const out: string[] = [];
    this.container.FullPaths.forEach((full, i) => {
      const entry = this.container.FileIndex[i];
      if (!entry?.content || full.endsWith('/')) return;
      const path = full.replace(/^Root Entry\//, '');
      if (!path.startsWith('\u0001')) out.push(path);
    });
    return out;
  }

  /** The part's actual path (part names are case-insensitive). */
  resolve(path: string): string | null {
    const lower = path.replace(/^\//, '').toLowerCase();
    return this.paths().find((p) => p.toLowerCase() === lower) ?? null;
  }

  private entry(path: string): CfbEntry | null {
    const actual = this.resolve(path);
    return actual ? (XLSX.CFB.find(this.container as never, `/${actual}`) as CfbEntry | null) : null;
  }

  read(path: string): string | null {
    const entry = this.entry(path);
    if (!entry?.content) return null;
    return utf8.decode(entry.content instanceof Uint8Array ? entry.content : Uint8Array.from(entry.content));
  }

  write(path: string, text: string): void {
    const bytes = encoder.encode(text);
    const entry = this.entry(path);
    if (entry) {
      entry.content = bytes;
      entry.size = bytes.length;
    } else {
      XLSX.CFB.utils.cfb_add(this.container as never, `/${path.replace(/^\//, '')}`, bytes);
    }
  }

  remove(path: string): void {
    const actual = this.resolve(path);
    if (actual) XLSX.CFB.utils.cfb_del(this.container as never, `/${actual}`);
  }

  files(): ZipFiles {
    const out: ZipFiles = {};
    this.container.FullPaths.forEach((full, i) => {
      const entry = this.container.FileIndex[i];
      if (entry?.content && !full.endsWith('/')) out[full.replace(/^Root Entry\//, '')] = { content: entry.content };
    });
    return out;
  }

  toBytes(): Uint8Array {
    return new Uint8Array(XLSX.CFB.write(this.container as never, { fileType: 'zip', type: 'array', compression: true }) as ArrayLike<number>);
  }
}

// Minimal package for new workbooks

let themeXml: string | null = null;

/** The standard Office theme (taken from a SheetJS-written workbook). */
function officeTheme(): string {
  if (themeXml === null) {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[]]), 'Sheet1');
    const bytes = new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer);
    themeXml = Package.open(bytes).read('xl/theme/theme1.xml') ?? '';
  }
  return themeXml;
}

function blankPackage(workbook: WorkbookModel): Package {
  const font = workbook.defaultFont ?? DEFAULT_FONT;
  const theme = officeTheme();
  const themeFont = font.name.toLowerCase() === 'calibri' && theme;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const parts: Record<string, string> = {
    '[Content_Types].xml':
      `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      (theme ? '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' : '') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>',
    '_rels/.rels':
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      `<Relationship Id="rId3" Type="${REL_NS}/extended-properties" Target="docProps/app.xml"/>` +
      '</Relationships>',
    'docProps/core.xml':
      `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
    'docProps/app.xml':
      `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Sheaf</Application></Properties>`,
    'xl/workbook.xml':
      `${XML_HEAD}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><workbookPr${workbook.date1904 ? ' date1904="1"' : ''}/>` +
      '<bookViews><workbookView activeTab="0"/></bookViews><sheets/><calcPr calcId="191029"/></workbook>',
    'xl/_rels/workbook.xml.rels':
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_NS}/styles" Target="styles.xml"/>` +
      (theme ? `<Relationship Id="rId2" Type="${REL_NS}/theme" Target="theme/theme1.xml"/>` : '') +
      '</Relationships>',
    'xl/styles.xml':
      `${XML_HEAD}<styleSheet xmlns="${MAIN_NS}">` +
      `<fonts count="1"><font><sz val="${font.size}"/>${themeFont ? '<color theme="1"/>' : ''}<name val="${escapeXml(font.name)}"/><family val="2"/>${themeFont ? '<scheme val="minor"/>' : ''}</font></fonts>` +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>',
  };
  if (theme) parts['xl/theme/theme1.xml'] = theme;
  return Package.create(parts);
}

function blankSheet(): string {
  return (
    `${XML_HEAD}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="A1"/>` +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
  );
}

// XML helpers

/** Raw attribute values of a start tag's attribute text, by name. */
function parseAttrs(text: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of text.matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)) out[m[1]] = m[2];
  return out;
}

function elementRegex(name: string, flags = ''): RegExp {
  return new RegExp(`<((?:[\\w.-]+:)?${name})(?=[\\s/>])[^>]*?(?:/>|>[\\s\\S]*?</\\1>)`, flags);
}

/** An element of the main SpreadsheetML namespace (written with the part's own prefix, if any). */
function mainRegex(prefix: string, name: string, flags = ''): RegExp {
  return new RegExp(`<(${prefix.replace(/\./g, '\\.')}${name})(?=[\\s/>])[^>]*?(?:/>|>[\\s\\S]*?</\\1>)`, flags);
}

function rangeRef(range: CellRange): string {
  const start = cellAddress(range.r0, range.c0);
  return range.r0 === range.r1 && range.c0 === range.c1 ? start : `${start}:${cellAddress(range.r1, range.c1)}`;
}

function columnOf(ref: string): number {
  let col = 0;
  for (const ch of ref.replace(/\$/g, '')) {
    const code = ch.toUpperCase().charCodeAt(0);
    if (code < 65 || code > 90) break;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

/** Replaces the first top-level worksheet element `name`, inserting it in schema order when missing (null removes it). */
function putElement(xml: string, prefix: string, name: string, element: string | null): string {
  const existing = mainRegex(prefix, name).exec(xml);
  if (existing) return xml.slice(0, existing.index) + (element ?? '') + xml.slice(existing.index + existing[0].length);
  if (element === null) return xml;
  const before = WORKSHEET_ORDER.slice(0, WORKSHEET_ORDER.indexOf(name)).reverse();
  for (const previous of before) {
    const matches = [...xml.matchAll(mainRegex(prefix, previous, 'g'))];
    const last = matches[matches.length - 1];
    if (last) {
      const at = last.index + last[0].length;
      return xml.slice(0, at) + element + xml.slice(at);
    }
  }
  return xml.replace(/(<(?:[\w.-]+:)?worksheet\b[^>]*>)/, (open) => open + element);
}

function numberText(value: number): string {
  return Number.isFinite(value) ? String(value) : '0';
}

/** The <c> element for a cell, or null when nothing needs storing. */
function cellXml(prefix: string, ref: string, cell: StoredCell | null, style: string | null): string | null {
  const s = style !== null && style !== '0' ? ` s="${style}"` : '';
  if (!cell || (cell.kind === CellKind.Empty && !cell.formula)) return s ? `<${prefix}c r="${ref}"${s}/>` : null;
  const f = cell.formula ? `<${prefix}f>${escapeXml(formulaForFile(cell.formula))}</${prefix}f>` : '';
  const v = (value: string) => `<${prefix}v>${value}</${prefix}v>`;
  switch (cell.kind) {
    case CellKind.Number:
    case CellKind.Date:
      return `<${prefix}c r="${ref}"${s}>${f}${v(numberText(cell.num))}</${prefix}c>`;
    case CellKind.Boolean:
      return `<${prefix}c r="${ref}"${s} t="b">${f}${v(cell.num ? '1' : '0')}</${prefix}c>`;
    case CellKind.Error:
      return `<${prefix}c r="${ref}"${s} t="e">${f}${v(escapeXml(cell.text))}</${prefix}c>`;
    case CellKind.String:
      return cell.formula
        ? `<${prefix}c r="${ref}"${s} t="str">${f}${v(escapeXml(cell.text))}</${prefix}c>`
        : `<${prefix}c r="${ref}"${s} t="inlineStr"><${prefix}is><${prefix}t xml:space="preserve">${escapeXml(cell.text)}</${prefix}t></${prefix}is></${prefix}c>`;
    default:
      // A formula whose result is empty.
      return `<${prefix}c r="${ref}"${s} t="str">${f}${v('')}</${prefix}c>`;
  }
}

/** Column width in characters (as stored) for a width in pixels. */
function widthChars(px: number): string {
  return String(Math.round((px / 7) * 256) / 256);
}

function sameMerges(a: readonly CellRange[], b: readonly CellRange[]): boolean {
  return a === b || (a.length === b.length && a.every((m, i) => m.r0 === b[i].r0 && m.c0 === b[i].c0 && m.r1 === b[i].r1 && m.c1 === b[i].c1));
}

function sameSizes(a: WorksheetModel['rowHeights'], b: WorksheetModel['rowHeights']): boolean {
  return a === b || (a.index.length === b.index.length && a.index.every((v, i) => v === b.index[i] && a.size[i] === b.size[i]));
}

// Worksheets

interface SheetInput {
  xml: string;
  /** The sheet as opened (null for a new sheet). */
  base: WorksheetModel | null;
  sheet: WorksheetModel;
  rows: IndexMap;
  cols: IndexMap;
  styles: readonly CellStyle[];
  styleWriter: StyleWriter;
  active: boolean;
  /** Rewrites formulas stored in the sheet XML (conditional formats, validations), or null when none need it. */
  fileFormula: ((formula: string) => string) | null;
}

/** Maps a range from file positions to current ones, or null when it was deleted. */
function mapCells(range: CellRange, rows: IndexMap, cols: IndexMap): CellRange | null {
  const r = mapRange(rows, range.r0, range.r1, EXCEL_MAX_ROWS);
  const c = mapRange(cols, range.c0, range.c1, EXCEL_MAX_COLS);
  return r && c ? { r0: r[0], r1: r[1], c0: c[0], c1: c[1] } : null;
}

/** A space-separated list of ranges (sqref), mapped; null when every range was deleted. */
function mapSqref(sqref: string, rows: IndexMap, cols: IndexMap): string | null {
  const out: string[] = [];
  for (const ref of sqref.trim().split(/\s+/)) {
    const range = decodeRange(ref);
    if (!range) {
      out.push(ref);
      continue;
    }
    const mapped = mapCells(range, rows, cols);
    if (mapped) out.push(rangeRef(mapped));
  }
  return out.length > 0 ? out.join(' ') : null;
}

function rowUnchanged(sheet: WorksheetModel, r: number, base: WorksheetModel, br: number): boolean {
  const [s0, s1] = rowSpan(sheet, r);
  const [b0, b1] = rowSpan(base, br);
  if (s1 - s0 !== b1 - b0) return false;
  if (sheet.cells === base.cells && sheet.strings === base.strings && sheet.formulas === base.formulas && r === br) return true;
  for (let k = 0; k < s1 - s0; k++) {
    const i = s0 + k;
    const j = b0 + k;
    const a = sheet.cells;
    const b = base.cells;
    if (a.cols[i] !== b.cols[j] || a.kinds[i] !== b.kinds[j] || a.styles[i] !== b.styles[j]) return false;
    if (!Object.is(a.numbers[i], b.numbers[j])) return false;
    if (stringAt(sheet.strings, a.texts[i]) !== stringAt(base.strings, b.texts[j])) return false;
    if (formulaAt(sheet, i) !== formulaAt(base, j)) return false;
  }
  return true;
}

const ROW_PATTERN = /<((?:[\w.-]+:)?row)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)/g;
const CELL_PATTERN = /<((?:[\w.-]+:)?c)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1>)/g;

/** Rewrites <sheetData>: rows and cells that didn't change are copied as they are. */
function writeSheetData(input: SheetInput): string {
  const { xml, base, sheet, rows, cols, styles, styleWriter } = input;
  const open = /<([\w.-]+:)?sheetData(\s[^>]*)?(\/?)>/.exec(xml);
  if (!open) return xml;
  const p = open[1] ?? '';
  const selfClosing = open[3] === '/';
  const dataStart = open.index + open[0].length;
  const closeTag = `</${p}sheetData>`;
  const dataEnd = selfClosing ? dataStart : xml.indexOf(closeTag, dataStart);
  const body = xml.slice(dataStart, dataEnd);
  const identity = isIdentity(rows) && isIdentity(cols);

  // Rows the model needs: those with cells or a custom height, in order.
  const modelRows: number[] = [];
  {
    const { rowStart } = sheet.cells;
    const heights = sheet.rowHeights.index;
    let h = 0;
    for (let r = 0; r < sheet.rowCount; r++) {
      while (h < heights.length && heights[h] < r) modelRows.push(heights[h++]);
      if (h < heights.length && heights[h] === r) h++;
      if (rowStart[r] < rowStart[r + 1] || (h > 0 && heights[h - 1] === r)) modelRows.push(r);
    }
    while (h < heights.length) modelRows.push(heights[h++]);
  }

  const out: string[] = [];
  const keepCell = (element: string, ref: string, current: StoredCell | null): string => {
    let next = getAttr(element, 'r') === ref ? element : setAttrs(element, { r: ref });
    const formula = /<((?:[\w.-]+:)?f)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1>)/.exec(next);
    if (formula && /\bt="shared"/.test(formula[2])) {
      // Shared formulas become plain ones (their group may be split or moved).
      const replacement = current?.formula ? `<${formula[1]}>${escapeXml(formulaForFile(current.formula))}</${formula[1]}>` : '';
      next = next.replace(formula[0], () => replacement);
    } else if (formula && !identity && /\bref="/.test(formula[2])) {
      const range = decodeRange(getAttr(formula[0], 'ref') ?? '');
      const mapped = range && mapCells(range, rows, cols);
      const replacement = mapped ? setAttrs(formula[0], { ref: rangeRef(mapped) }) : `<${formula[1]}>${escapeXml(formulaForFile(current?.formula ?? ''))}</${formula[1]}>`;
      next = next.replace(formula[0], () => replacement);
    }
    return next;
  };

  const buildCells = (r: number, br: number, inner: string | null): string => {
    const xmlCells = new Map<number, { xml: string; s: string | undefined }>();
    if (inner) {
      let col = -1;
      for (const m of inner.matchAll(CELL_PATTERN)) {
        const attrs = parseAttrs(m[2]);
        col = attrs.r ? columnOf(attrs.r) : col + 1;
        const current = fromBase(cols, col);
        if (current >= 0 && current < EXCEL_MAX_COLS) xmlCells.set(current, { xml: m[0], s: attrs.s });
      }
    }
    const columns = new Set<number>(xmlCells.keys());
    const [s0, s1] = rowSpan(sheet, r);
    for (let i = s0; i < s1; i++) columns.add(sheet.cells.cols[i]);
    const parts: string[] = [];
    for (const c of [...columns].sort((a, b) => a - b)) {
      const current = readStoredCell(sheet, r, c);
      const bc = toBase(cols, c);
      const before = base && br >= 0 && bc >= 0 ? readStoredCell(base, br, bc) : null;
      const kept = xmlCells.get(c);
      const ref = cellAddress(r, c);
      if (kept && sameCell(current, before)) {
        parts.push(keepCell(kept.xml, ref, current));
        continue;
      }
      if (!current) continue;
      let style: string | null;
      if (kept && (before ? current.style === before.style : current.style === 0)) style = kept.s ?? null;
      else style = String(styleWriter.xfFor(kept ? Number(kept.s ?? 0) : 0, before ? (styles[before.style] ?? {}) : {}, styles[current.style] ?? {}));
      const element = cellXml(p, ref, current, style);
      if (element) parts.push(element);
    }
    return parts.join('');
  };

  const emitRow = (r: number, baseRow: { br: number; attrs: string; inner: string | null; xml: string } | null) => {
    const br = baseRow?.br ?? toBase(rows, r);
    const height = sizeAt(sheet.rowHeights, r);
    const baseHeight = base && br >= 0 ? sizeAt(base.rowHeights, br) : undefined;
    const heightChanged = height !== baseHeight;
    if (baseRow && identity && !heightChanged && !/\bt="shared"/.test(baseRow.inner ?? '') && base && rowUnchanged(sheet, r, base, br)) {
      out.push(baseRow.xml);
      return;
    }
    let row = `<${p}row${baseRow ? baseRow.attrs.replace(/\s*\/$/, '') : ''}/>`;
    row = setAttrs(row, { r: String(r + 1), spans: null });
    if (heightChanged) {
      row = setAttrs(
        row,
        height === undefined
          ? { ht: null, customHeight: null, hidden: null }
          : height === 0
            ? { hidden: '1' }
            : { ht: String(Math.round(height * 0.75 * 100) / 100), customHeight: '1', hidden: null },
      );
    }
    const cells = buildCells(r, br, baseRow?.inner ?? null);
    const start = row.slice(0, -2);
    if (!cells && !/\s(?!r=|spans=)[\w:.-]+="/.test(start.replace(/^<[\w.:-]+/, ''))) return;
    out.push(cells ? `${start}>${cells}</${p}row>` : row);
  };

  let m = 0;
  let previous = -1;
  for (const match of body.matchAll(ROW_PATTERN)) {
    const attrs = match[2];
    const rAttr = parseAttrs(attrs).r;
    const br = rAttr ? Number(rAttr) - 1 : previous + 1;
    previous = br;
    const r = fromBase(rows, br);
    if (r < 0 || r >= EXCEL_MAX_ROWS) continue;
    while (m < modelRows.length && modelRows[m] < r) emitRow(modelRows[m++], null);
    if (m < modelRows.length && modelRows[m] === r) m++;
    emitRow(r, { br, attrs, inner: match[4] ?? null, xml: match[0] });
  }
  while (m < modelRows.length) emitRow(modelRows[m++], null);

  const sheetData = `<${p}sheetData${open[2] ?? ''}>${out.join('')}</${p}sheetData>`;
  return xml.slice(0, open.index) + sheetData + xml.slice(selfClosing ? dataStart : dataEnd + closeTag.length);
}

function writeCols(input: SheetInput, xml: string, prefix: string): string {
  const { base, sheet, cols } = input;
  if (isIdentity(cols) && base && sameSizes(sheet.columnWidths, base.columnWidths)) return xml;
  const existing = mainRegex(prefix, 'cols').exec(xml);
  const byColumn = new Map<number, Record<string, string>>();
  for (const m of (existing?.[0] ?? '').matchAll(/<(?:[\w.-]+:)?col\b([^>]*?)\/?>/g)) {
    const attrs = parseAttrs(m[1]);
    const min = Number(attrs.min) - 1;
    const max = Math.min(Number(attrs.max) - 1, EXCEL_MAX_COLS - 1);
    delete attrs.min;
    delete attrs.max;
    for (let c = Math.max(0, min); c <= max; c++) {
      const current = fromBase(cols, c);
      if (current >= 0 && current < EXCEL_MAX_COLS) byColumn.set(current, attrs);
    }
  }
  const touched = new Set<number>(sheet.columnWidths.index);
  if (base) base.columnWidths.index.forEach((c) => touched.add(fromBase(cols, c)));
  for (const c of touched) {
    if (c < 0) continue;
    const width = sizeAt(sheet.columnWidths, c);
    const bc = toBase(cols, c);
    const was = base && bc >= 0 ? sizeAt(base.columnWidths, bc) : undefined;
    if (width === was) continue;
    const attrs = { ...(byColumn.get(c) ?? {}) };
    if (width === undefined) {
      delete attrs.width;
      delete attrs.customWidth;
      delete attrs.hidden;
    } else if (width === 0) {
      attrs.hidden = '1';
    } else {
      attrs.width = widthChars(width);
      attrs.customWidth = '1';
      delete attrs.hidden;
    }
    if (Object.keys(attrs).length > 0) byColumn.set(c, attrs);
    else byColumn.delete(c);
  }
  const columns = [...byColumn.keys()].sort((a, b) => a - b);
  const elements: string[] = [];
  for (let k = 0; k < columns.length; ) {
    const first = columns[k];
    const attrs = byColumn.get(first)!;
    const key = JSON.stringify(attrs);
    let last = first;
    while (k + 1 < columns.length && columns[k + 1] === last + 1 && JSON.stringify(byColumn.get(columns[k + 1])) === key) last = columns[++k];
    k++;
    const withWidth = attrs.width ? attrs : { width: widthChars(sheet.defaultColumnWidth), ...attrs };
    const text = Object.entries(withWidth).map(([name, value]) => ` ${name}="${value}"`).join('');
    elements.push(`<${prefix}col min="${first + 1}" max="${last + 1}"${text}/>`);
  }
  return putElement(xml, prefix, 'cols', elements.length > 0 ? `<${prefix}cols>${elements.join('')}</${prefix}cols>` : null);
}

function writeSheetView(input: SheetInput, xml: string, prefix: string): string {
  const { base, sheet } = input;
  const frozenChanged = !base || base.frozen.rows !== sheet.frozen.rows || base.frozen.columns !== sheet.frozen.columns;
  const gridlinesChanged = !base || base.showGridlines !== sheet.showGridlines;
  const moved = !isIdentity(input.rows) || !isIdentity(input.cols);
  let view = mainRegex(prefix, 'sheetView').exec(xml);
  if (!view) {
    if (!frozenChanged && !gridlinesChanged && !input.active) return xml;
    xml = putElement(xml, prefix, 'sheetViews', `<${prefix}sheetViews><${prefix}sheetView workbookViewId="0"/></${prefix}sheetViews>`);
    view = mainRegex(prefix, 'sheetView').exec(xml)!;
  }
  let element = view[0];
  const tag = view[1];
  element = setAttrs(element, { tabSelected: input.active ? '1' : null });
  if (gridlinesChanged) element = setAttrs(element, { showGridLines: sheet.showGridlines ? null : '0' });
  if (moved) element = setAttrs(element, { topLeftCell: null });
  if (frozenChanged || moved) {
    const children = /^<[^>]*[^/]>([\s\S]*)<\/[^>]+>$/.exec(element)?.[1] ?? '';
    let rest = children.replace(mainRegex(prefix, 'pane', 'g'), '').replace(mainRegex(prefix, 'selection', 'g'), '');
    const { rows, columns } = sheet.frozen;
    let pane = '';
    if (rows > 0 || columns > 0) {
      const topLeft = cellAddress(rows, columns);
      const active = rows > 0 && columns > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
      pane =
        `<${prefix}pane${columns > 0 ? ` xSplit="${columns}"` : ''}${rows > 0 ? ` ySplit="${rows}"` : ''} topLeftCell="${topLeft}" activePane="${active}" state="frozen"/>` +
        `<${prefix}selection pane="${active}" activeCell="${topLeft}" sqref="${topLeft}"/>`;
    }
    rest = pane + rest;
    const start = element.slice(0, element.search(/\/?>/));
    element = rest ? `${start}>${rest}</${tag}>` : `${start}/>`;
  }
  return xml.slice(0, view.index) + element + xml.slice(view.index + view[0].length);
}

/** Moves ranges and formulas stored with the sheet (conditional formats, validations, links…). */
function writeSheetRanges(input: SheetInput, xml: string, prefix: string): string {
  const { rows, cols, fileFormula } = input;
  const moved = !isIdentity(rows) || !isIdentity(cols);
  if (!moved && !fileFormula) return xml;
  const sqref = (text: string) => (moved ? mapSqref(text, rows, cols) : text);
  const formulas = (element: string, names: string) =>
    fileFormula
      ? element.replace(new RegExp(`<((?:[\\w.-]+:)?(?:${names}))>([^<]*)</\\1>`, 'g'), (_, tag: string, text: string) => `<${tag}>${escapeXml(fileFormula(unescapeXml(text)))}</${tag}>`)
      : element;
  const listed = (block: string, item: string) => [...block.replace(/^<[^>]*>/, '').matchAll(mainRegex(prefix, item, 'g'))].length;

  xml = xml.replace(mainRegex(prefix, 'conditionalFormatting', 'g'), (element) => {
    const ref = sqref(getAttr(element, 'sqref') ?? '');
    return ref === null ? '' : formulas(setAttrs(element, { sqref: ref }), 'formula');
  });
  xml = xml.replace(mainRegex(prefix, 'dataValidations'), (block) => {
    const next = block.replace(mainRegex(prefix, 'dataValidation', 'g'), (element) => {
      const ref = sqref(getAttr(element, 'sqref') ?? '');
      return ref === null ? '' : formulas(setAttrs(element, { sqref: ref }), 'formula1|formula2');
    });
    const count = listed(next, 'dataValidation');
    return count > 0 ? setAttrs(next, { count: String(count) }) : '';
  });
  if (moved) {
    xml = xml.replace(mainRegex(prefix, 'hyperlinks'), (block) => {
      const next = block.replace(mainRegex(prefix, 'hyperlink', 'g'), (element) => {
        const ref = sqref(getAttr(element, 'ref') ?? '');
        return ref === null ? '' : setAttrs(element, { ref });
      });
      return listed(next, 'hyperlink') > 0 ? next : '';
    });
    for (const [name, attr] of [['autoFilter', 'ref'], ['ignoredError', 'sqref'], ['protectedRange', 'sqref']] as const) {
      xml = xml.replace(mainRegex(prefix, name, 'g'), (element) => {
        const ref = sqref(getAttr(element, attr) ?? '');
        return ref === null ? '' : setAttrs(element, { [attr]: ref });
      });
    }
  }
  // Excel 2010+ extensions (x14:…) keep their ranges in <xm:sqref> and formulas in <xm:f>.
  xml = xml.replace(/<([\w.-]+:(?:conditionalFormatting|dataValidation))(?=[\s>])[^>]*>[\s\S]*?<\/\1>/g, (element) => {
    let dropped = false;
    let next = element.replace(/<((?:[\w.-]+:)?sqref)>([^<]*)<\/\1>/g, (_, tag: string, text: string) => {
      const ref = sqref(text);
      if (ref === null) dropped = true;
      return `<${tag}>${ref ?? ''}</${tag}>`;
    });
    if (fileFormula) next = next.replace(/<((?:[\w.-]+:)?f)>([^<]*)<\/\1>/g, (_, tag: string, text: string) => `<${tag}>${escapeXml(fileFormula(unescapeXml(text)))}</${tag}>`);
    return dropped ? '' : next;
  });
  return xml;
}

function writeSheet(input: SheetInput): string {
  const { base, sheet } = input;
  const prefix = /<([\w.-]+:)?worksheet\b/.exec(input.xml)?.[1] ?? '';
  const moved = !isIdentity(input.rows) || !isIdentity(input.cols);
  const cellsChanged =
    !base || moved || sheet.cells !== base.cells || sheet.strings !== base.strings || sheet.formulas !== base.formulas || !sameSizes(sheet.rowHeights, base.rowHeights);
  let xml = input.xml;
  if (cellsChanged) {
    xml = writeSheetData(input);
    const dimension = sheet.rowCount > 0 && sheet.columnCount > 0 ? `A1:${cellAddress(sheet.rowCount - 1, sheet.columnCount - 1)}` : 'A1';
    xml = putElement(xml, prefix, 'dimension', `<${prefix}dimension ref="${dimension}"/>`);
  }
  xml = writeCols(input, xml, prefix);
  if (!base || moved || !sameMerges(sheet.merges, base.merges)) {
    const merges = sheet.merges.map((m) => `<${prefix}mergeCell ref="${rangeRef(m)}"/>`);
    xml = putElement(xml, prefix, 'mergeCells', merges.length > 0 ? `<${prefix}mergeCells count="${merges.length}">${merges.join('')}</${prefix}mergeCells>` : null);
  }
  xml = writeSheetView(input, xml, prefix);
  return writeSheetRanges(input, xml, prefix);
}

/** A copied sheet can't share the original's drawings, comments or tables. */
function withoutRelationships(xml: string): string {
  const prefix = /<([\w.-]+:)?worksheet\b/.exec(xml)?.[1] ?? '';
  for (const name of ['drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'tableParts']) {
    xml = xml.replace(mainRegex(prefix, name, 'g'), '');
  }
  xml = xml.replace(mainRegex(prefix, 'hyperlink', 'g'), (element) => (/\s[\w.-]+:id="/.test(element) ? '' : element));
  return xml.replace(mainRegex(prefix, 'hyperlinks'), (block) => (mainRegex(prefix, 'hyperlink').test(block.replace(/^<[^>]*>/, '')) ? block : ''));
}

// Parts that follow a sheet's rows and columns

/** Index after a deletion: the nearest surviving index (the next one for a start, the previous one for an end). */
function nearest(map: IndexMap, index: number, limit: number, start: boolean): number {
  const direct = fromBase(map, index);
  if (direct >= 0 && direct < limit) return direct;
  const span = start ? mapRange(map, index, limit - 1, limit) : mapRange(map, 0, index, limit);
  return span ? (start ? span[0] : span[1]) : 0;
}

function remapDrawing(xml: string, rows: IndexMap, cols: IndexMap): string {
  return xml.replace(/<((?:[\w.-]+:)?(from|to))>([\s\S]*?)<\/\1>/g, (_, tag: string, which: string, inner: string) => {
    const start = which === 'from';
    const next = inner
      .replace(/<((?:[\w.-]+:)?col)>(\d+)<\/\1>/, (_m, t: string, v: string) => `<${t}>${nearest(cols, Number(v), EXCEL_MAX_COLS, start)}</${t}>`)
      .replace(/<((?:[\w.-]+:)?row)>(\d+)<\/\1>/, (_m, t: string, v: string) => `<${t}>${nearest(rows, Number(v), EXCEL_MAX_ROWS, start)}</${t}>`);
    return `<${tag}>${next}</${tag}>`;
  });
}

function remapCellRefs(xml: string, element: string, rows: IndexMap, cols: IndexMap): string {
  return xml.replace(elementRegex(element, 'g'), (el) => {
    const range = decodeRange(getAttr(el, 'ref') ?? '');
    if (!range) return el;
    const r = fromBase(rows, range.r0);
    const c = fromBase(cols, range.c0);
    return r < 0 || c < 0 || r >= EXCEL_MAX_ROWS || c >= EXCEL_MAX_COLS ? '' : setAttrs(el, { ref: cellAddress(r, c) });
  });
}

function remapVml(xml: string, rows: IndexMap, cols: IndexMap): string {
  return xml
    .replace(/<((?:[\w.-]+:)?Row)>(\d+)<\/\1>/g, (m, tag: string, v: string) => {
      const r = fromBase(rows, Number(v));
      return r >= 0 ? `<${tag}>${r}</${tag}>` : m;
    })
    .replace(/<((?:[\w.-]+:)?Column)>(\d+)<\/\1>/g, (m, tag: string, v: string) => {
      const c = fromBase(cols, Number(v));
      return c >= 0 ? `<${tag}>${c}</${tag}>` : m;
    });
}

/** Table column names from the header cells: unique, and never empty. */
function tableColumnNames(sheet: WorksheetModel, table: TableInfo): string[] {
  const taken = new Set<string>();
  return table.columns.map((name, k) => {
    let text = name;
    if (table.headerRow) {
      const cell = readStoredCell(sheet, table.range.r0, table.range.c0 + k);
      text = cell && cell.kind !== CellKind.Empty ? cell.text : '';
    }
    let candidate = text.trim() || `Column${k + 1}`;
    for (let n = 2; taken.has(candidate.toLowerCase()); n++) candidate = `${text.trim() || 'Column'}${n}`;
    taken.add(candidate.toLowerCase());
    return candidate;
  });
}

function writeTable(xml: string, sheet: WorksheetModel, table: TableInfo, rows: IndexMap, cols: IndexMap): string {
  const tableTag = /<((?:[\w.-]+:)?table)\b[^>]*>/.exec(xml);
  if (!tableTag) return xml;
  const totals = Number(getAttr(tableTag[0], 'totalsRowCount') ?? 0);
  const ref = rangeRef(table.range);
  xml = xml.replace(tableTag[0], () => setAttrs(tableTag[0], { ref }));
  const filterRange = { ...table.range, r1: table.range.r1 - (totals > 0 ? totals : 0) };
  xml = xml.replace(elementRegex('autoFilter'), (element) => setAttrs(element, { ref: rangeRef(filterRange) }));
  xml = xml.replace(elementRegex('sortState'), (element) => {
    const range = decodeRange(getAttr(element, 'ref') ?? '');
    const mapped = range && mapCells(range, rows, cols);
    return mapped ? setAttrs(element, { ref: rangeRef(mapped) }) : '';
  });
  const names = tableColumnNames(sheet, table);
  xml = xml.replace(elementRegex('tableColumns'), (block) => {
    const p = /^<([\w.-]+:)?/.exec(block)?.[1] ?? '';
    const existing = [...block.matchAll(elementRegex('tableColumn', 'g'))].map((m) => m[0]);
    let maxId = existing.reduce((max, el) => Math.max(max, Number(getAttr(el, 'id') ?? 0)), 0);
    const byName = new Map(existing.map((el) => [unescapeXml(getAttr(el, 'name') ?? '').toLowerCase(), el]));
    const oldNames = table.columns.map((n) => n.toLowerCase());
    const columns = names.map((name, k) => {
      const previous = byName.get(oldNames[k]) ?? byName.get(name.toLowerCase());
      return previous ? setAttrs(previous, { name: escapeXml(name) }) : `<${p}tableColumn id="${++maxId}" name="${escapeXml(name)}"/>`;
    });
    return `<${p}tableColumns count="${columns.length}">${columns.join('')}</${p}tableColumns>`;
  });
  return xml;
}

// Workbook

export interface WriteXlsxInput {
  /** The .xlsx/.xlsm the workbook was opened from, or null to start from an empty package. */
  original: Uint8Array | null;
  /** The workbook as opened from `original`. */
  baseline: WorkbookModel | null;
  workbook: WorkbookModel;
}

export function writeXlsx({ original, baseline, workbook }: WriteXlsxInput): Uint8Array {
  const pkg = original ? Package.open(original) : blankPackage(workbook);
  const parts = readWorkbookParts(pkg.files());
  if (!parts) throw new Error('Workbook part not found');
  const workbookDir = parts.workbookPath.includes('/') ? parts.workbookPath.slice(0, parts.workbookPath.lastIndexOf('/')) : '';
  const baseSheets = (original && baseline?.sheets) || [];
  const styles = workbook.styles;

  // Where each sheet comes from in the package.
  const sources = workbook.sheets.map((sheet) => {
    const origin = sheet.origin;
    if (!original || !origin || origin.base === null || !parts.sheets[origin.base]?.path) return null;
    return { base: origin.base, copy: origin.copy, path: parts.sheets[origin.base].path!, rows: origin.rows, cols: origin.cols };
  });
  const kept = new Map<number, number>(); // file sheet index → current index
  sources.forEach((source, index) => {
    if (source && !source.copy && !kept.has(source.base)) kept.set(source.base, index);
  });

  // Sheet names and positions as the file knew them → now.
  const fileNames = parts.sheets.map((s) => s.name);
  const byFileName = new Map(fileNames.map((name, i) => [name.toLowerCase(), i]));
  const changes: SheetChanges = {
    currentName(name) {
      const index = byFileName.get(name.toLowerCase());
      if (index === undefined) return undefined;
      const current = kept.get(index);
      return current === undefined ? null : workbook.sheets[current].name;
    },
    axes(name) {
      const index = byFileName.get(name.toLowerCase());
      const current = index === undefined ? undefined : kept.get(index);
      const source = current === undefined ? null : sources[current];
      if (!source || (isIdentity(source.rows) && isIdentity(source.cols))) return null;
      return { rows: source.rows, cols: source.cols } satisfies AxisChange;
    },
  };
  const renamedOrDeleted = fileNames.some((name, i) => {
    const current = kept.get(i);
    return current === undefined || workbook.sheets[current].name !== name;
  });
  const anyMoved = sources.some((s) => s && !s.copy && (!isIdentity(s.rows) || !isIdentity(s.cols)));
  const reordered = workbook.sheets.length !== fileNames.length || sources.some((s, i) => !s || s.copy || s.base !== i);

  const stylesPath =
    readRelationships(pkg.files(), parts.workbookPath).find((rel) => /\/styles$/.test(rel.type))?.target ??
    pkg.paths().find((p) => /(^|\/)styles\.xml$/i.test(p)) ??
    null;
  const styleWriter = new StyleWriter((stylesPath && pkg.read(stylesPath)) || '<styleSheet/>', workbook.defaultFont ?? DEFAULT_FONT);

  const active = Math.max(0, Math.min(workbook.activeSheetIndex, workbook.sheets.length - 1));
  const usedPaths = new Set<string>();
  const newPath = () => {
    for (let n = 1; ; n++) {
      const path = `${workbookDir ? `${workbookDir}/` : ''}worksheets/sheet${n}.xml`;
      if (!usedPaths.has(path.toLowerCase()) && !pkg.resolve(path)) return path;
    }
  };
  for (const source of sources) if (source && !source.copy) usedPaths.add(source.path.toLowerCase());

  const outputs = workbook.sheets.map((sheet, index) => {
    const source = sources[index];
    const ownFile = source && !source.copy && kept.get(source.base) === index;
    let path: string;
    let xml: string;
    let base: WorksheetModel | null = null;
    if (source && ownFile) {
      path = source.path;
      xml = pkg.read(path) ?? blankSheet();
      base = baseSheets[source.base] ?? null;
    } else {
      path = newPath();
      usedPaths.add(path.toLowerCase());
      xml = source ? withoutRelationships(pkg.read(source.path) ?? blankSheet()) : blankSheet();
      base = source ? (baseSheets[source.base] ?? null) : null;
    }
    const isNew = !ownFile;
    if (sheet.kind !== 'worksheet') return { path, isNew, sheet };

    const rows = source?.rows ?? IDENTITY;
    const cols = source?.cols ?? IDENTITY;
    // Header cells of tables must match their column names.
    let model = sheet;
    for (const table of sheet.tables ?? []) {
      if (!table.headerRow) continue;
      const names = tableColumnNames(sheet, table);
      const writes = new Map<number, StoredCell | null>();
      names.forEach((name, k) => {
        const cell = readStoredCell(model, table.range.r0, table.range.c0 + k);
        if (!cell || cell.text !== name) writes.set(cellKey(table.range.r0, table.range.c0 + k), { kind: CellKind.String, num: 0, text: name, style: cell?.style ?? 0 });
      });
      model = applyWrites(model, writes);
    }
    const homeName = source ? fileNames[source.base] : null;
    const fileFormula =
      source && (renamedOrDeleted || anyMoved)
        ? (formula: string) =>
            remapFileFormula(formula, homeName, {
              currentName: changes.currentName,
              axes: (name) => (homeName && name.toLowerCase() === homeName.toLowerCase() ? (isIdentity(rows) && isIdentity(cols) ? null : { rows, cols }) : changes.axes(name)),
            })
        : null;
    const next = writeSheet({ xml, base, sheet: model, rows, cols, styles, styleWriter, active: index === active, fileFormula });
    if (next !== xml || isNew) pkg.write(path, next);

    // Tables, drawings and comments follow moved rows and columns.
    const moved = !isIdentity(rows) || !isIdentity(cols);
    const tablesChanged = Boolean(sheet.tables?.length) && (moved || !base || model.cells !== base.cells || model.strings !== base.strings);
    if (!isNew && (moved || tablesChanged)) {
      for (const rel of readRelationships(pkg.files(), path)) {
        if (rel.external) continue;
        const partXml = pkg.read(rel.target);
        if (partXml === null) continue;
        let updated = partXml;
        if (/\/table$/.test(rel.type)) {
          const name = unescapeXml(getAttr(/<(?:[\w.-]+:)?table\b[^>]*>/.exec(partXml)?.[0] ?? '', 'displayName') ?? '').toLowerCase();
          const table = sheet.tables?.find((t) => t.name.toLowerCase() === name);
          if (table) updated = writeTable(partXml, model, table, rows, cols);
          else if (base?.tables?.some((t) => t.name.toLowerCase() === name)) {
            removeTablePart(pkg, path, rel.id, rel.target);
            continue;
          }
        } else if (moved) {
          if (/\/drawing$/.test(rel.type)) updated = remapDrawing(partXml, rows, cols);
          else if (/\/comments$/.test(rel.type)) updated = remapCellRefs(partXml, 'comment', rows, cols);
          else if (/\/vmlDrawing$/.test(rel.type)) updated = remapVml(partXml, rows, cols);
          else if (/threadedComment$/.test(rel.type)) updated = remapCellRefs(partXml, 'threadedComment', rows, cols);
        }
        if (updated !== partXml) pkg.write(rel.target, updated);
      }
    }
    return { path, isNew, sheet };
  });

  // workbook.xml: the sheet list, defined names, active sheet.
  let workbookXml = pkg.read(parts.workbookPath)!;
  const relsPath = relsPathOf(parts.workbookPath);
  let relsXml = pkg.read(relsPath) ?? `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  let typesXml = pkg.read('[Content_Types].xml') ?? '';
  const sheetElements = [...workbookXml.matchAll(/<((?:[\w.-]+:)?sheet)\b[^>]*?\/>/g)].map((m) => m[0]);
  const relPrefix = /\sxmlns:([\w.-]+)="[^"]*\/relationships"/.exec(workbookXml)?.[1] ?? 'r';
  const sheetTag = /<((?:[\w.-]+:)?sheets)\b/.exec(workbookXml)?.[1] ?? 'sheets';
  const sheetPrefix = sheetTag.includes(':') ? sheetTag.slice(0, sheetTag.indexOf(':') + 1) : '';

  if (reordered || renamedOrDeleted || workbook.sheets.some((s, i) => s.visibility !== (baseSheets[sources[i]?.base ?? -1]?.visibility ?? 'visible'))) {
    let maxSheetId = sheetElements.reduce((max, el) => Math.max(max, Number(getAttr(el, 'sheetId') ?? 0)), 0);
    const relIds = [...relsXml.matchAll(/\sId="([^"]*)"/g)].map((m) => m[1]);
    let nextRel = relIds.reduce((max, id) => Math.max(max, Number(/^rId(\d+)$/.exec(id)?.[1] ?? 0)), 0);
    const elements = workbook.sheets.map((sheet, index) => {
      const source = sources[index];
      const state = sheet.visibility === 'visible' ? null : sheet.visibility;
      if (source && !outputs[index].isNew && sheetElements[source.base]) {
        return setAttrs(sheetElements[source.base], { name: escapeXml(sheet.name), state });
      }
      const id = `rId${++nextRel}`;
      const target = workbookDir ? outputs[index].path.slice(workbookDir.length + 1) : outputs[index].path;
      relsXml = relsXml.replace(/<\/((?:[\w.-]+:)?Relationships)>/, (close) => `<Relationship Id="${id}" Type="${WORKSHEET_REL}" Target="${target}"/>${close}`);
      typesXml = typesXml.replace(/<\/Types>/, `<Override PartName="/${outputs[index].path}" ContentType="${WORKSHEET_TYPE}"/></Types>`);
      return `<${sheetPrefix}sheet name="${escapeXml(sheet.name)}" sheetId="${++maxSheetId}"${state ? ` state="${state}"` : ''} ${relPrefix}:id="${id}"/>`;
    });
    workbookXml = workbookXml.replace(/<((?:[\w.-]+:)?sheets)\b[^>]*?(?:\/>|>[\s\S]*?<\/\1>)/, () => `<${sheetTag}>${elements.join('')}</${sheetTag}>`);

    // Sheets that were deleted.
    parts.sheets.forEach((fileSheet, index) => {
      if (kept.has(index) || !fileSheet.path) return;
      const relId = new RegExp(`\\s${relPrefix}:id="([^"]*)"`).exec(sheetElements[index] ?? '')?.[1];
      if (relId) relsXml = relsXml.replace(new RegExp(`<(?:[\\w.-]+:)?Relationship\\b[^>]*\\sId="${relId}"[^>]*/>`), '');
      const partName = `/${fileSheet.path}`.toLowerCase();
      typesXml = typesXml.replace(/<Override\b[^>]*\/>/g, (el) => ((getAttr(el, 'PartName') ?? '').toLowerCase() === partName ? '' : el));
      if (!outputs.some((o) => o.path.toLowerCase() === fileSheet.path!.toLowerCase())) {
        pkg.remove(fileSheet.path);
        pkg.remove(relsPathOf(fileSheet.path));
      }
    });

    // Defined names: sheet-scoped ones follow their sheet; formulas follow renames and moves.
    workbookXml = workbookXml.replace(/<((?:[\w.-]+:)?definedName)\b([^>]*)>([\s\S]*?)<\/\1>/g, (_element, tag: string, attrs: string, text: string) => {
      const local = parseAttrs(attrs).localSheetId;
      let home: string | null = null;
      let next = `<${tag}${attrs}>`;
      if (local !== undefined) {
        const current = kept.get(Number(local));
        if (current === undefined) return '';
        home = fileNames[Number(local)] ?? null;
        next = setAttrs(next, { localSheetId: String(current) });
      }
      return `${next}${escapeXml(remapFileFormula(unescapeXml(text), home, changes))}</${tag}>`;
    });
    workbookXml = workbookXml.replace(/<((?:[\w.-]+:)?definedNames)\b[^>]*>\s*<\/\1>/, '');
  } else if (anyMoved) {
    workbookXml = workbookXml.replace(/<((?:[\w.-]+:)?definedName)\b([^>]*)>([\s\S]*?)<\/\1>/g, (_, tag: string, attrs: string, text: string) => {
      const local = parseAttrs(attrs).localSheetId;
      const home = local !== undefined ? (fileNames[Number(local)] ?? null) : null;
      return `<${tag}${attrs}>${escapeXml(remapFileFormula(unescapeXml(text), home, changes))}</${tag}>`;
    });
  }

  // Excel opens on the sheet that was active.
  let activeTab = active;
  if (workbook.sheets[activeTab]?.visibility !== 'visible') activeTab = Math.max(0, workbook.sheets.findIndex((s) => s.visibility === 'visible'));
  if (/<(?:[\w.-]+:)?workbookView\b/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<((?:[\w.-]+:)?workbookView)\b[^>]*?\/?>/, (element) =>
      setAttrs(element, { activeTab: activeTab > 0 ? String(activeTab) : null, firstSheet: null }),
    );
  } else if (activeTab > 0) {
    workbookXml = workbookXml.replace(/<((?:[\w.-]+:)?)sheets\b/, (open, p: string) => `<${p}bookViews><${p}workbookView activeTab="${activeTab}"/></${p}bookViews>${open}`);
  }
  workbookXml = withFullCalcOnLoad(workbookXml);
  pkg.write(parts.workbookPath, workbookXml);
  pkg.write(relsPath, relsXml);
  if (typesXml) pkg.write('[Content_Types].xml', typesXml);

  // Charts and pivot tables refer to sheets by name and position.
  if (renamedOrDeleted || anyMoved) {
    for (const path of pkg.paths()) {
      if (/(^|\/)charts\/chart[^/]*\.xml$/i.test(path)) {
        const xml = pkg.read(path)!;
        const next = xml.replace(/<((?:[\w.-]+:)?f)>([^<]*)<\/\1>/g, (_, tag: string, text: string) => `<${tag}>${escapeXml(remapFileFormula(unescapeXml(text), null, changes))}</${tag}>`);
        if (next !== xml) pkg.write(path, next);
      } else if (/pivotCacheDefinition[^/]*\.xml$/i.test(path)) {
        const xml = pkg.read(path)!;
        const next = xml.replace(/<((?:[\w.-]+:)?worksheetSource)\b[^>]*?\/?>/g, (element) => {
          const name = unescapeXml(getAttr(element, 'sheet') ?? '');
          const current = name ? changes.currentName(name) : undefined;
          if (!current) return element;
          const axes = changes.axes(name);
          const range = axes ? decodeRange(getAttr(element, 'ref') ?? '') : null;
          const mapped = range && axes ? mapCells(range, axes.rows ?? IDENTITY, axes.cols ?? IDENTITY) : null;
          return setAttrs(element, { sheet: escapeXml(current), ...(mapped ? { ref: rangeRef(mapped) } : {}) });
        });
        if (next !== xml) pkg.write(path, next);
      }
    }
  }

  if (stylesPath) pkg.write(stylesPath, styleWriter.toXml());

  // Sheet titles in docProps/app.xml are optional; drop them rather than keep a stale list.
  if (reordered || renamedOrDeleted) {
    const app = pkg.read('docProps/app.xml');
    if (app) {
      const next = app.replace(elementRegex('HeadingPairs'), '').replace(elementRegex('TitlesOfParts'), '');
      if (next !== app) pkg.write('docProps/app.xml', next);
    }
  }

  // The calculation chain may be stale; Excel rebuilds it when it's missing.
  const calcChain = pkg.paths().find((k) => /(^|\/)calcChain\.xml$/i.test(k));
  if (calcChain) {
    pkg.remove(calcChain);
    const types = pkg.read('[Content_Types].xml');
    if (types) pkg.write('[Content_Types].xml', types.replace(/<Override[^>]*calcChain\.xml"[^>]*\/>/i, ''));
    const rels = pkg.read(relsPath);
    if (rels) pkg.write(relsPath, rels.replace(/<Relationship[^>]*calcChain\.xml"[^>]*\/>/i, ''));
  }

  return pkg.toBytes();
}

/** Removes a table the sheet no longer has (its rows or columns were deleted). */
function removeTablePart(pkg: Package, sheetPath: string, relId: string, tablePath: string): void {
  const sheetXml = pkg.read(sheetPath);
  if (sheetXml) {
    let next = sheetXml.replace(new RegExp(`<(?:[\\w.-]+:)?tablePart\\b[^>]*:id="${relId}"[^>]*/>`), '');
    next = next.replace(elementRegex('tableParts'), (block) => {
      const count = [...block.matchAll(elementRegex('tablePart', 'g'))].length;
      return count > 0 ? setAttrs(block, { count: String(count) }) : '';
    });
    pkg.write(sheetPath, next);
  }
  const relsPath = relsPathOf(sheetPath);
  const rels = pkg.read(relsPath);
  if (rels) pkg.write(relsPath, rels.replace(new RegExp(`<(?:[\\w.-]+:)?Relationship\\b[^>]*\\sId="${relId}"[^>]*/>`), ''));
  pkg.remove(tablePath);
  const types = pkg.read('[Content_Types].xml');
  if (types) {
    const partName = `/${tablePath}`.toLowerCase();
    pkg.write('[Content_Types].xml', types.replace(/<Override\b[^>]*\/>/g, (el) => ((getAttr(el, 'PartName') ?? '').toLowerCase() === partName ? '' : el)));
  }
}

/** Makes Excel recalculate on open (values we couldn't compute are refreshed there). */
function withFullCalcOnLoad(workbookXml: string): string {
  if (/<(?:[\w.-]+:)?calcPr\b/.test(workbookXml)) {
    return workbookXml.replace(/<((?:[\w.-]+:)?calcPr)\b[^>]*?\/?>/, (element) => setAttrs(element, { fullCalcOnLoad: '1' }));
  }
  const prefix = /<([\w.-]+:)?workbook\b/.exec(workbookXml)?.[1] ?? '';
  const calcPr = `<${prefix}calcPr fullCalcOnLoad="1"/>`;
  for (const anchor of ['definedNames', 'externalReferences', 'functionGroups', 'sheets']) {
    const close = new RegExp(`</(?:[\\w.-]+:)?${anchor}>`).exec(workbookXml);
    if (close) return workbookXml.slice(0, close.index + close[0].length) + calcPr + workbookXml.slice(close.index + close[0].length);
  }
  return workbookXml;
}
