import type { BorderEdge, CellRange, CellStyle, TableInfo } from '../../models/workbook';
import { applyTint, INDEXED_COLORS, normalizeHex } from '../../utils/color';

/**
 * Reads details SheetJS Community Edition doesn't expose, directly from the
 * XML parts of an .xlsx package: freeze panes, gridline visibility, default
 * sizes, the active sheet, cell borders and each cell's style index.
 */

type ZipEntry = { content?: Uint8Array | ArrayLike<number> } | undefined;
export type ZipFiles = Record<string, ZipEntry>;

const utf8 = new TextDecoder('utf-8');

function findEntry(files: ZipFiles, path: string): ZipEntry {
  const normalized = path.replace(/^\/+/, '');
  if (files[normalized]) return files[normalized];
  const lower = normalized.toLowerCase();
  for (const key of Object.keys(files)) {
    if (key.replace(/^\/+/, '').toLowerCase() === lower) return files[key];
  }
  return undefined;
}

export function zipEntrySize(files: ZipFiles, path: string): number {
  return findEntry(files, path)?.content?.length ?? 0;
}

/** Decodes a part (or its first `maxBytes`) as UTF-8. */
export function readZipText(files: ZipFiles, path: string, maxBytes = Infinity): string | null {
  const content = findEntry(files, path)?.content;
  if (!content) return null;
  const bytes = content instanceof Uint8Array ? content : Uint8Array.from(content);
  return utf8.decode(bytes.length > maxBytes ? bytes.subarray(0, maxBytes) : bytes);
}

const ENTITY = /&(?:#x([0-9a-f]+)|#(\d+)|(amp|lt|gt|quot|apos));/gi;
const NAMED_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

function decodeEntities(text: string): string {
  return text.replace(ENTITY, (_, hex: string, dec: string, name: string) =>
    hex ? String.fromCodePoint(parseInt(hex, 16)) : dec ? String.fromCodePoint(Number(dec)) : NAMED_ENTITIES[name.toLowerCase()],
  );
}

const ATTRIBUTE = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Attributes of a start tag, keyed by local name (namespace prefixes dropped). */
function parseAttributes(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const match of tag.matchAll(ATTRIBUTE)) {
    const name = match[1];
    const local = name.includes(':') && !name.startsWith('xmlns') ? name.slice(name.indexOf(':') + 1) : name;
    out[local] = decodeEntities(match[2] ?? match[3] ?? '');
  }
  return out;
}

/** First start tag with the given local name, e.g. `<x:pane .../>`. */
function findTag(xml: string, localName: string): Record<string, string> | null {
  const match = new RegExp(`<(?:[\\w.-]+:)?${localName}(?=[\\s/>])[^>]*>`).exec(xml);
  return match ? parseAttributes(match[0]) : null;
}

function allTags(xml: string, localName: string): Record<string, string>[] {
  const pattern = new RegExp(`<(?:[\\w.-]+:)?${localName}(?=[\\s/>])[^>]*>`, 'g');
  return [...xml.matchAll(pattern)].map((m) => parseAttributes(m[0]));
}

export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const segments = (baseDir ? `${baseDir}/${target}` : target).split('/');
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === '..') out.pop();
    else if (segment && segment !== '.') out.push(segment);
  }
  return out.join('/');
}

export interface WorkbookParts {
  /** Worksheet part path per sheet, in workbook order. */
  sheets: { name: string; path: string | null }[];
  /** Index of the sheet that was active when the file was saved. */
  activeTab: number;
  workbookPath: string;
}

export function readWorkbookParts(files: ZipFiles): WorkbookParts | null {
  // The package relationships point at the workbook part (normally xl/workbook.xml).
  const rootRels = readZipText(files, '_rels/.rels');
  let workbookPath = 'xl/workbook.xml';
  if (rootRels) {
    const officeDoc = allTags(rootRels, 'Relationship').find((r) => /\/officeDocument$/.test(r.Type ?? ''));
    if (officeDoc?.Target) workbookPath = resolvePath('', officeDoc.Target);
  }
  const workbookXml = readZipText(files, workbookPath);
  if (!workbookXml) return null;

  const dir = workbookPath.includes('/') ? workbookPath.slice(0, workbookPath.lastIndexOf('/')) : '';
  const relsPath = `${dir ? `${dir}/` : ''}_rels/${workbookPath.slice(dir ? dir.length + 1 : 0)}.rels`;
  const rels = readZipText(files, relsPath) ?? '';
  const targets = new Map<string, string>();
  for (const rel of allTags(rels, 'Relationship')) {
    if (rel.Id && rel.Target) targets.set(rel.Id, resolvePath(dir, rel.Target));
  }

  const sheets = allTags(workbookXml, 'sheet').map((attrs) => ({
    name: attrs.name ?? '',
    path: attrs.id ? (targets.get(attrs.id) ?? null) : null,
  }));
  const view = findTag(workbookXml, 'workbookView');
  const activeTab = Number(view?.activeTab ?? 0);
  return { sheets, activeTab: Number.isInteger(activeTab) && activeTab >= 0 ? activeTab : 0, workbookPath };
}

export interface SheetViewInfo {
  frozenRows: number;
  frozenColumns: number;
  showGridlines: boolean;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
}

/** Excel column width (characters of the default font) → pixels, per ECMA-376 18.3.1.13. */
export function columnWidthToPx(width: number, maxDigitWidth = 7): number {
  return Math.floor(((256 * width + Math.floor(128 / maxDigitWidth)) / 256) * maxDigitWidth);
}

export function pointsToPx(points: number): number {
  return Math.round((points * 4) / 3);
}

/** Reads sheetView/pane/sheetFormatPr, which all precede <sheetData>. */
export function readSheetViewInfo(xmlHead: string): SheetViewInfo {
  const dataStart = xmlHead.search(/<(?:[\w.-]+:)?sheetData[\s/>]/);
  const head = dataStart >= 0 ? xmlHead.slice(0, dataStart) : xmlHead;
  const info: SheetViewInfo = { frozenRows: 0, frozenColumns: 0, showGridlines: true };

  const view = findTag(head, 'sheetView');
  if (view?.showGridLines === '0' || view?.showGridLines === 'false') info.showGridlines = false;

  const pane = findTag(head, 'pane');
  if (pane && (pane.state === 'frozen' || pane.state === 'frozenSplit')) {
    const rows = Math.floor(Number(pane.ySplit ?? 0));
    const cols = Math.floor(Number(pane.xSplit ?? 0));
    info.frozenRows = Number.isFinite(rows) && rows > 0 ? rows : 0;
    info.frozenColumns = Number.isFinite(cols) && cols > 0 ? cols : 0;
  }

  const format = findTag(head, 'sheetFormatPr');
  if (format) {
    const defaultWidth = Number(format.defaultColWidth);
    const baseWidth = Number(format.baseColWidth);
    if (defaultWidth > 0) info.defaultColumnWidth = columnWidthToPx(defaultWidth);
    else if (baseWidth > 0) info.defaultColumnWidth = Math.ceil((baseWidth * 7 + 5) / 8) * 8;
    const rowHeight = Number(format.defaultRowHeight);
    if (rowHeight > 0) info.defaultRowHeight = pointsToPx(rowHeight);
  }
  return info;
}

/** Cell style indexes in document (row-major) order, for cells with a non-default style. */
export interface CellStyleRefs {
  count: number;
  rows: Uint32Array;
  cols: Uint16Array;
  xf: Uint32Array;
}

const CELL_OR_ROW = /<(?:[A-Za-z_][\w.-]*:)?(row|c)(?=[\s/>])/g;

function attributeValue(attrs: string, name: string): string | null {
  let from = 0;
  for (;;) {
    const at = attrs.indexOf(name, from);
    if (at < 0) return null;
    from = at + name.length;
    const before = at === 0 ? 32 : attrs.charCodeAt(at - 1);
    if (before !== 32 && before !== 9 && before !== 10 && before !== 13) continue;
    let p = from;
    while (attrs.charCodeAt(p) === 32) p++;
    if (attrs.charCodeAt(p) !== 61 /* = */) continue;
    p++;
    while (attrs.charCodeAt(p) === 32) p++;
    const quote = attrs.charCodeAt(p);
    if (quote !== 34 && quote !== 39) return null;
    const end = attrs.indexOf(quote === 34 ? '"' : "'", p + 1);
    return end < 0 ? null : attrs.slice(p + 1, end);
  }
}

/** "AB12" → [11, 27] (row, col), zero-based. */
function decodeCellRef(ref: string): [number, number] | null {
  let col = 0;
  let i = 0;
  for (; i < ref.length; i++) {
    const code = ref.charCodeAt(i) | 0x20;
    if (code < 97 || code > 122) break;
    col = col * 26 + (code - 96);
  }
  if (i === 0 || i === ref.length) return null;
  const row = Number(ref.slice(i));
  return Number.isInteger(row) && row > 0 ? [row - 1, col - 1] : null;
}

export function scanCellStyleRefs(xml: string): CellStyleRefs {
  let capacity = 4096;
  let rows = new Uint32Array(capacity);
  let cols = new Uint16Array(capacity);
  let xf = new Uint32Array(capacity);
  let count = 0;
  let sorted = true;

  const start = xml.search(/<(?:[\w.-]+:)?sheetData[\s/>]/);
  if (start < 0) return { count: 0, rows, cols, xf };
  const end = xml.lastIndexOf('sheetData>');

  let row = -1;
  let col = -1;
  let lastKey = -1;
  CELL_OR_ROW.lastIndex = start;
  for (let m = CELL_OR_ROW.exec(xml); m; m = CELL_OR_ROW.exec(xml)) {
    if (end >= 0 && m.index > end) break;
    const attrStart = m.index + m[0].length;
    const close = xml.indexOf('>', attrStart);
    if (close < 0) break;
    const attrs = xml.slice(attrStart, close);
    CELL_OR_ROW.lastIndex = close + 1;

    if (m[1] === 'row') {
      const r = attributeValue(attrs, 'r');
      row = r !== null ? Number(r) - 1 : row + 1;
      col = -1;
      continue;
    }

    const ref = attributeValue(attrs, 'r');
    const decoded = ref !== null ? decodeCellRef(ref) : null;
    if (decoded) {
      row = decoded[0];
      col = decoded[1];
    } else {
      col++;
    }
    const s = attributeValue(attrs, 's');
    if (s === null || s === '0') continue;
    const styleIndex = Number(s);
    if (!Number.isInteger(styleIndex) || styleIndex < 0 || row < 0 || col < 0 || col > 0xffff) continue;

    if (count === capacity) {
      capacity *= 2;
      const r2 = new Uint32Array(capacity);
      r2.set(rows);
      rows = r2;
      const c2 = new Uint16Array(capacity);
      c2.set(cols);
      cols = c2;
      const x2 = new Uint32Array(capacity);
      x2.set(xf);
      xf = x2;
    }
    const key = row * 65536 + col;
    if (key <= lastKey) sorted = false;
    lastKey = key;
    rows[count] = row;
    cols[count] = col;
    xf[count] = styleIndex;
    count++;
  }

  if (!sorted) {
    const order = Array.from({ length: count }, (_, k) => k).sort((a, b) => rows[a] - rows[b] || cols[a] - cols[b]);
    rows = Uint32Array.from(order, (k) => rows[k]);
    cols = Uint16Array.from(order, (k) => cols[k]);
    xf = Uint32Array.from(order, (k) => xf[k]);
  }
  return { count, rows, cols, xf };
}

type BorderSet = NonNullable<CellStyle['borders']>;

const BORDER_WIDTHS: Record<string, BorderEdge['width']> = {
  thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, slantDashDot: 2,
  medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2,
  thick: 3, double: 3,
};

function borderStyle(style: string): BorderEdge['style'] {
  if (style === 'double') return 'double';
  if (style === 'dotted' || style === 'hair') return 'dotted';
  if (/dash/i.test(style)) return 'dashed';
  return 'solid';
}

/** Resolves a <color .../> element's attributes to #rrggbb. */
export function resolveColor(attrs: Record<string, string> | null, themeColors: readonly string[]): string | undefined {
  if (!attrs || attrs.auto === '1' || attrs.auto === 'true') return undefined;
  let color: string | undefined;
  if (attrs.rgb) color = normalizeHex(attrs.rgb);
  else if (attrs.indexed !== undefined) color = INDEXED_COLORS[Number(attrs.indexed)];
  else if (attrs.theme !== undefined) color = themeColors[Number(attrs.theme)];
  if (color && attrs.tint) color = applyTint(color, Number(attrs.tint));
  return color;
}

/** Parses <borders> from styles.xml. SheetJS CE records borders without their edges. */
export function readBorders(stylesXml: string, themeColors: readonly string[]): (BorderSet | undefined)[] {
  const block = /<(?:[\w.-]+:)?borders(?=[\s>])[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?borders>/.exec(stylesXml)?.[1];
  if (!block) return [];
  const result: (BorderSet | undefined)[] = [];
  const borderPattern = /<(?:[\w.-]+:)?border(?=[\s/>])[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?border>)/g;
  const edgePattern = /<(?:[\w.-]+:)?(left|right|top|bottom|start|end)(?=[\s/>])([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?\1>)/g;

  for (const border of block.matchAll(borderPattern)) {
    const inner = border[1] ?? '';
    let set: BorderSet | undefined;
    for (const edge of inner.matchAll(edgePattern)) {
      const style = parseAttributes(edge[2]).style;
      if (!style || style === 'none') continue;
      const colorTag = edge[3] ? findTag(edge[3], 'color') : null;
      const side = edge[1] === 'start' ? 'left' : edge[1] === 'end' ? 'right' : (edge[1] as keyof BorderSet);
      (set ??= {})[side] = {
        width: BORDER_WIDTHS[style] ?? 1,
        style: borderStyle(style),
        color: resolveColor(colorTag, themeColors),
      };
    }
    result.push(set);
  }
  return result;
}

/** The relationships part of a part: xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels. */
export function relsPathOf(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

export interface Relationship {
  id: string;
  type: string;
  /** Resolved package path (for internal targets). */
  target: string;
  external: boolean;
}

/** Relationships of a part, with targets resolved to package paths. */
export function readRelationships(files: ZipFiles, partPath: string): Relationship[] {
  const xml = readZipText(files, relsPathOf(partPath));
  if (!xml) return [];
  const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : '';
  return allTags(xml, 'Relationship').map((rel) => {
    const external = rel.TargetMode === 'External';
    return { id: rel.Id ?? '', type: rel.Type ?? '', target: external ? (rel.Target ?? '') : resolvePath(dir, rel.Target ?? ''), external };
  });
}

/** "B2:D10" → range (zero-based); null when it isn't a cell range. */
export function decodeRange(ref: string): CellRange | null {
  const [first, second = first] = ref.split(':');
  const a = decodeCellRef(first.replace(/\$/g, ''));
  const b = decodeCellRef(second.replace(/\$/g, ''));
  if (!a || !b) return null;
  return { r0: Math.min(a[0], b[0]), c0: Math.min(a[1], b[1]), r1: Math.max(a[0], b[0]), c1: Math.max(a[1], b[1]) };
}

/** Excel tables on a worksheet (from its table parts). */
export function readSheetTables(files: ZipFiles, sheetPath: string): TableInfo[] {
  const tables: TableInfo[] = [];
  for (const rel of readRelationships(files, sheetPath)) {
    if (rel.external || !/\/table$/.test(rel.type)) continue;
    const xml = readZipText(files, rel.target);
    const table = xml ? findTag(xml, 'table') : null;
    const range = table?.ref ? decodeRange(table.ref) : null;
    if (!xml || !table || !range) continue;
    tables.push({
      name: table.displayName ?? table.name ?? '',
      range,
      columns: allTags(xml, 'tableColumn').map((column) => column.name ?? ''),
      headerRow: table.headerRowCount !== '0',
    });
  }
  return tables;
}
