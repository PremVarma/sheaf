import type { BorderEdge, CellStyle, HorizontalAlign, VerticalAlign } from '../../models/workbook';
import { INDEXED_COLORS, normalizeHex } from '../../utils/color';

/** Style ids are stored in a Uint16Array. */
export const MAX_STYLES = 0xffff;

const EDGES = ['top', 'right', 'bottom', 'left'] as const;

function canonicalEdge(edge: BorderEdge): BorderEdge {
  return edge.color ? { width: edge.width, style: edge.style, color: edge.color } : { width: edge.width, style: edge.style };
}

/**
 * The same style with properties in a fixed order and defaults left out, so
 * equal styles compare (and intern) equal however they were built.
 */
export function canonicalStyle(style: CellStyle): CellStyle {
  const out: CellStyle = {};
  if (style.bold) out.bold = true;
  if (style.italic) out.italic = true;
  if (style.underline) out.underline = true;
  if (style.strike) out.strike = true;
  if (style.fontName) out.fontName = style.fontName;
  if (style.color) out.color = style.color;
  if (style.fontScale !== undefined && style.fontScale !== 1) out.fontScale = style.fontScale;
  if (style.fill) out.fill = style.fill;
  if (style.hAlign) out.hAlign = style.hAlign;
  if (style.vAlign && style.vAlign !== 'bottom') out.vAlign = style.vAlign;
  if (style.wrap) out.wrap = true;
  if (style.indent) out.indent = style.indent;
  if (style.borders) {
    const borders: NonNullable<CellStyle['borders']> = {};
    for (const edge of EDGES) if (style.borders[edge]) borders[edge] = canonicalEdge(style.borders[edge]!);
    if (Object.keys(borders).length > 0) out.borders = borders;
  }
  if (style.numberFormat && style.numberFormat !== 'General') out.numberFormat = style.numberFormat;
  return out;
}

export function styleKey(style: CellStyle): string {
  return JSON.stringify(canonicalStyle(style));
}

/** Deduplicating table of cell styles. Index 0 is always the default style. */
export class StyleTable {
  readonly styles: CellStyle[] = [{}];
  private readonly ids = new Map<string, number>([['{}', 0]]);

  intern(style: CellStyle): number {
    const canonical = canonicalStyle(style);
    const key = JSON.stringify(canonical);
    let id = this.ids.get(key);
    if (id === undefined) {
      if (this.styles.length > MAX_STYLES) return 0; // Uint16 storage; never reached in practice
      id = this.styles.length;
      this.styles.push(canonical);
      this.ids.set(key, id);
    }
    return id;
  }

  private readonly withFormat = new Map<string, number>();

  /** The same style with a number format added (cached). */
  withNumberFormat(id: number, numberFormat: string): number {
    const key = `${id}|${numberFormat}`;
    let result = this.withFormat.get(key);
    if (result === undefined) {
      const base = this.styles[id] ?? {};
      result = base.numberFormat === numberFormat ? id : this.intern({ ...base, numberFormat });
      this.withFormat.set(key, result);
    }
    return result;
  }

  /** Whether an empty cell with this style is still worth drawing (fill or borders). */
  isVisibleWhenEmpty(id: number): boolean {
    const style = this.styles[id];
    return Boolean(style && (style.fill || style.borders));
  }
}

interface SheetJsColor {
  rgb?: string;
  indexed?: number;
  index?: number;
  theme?: number;
  auto?: boolean | number;
}
interface SheetJsFont {
  name?: string;
  bold?: number | boolean;
  italic?: number | boolean;
  underline?: number;
  strike?: number | boolean;
  sz?: number;
  color?: SheetJsColor;
}
interface SheetJsFill {
  patternType?: string;
  fgColor?: SheetJsColor;
  bgColor?: SheetJsColor;
}
interface SheetJsXf {
  fontId?: number;
  fillId?: number;
  borderId?: number;
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean; indent?: string | number };
}
export interface SheetJsStyles {
  Fonts?: SheetJsFont[];
  Fills?: SheetJsFill[];
  CellXf?: SheetJsXf[];
}

function colorOf(color: SheetJsColor | undefined): string | undefined {
  if (!color || color.auto) return undefined;
  const fromRgb = normalizeHex(color.rgb);
  if (fromRgb) return fromRgb;
  const indexed = color.indexed ?? color.index;
  // Indexes 64+ are "system foreground/background", i.e. automatic.
  return indexed !== undefined && indexed < 64 ? INDEXED_COLORS[indexed] : undefined;
}

/** Fill color from a SheetJS fill object (also used for .xls cells). */
export function fillColor(fill: SheetJsFill | undefined): string | undefined {
  if (!fill || !fill.patternType || fill.patternType === 'none') return undefined;
  return colorOf(fill.fgColor) ?? (fill.patternType === 'solid' ? colorOf(fill.bgColor) : undefined);
}

const H_ALIGN = new Set<HorizontalAlign>(['left', 'center', 'right', 'justify', 'fill', 'centerContinuous', 'distributed']);
const V_ALIGN = new Set<VerticalAlign>(['top', 'center', 'bottom']);

/**
 * Maps xlsx cellXfs indexes to interned style ids, caching per index.
 * Combines SheetJS's parsed fonts/fills/alignment with our own border parse.
 */
export function createXfResolver(
  styles: SheetJsStyles,
  borders: (CellStyle['borders'] | undefined)[],
  table: StyleTable,
): (xfIndex: number) => number {
  const cache = new Map<number, number>();
  const defaultSize = styles.Fonts?.[0]?.sz ?? 11;
  const defaultName = styles.Fonts?.[0]?.name?.toLowerCase();

  return (xfIndex) => {
    const cached = cache.get(xfIndex);
    if (cached !== undefined) return cached;
    const xf = styles.CellXf?.[xfIndex];
    let id = 0;
    if (xf) {
      const style: CellStyle = {};
      const font = styles.Fonts?.[xf.fontId ?? 0];
      if (font) {
        if (font.bold) style.bold = true;
        if (font.italic) style.italic = true;
        if (font.underline) style.underline = true;
        if (font.strike) style.strike = true;
        if (font.name && font.name.toLowerCase() !== defaultName) style.fontName = font.name;
        const color = colorOf(font.color);
        if (color && color !== '#000000') style.color = color;
        if (font.sz && Math.abs(font.sz - defaultSize) > 0.01) {
          style.fontScale = Math.round(Math.min(40, Math.max(0.1, font.sz / defaultSize)) * 10_000) / 10_000;
        }
      }
      const fill = fillColor(styles.Fills?.[xf.fillId ?? 0]);
      if (fill) style.fill = fill;

      const alignment = xf.alignment;
      if (alignment) {
        if (alignment.horizontal && H_ALIGN.has(alignment.horizontal as HorizontalAlign)) {
          style.hAlign = alignment.horizontal as HorizontalAlign;
        }
        if (alignment.vertical && V_ALIGN.has(alignment.vertical as VerticalAlign) && alignment.vertical !== 'bottom') {
          style.vAlign = alignment.vertical as VerticalAlign;
        }
        if (alignment.wrapText) style.wrap = true;
        const indent = Number(alignment.indent);
        if (indent > 0) style.indent = Math.min(indent, 15);
      }
      const border = borders[xf.borderId ?? 0];
      if (border) style.borders = border;
      id = table.intern(style);
    }
    cache.set(xfIndex, id);
    return id;
  };
}

/** Style ids for .xls cells, where SheetJS only reports fills. */
export function createFillResolver(table: StyleTable): (fill: SheetJsFill) => number {
  const cache = new WeakMap<object, number>();
  return (fill) => {
    const cached = cache.get(fill);
    if (cached !== undefined) return cached;
    const color = fillColor(fill);
    const id = color ? table.intern({ fill: color }) : 0;
    cache.set(fill, id);
    return id;
  };
}
