import type { BorderEdge, CellStyle } from '../../models/workbook';
import { canonicalStyle, styleKey } from '../workbook/styles';
import { escapeXml } from './xmlText';

/**
 * Adds cell formats (xf) to styles.xml for cells whose formatting changed.
 * Each new format starts from the cell's existing one and changes only what
 * the user changed, so everything the viewer doesn't model (theme fonts,
 * protection, rotation…) is kept.
 */

type Block = 'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs';
const ITEM: Record<Block, string> = { numFmts: 'numFmt', fonts: 'font', fills: 'fill', borders: 'border', cellXfs: 'xf' };
/** Order of the blocks inside <styleSheet>. */
const BLOCK_ORDER: Block[] = ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs'];
const FONT_CHILD_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
const FONT_KEYS = ['bold', 'italic', 'underline', 'strike', 'fontName', 'color', 'fontScale'] as const;
const ALIGN_KEYS = ['hAlign', 'vAlign', 'wrap', 'indent'] as const;
const SIDES = ['left', 'right', 'top', 'bottom'] as const;

const DEFAULT_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

function elementPattern(name: string, global = true): RegExp {
  return new RegExp(`<((?:[\\w.-]+:)?${name})(?=[\\s/>])[^>]*?(?:/>|>[\\s\\S]*?</\\1>)`, global ? 'g' : '');
}

/** Attribute value of an element's start tag. */
export function getAttr(element: string, name: string): string | undefined {
  const start = /^<[^>]*>/.exec(element)?.[0] ?? element;
  return new RegExp(`\\s${name}="([^"]*)"`).exec(start)?.[1];
}

/** Index where an element's start tag closes (its "/>" or ">"), skipping quoted values. */
function startTagEnd(element: string): number {
  let quote = '';
  for (let i = 1; i < element.length; i++) {
    const ch = element[i];
    if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") quote = ch;
    else if (ch === '>') return element[i - 1] === '/' ? i - 1 : i;
  }
  return element.length;
}

/** Sets (or with null removes) attributes on an element's start tag. */
export function setAttrs(element: string, attrs: Record<string, string | null>): string {
  const end = startTagEnd(element);
  let start = element.slice(0, end);
  const rest = element.slice(end);
  for (const [name, value] of Object.entries(attrs)) {
    const pattern = new RegExp(`\\s${name.replace('.', '\\.')}="[^"]*"`);
    if (value === null) start = start.replace(pattern, '');
    else if (pattern.test(start)) start = start.replace(pattern, ` ${name}="${value}"`);
    else start += ` ${name}="${value}"`;
  }
  return start + rest;
}

function prefixOf(element: string): string {
  return /^<([\w.-]+:)?/.exec(element)?.[1] ?? '';
}

/** Children of an element keyed by local name, in document order. */
function childElements(element: string): { name: string; xml: string }[] {
  const inner = /^<[^>]*[^/]>([\s\S]*)<\/[^>]+>$/.exec(element)?.[1];
  if (!inner) return [];
  const out: { name: string; xml: string }[] = [];
  const pattern = /<((?:[\w.-]+:)?([\w.-]+))(?=[\s/>])[^>]*?(?:\/>|>[\s\S]*?<\/\1>)/g;
  for (const match of inner.matchAll(pattern)) out.push({ name: match[2], xml: match[0] });
  return out;
}

/** The start tag without its closing "/>" or ">". */
function startTag(element: string): string {
  return element.slice(0, startTagEnd(element));
}

function argb(color: string): string {
  return `FF${color.replace('#', '').toUpperCase()}`;
}

function borderStyleName(edge: BorderEdge): string {
  if (edge.style === 'double') return 'double';
  if (edge.style === 'dotted') return 'dotted';
  if (edge.style === 'dashed') return edge.width >= 2 ? 'mediumDashed' : 'dashed';
  return edge.width >= 3 ? 'thick' : edge.width === 2 ? 'medium' : 'thin';
}

const same = (a: unknown, b: unknown) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

/** Built-in number formats that mean the same everywhere (date ones follow the system locale, so they're written out). */
const BUILTIN_FORMATS: Record<string, number> = {
  General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '# ?/?': 12, '# ??/??': 13, '@': 49,
};

function builtinFormatId(code: string): number | undefined {
  return Object.hasOwn(BUILTIN_FORMATS, code) ? BUILTIN_FORMATS[code] : undefined;
}

export class StyleWriter {
  private xml: string;
  private readonly items: Record<Block, string[]>;
  private readonly prefix: string;
  private readonly lookup: Record<Block, Map<string, number>>;
  private readonly customFormats = new Map<string, number>();
  private nextFormatId = 164;
  private changed = false;
  private readonly memo = new Map<string, number>();
  private readonly defaultFont: { name: string; size: number };

  constructor(stylesXml: string, defaultFont: { name: string; size: number }) {
    this.xml = stylesXml;
    this.defaultFont = defaultFont;
    this.prefix = /<([\w.-]+:)?styleSheet\b/.exec(stylesXml)?.[1] ?? '';
    this.items = { numFmts: [], fonts: [], fills: [], borders: [], cellXfs: [] };
    this.lookup = { numFmts: new Map(), fonts: new Map(), fills: new Map(), borders: new Map(), cellXfs: new Map() };
    for (const block of BLOCK_ORDER) {
      const inner = new RegExp(`<((?:[\\w.-]+:)?${block})\\b[^>]*?(?:/>|>([\\s\\S]*?)</\\1>)`).exec(stylesXml)?.[2] ?? '';
      this.items[block] = [...inner.matchAll(elementPattern(ITEM[block]))].map((m) => m[0]);
      this.items[block].forEach((xml, i) => {
        if (!this.lookup[block].has(xml)) this.lookup[block].set(xml, i);
      });
    }
    for (const numFmt of this.items.numFmts) {
      const id = Number(getAttr(numFmt, 'numFmtId'));
      const code = (getAttr(numFmt, 'formatCode') ?? '').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
      if (!this.customFormats.has(code)) this.customFormats.set(code, id);
      if (id >= this.nextFormatId) this.nextFormatId = id + 1;
    }
    if (this.items.cellXfs.length === 0) this.items.cellXfs.push(DEFAULT_XF);
  }

  /** An xf index with `base`'s format changed from style `from` to style `to`. */
  xfFor(base: number, from: CellStyle, to: CellStyle): number {
    const key = `${base}|${styleKey(from)}|${styleKey(to)}`;
    const cached = this.memo.get(key);
    if (cached !== undefined) return cached;
    const f = canonicalStyle(from);
    const t = canonicalStyle(to);
    let xf = this.items.cellXfs[base] ?? this.items.cellXfs[0];

    if (FONT_KEYS.some((k) => !same(f[k], t[k]))) {
      const fontId = Number(getAttr(xf, 'fontId') ?? 0);
      const font = this.patchFont(this.items.fonts[fontId] ?? this.items.fonts[0] ?? `<${this.prefix}font/>`, f, t);
      xf = setAttrs(xf, { fontId: String(this.add('fonts', font)), applyFont: '1' });
    }
    if (f.fill !== t.fill) {
      const p = this.prefix;
      const fillId = t.fill
        ? this.add('fills', `<${p}fill><${p}patternFill patternType="solid"><${p}fgColor rgb="${argb(t.fill)}"/><${p}bgColor indexed="64"/></${p}patternFill></${p}fill>`)
        : 0;
      xf = setAttrs(xf, { fillId: String(fillId), applyFill: '1' });
    }
    if (!same(f.borders, t.borders)) {
      const borderId = Number(getAttr(xf, 'borderId') ?? 0);
      const border = this.patchBorder(this.items.borders[borderId] ?? this.items.borders[0], f, t);
      xf = setAttrs(xf, { borderId: String(this.add('borders', border)), applyBorder: '1' });
    }
    if ((f.numberFormat ?? 'General') !== (t.numberFormat ?? 'General')) {
      xf = setAttrs(xf, { numFmtId: String(this.formatId(t.numberFormat ?? 'General')), applyNumberFormat: '1' });
    }
    if (ALIGN_KEYS.some((k) => !same(f[k], t[k]))) xf = this.patchAlignment(xf, f, t);

    const id = this.add('cellXfs', xf);
    this.memo.set(key, id);
    return id;
  }

  /** styles.xml with the added formats (unchanged text when nothing was added). */
  toXml(): string {
    if (!this.changed) return this.xml;
    let xml = this.xml;
    for (const block of BLOCK_ORDER) {
      const items = this.items[block];
      const pattern = new RegExp(`<((?:[\\w.-]+:)?${block})\\b([^>]*?)(?:/>|>[\\s\\S]*?</\\1>)`);
      const existing = pattern.exec(xml);
      if (existing) {
        const attrs = /\scount="\d+"/.test(existing[2]) ? existing[2].replace(/\scount="\d+"/, ` count="${items.length}"`) : `${existing[2]} count="${items.length}"`;
        xml = xml.replace(existing[0], () => `<${existing[1]}${attrs}>${items.join('')}</${existing[1]}>`);
      } else if (items.length > 0) {
        const element = `<${this.prefix}${block} count="${items.length}">${items.join('')}</${this.prefix}${block}>`;
        // Insert before the first later block that exists, or after the opening tag.
        const later = BLOCK_ORDER.slice(BLOCK_ORDER.indexOf(block) + 1)
          .map((b) => new RegExp(`<(?:[\\w.-]+:)?${b}\\b`).exec(xml))
          .find(Boolean);
        if (later) xml = xml.slice(0, later.index) + element + xml.slice(later.index);
        else xml = xml.replace(/(<(?:[\w.-]+:)?styleSheet\b[^>]*>)/, (open) => open + element);
      }
    }
    return xml;
  }

  private add(block: Block, xml: string): number {
    const existing = this.lookup[block].get(xml);
    if (existing !== undefined) return existing;
    this.items[block].push(xml);
    const id = this.items[block].length - 1;
    this.lookup[block].set(xml, id);
    this.changed = true;
    return id;
  }

  private formatId(code: string): number {
    const builtin = builtinFormatId(code);
    if (builtin !== undefined) return builtin;
    let id = this.customFormats.get(code);
    if (id === undefined) {
      id = this.nextFormatId++;
      this.customFormats.set(code, id);
      this.items.numFmts.push(`<${this.prefix}numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`);
      this.changed = true;
    }
    return id;
  }

  private patchFont(font: string, from: CellStyle, to: CellStyle): string {
    const p = prefixOf(font);
    const children = new Map<string, string>();
    const others: string[] = [];
    for (const child of childElements(font)) {
      if (FONT_CHILD_ORDER.includes(child.name)) children.set(child.name, child.xml);
      else others.push(child.xml);
    }
    const flag = (key: 'bold' | 'italic' | 'underline' | 'strike', name: string) => {
      if (from[key] === to[key]) return;
      if (to[key]) children.set(name, `<${p}${name}/>`);
      else children.delete(name);
    };
    flag('bold', 'b');
    flag('italic', 'i');
    flag('underline', 'u');
    flag('strike', 'strike');
    if (from.color !== to.color) {
      if (to.color) children.set('color', `<${p}color rgb="${argb(to.color)}"/>`);
      else children.delete('color');
    }
    if ((from.fontScale ?? 1) !== (to.fontScale ?? 1)) {
      const size = Math.round(this.defaultFont.size * (to.fontScale ?? 1) * 2) / 2;
      children.set('sz', `<${p}sz val="${size}"/>`);
    }
    if (from.fontName !== to.fontName) {
      // A theme font (scheme) would override the chosen name.
      children.delete('scheme');
      children.delete('charset');
      children.delete('family');
      if (to.fontName) children.set('name', `<${p}name val="${escapeXml(to.fontName)}"/>`);
      else {
        const defaults = childElements(this.items.fonts[0] ?? '');
        for (const name of ['name', 'family', 'charset', 'scheme']) {
          const child = defaults.find((c) => c.name === name);
          if (child) children.set(name, child.xml);
        }
        if (!children.has('name')) children.set('name', `<${p}name val="${escapeXml(this.defaultFont.name)}"/>`);
      }
    }
    const ordered = FONT_CHILD_ORDER.filter((name) => children.has(name)).map((name) => children.get(name)!);
    const open = startTag(font);
    return `${open}>${ordered.join('')}${others.join('')}</${p}font>`;
  }

  private patchBorder(border: string | undefined, from: CellStyle, to: CellStyle): string {
    const p = border ? prefixOf(border) : this.prefix;
    const edges = new Map<string, string>();
    const extra: string[] = [];
    for (const child of border ? childElements(border) : []) {
      const name = child.name === 'start' ? 'left' : child.name === 'end' ? 'right' : child.name;
      if ((SIDES as readonly string[]).includes(name)) {
        edges.set(name, child.xml.replace(new RegExp(`^<${p}(start|end)\\b`), `<${p}${name}`).replace(new RegExp(`</${p}(start|end)>$`), `</${p}${name}>`));
      } else if (name === 'diagonal') edges.set('diagonal', child.xml);
      else extra.push(child.xml);
    }
    for (const side of SIDES) {
      const before = from.borders?.[side];
      const after = to.borders?.[side];
      if (same(before, after)) continue;
      edges.set(
        side,
        after
          ? `<${p}${side} style="${borderStyleName(after)}"><${p}color ${after.color ? `rgb="${argb(after.color)}"` : 'indexed="64"'}/></${p}${side}>`
          : `<${p}${side}/>`,
      );
    }
    const open = border ? startTag(border) : `<${p}border`;
    const ordered = [...SIDES, 'diagonal'].map((side) => edges.get(side) ?? `<${p}${side}/>`);
    return `${open}>${ordered.join('')}${extra.join('')}</${p}border>`;
  }

  private patchAlignment(xf: string, from: CellStyle, to: CellStyle): string {
    const p = prefixOf(xf);
    const children = childElements(xf);
    const current = children.find((c) => c.name === 'alignment')?.xml;
    const updates: Record<string, string | null> = {};
    if (from.hAlign !== to.hAlign) updates.horizontal = to.hAlign ?? null;
    if (from.vAlign !== to.vAlign) updates.vertical = to.vAlign ?? null;
    if (from.wrap !== to.wrap) updates.wrapText = to.wrap ? '1' : null;
    if (from.indent !== to.indent) updates.indent = to.indent ? String(to.indent) : null;
    const alignment = setAttrs(`${current ? startTag(current) : `<${p}alignment`}/>`, updates);
    const inner = [
      ...(/\s[\w:.-]+="/.test(alignment) ? [alignment] : []),
      ...children.filter((c) => c.name !== 'alignment').map((c) => c.xml),
    ];
    const open = startTag(setAttrs(`${startTag(xf)}/>`, { applyAlignment: '1' }));
    return inner.length > 0 ? `${open}>${inner.join('')}</${p}xf>` : `${open}/>`;
  }
}
