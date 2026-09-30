import type { CSSProperties } from 'react';
import { CellKind, type BorderEdge, type CellStyle } from '../../models/workbook';
import { adaptForDarkBackground, readableTextOn } from '../../utils/color';

/** Precomputed rendering data for one workbook style (per color scheme). */
export interface CellCss {
  /** Inline style merged into each cell element. */
  style: CSSProperties;
  /** Class suffix for vertical alignment / wrapping / fill. */
  extraClass: string;
  /** Horizontal alignment class, or null to align by value type. */
  alignClass: string | null;
  borderStyle: CSSProperties | null;
  bold: boolean;
  italic: boolean;
  fontScale: number;
  /** CSS family of the cell's own font (quoted), or undefined for the default font. */
  fontFamily: string | undefined;
  /** Indent in px at 100%. */
  indentPx: number;
  wrap: boolean;
  /** Filled cells cover their gridlines, as in Excel. */
  opaque: boolean;
  /** Left-aligned text that may spill into empty neighbors. */
  canOverflow: boolean;
}

/** Excel's "General" alignment by value type. */
export const KIND_ALIGN_CLASS: Record<number, string> = {
  [CellKind.Empty]: 'xv-al',
  [CellKind.String]: 'xv-al',
  [CellKind.Number]: 'xv-ar',
  [CellKind.Date]: 'xv-ar',
  [CellKind.Boolean]: 'xv-ac',
  [CellKind.Error]: 'xv-ac',
};

const ALIGN_CLASS: Record<string, string> = {
  left: 'xv-al',
  fill: 'xv-al',
  justify: 'xv-al',
  center: 'xv-ac',
  centerContinuous: 'xv-ac',
  distributed: 'xv-ac',
  right: 'xv-ar',
};

function edgeCss(edge: BorderEdge | undefined, dark: boolean): string | undefined {
  if (!edge) return undefined;
  const color = edge.color ? (dark ? adaptForDarkBackground(edge.color) : edge.color) : 'var(--xv-border-auto)';
  return `${edge.width}px ${edge.style} ${color}`;
}

export function buildCellCss(styles: readonly CellStyle[], dark: boolean): CellCss[] {
  return styles.map((style) => {
    const css: CSSProperties = {};
    let fill = style.fill;
    // A white fill is usually there to hide gridlines; keep that intent in dark mode.
    if (fill === '#ffffff' && dark) fill = undefined;
    const opaque = Boolean(style.fill);
    if (fill) css.background = fill;
    else if (opaque) css.background = 'var(--xv-grid-bg)';

    if (style.color) css.color = dark && !fill ? adaptForDarkBackground(style.color) : style.color;
    else if (fill) css.color = readableTextOn(fill);

    if (style.bold) css.fontWeight = 600;
    if (style.italic) css.fontStyle = 'italic';
    const decorations = [style.underline && 'underline', style.strike && 'line-through'].filter(Boolean);
    if (decorations.length) css.textDecoration = decorations.join(' ');
    const fontScale = style.fontScale ?? 1;
    if (fontScale !== 1) css.fontSize = `calc(var(--xv-font) * ${fontScale})`;
    const fontFamily = style.fontName ? `"${style.fontName.replace(/["\\]/g, '')}"` : undefined;
    if (fontFamily) css.fontFamily = `${fontFamily}, var(--font-cell)`;
    if (style.indent) css.paddingLeft = `calc(${3 + style.indent * 9}px * var(--xv-zoom))`;

    const borders = style.borders;
    const borderStyle: CSSProperties | null = borders
      ? {
          borderTop: edgeCss(borders.top, dark),
          borderRight: edgeCss(borders.right, dark),
          borderBottom: edgeCss(borders.bottom, dark),
          borderLeft: edgeCss(borders.left, dark),
        }
      : null;

    const wrap = Boolean(style.wrap || style.hAlign === 'justify' || style.hAlign === 'distributed');
    const alignClass = style.hAlign ? (ALIGN_CLASS[style.hAlign] ?? null) : null;
    let extraClass = '';
    if (style.vAlign === 'top') extraClass += ' xv-vt';
    else if (style.vAlign === 'center') extraClass += ' xv-vc';
    if (wrap) extraClass += ' xv-wrap';
    if (opaque) extraClass += ' xv-fill';

    return {
      style: css,
      extraClass,
      alignClass,
      borderStyle,
      bold: Boolean(style.bold),
      italic: Boolean(style.italic),
      fontScale,
      fontFamily,
      indentPx: (style.indent ?? 0) * 9,
      wrap,
      opaque,
      canOverflow: !wrap && (!style.hAlign || style.hAlign === 'left'),
    };
  });
}
