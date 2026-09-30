/**
 * Text width measurement for text overflow and column auto-fit.
 *
 * Works on the main thread (canvas) and in the parser worker (OffscreenCanvas).
 * Falls back to a character-count estimate where no canvas is available (tests).
 * Widths are measured once at the base font size and scaled by the caller.
 */

export const CELL_FONT_FAMILY =
  '-apple-system, BlinkMacSystemFont, "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif';

/** Cell font size at 100% zoom, in CSS pixels. Matches Excel's Calibri 11 digit width. */
export const BASE_FONT_PX = 12;

type Context2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

let context: Context2D | null | undefined;
const cache = new Map<string, number>();
const MAX_CACHE_ENTRIES = 50_000;

function getContext(): Context2D | null {
  if (context !== undefined) return context;
  context = null;
  try {
    if (typeof OffscreenCanvas !== 'undefined') {
      context = new OffscreenCanvas(1, 1).getContext('2d');
    } else if (typeof document !== 'undefined') {
      context = document.createElement('canvas').getContext('2d');
    }
  } catch {
    context = null;
  }
  return context;
}

/**
 * Width in CSS pixels of `text` rendered in the cell font at BASE_FONT_PX.
 * `tabular` measures digits at their fixed (tabular-nums) width, as numbers are drawn.
 * `family` is a CSS font family tried before the cell font (a cell's own font).
 */
export function measureText(text: string, bold = false, italic = false, tabular = false, family?: string): number {
  if (!text) return 0;
  if (tabular) text = text.replace(/[0-9]/g, '0');
  const key = `${bold ? 'b' : ''}${italic ? 'i' : ''}${family ?? ''}|${text}`;
  const cached = cache.get(key);
  if (cached !== undefined) return cached;

  const ctx = getContext();
  let width: number;
  if (ctx) {
    ctx.font = `${italic ? 'italic ' : ''}${bold ? '600 ' : '400 '}${BASE_FONT_PX}px ${family ? `${family}, ` : ''}${CELL_FONT_FAMILY}`;
    width = ctx.measureText(text).width;
  } else {
    width = text.length * BASE_FONT_PX * (bold ? 0.6 : 0.55);
  }
  if (cache.size >= MAX_CACHE_ENTRIES) cache.clear();
  cache.set(key, width);
  return width;
}
