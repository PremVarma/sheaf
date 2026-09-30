/** Colors from spreadsheet files, normalized for CSS and adapted for dark mode. */

/** "FF00FF00" (ARGB) or "00FF00" → "#00ff00"; undefined for anything else. */
export function normalizeHex(value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  const hex = value.replace(/^#/, '');
  if (/^[0-9a-fA-F]{8}$/.test(hex)) return `#${hex.slice(2).toLowerCase()}`;
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toLowerCase()}`;
  return undefined;
}

function channels(hex: string): [number, number, number] {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function toHex(r: number, g: number, b: number): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)));
  return `#${((1 << 24) | (clamp(r) << 16) | (clamp(g) << 8) | clamp(b)).toString(16).slice(1)}`;
}

/** WCAG relative luminance, 0 (black) … 1 (white). */
export function relativeLuminance(hex: string): number {
  const [r, g, b] = channels(hex).map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** Black or white, whichever reads better on the given fill. */
export function readableTextOn(fill: string): string {
  return relativeLuminance(fill) > 0.4 ? '#000000' : '#ffffff';
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/** Excel theme tint: negative darkens, positive lightens (ECMA-376 18.8.19). */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = rgbToHsl(...channels(hex));
  const lum = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return toHex(...hslToRgb(h, s, lum));
}

/**
 * Dark mode: text colors chosen for a white sheet can vanish on a dark one.
 * Very dark colors are mirrored in lightness, keeping their hue.
 */
export function adaptForDarkBackground(hex: string): string {
  const [h, s, l] = rgbToHsl(...channels(hex));
  if (l >= 0.55) return hex;
  return toHex(...hslToRgb(h, s, Math.max(0.65, 1 - l)));
}

/** Excel's default indexed color palette (BIFF8 / styles.xml `indexed`). */
export const INDEXED_COLORS: readonly string[] = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];
