import { applyTint } from '../../utils/color';

/** Excel's Office theme colors, with its lighter and darker variations. */
const THEME = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
const STANDARD = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];
const NAMES: Record<string, string> = {
  '#ffffff': 'White',
  '#000000': 'Black',
  '#e7e6e6': 'Light gray',
  '#44546a': 'Blue gray',
  '#4472c4': 'Blue',
  '#ed7d31': 'Orange',
  '#a5a5a5': 'Gray',
  '#ffc000': 'Gold',
  '#5b9bd5': 'Light blue',
  '#70ad47': 'Green',
  '#c00000': 'Dark red',
  '#ff0000': 'Red',
  '#ffff00': 'Yellow',
  '#92d050': 'Light green',
  '#00b050': 'Green',
  '#00b0f0': 'Sky blue',
  '#0070c0': 'Blue',
  '#002060': 'Dark blue',
  '#7030a0': 'Purple',
};

function shades(color: string, column: number): string[] {
  if (column === 0) return [-0.05, -0.15, -0.25, -0.35, -0.5].map((t) => applyTint(color, t));
  if (column === 1) return [0.5, 0.35, 0.25, 0.15, 0.05].map((t) => applyTint(color, t));
  return [0.8, 0.6, 0.4, -0.25, -0.5].map((t) => applyTint(color, t));
}

const THEME_ROWS: string[][] = [THEME, ...[0, 1, 2, 3, 4].map((row) => THEME.map((color, column) => shades(color, column)[row]))];

interface ColorPaletteProps {
  value: string | undefined;
  /** "Automatic" (font color) or "No Fill". */
  noneLabel: string;
  onPick(color: string | null): void;
  /** Opens the system color picker. */
  onMore(): void;
}

export function ColorPalette({ value, noneLabel, onPick, onMore }: ColorPaletteProps) {
  const swatch = (color: string, label: string) => (
    <button
      key={color + label}
      type="button"
      title={label}
      aria-label={label}
      aria-pressed={value === color}
      className={`h-[18px] w-[18px] rounded-[3px] border outline-none transition-transform hover:scale-110 focus-visible:ring-2 focus-visible:ring-accent ${
        value === color ? 'border-accent ring-2 ring-accent/60' : 'border-black/10 dark:border-white/15'
      }`}
      style={{ background: color }}
      onClick={() => onPick(color)}
    />
  );
  return (
    <div className="w-[236px] p-2" role="dialog" aria-label="Colors">
      <button
        type="button"
        className="mb-2 flex h-7 w-full items-center gap-2 rounded-md px-2 text-[12.5px] text-fg hover:bg-hover"
        onClick={() => onPick(null)}
      >
        <span className="relative h-4 w-4 rounded-[3px] border border-line-strong bg-surface">
          <span className="absolute left-1/2 top-[-2px] h-[20px] w-px rotate-45 bg-danger" />
        </span>
        {noneLabel}
      </button>
      <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-muted">Theme colors</div>
      <div className="grid grid-cols-10 gap-[4px]">
        {THEME_ROWS.flatMap((row, r) => row.map((color, c) => swatch(color, r === 0 ? (NAMES[color] ?? color) : `${NAMES[THEME[c]] ?? THEME[c]} (${color})`)))}
      </div>
      <div className="mb-1 mt-2.5 text-[11px] font-semibold uppercase tracking-wide text-muted">Standard colors</div>
      <div className="grid grid-cols-10 gap-[4px]">{STANDARD.map((color) => swatch(color, NAMES[color] ?? color))}</div>
      <button type="button" className="mt-2.5 flex h-7 w-full items-center rounded-md px-2 text-[12.5px] text-fg hover:bg-hover" onClick={onMore}>
        More colors…
      </button>
    </div>
  );
}
