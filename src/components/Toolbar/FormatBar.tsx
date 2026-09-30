import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { DEFAULT_FONT, type CellStyle } from '../../models/workbook';
import { CURRENCY_SYMBOLS, currencyFormat, formatPresets, localCurrency, presetFor } from '../../services/editing/numberFormats';
import { getCell } from '../../services/workbook/sheetService';
import { useController, usePlatform, useViewer } from '../../state/AppContext';
import { activeCellStyle, FONT_SIZES, type BorderPreset } from '../../state/controller';
import type { ViewerState } from '../../state/viewerState';
import { Icon, type IconName } from '../Icon';
import { Menu, Popover, type MenuItem } from '../Menu/Menu';
import { shortcutLabel } from '../ui';
import { ColorPalette } from './ColorPicker';

/** Fonts offered in the font list (the workbook's default font is added when missing). */
const FONTS = ['Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Helvetica', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];

const noFocus = (event: React.MouseEvent) => event.preventDefault();

type Formatter = (value: number, format?: string) => string;
let formatter: Formatter | null = null;

/** Number format previews need SheetJS's formatter, which is loaded on first use to keep startup light. */
function useFormatter(): [Formatter | null, () => void] {
  const [format, setFormat] = useState<Formatter | null>(() => formatter);
  const load = () => {
    if (formatter) return;
    void import('../../services/workbook/numberFormat').then((module) => {
      formatter = (value, code) => module.formatNumber(value, code);
      setFormat(() => formatter);
    });
  };
  return [format, load];
}

function activeNumber(state: ViewerState): number | null {
  const view = state.sheetViews[state.activeSheet];
  const sheet = state.workbook?.sheets[state.activeSheet];
  if (!view || !sheet) return null;
  const value = getCell(sheet, view.selection.anchor.row, view.selection.anchor.col)?.value;
  return typeof value === 'number' ? value : null;
}

function Divider() {
  return <div className="mx-1 h-5 w-px shrink-0 bg-line" aria-hidden="true" />;
}

interface ButtonProps {
  label: string;
  title?: string;
  icon?: IconName;
  active?: boolean;
  disabled?: boolean;
  onClick(): void;
  children?: ReactNode;
  className?: string;
}

/** A toolbar button that leaves the keyboard focus where it was (the grid or the cell being edited). */
function FormatButton({ label, title, icon, active = false, disabled, onClick, children, className = '' }: ButtonProps) {
  return (
    <button
      type="button"
      aria-label={label}
      title={title ?? label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={noFocus}
      onClick={onClick}
      className={`inline-flex h-7 min-w-7 shrink-0 pointer-coarse:h-9 pointer-coarse:min-w-9 items-center justify-center rounded-md px-1 text-fg transition-colors hover:bg-hover active:bg-pressed disabled:pointer-events-none disabled:opacity-40 ${
        active ? 'bg-accent-soft text-accent' : ''
      } ${className}`}
    >
      {icon && <Icon name={icon} size={16} />}
      {children}
    </button>
  );
}

/** A button that opens a menu or panel below it. */
function DropdownButton({
  label,
  title,
  disabled,
  children,
  menu,
  panel,
  className = '',
  split,
}: {
  label: string;
  title?: string;
  disabled?: boolean;
  children: ReactNode;
  menu?: () => MenuItem[];
  panel?: (close: () => void) => ReactNode;
  className?: string;
  /** Main action for a split button (the arrow opens the menu). */
  split?: () => void;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const controller = useController();
  const close = () => {
    setOpen(false);
    controller.focusGrid();
  };
  const arrow = (
    <button
      type="button"
      aria-label={split ? `${label} options` : label}
      title={title ?? label}
      aria-haspopup="menu"
      aria-expanded={open}
      disabled={disabled}
      onMouseDown={noFocus}
      onClick={() => setOpen((v) => !v)}
      className={`inline-flex h-7 shrink-0 pointer-coarse:h-9 items-center justify-center gap-0.5 rounded-md text-fg transition-colors hover:bg-hover disabled:pointer-events-none disabled:opacity-40 ${
        open ? 'bg-pressed' : ''
      } ${split ? 'w-4 rounded-l-none' : `px-1 ${className}`}`}
    >
      {!split && children}
      <Icon name="chevronDown" size={11} className="opacity-60" />
    </button>
  );
  return (
    <div ref={ref} className="flex shrink-0 items-center">
      {split && (
        <button
          type="button"
          aria-label={label}
          title={title ?? label}
          disabled={disabled}
          onMouseDown={noFocus}
          onClick={split}
          className={`inline-flex h-7 shrink-0 pointer-coarse:h-9 items-center justify-center rounded-md rounded-r-none px-1 text-fg transition-colors hover:bg-hover disabled:pointer-events-none disabled:opacity-40 ${className}`}
        >
          {children}
        </button>
      )}
      {arrow}
      {open && menu && <Menu items={menu()} placement={{ anchor: ref.current!, side: 'below' }} onClose={close} ignore={[ref.current]} label={label} />}
      {open && panel && (
        <Popover placement={{ anchor: ref.current!, side: 'below' }} onClose={close} ignore={[ref.current]}>
          {panel(close)}
        </Popover>
      )}
    </div>
  );
}

/** A small picture of a cell with some borders drawn (for the borders menu). */
function BorderGlyph({ sides, thick = false, double = false }: { sides: string; thick?: boolean; double?: boolean }) {
  const has = (side: string) => sides.includes(side);
  const solid = (side: string) => ({ strokeWidth: has(side) ? (thick ? 2.4 : 1.4) : 0.8, strokeDasharray: has(side) ? undefined : '1.5 1.5', opacity: has(side) ? 1 : 0.45 });
  return (
    <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" aria-hidden="true">
      <line x1="2" y1="2" x2="14" y2="2" {...solid('t')} />
      <line x1="2" y1="14" x2="14" y2="14" {...solid('b')} />
      {double && has('b') && <line x1="2" y1="11.5" x2="14" y2="11.5" strokeWidth={1.2} />}
      <line x1="2" y1="2" x2="2" y2="14" {...solid('l')} />
      <line x1="14" y1="2" x2="14" y2="14" {...solid('r')} />
      <line x1="8" y1="2" x2="8" y2="14" {...solid('v')} />
      <line x1="2" y1="8" x2="14" y2="8" {...solid('h')} />
    </svg>
  );
}

function FontSizeBox({ size, disabled, onSize }: { size: number; disabled: boolean; onSize(size: number): void }) {
  const [draft, setDraft] = useState(String(size));
  useEffect(() => setDraft(String(size)), [size]);
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const controller = useController();
  const commit = () => {
    const value = Number(draft);
    if (Number.isFinite(value) && value >= 1 && value <= 409 && value !== size) onSize(Math.round(value * 2) / 2);
    else setDraft(String(size));
  };
  return (
    <div ref={ref} className="flex h-7 shrink-0 pointer-coarse:h-9 items-center rounded-md border border-control-line bg-control">
      <input
        aria-label="Font size"
        title="Font size"
        className="h-full w-9 bg-transparent pl-1.5 text-[12.5px] tabular-nums text-fg outline-none disabled:opacity-40"
        value={draft}
        disabled={disabled}
        inputMode="decimal"
        onChange={(event) => setDraft(event.target.value)}
        onFocus={(event) => event.target.select()}
        onBlur={commit}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            commit();
            controller.focusGrid();
          } else if (event.key === 'Escape') {
            setDraft(String(size));
            controller.focusGrid();
          }
        }}
      />
      <button
        type="button"
        aria-label="Font sizes"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onMouseDown={noFocus}
        onClick={() => setOpen((v) => !v)}
        className="flex h-full w-4 items-center justify-center rounded-r-md text-fg hover:bg-hover disabled:opacity-40"
      >
        <Icon name="chevronDown" size={11} className="opacity-60" />
      </button>
      {open && (
        <Menu
          label="Font sizes"
          placement={{ anchor: ref.current!, side: 'below' }}
          ignore={[ref.current]}
          onClose={() => {
            setOpen(false);
            controller.focusGrid();
          }}
          items={FONT_SIZES.map((s) => ({ label: String(s), checked: s === size, onSelect: () => onSize(s) }))}
        />
      )}
    </div>
  );
}

/** Whether a horizontally scrolling element has more content to the left or right. */
function useScrollEdges(ref: React.RefObject<HTMLDivElement | null>): { left: boolean; right: boolean } {
  const [edges, setEdges] = useState({ left: false, right: false });
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const update = () => {
      const left = el.scrollLeft > 1;
      const right = el.scrollLeft + el.clientWidth < el.scrollWidth - 1;
      setEdges((prev) => (prev.left === left && prev.right === right ? prev : { left, right }));
    };
    update();
    el.addEventListener('scroll', update, { passive: true });
    const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(update);
    observer?.observe(el);
    return () => {
      el.removeEventListener('scroll', update);
      observer?.disconnect();
    };
  }, [ref]);
  return edges;
}

/** Fades the edges that have more content beyond them. */
function edgeFade({ left, right }: { left: boolean; right: boolean }): React.CSSProperties | undefined {
  if (!left && !right) return undefined;
  const mask = `linear-gradient(to right, ${left ? 'transparent' : 'black'}, black 32px, black calc(100% - 32px), ${right ? 'transparent' : 'black'})`;
  return { maskImage: mask, WebkitMaskImage: mask };
}

/** Formatting toolbar: font, colors, borders, alignment, merging, number formats, rows/columns and sorting. */
export function FormatBar({ className = '' }: { className?: string }) {
  const controller = useController();
  const { isMac } = usePlatform();
  const barRef = useRef<HTMLDivElement>(null);
  const edges = useScrollEdges(barRef);
  const enabled = useViewer((s) => s.workbook?.sheets[s.activeSheet]?.kind === 'worksheet');
  const style: CellStyle = useViewer(activeCellStyle);
  const defaultFont = useViewer((s) => s.workbook?.defaultFont ?? DEFAULT_FONT);
  const sample = useViewer(activeNumber);
  const merged = useViewer((s) => {
    const view = s.sheetViews[s.activeSheet];
    const sheet = s.workbook?.sheets[s.activeSheet];
    if (!view || !sheet) return false;
    const { row, col } = view.selection.anchor;
    return sheet.merges.some((m) => row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1);
  });
  const [textColor, setTextColor] = useState('#c00000');
  const [fillColor, setFillColor] = useState('#ffff00');
  const textInput = useRef<HTMLInputElement>(null);
  const fillInput = useRef<HTMLInputElement>(null);
  const presets = useMemo(() => formatPresets(), []);
  const [formatNumber, loadFormatter] = useFormatter();
  const disabled = !enabled;
  const k = (key: string, shift = false) => shortcutLabel(key, isMac, shift);

  // The system color panel reports its final choice with a "change" event.
  useEffect(() => {
    const listen = (input: HTMLInputElement | null, apply: (color: string) => void) => {
      if (!input) return () => undefined;
      const onChange = () => apply(input.value.toLowerCase());
      input.addEventListener('change', onChange);
      return () => input.removeEventListener('change', onChange);
    };
    const a = listen(textInput.current, (color) => {
      setTextColor(color);
      void controller.setTextColor(color);
    });
    const b = listen(fillInput.current, (color) => {
      setFillColor(color);
      void controller.setFillColor(color);
    });
    return () => {
      a();
      b();
    };
  }, [controller]);

  const fontName = style.fontName ?? defaultFont.name;
  const fonts = FONTS.includes(fontName) ? FONTS : [...FONTS, fontName].sort();
  const fontSize = Math.round(defaultFont.size * (style.fontScale ?? 1) * 2) / 2;
  const preset = presetFor(style.numberFormat, presets);
  const run = (action: () => Promise<void>) => () => void action();

  const borderItems = (): MenuItem[] => {
    const item = (label: string, preset: BorderPreset, glyph: ReactNode): MenuItem => ({ label, icon: glyph, onSelect: () => void controller.applyBorders(preset) });
    return [
      item('Bottom Border', 'bottom', <BorderGlyph sides="b" />),
      item('Top Border', 'top', <BorderGlyph sides="t" />),
      item('Left Border', 'left', <BorderGlyph sides="l" />),
      item('Right Border', 'right', <BorderGlyph sides="r" />),
      { kind: 'separator' },
      item('No Border', 'none', <BorderGlyph sides="" />),
      item('All Borders', 'all', <BorderGlyph sides="tblrvh" />),
      item('Outside Borders', 'outside', <BorderGlyph sides="tblr" />),
      item('Inside Borders', 'inside', <BorderGlyph sides="vh" />),
      item('Thick Outside Borders', 'thickOutside', <BorderGlyph sides="tblr" thick />),
      { kind: 'separator' },
      item('Thick Bottom Border', 'thickBottom', <BorderGlyph sides="b" thick />),
      item('Double Bottom Border', 'doubleBottom', <BorderGlyph sides="b" double />),
      item('Top and Bottom Border', 'topBottom', <BorderGlyph sides="tb" />),
    ];
  };

  const preview = (value: number, format?: string) => {
    if (!formatNumber) {
      loadFormatter();
      return undefined;
    }
    const text = formatNumber(value, format);
    return text.length > 18 ? undefined : text;
  };

  const numberItems = (): MenuItem[] => [
    ...presets.map((p): MenuItem => ({
      label: p.label,
      checked: preset?.id === p.id,
      detail: sample !== null && p.id !== 'text' ? preview(sample, p.format) : undefined,
      onSelect: () => void controller.setNumberFormat(p.format),
    })),
    { kind: 'separator' },
    {
      kind: 'submenu',
      label: 'Currency',
      items: CURRENCY_SYMBOLS.map((symbol) => ({
        label: symbol,
        detail: preview(sample ?? 1234.5, currencyFormat(symbol)),
        checked: style.numberFormat === currencyFormat(symbol),
        onSelect: () => void controller.setNumberFormat(currencyFormat(symbol)),
      })),
    },
  ];

  const mergeItems = (): MenuItem[] => [
    { label: 'Merge & Center', onSelect: () => void controller.merge('center') },
    { label: 'Merge Across', onSelect: () => void controller.merge('across') },
    { label: 'Merge Cells', onSelect: () => void controller.merge('merge') },
    { label: 'Unmerge Cells', disabled: !merged, onSelect: () => void controller.unmerge() },
  ];

  const insertItems = (): MenuItem[] => [
    { label: 'Insert Rows Above', icon: <Icon name="rowInsert" size={14} />, shortcut: isMac ? '⌥⌘=' : 'Ctrl+Alt+=', onSelect: () => void controller.insertRows('above') },
    { label: 'Insert Rows Below', onSelect: () => void controller.insertRows('below') },
    { label: 'Insert Columns Left', icon: <Icon name="columnInsert" size={14} />, onSelect: () => void controller.insertColumns('left') },
    { label: 'Insert Columns Right', onSelect: () => void controller.insertColumns('right') },
    { kind: 'separator' },
    { label: 'Insert Sheet', icon: <Icon name="sheet" size={14} />, shortcut: '⇧F11', onSelect: () => void controller.addSheet() },
  ];

  const deleteItems = (): MenuItem[] => [
    { label: 'Delete Rows', icon: <Icon name="rowDelete" size={14} />, shortcut: isMac ? '⌥⌘−' : 'Ctrl+Alt+−', onSelect: () => void controller.deleteRows() },
    { label: 'Delete Columns', icon: <Icon name="columnDelete" size={14} />, onSelect: () => void controller.deleteColumns() },
    { kind: 'separator' },
    { label: 'Hide Rows', onSelect: () => void controller.setRowsHidden(true) },
    { label: 'Unhide Rows', onSelect: () => void controller.setRowsHidden(false) },
    { label: 'Hide Columns', onSelect: () => void controller.setColumnsHidden(true) },
    { label: 'Unhide Columns', onSelect: () => void controller.setColumnsHidden(false) },
    { kind: 'separator' },
    { label: 'Delete Sheet', danger: true, onSelect: () => void controller.deleteSheet(controller.store.getState().activeSheet) },
  ];

  const verticalItems = (): MenuItem[] => [
    { label: 'Top', icon: <Icon name="alignTop" size={14} />, checked: style.vAlign === 'top', onSelect: () => void controller.setVerticalAlignment('top') },
    { label: 'Middle', icon: <Icon name="alignMiddle" size={14} />, checked: style.vAlign === 'center', onSelect: () => void controller.setVerticalAlignment('center') },
    { label: 'Bottom', icon: <Icon name="alignBottom" size={14} />, checked: !style.vAlign || style.vAlign === 'bottom', onSelect: () => void controller.setVerticalAlignment('bottom') },
    { kind: 'separator' },
    { label: 'Increase Indent', icon: <Icon name="indentMore" size={14} />, onSelect: () => void controller.stepIndent(1) },
    { label: 'Decrease Indent', icon: <Icon name="indentLess" size={14} />, onSelect: () => void controller.stepIndent(-1) },
  ];

  return (
    <div
      ref={barRef}
      className={`no-scrollbar flex h-9 items-center gap-0.5 overflow-x-auto border-t border-line/70 px-2 pointer-coarse:h-11 ${className}`}
      style={edgeFade(edges)}
      role="toolbar"
      aria-label="Formatting"
      onWheel={(event) => {
        // A mouse wheel scrolls the bar sideways when it doesn't fit.
        if (event.deltaX === 0 && event.deltaY !== 0) event.currentTarget.scrollLeft += event.deltaY;
      }}
    >
      <select
        aria-label="Font"
        title="Font"
        disabled={disabled}
        value={fontName}
        onChange={(event) => {
          void controller.setFontName(event.target.value).then(() => controller.focusGrid());
        }}
        className="h-7 w-[118px] shrink-0 pointer-coarse:h-9 rounded-md border border-control-line bg-control px-1.5 text-[12.5px] text-fg outline-none disabled:opacity-40"
      >
        {fonts.map((font) => (
          <option key={font} value={font}>
            {font}
          </option>
        ))}
      </select>
      <div className="w-1 shrink-0" />
      <FontSizeBox size={fontSize} disabled={disabled} onSize={(size) => void controller.setFontSize(size)} />
      <FormatButton label="Increase font size" title={`Increase font size (${k('.', true)})`} disabled={disabled} onClick={run(() => controller.stepFontSize(1))}>
        <span className="text-[14px] font-semibold leading-none">A</span>
        <span className="-ml-px mb-2 text-[8px] leading-none">▲</span>
      </FormatButton>
      <FormatButton label="Decrease font size" title={`Decrease font size (${k(',', true)})`} disabled={disabled} onClick={run(() => controller.stepFontSize(-1))}>
        <span className="text-[11px] font-semibold leading-none">A</span>
        <span className="-ml-px mb-1.5 text-[8px] leading-none">▼</span>
      </FormatButton>
      <Divider />

      <FormatButton label="Bold" title={`Bold (${k('B')})`} active={Boolean(style.bold)} disabled={disabled} onClick={run(() => controller.toggleStyle('bold'))}>
        <span className="w-4 text-[14px] font-bold">B</span>
      </FormatButton>
      <FormatButton label="Italic" title={`Italic (${k('I')})`} active={Boolean(style.italic)} disabled={disabled} onClick={run(() => controller.toggleStyle('italic'))}>
        <span className="w-4 font-serif text-[15px] italic">I</span>
      </FormatButton>
      <FormatButton label="Underline" title={`Underline (${k('U')})`} active={Boolean(style.underline)} disabled={disabled} onClick={run(() => controller.toggleStyle('underline'))}>
        <span className="w-4 text-[14px] underline underline-offset-2">U</span>
      </FormatButton>
      <FormatButton
        label="Strikethrough"
        title={`Strikethrough (${isMac ? '⇧⌘X' : 'Ctrl+5'})`}
        active={Boolean(style.strike)}
        disabled={disabled}
        onClick={run(() => controller.toggleStyle('strike'))}
      >
        <span className="w-4 text-[14px] line-through">S</span>
      </FormatButton>
      <Divider />

      <DropdownButton
        label="Font color"
        disabled={disabled}
        split={run(() => controller.setTextColor(textColor))}
        panel={(close) => (
          <ColorPalette
            value={style.color}
            noneLabel="Automatic"
            onMore={() => {
              close();
              textInput.current?.click();
            }}
            onPick={(color) => {
              close();
              if (color) setTextColor(color);
              void controller.setTextColor(color);
            }}
          />
        )}
      >
        <span className="flex w-5 flex-col items-center">
          <span className="text-[14px] font-semibold leading-[14px]">A</span>
          <span className="mt-0.5 h-[3px] w-4 rounded-sm" style={{ background: textColor }} />
        </span>
      </DropdownButton>
      <DropdownButton
        label="Fill color"
        disabled={disabled}
        split={run(() => controller.setFillColor(fillColor))}
        panel={(close) => (
          <ColorPalette
            value={style.fill}
            noneLabel="No Fill"
            onMore={() => {
              close();
              fillInput.current?.click();
            }}
            onPick={(color) => {
              close();
              if (color) setFillColor(color);
              void controller.setFillColor(color);
            }}
          />
        )}
      >
        <span className="flex w-5 flex-col items-center">
          <Icon name="fill" size={14} />
          <span className="mt-0.5 h-[3px] w-4 rounded-sm" style={{ background: fillColor }} />
        </span>
      </DropdownButton>
      <input ref={textInput} type="color" className="sr-only" tabIndex={-1} aria-hidden="true" defaultValue="#c00000" />
      <input ref={fillInput} type="color" className="sr-only" tabIndex={-1} aria-hidden="true" defaultValue="#ffff00" />
      <DropdownButton label="Borders" disabled={disabled} split={run(() => controller.applyBorders('all'))} menu={borderItems}>
        <Icon name="borders" size={16} />
      </DropdownButton>
      <Divider />

      <FormatButton label="Align left" title={`Align left (${k('L', true)})`} icon="alignLeft" active={style.hAlign === 'left'} disabled={disabled} onClick={run(() => controller.setAlignment('left'))} />
      <FormatButton label="Center" title={`Center (${k('E', true)})`} icon="alignCenter" active={style.hAlign === 'center'} disabled={disabled} onClick={run(() => controller.setAlignment('center'))} />
      <FormatButton label="Align right" title={`Align right (${k('R', true)})`} icon="alignRight" active={style.hAlign === 'right'} disabled={disabled} onClick={run(() => controller.setAlignment('right'))} />
      <DropdownButton label="Vertical alignment and indent" disabled={disabled} menu={verticalItems}>
        <Icon name={style.vAlign === 'top' ? 'alignTop' : style.vAlign === 'center' ? 'alignMiddle' : 'alignBottom'} size={16} />
      </DropdownButton>
      <FormatButton label="Wrap text" icon="wrap" active={Boolean(style.wrap)} disabled={disabled} onClick={run(() => controller.toggleWrap())} />
      <DropdownButton label="Merge & Center" disabled={disabled} split={run(() => controller.toggleMergeCenter())} menu={mergeItems}>
        <span className={`flex items-center rounded ${merged ? 'text-accent' : ''}`}>
          <Icon name="merge" size={16} />
        </span>
      </DropdownButton>
      <Divider />

      <DropdownButton label="Number format" title="Number format" disabled={disabled} menu={numberItems} className="w-[92px] justify-between pl-2 text-[12.5px]">
        <span className="truncate">{preset?.label ?? 'Custom'}</span>
      </DropdownButton>
      <FormatButton
        label="Currency format"
        title={`Currency (${isMac ? '⇧⌘4' : 'Ctrl+Shift+4'})`}
        disabled={disabled}
        onClick={run(() => controller.setNumberFormat(currencyFormat(localCurrency())))}
      >
        <span className="w-4 text-[13.5px] font-medium">{localCurrency()}</span>
      </FormatButton>
      <FormatButton label="Percent format" title={`Percent (${isMac ? '⇧⌘5' : 'Ctrl+Shift+5'})`} disabled={disabled} onClick={run(() => controller.setNumberFormat('0%'))}>
        <span className="w-4 text-[13.5px] font-medium">%</span>
      </FormatButton>
      <FormatButton label="Thousands separator" title="Thousands separator" disabled={disabled} onClick={run(() => controller.setNumberFormat('#,##0.00'))}>
        <span className="w-4 text-[15px] font-semibold leading-none">,</span>
      </FormatButton>
      <FormatButton label="Increase decimals" title="Increase decimals" disabled={disabled} onClick={run(() => controller.stepDecimals(1))}>
        <span className="text-[11px] font-medium tabular-nums">.0</span>
        <span className="text-[9px]">→</span>
        <span className="text-[11px] font-medium tabular-nums">.00</span>
      </FormatButton>
      <FormatButton label="Decrease decimals" title="Decrease decimals" disabled={disabled} onClick={run(() => controller.stepDecimals(-1))}>
        <span className="text-[11px] font-medium tabular-nums">.00</span>
        <span className="text-[9px]">→</span>
        <span className="text-[11px] font-medium tabular-nums">.0</span>
      </FormatButton>
      <Divider />

      <DropdownButton label="Insert" title="Insert rows, columns or a sheet" disabled={disabled} menu={insertItems} className="gap-1 px-1.5 text-[12.5px]">
        <Icon name="rowInsert" size={16} />
        <span>Insert</span>
      </DropdownButton>
      <DropdownButton label="Delete" title="Delete or hide rows and columns" disabled={disabled} menu={deleteItems} className="gap-1 px-1.5 text-[12.5px]">
        <Icon name="rowDelete" size={16} />
        <span>Delete</span>
      </DropdownButton>
      <Divider />
      <FormatButton label="Sort A to Z" title="Sort A → Z" icon="sortAsc" disabled={disabled} onClick={run(() => controller.sort(true))} />
      <FormatButton label="Sort Z to A" title="Sort Z → A" icon="sortDesc" disabled={disabled} onClick={run(() => controller.sort(false))} />
      <FormatButton label="Clear formatting" title={`Clear formatting (${k('\\')})`} icon="eraser" disabled={disabled} onClick={run(() => controller.clearFormats())} />
    </div>
  );
}
