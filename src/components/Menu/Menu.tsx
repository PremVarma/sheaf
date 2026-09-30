import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

/** Menus and floating panels: toolbar dropdowns, right-click menus, the color picker. */

export type MenuItem =
  | {
      kind?: 'item';
      label: string;
      shortcut?: string;
      icon?: ReactNode;
      checked?: boolean;
      disabled?: boolean;
      danger?: boolean;
      /** Text shown on the right instead of a shortcut (e.g. a formatted sample). */
      detail?: string;
      onSelect(): void;
    }
  | { kind: 'submenu'; label: string; icon?: ReactNode; disabled?: boolean; items: MenuItem[] }
  | { kind: 'separator' }
  | { kind: 'label'; label: string };

export type Placement =
  | { x: number; y: number }
  /** Below (or beside, for submenus) an element. */
  | { anchor: HTMLElement; side?: 'below' | 'right'; align?: 'start' | 'end' };

/** Keeps a floating element inside the window. */
function place(placement: Placement, width: number, height: number): { left: number; top: number } {
  const margin = 6;
  const vw = window.innerWidth || 1024;
  const vh = window.innerHeight || 768;
  let left: number;
  let top: number;
  if ('anchor' in placement) {
    const rect = placement.anchor.getBoundingClientRect();
    if (placement.side === 'right') {
      left = rect.right - 2;
      top = rect.top - 4;
      if (left + width > vw - margin) left = rect.left - width + 2;
    } else {
      left = placement.align === 'end' ? rect.right - width : rect.left;
      top = rect.bottom + 4;
      if (top + height > vh - margin && rect.top - height - 4 > margin) top = rect.top - height - 4;
    }
  } else {
    left = placement.x;
    top = placement.y;
    if (top + height > vh - margin) top = Math.max(margin, placement.y - height);
  }
  return {
    left: Math.max(margin, Math.min(left, vw - width - margin)),
    top: Math.max(margin, Math.min(top, vh - height - margin)),
  };
}

interface PopoverProps {
  placement: Placement;
  onClose(): void;
  children: ReactNode;
  className?: string;
  /** Elements that don't count as "outside" (the button that opened it). */
  ignore?: (HTMLElement | null)[];
  role?: string;
  label?: string;
}

/** A floating panel that closes on Escape or a click outside it. */
export function Popover({ placement, onClose, children, className = '', ignore = [], role, label }: PopoverProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState<{ left: number; top: number } | null>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  const ignoreRef = useRef(ignore);
  ignoreRef.current = ignore;

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    setPosition(place(placement, el.offsetWidth, el.offsetHeight));
    // eslint-disable-next-line react-hooks/exhaustive-deps -- measured once when opened
  }, []);

  useEffect(() => {
    const onPointer = (event: PointerEvent) => {
      const target = event.target as Node;
      if (ref.current?.contains(target)) return;
      if (ignoreRef.current.some((el) => el?.contains(target))) return;
      // Clicks inside another open menu (a submenu) don't close this one.
      if ((target as HTMLElement).closest?.('[data-floating]')) return;
      onCloseRef.current();
    };
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        onCloseRef.current();
      }
    };
    const onBlur = () => onCloseRef.current();
    window.addEventListener('pointerdown', onPointer, true);
    window.addEventListener('keydown', onKey, true);
    window.addEventListener('blur', onBlur);
    return () => {
      window.removeEventListener('pointerdown', onPointer, true);
      window.removeEventListener('keydown', onKey, true);
      window.removeEventListener('blur', onBlur);
    };
  }, []);

  const style: CSSProperties = position ? { left: position.left, top: position.top } : { left: -9999, top: -9999 };
  return createPortal(
    <div
      ref={ref}
      data-floating=""
      role={role}
      aria-label={label}
      className={`fixed z-[60] rounded-lg border border-line bg-surface shadow-xl ${className}`}
      style={style}
    >
      {children}
    </div>,
    document.body,
  );
}

interface MenuProps {
  items: MenuItem[];
  placement: Placement;
  onClose(): void;
  label?: string;
  ignore?: (HTMLElement | null)[];
  /** Focus the first item (keyboard-opened menus). */
  autoFocus?: boolean;
}

/** A list of commands with keyboard navigation (arrows, Enter, Escape, → for submenus). */
export function Menu({ items, placement, onClose, label, ignore, autoFocus = true }: MenuProps) {
  const listRef = useRef<HTMLDivElement>(null);
  const [submenu, setSubmenu] = useState<{ index: number; anchor: HTMLElement } | null>(null);
  const hoverTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (!autoFocus) return;
    const first = listRef.current?.querySelector<HTMLElement>('[role^="menuitem"]:not([aria-disabled="true"])');
    first?.focus({ preventScroll: true });
  }, [autoFocus]);

  useEffect(() => () => clearTimeout(hoverTimer.current), []);

  const focusable = () => Array.from(listRef.current?.querySelectorAll<HTMLElement>(':scope > [role^="menuitem"]:not([aria-disabled="true"])') ?? []);

  const onKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const list = focusable();
    const index = list.indexOf(document.activeElement as HTMLElement);
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      const step = event.key === 'ArrowDown' ? 1 : -1;
      list[(index + step + list.length) % list.length]?.focus();
    } else if (event.key === 'Home' || event.key === 'End') {
      event.preventDefault();
      (event.key === 'Home' ? list[0] : list[list.length - 1])?.focus();
    } else if (event.key === 'ArrowRight') {
      const el = document.activeElement as HTMLElement;
      const i = Number(el?.dataset.index);
      if (items[i]?.kind === 'submenu') {
        event.preventDefault();
        setSubmenu({ index: i, anchor: el });
      }
    }
  };

  return (
    <Popover placement={placement} onClose={onClose} ignore={ignore} className="min-w-44 max-w-80 p-1" role="presentation">
      <div ref={listRef} role="menu" aria-label={label} tabIndex={-1} className="flex flex-col outline-none" onKeyDown={onKeyDown}>
        {items.map((item, index) => {
          if (item.kind === 'separator') return <div key={index} role="separator" className="my-1 h-px bg-line" />;
          if (item.kind === 'label') {
            return (
              <div key={index} className="px-2 pb-0.5 pt-1.5 text-[11px] font-semibold uppercase tracking-wide text-muted">
                {item.label}
              </div>
            );
          }
          const disabled = Boolean(item.disabled);
          const isSubmenu = item.kind === 'submenu';
          const open = submenu?.index === index;
          return (
            <button
              key={index}
              type="button"
              role={item.kind !== 'submenu' && item.checked !== undefined ? 'menuitemcheckbox' : 'menuitem'}
              aria-checked={item.kind !== 'submenu' && item.checked !== undefined ? item.checked : undefined}
              aria-disabled={disabled || undefined}
              aria-haspopup={isSubmenu ? 'menu' : undefined}
              aria-expanded={isSubmenu ? open : undefined}
              data-index={index}
              tabIndex={-1}
              className={`flex h-7 w-full shrink-0 items-center gap-2 rounded-md px-2 text-left text-[12.5px] outline-none ${
                disabled
                  ? 'cursor-default text-muted opacity-50'
                  : `${item.kind !== 'submenu' && item.danger ? 'text-danger' : 'text-fg'} hover:bg-accent hover:text-accent-fg focus:bg-accent focus:text-accent-fg ${open ? 'bg-hover' : ''}`
              }`}
              onPointerEnter={(event) => {
                if (disabled) return;
                const el = event.currentTarget;
                el.focus({ preventScroll: true });
                clearTimeout(hoverTimer.current);
                hoverTimer.current = setTimeout(() => setSubmenu(isSubmenu ? { index, anchor: el } : null), isSubmenu ? 120 : 200);
              }}
              onClick={(event) => {
                if (disabled) return;
                if (item.kind === 'submenu') {
                  setSubmenu({ index, anchor: event.currentTarget });
                  return;
                }
                onClose();
                item.onSelect();
              }}
            >
              <span className="flex w-4 shrink-0 items-center justify-center">
                {item.kind !== 'submenu' && item.checked ? '✓' : item.icon}
              </span>
              <span className="min-w-0 flex-1 truncate">{item.label}</span>
              {item.kind !== 'submenu' && (item.detail ?? item.shortcut) && (
                <span className="ml-4 shrink-0 text-[11.5px] opacity-60">{item.detail ?? item.shortcut}</span>
              )}
              {isSubmenu && <span className="ml-4 shrink-0 opacity-60">›</span>}
            </button>
          );
        })}
      </div>
      {submenu && items[submenu.index]?.kind === 'submenu' && (
        <Menu
          items={(items[submenu.index] as Extract<MenuItem, { kind: 'submenu' }>).items}
          placement={{ anchor: submenu.anchor, side: 'right' }}
          onClose={() => {
            setSubmenu(null);
            onClose();
          }}
          autoFocus={false}
        />
      )}
    </Popover>
  );
}
