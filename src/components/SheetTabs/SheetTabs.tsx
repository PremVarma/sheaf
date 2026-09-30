import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { useController, useViewer } from '../../state/AppContext';
import { Icon } from '../Icon';
import { Menu, type MenuItem } from '../Menu/Menu';
import { ToolButton } from '../ui';

/** Inline editor for a sheet name (double-click a tab, or Rename in its menu). */
function RenameField({ name, onDone }: { name: string; onDone(name: string | null): Promise<string | null> }) {
  const [value, setValue] = useState(name);
  const [error, setError] = useState<string | null>(null);
  const ref = useRef<HTMLInputElement>(null);
  const done = useRef(false);
  useLayoutEffect(() => {
    ref.current?.focus();
    ref.current?.select();
  }, []);
  const finish = async (next: string | null) => {
    if (done.current) return;
    const problem = await onDone(next);
    if (problem) {
      setError(problem);
      ref.current?.focus();
      return;
    }
    done.current = true;
  };
  return (
    <span className="relative flex items-center">
      <input
        ref={ref}
        aria-label="Sheet name"
        aria-invalid={error ? true : undefined}
        value={value}
        maxLength={31}
        spellCheck={false}
        className={`h-6 w-36 rounded border bg-control px-1.5 text-[12.5px] text-fg outline-none ${error ? 'border-danger' : 'border-accent'}`}
        onChange={(event) => {
          setValue(event.target.value);
          setError(null);
        }}
        onKeyDown={(event) => {
          event.stopPropagation();
          if (event.key === 'Enter') {
            event.preventDefault();
            void finish(value);
          } else if (event.key === 'Escape') {
            event.preventDefault();
            void finish(null);
          }
        }}
        onBlur={() => void finish(error ? null : value)}
      />
      {error && (
        <span role="alert" className="absolute bottom-8 left-0 z-50 w-64 rounded-md border border-danger/40 bg-surface px-2 py-1.5 text-[11.5px] text-danger shadow-lg">
          {error}
        </span>
      )}
    </span>
  );
}

/**
 * Excel-style sheet tabs: switch, add (+), rename (double-click), reorder
 * (drag), and a right-click menu to duplicate, delete, move, hide and unhide.
 */
export function SheetTabs() {
  const controller = useController();
  const workbook = useViewer((s) => s.workbook);
  const active = useViewer((s) => s.activeSheet);
  const stripRef = useRef<HTMLDivElement>(null);
  const [overflow, setOverflow] = useState({ left: false, right: false });
  const [listOpen, setListOpen] = useState(false);
  const listRef = useRef<HTMLDivElement>(null);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [menu, setMenu] = useState<{ index: number; x: number; y: number } | null>(null);
  const [drop, setDrop] = useState<{ from: number; to: number; x: number } | null>(null);
  const drag = useRef<{ index: number; startX: number; moved: boolean } | null>(null);

  const sheets = workbook?.sheets ?? [];
  const tabs = sheets.filter((s) => s.visibility === 'visible' || s.index === active);
  const hidden = sheets.filter((s) => s.visibility !== 'visible');
  const editable = workbook !== null;

  useEffect(() => {
    controller.requestRename = (index) => setRenaming(index);
    return () => {
      controller.requestRename = null;
    };
  }, [controller]);

  const measure = () => {
    const strip = stripRef.current;
    if (!strip) return;
    const next = { left: strip.scrollLeft > 0, right: strip.scrollLeft + strip.clientWidth < strip.scrollWidth - 1 };
    setOverflow((prev) => (prev.left === next.left && prev.right === next.right ? prev : next));
  };

  useLayoutEffect(() => {
    const strip = stripRef.current;
    if (!strip) return;
    measure();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(measure);
    observer.observe(strip);
    return () => observer.disconnect();
  }, [workbook]);

  useEffect(() => {
    const tab = stripRef.current?.querySelector<HTMLElement>('[aria-selected="true"]');
    tab?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }, [active, workbook]);

  const scrollBy = (direction: 1 | -1) => {
    const strip = stripRef.current;
    strip?.scrollBy({ left: direction * Math.max(120, strip.clientWidth * 0.6), behavior: 'smooth' });
  };

  /** Where a dragged tab would land, from the pointer position. */
  const dropTarget = (clientX: number, from: number): { to: number; x: number } => {
    const strip = stripRef.current;
    const elements = Array.from(strip?.querySelectorAll<HTMLElement>('[data-sheet-index]') ?? []);
    const stripLeft = strip?.getBoundingClientRect().left ?? 0;
    for (const el of elements) {
      const rect = el.getBoundingClientRect();
      if (clientX < rect.left + rect.width / 2) {
        const index = Number(el.dataset.sheetIndex);
        return { to: index > from ? index - 1 : index, x: rect.left - stripLeft + (strip?.scrollLeft ?? 0) };
      }
    }
    const last = elements[elements.length - 1];
    const rect = last?.getBoundingClientRect();
    return { to: sheets.length - 1, x: rect ? rect.right - stripLeft + (strip?.scrollLeft ?? 0) : 0 };
  };

  if (!workbook) return <div className="h-8 shrink-0 border-t border-line bg-tabs pointer-coarse:h-10" />;

  const tabMenu = (index: number): MenuItem[] => {
    const sheet = sheets[index];
    const visibleCount = sheets.filter((s) => s.visibility === 'visible').length;
    return [
      { label: 'Insert Sheet', icon: <Icon name="plus" size={14} />, shortcut: '⇧F11', onSelect: () => void controller.addSheet() },
      { label: 'Duplicate', onSelect: () => void controller.duplicateSheet(index) },
      { label: 'Rename', onSelect: () => setRenaming(index) },
      { label: 'Delete', danger: true, disabled: visibleCount <= 1 && sheet?.visibility === 'visible', onSelect: () => void controller.deleteSheet(index) },
      { kind: 'separator' },
      { label: 'Move Left', disabled: index === 0, onSelect: () => void controller.moveSheet(index, index - 1) },
      { label: 'Move Right', disabled: index === sheets.length - 1, onSelect: () => void controller.moveSheet(index, index + 1) },
      { kind: 'separator' },
      { label: 'Hide', disabled: visibleCount <= 1, onSelect: () => void controller.setSheetHidden(index, true) },
      hidden.length > 0
        ? { kind: 'submenu', label: 'Unhide', items: hidden.map((s) => ({ label: s.name, onSelect: () => void controller.setSheetHidden(s.index, false) })) }
        : { label: 'Unhide', disabled: true, onSelect: () => undefined },
    ];
  };

  return (
    <div className="relative flex h-8 shrink-0 items-stretch border-t border-line bg-tabs pointer-coarse:h-10">
      <div className="flex items-center gap-0.5 px-1.5">
        <ToolButton icon="chevronLeft" label="Scroll sheet tabs left" disabled={!overflow.left} onClick={() => scrollBy(-1)} />
        <ToolButton icon="chevronRight" label="Scroll sheet tabs right" disabled={!overflow.right} onClick={() => scrollBy(1)} />
        <div ref={listRef} className="relative">
          <ToolButton icon="list" label="All sheets" aria-haspopup="menu" aria-expanded={listOpen} onClick={() => setListOpen((v) => !v)} />
          {listOpen && (
            <Menu
              label="All sheets"
              placement={{ anchor: listRef.current!, side: 'below' }}
              ignore={[listRef.current]}
              onClose={() => setListOpen(false)}
              items={sheets.map((sheet) => ({
                label: sheet.visibility !== 'visible' ? `${sheet.name} (hidden)` : sheet.name,
                checked: sheet.index === active,
                onSelect: () => controller.activateSheet(sheet.index),
              }))}
            />
          )}
        </div>
      </div>

      <div
        ref={stripRef}
        role="tablist"
        aria-label="Worksheets"
        className="no-scrollbar relative flex min-w-0 items-stretch overflow-x-auto"
        onScroll={measure}
        onWheel={(event) => {
          if (Math.abs(event.deltaY) > Math.abs(event.deltaX)) stripRef.current?.scrollBy({ left: event.deltaY });
        }}
      >
        {tabs.map((sheet) => {
          const selected = sheet.index === active;
          if (renaming === sheet.index) {
            return (
              <div key={sheet.id} className="flex shrink-0 items-center border-r border-line bg-surface px-1.5">
                <RenameField
                  name={sheet.name}
                  onDone={async (name) => {
                    if (name === null || name.trim() === sheet.name) {
                      setRenaming(null);
                      controller.focusGrid();
                      return null;
                    }
                    const error = await controller.renameSheet(sheet.index, name);
                    if (!error) {
                      setRenaming(null);
                      controller.focusGrid();
                    }
                    return error;
                  }}
                />
              </div>
            );
          }
          return (
            <button
              key={sheet.id}
              type="button"
              role="tab"
              aria-selected={selected}
              data-sheet-index={sheet.index}
              title={sheet.visibility !== 'visible' ? `${sheet.name} (hidden sheet)` : `${sheet.name} — double-click to rename, drag to move`}
              className={`relative flex max-w-56 shrink-0 items-center gap-1.5 border-r border-line px-3.5 text-[12.5px] transition-colors ${
                selected ? 'bg-surface font-semibold text-accent' : 'text-muted hover:bg-hover hover:text-fg'
              } ${drop?.from === sheet.index ? 'opacity-50' : ''}`}
              onClick={() => {
                if (!drag.current?.moved) controller.activateSheet(sheet.index);
              }}
              onDoubleClick={() => editable && setRenaming(sheet.index)}
              onContextMenu={(event) => {
                event.preventDefault();
                controller.activateSheet(sheet.index);
                setMenu({ index: sheet.index, x: event.clientX, y: event.clientY });
              }}
              onPointerDown={(event) => {
                if (event.button !== 0) return;
                drag.current = { index: sheet.index, startX: event.clientX, moved: false };
                const onMove = (e: PointerEvent) => {
                  const d = drag.current;
                  if (!d) return;
                  if (!d.moved && Math.abs(e.clientX - d.startX) < 6) return;
                  d.moved = true;
                  setDrop({ from: d.index, ...dropTarget(e.clientX, d.index) });
                };
                const onUp = (e: PointerEvent) => {
                  window.removeEventListener('pointermove', onMove);
                  window.removeEventListener('pointerup', onUp);
                  const d = drag.current;
                  setDrop(null);
                  if (d?.moved) {
                    const { to } = dropTarget(e.clientX, d.index);
                    if (to !== d.index) void controller.moveSheet(d.index, to);
                    // Swallow the click that follows the drag.
                    setTimeout(() => {
                      drag.current = null;
                    }, 0);
                  } else drag.current = null;
                };
                window.addEventListener('pointermove', onMove);
                window.addEventListener('pointerup', onUp);
              }}
            >
              {selected && <span className="absolute inset-x-0 bottom-0 h-0.5 bg-accent" aria-hidden="true" />}
              {sheet.visibility !== 'visible' && <Icon name="eyeOff" size={12} />}
              <span className="truncate">{sheet.name}</span>
            </button>
          );
        })}
        {drop && <span className="pointer-events-none absolute bottom-1 top-1 w-0.5 rounded bg-accent" style={{ left: drop.x - 1 }} aria-hidden="true" />}
      </div>
      <div className="flex items-center px-1">
        <ToolButton icon="plus" label="New sheet" title="New sheet (⇧F11)" disabled={!editable} onClick={() => void controller.addSheet()} />
      </div>
      <div className="flex-1" />
      {hidden.length > 0 && (
        <div className="flex items-center px-3 text-[11.5px] text-muted max-sm:hidden" title="Right-click a tab to unhide sheets">
          {hidden.length} hidden
        </div>
      )}
      {menu && (
        <Menu
          label={`Sheet ${sheets[menu.index]?.name ?? ''}`}
          items={tabMenu(menu.index)}
          placement={{ x: menu.x, y: menu.y }}
          onClose={() => setMenu(null)}
        />
      )}
    </div>
  );
}
