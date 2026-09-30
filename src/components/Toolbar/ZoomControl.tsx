import { useEffect, useRef, useState } from 'react';
import { useController, usePlatform, useViewer } from '../../state/AppContext';
import { MAX_ZOOM, MIN_ZOOM, ZOOM_LEVELS } from '../../state/viewerState';
import { shortcutLabel, ToolButton } from '../ui';

/** − 100% + with a preset menu on the percentage. */
export function ZoomControl() {
  const controller = useController();
  const { isMac } = usePlatform();
  const zoom = useViewer((s) => s.zoom);
  const enabled = useViewer((s) => s.workbook !== null);
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false);
    };
    window.addEventListener('pointerdown', onPointerDown);
    window.addEventListener('keydown', onKeyDown);
    return () => {
      window.removeEventListener('pointerdown', onPointerDown);
      window.removeEventListener('keydown', onKeyDown);
    };
  }, [open]);

  const percent = `${Math.round(zoom * 100)}%`;
  return (
    <div ref={rootRef} className="relative flex shrink-0 items-center" role="group" aria-label="Zoom">
      <ToolButton
        icon="minus"
        label="Zoom out"
        title={`Zoom out (${shortcutLabel('−', isMac)})`}
        disabled={!enabled || zoom <= MIN_ZOOM}
        onClick={() => controller.stepZoom(-1)}
      />
      <button
        type="button"
        className="h-7 w-12 rounded-md text-center text-[12px] tabular-nums text-fg hover:bg-hover disabled:opacity-40"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={`Zoom ${percent}`}
        title="Zoom level"
        disabled={!enabled}
        onClick={() => setOpen((value) => !value)}
      >
        {percent}
      </button>
      <ToolButton
        icon="plus"
        label="Zoom in"
        title={`Zoom in (${shortcutLabel('+', isMac)})`}
        disabled={!enabled || zoom >= MAX_ZOOM}
        onClick={() => controller.stepZoom(1)}
      />
      {open && (
        <div role="menu" className="absolute right-0 top-8 z-50 w-32 rounded-lg border border-line bg-surface p-1 shadow-xl">
          {ZOOM_LEVELS.map((level) => (
            <button
              key={level}
              type="button"
              role="menuitemradio"
              aria-checked={Math.abs(level - zoom) < 0.001}
              className="flex h-7 w-full items-center justify-between rounded-md px-2 text-[12.5px] text-fg hover:bg-accent hover:text-accent-fg"
              onClick={() => {
                controller.setZoom(level);
                setOpen(false);
              }}
            >
              <span>{Math.round(level * 100)}%</span>
              {Math.abs(level - zoom) < 0.001 && <span aria-hidden="true">✓</span>}
            </button>
          ))}
          <div className="my-1 border-t border-line" />
          <button
            type="button"
            role="menuitem"
            className="flex h-7 w-full items-center justify-between rounded-md px-2 text-[12.5px] text-fg hover:bg-accent hover:text-accent-fg"
            onClick={() => {
              controller.setZoom(1);
              setOpen(false);
            }}
          >
            <span>Actual size</span>
            <span className="text-[11px] opacity-70">{shortcutLabel('0', isMac)}</span>
          </button>
        </div>
      )}
    </div>
  );
}
