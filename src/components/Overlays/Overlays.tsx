import { useEffect, useRef, useState } from 'react';
import { LARGE_FILE_BYTES } from '../../services/workbook/formatDetection';
import { useController, useViewer } from '../../state/AppContext';
import { formatBytes } from '../../utils/format';
import { Icon } from '../Icon';

/** Progress while a file is read and parsed; appears after a short delay to avoid flicker. */
export function LoadingOverlay() {
  const controller = useController();
  const loading = useViewer((s) => s.loading);
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    if (!loading) {
      setVisible(false);
      return;
    }
    const timer = setTimeout(() => setVisible(true), 150);
    return () => clearTimeout(timer);
  }, [loading]);

  if (!loading || !visible) return null;
  const stage = loading.stage === 'parsing' ? 'Reading workbook…' : 'Loading file…';
  const large = (loading.size ?? 0) >= LARGE_FILE_BYTES;

  return (
    <div className="absolute inset-0 z-40 flex items-center justify-center bg-app/70 backdrop-blur-[2px]" role="dialog" aria-modal="true" aria-label="Opening file">
      <div className="flex w-80 flex-col items-center rounded-2xl border border-line bg-surface px-6 py-6 text-center shadow-2xl">
        <div className="h-7 w-7 animate-spin rounded-full border-[3px] border-accent/25 border-t-accent" aria-hidden="true" />
        <p className="mt-4 w-full truncate text-[13px] font-semibold text-fg" title={loading.fileName}>
          Opening “{loading.fileName}”
        </p>
        <p className="mt-1 text-[12px] text-muted" aria-live="polite">
          {stage}
          {loading.size ? ` · ${formatBytes(loading.size)}` : ''}
        </p>
        {large && <p className="mt-2 text-[11.5px] text-muted">Large workbooks can take a little while.</p>}
        <button
          type="button"
          className="mt-5 h-7 rounded-md border border-control-line bg-control px-4 text-[12.5px] text-fg hover:bg-hover"
          onClick={() => controller.cancelLoad()}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}

/** Friendly error for files that couldn't be opened. Never shows technical details. */
export function ErrorDialog() {
  const controller = useController();
  const error = useViewer((s) => s.error);
  const buttonRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    if (!error) return;
    buttonRef.current?.focus();
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' || event.key === 'Enter') {
        event.preventDefault();
        event.stopPropagation();
        controller.dismissError();
      }
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [error, controller]);

  if (!error) return null;
  return (
    <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/25 p-6">
      <div
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="error-title"
        aria-describedby="error-message"
        className="w-full max-w-sm rounded-2xl border border-line bg-surface p-6 shadow-2xl"
      >
        <div className="flex items-start gap-3">
          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-warning/15 text-warning">
            <Icon name="warning" size={18} />
          </div>
          <div className="min-w-0">
            <h2 id="error-title" className="break-words text-[14px] font-semibold text-fg">
              {error.title}
            </h2>
            <p id="error-message" className="mt-1.5 text-[12.5px] leading-relaxed text-muted">
              {error.message}
            </p>
          </div>
        </div>
        <div className="mt-6 flex justify-end gap-2">
          <button
            type="button"
            className="h-7 rounded-md border border-control-line bg-control px-3.5 text-[12.5px] text-fg hover:bg-hover"
            onClick={() => {
              controller.dismissError();
              void controller.openDialog();
            }}
          >
            Open Another File…
          </button>
          <button
            ref={buttonRef}
            type="button"
            className="h-7 rounded-md bg-accent px-4 text-[12.5px] font-medium text-accent-fg hover:brightness-110"
            onClick={() => controller.dismissError()}
          >
            OK
          </button>
        </div>
      </div>
    </div>
  );
}

/** Full-window hint while a file is dragged over the app. */
export function DropOverlay() {
  const active = useViewer((s) => s.dragActive);
  if (!active) return null;
  return (
    <div className="pointer-events-none absolute inset-3 z-50 flex items-center justify-center rounded-2xl border-2 border-dashed border-accent bg-accent-soft/85">
      <div className="flex flex-col items-center text-accent">
        <Icon name="drop" size={36} />
        <p className="mt-2 text-[15px] font-semibold">Drop to open</p>
      </div>
    </div>
  );
}

/** Questions and notices from the app (delete a sheet? merge cells? can't insert…). */
export function PromptDialog() {
  const controller = useController();
  const dialog = useViewer((s) => s.dialog);
  const confirmRef = useRef<HTMLButtonElement>(null);
  const cancelRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    if (!dialog) return;
    // Destructive actions start on Cancel so Enter doesn't delete by accident.
    (dialog.destructive ? cancelRef.current : confirmRef.current)?.focus();
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        controller.resolveDialog(dialog.id, false);
      }
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [dialog, controller]);

  if (!dialog) return null;
  return (
    <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/25 p-6">
      <div
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="prompt-title"
        aria-describedby="prompt-message"
        className="w-full max-w-sm rounded-2xl border border-line bg-surface p-6 shadow-2xl"
      >
        <h2 id="prompt-title" className="text-[14px] font-semibold text-fg">
          {dialog.title}
        </h2>
        <p id="prompt-message" className="mt-1.5 text-[12.5px] leading-relaxed text-muted">
          {dialog.message}
        </p>
        <div className="mt-6 flex justify-end gap-2">
          {dialog.kind === 'confirm' && (
            <button
              ref={cancelRef}
              type="button"
              className="h-7 rounded-md border border-control-line bg-control px-3.5 text-[12.5px] text-fg hover:bg-hover"
              onClick={() => controller.resolveDialog(dialog.id, false)}
            >
              {dialog.cancelLabel ?? 'Cancel'}
            </button>
          )}
          <button
            ref={confirmRef}
            type="button"
            className={`h-7 rounded-md px-4 text-[12.5px] font-medium hover:brightness-110 ${
              dialog.destructive ? 'bg-danger text-white' : 'bg-accent text-accent-fg'
            }`}
            onClick={() => controller.resolveDialog(dialog.id, true)}
          >
            {dialog.confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
