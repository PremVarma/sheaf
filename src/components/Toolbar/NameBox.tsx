import { useEffect, useRef, useState } from 'react';
import { mergeAt } from '../../services/workbook/sheetService';
import { useController, useViewer } from '../../state/AppContext';
import { normalizeRange, rangeAddress } from '../../utils/cellAddress';

/** Shows the selection's address; type a reference such as A152 and press Enter to go there. */
export function NameBox() {
  const controller = useController();
  const enabled = useViewer((s) => s.workbook !== null);
  const address = useViewer((s) => {
    const view = s.sheetViews[s.activeSheet];
    const sheet = s.workbook?.sheets[s.activeSheet];
    if (!view || !sheet) return '';
    const range = normalizeRange(view.selection.anchor, view.selection.focus);
    // A merged cell reads as its top-left address, like Excel.
    const merge = mergeAt(sheet, range.r0, range.c0);
    if (merge && merge.r0 <= range.r0 && merge.r1 >= range.r1 && merge.c0 <= range.c0 && merge.c1 >= range.c1) {
      return rangeAddress({ r0: merge.r0, c0: merge.c0, r1: merge.r0, c1: merge.c0 });
    }
    return rangeAddress(range);
  });
  const focusRequest = useViewer((s) => s.focusRequest);
  const [draft, setDraft] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (focusRequest?.target !== 'nameBox') return;
    inputRef.current?.focus();
    inputRef.current?.select();
  }, [focusRequest]);

  const submit = () => {
    const message = controller.goToReference(draft ?? address);
    if (message) {
      setError(message);
      inputRef.current?.select();
    } else {
      setDraft(null);
      setError(null);
    }
  };

  return (
    <div className="relative shrink-0">
      <input
        ref={inputRef}
        className={`h-6 w-24 rounded-md border bg-control px-2 text-[12px] tabular-nums text-fg outline-none placeholder:text-muted focus:border-accent focus:ring-2 focus:ring-accent/25 disabled:opacity-50 ${
          error ? 'border-danger ring-2 ring-danger/20' : 'border-control-line'
        }`}
        value={draft ?? address}
        placeholder="A1"
        disabled={!enabled}
        spellCheck={false}
        autoComplete="off"
        aria-label="Name Box: go to a cell"
        aria-invalid={error ? true : undefined}
        aria-describedby={error ? 'namebox-error' : undefined}
        title="Go to a cell, e.g. A152"
        onFocus={(event) => {
          setDraft(address);
          const input = event.currentTarget;
          setTimeout(() => input.select(), 0);
        }}
        onChange={(event) => {
          setDraft(event.target.value);
          setError(null);
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            submit();
          } else if (event.key === 'Escape') {
            event.preventDefault();
            setDraft(null);
            setError(null);
            controller.focusGrid();
          }
        }}
        onBlur={() => {
          setDraft(null);
          setError(null);
        }}
      />
      {error && (
        <div
          id="namebox-error"
          role="alert"
          className="absolute left-0 top-8 z-50 w-64 rounded-md border border-line bg-surface px-3 py-2 text-[12px] text-fg shadow-lg"
        >
          {error}
        </div>
      )}
    </div>
  );
}
