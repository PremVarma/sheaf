import { useEffect, useRef, useState } from 'react';
import { editTextAt } from '../../services/editing/editText';
import { useController, useViewer } from '../../state/AppContext';

/**
 * The active cell's contents: its formula, raw number or text. Editable: typing
 * here edits the cell (shared with the in-cell editor); Enter confirms, Esc cancels.
 */
export function FormulaBar() {
  const controller = useController();
  const editing = useViewer((s) => (s.editing && s.editing.sheet === s.activeSheet ? s.editing : null));
  const canEdit = useViewer((s) => Boolean(s.workbook && s.workbook.sheets[s.activeSheet]?.kind === 'worksheet'));
  const shown = useViewer((s) => {
    const view = s.sheetViews[s.activeSheet];
    const sheet = s.workbook?.sheets[s.activeSheet];
    if (!view || !sheet || !s.workbook) return '';
    const { row, col } = view.selection.anchor;
    return s.editing ? '' : editTextAt(s.workbook, s.activeSheet, row, col);
  });
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [focused, setFocused] = useState(false);

  // Moving focus here from the in-cell editor keeps the edit going.
  useEffect(() => {
    if (editing?.source === 'formulaBar' && document.activeElement !== inputRef.current) inputRef.current?.focus();
  }, [editing?.source]);

  const value = editing ? editing.draft : shown;
  const isFormula = value.startsWith('=');

  return (
    <div className="flex min-w-0 flex-1 items-center gap-2" data-formula-bar="">
      <span className="select-none font-serif text-[13px] italic text-muted" aria-hidden="true">
        fx
      </span>
      {/* A one-row textarea (not an input) so line breaks inside cells survive editing. */}
      <textarea
        ref={inputRef}
        rows={1}
        className={`h-6 min-w-0 flex-1 resize-none overflow-hidden whitespace-pre rounded-sm border bg-transparent px-1.5 py-[3px] text-[12.5px] leading-4 text-fg outline-none ${
          focused ? 'border-accent bg-control' : 'border-transparent'
        } ${isFormula ? 'font-mono text-[12px]' : ''}`}
        value={value}
        disabled={!canEdit}
        spellCheck={false}
        autoComplete="off"
        aria-label="Cell contents"
        data-testid="formula-bar"
        onFocus={() => {
          setFocused(true);
          controller.startEdit('edit', undefined, 'formulaBar');
        }}
        onBlur={(event) => {
          setFocused(false);
          const next = event.relatedTarget as HTMLElement | null;
          if (next?.closest('[data-cell-editor]')) return;
          void controller.commitEdit('none');
        }}
        onChange={(event) => controller.setDraft(event.target.value, 'formulaBar')}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && !event.altKey) {
            event.preventDefault();
            void controller.commitEdit(event.shiftKey ? 'up' : 'down');
          } else if (event.key === 'Tab') {
            event.preventDefault();
            void controller.commitEdit(event.shiftKey ? 'left' : 'right');
          } else if (event.key === 'Escape') {
            event.preventDefault();
            controller.cancelEdit();
          }
        }}
      />
    </div>
  );
}
