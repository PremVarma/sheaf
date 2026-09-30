import { useLayoutEffect, useRef, type CSSProperties } from 'react';
import type { CommitMove } from '../../state/controller';
import { BASE_FONT_PX, measureText } from '../../utils/textMeasure';
import type { CellCss } from './cellCss';

export interface CellEditorProps {
  /** Cell rectangle in the grid container's coordinates. */
  rect: { left: number; top: number; width: number; height: number };
  /** Right edge the editor may grow to. */
  maxRight: number;
  draft: string;
  mode: 'enter' | 'edit';
  zoom: number;
  css: CellCss;
  /** Whether the editor owns the keyboard (false while the formula bar is being used). */
  focused: boolean;
  isMac: boolean;
  onChange(text: string): void;
  onCommit(move: CommitMove): void;
  onCancel(): void;
}

const ARROWS: Record<string, CommitMove> = { ArrowDown: 'down', ArrowUp: 'up', ArrowLeft: 'left', ArrowRight: 'right' };

/** Text box over the active cell while it's being edited. Grows with its content like Excel's. */
export function CellEditor(props: CellEditorProps) {
  const { rect, draft, zoom, css } = props;
  const ref = useRef<HTMLTextAreaElement>(null);
  const caret = useRef<number | null>(null);

  useLayoutEffect(() => {
    const element = ref.current;
    if (!props.focused || !element) return;
    element.focus({ preventScroll: true });
    element.setSelectionRange(element.value.length, element.value.length);
  }, [props.focused]);

  // Keep the caret after an inserted line break (a controlled value would move it to the end).
  useLayoutEffect(() => {
    if (caret.current === null || !ref.current) return;
    ref.current.setSelectionRange(caret.current, caret.current);
    caret.current = null;
  }, [draft]);

  const lines = draft.split('\n');
  const scale = css.fontScale * zoom;
  const widest = Math.max(0, ...lines.map((line) => measureText(line, css.bold, css.italic, false, css.fontFamily))) * scale;
  const width = Math.min(Math.max(rect.width + 1, widest + 14 * zoom), Math.max(rect.width + 1, props.maxRight - rect.left));
  const lineHeight = Math.round(BASE_FONT_PX * 1.3 * scale);
  const height = Math.max(rect.height + 1, lines.length * lineHeight + Math.round(6 * zoom));

  const style: CSSProperties = {
    left: rect.left - 1,
    top: rect.top - 1,
    width,
    height,
    fontSize: BASE_FONT_PX * scale,
    lineHeight: `${lineHeight}px`,
    fontWeight: css.bold ? 600 : undefined,
    fontStyle: css.italic ? 'italic' : undefined,
    fontFamily: css.fontFamily ? `${css.fontFamily}, var(--font-cell)` : undefined,
    textDecoration: css.style.textDecoration,
    padding: `${Math.max(1, Math.round(2 * zoom))}px ${Math.round(3 * zoom)}px`,
  };

  return (
    <textarea
      ref={ref}
      className="xv-editor"
      style={style}
      value={draft}
      rows={1}
      spellCheck={false}
      autoComplete="off"
      aria-label="Edit cell"
      data-cell-editor=""
      onChange={(event) => props.onChange(event.target.value)}
      onKeyDown={(event) => {
        if (event.nativeEvent.isComposing) return;
        const { key } = event;
        if (key === 'Enter') {
          // Option/Alt+Enter (and Ctrl+Enter on Windows) insert a line break, like Excel.
          if (event.altKey || (!props.isMac && event.ctrlKey)) {
            event.preventDefault();
            const el = event.currentTarget;
            caret.current = el.selectionStart + 1;
            props.onChange(`${el.value.slice(0, el.selectionStart)}\n${el.value.slice(el.selectionEnd)}`);
            return;
          }
          event.preventDefault();
          props.onCommit(event.shiftKey ? 'up' : 'down');
        } else if (key === 'Tab') {
          event.preventDefault();
          props.onCommit(event.shiftKey ? 'left' : 'right');
        } else if (key === 'Escape') {
          event.preventDefault();
          props.onCancel();
        } else if (props.mode === 'enter' && ARROWS[key] && !event.shiftKey && !event.metaKey && !event.ctrlKey) {
          // Typing into a cell: arrows finish the entry and move, as in Excel.
          event.preventDefault();
          props.onCommit(ARROWS[key]);
        }
      }}
      onBlur={(event) => {
        // Moving to the formula bar keeps the edit going.
        const next = event.relatedTarget as HTMLElement | null;
        if (next?.closest('[data-formula-bar]')) return;
        props.onCommit('none');
      }}
    />
  );
}
