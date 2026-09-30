import { useController, usePlatform, useViewer } from '../../state/AppContext';
import { SearchBar } from '../SearchBar/SearchBar';
import { shortcutLabel, ToolButton } from '../ui';
import { FormatBar } from './FormatBar';
import { FormulaBar } from './FormulaBar';
import { NameBox } from './NameBox';
import { ZoomControl } from './ZoomControl';

export function Toolbar() {
  const controller = useController();
  const { isMac } = usePlatform();
  const hasWorkbook = useViewer((s) => s.workbook !== null);
  const infoOpen = useViewer((s) => s.infoOpen);
  const dirty = useViewer((s) => s.document.dirty);
  const saving = useViewer((s) => s.document.saving);
  const canUndo = useViewer((s) => s.document.canUndo);
  const canRedo = useViewer((s) => s.document.canRedo);
  // On phones an active search takes the whole row.
  const compact = useViewer((s) => s.search.query !== '') ? 'max-sm:hidden' : '';

  return (
    <header className="shrink-0 border-b border-line bg-toolbar">
      <div className="flex h-11 items-center gap-1 px-3">
        <div className={`flex items-center gap-1 ${compact}`}>
          <ToolButton icon="newFile" label="New workbook" title={`New workbook (${shortcutLabel('N', isMac)})`} onClick={() => void controller.newWorkbook()} />
          <ToolButton icon="open" label="Open" title={`Open a workbook (${shortcutLabel('O', isMac)})`} className="px-2.5" onClick={() => void controller.openDialog()}>
            <span className="text-[12.5px] max-sm:hidden">Open</span>
          </ToolButton>
          <ToolButton
            icon="save"
            label="Save"
            title={`Save (${shortcutLabel('S', isMac)})`}
            disabled={!hasWorkbook || saving}
            active={dirty}
            onClick={() => void controller.save()}
          />
          <div className="mx-1 h-5 w-px bg-line" aria-hidden="true" />
          <ToolButton icon="undo" label="Undo" title={`Undo (${shortcutLabel('Z', isMac)})`} disabled={!canUndo} onClick={() => void controller.undo()} />
          <ToolButton icon="redo" label="Redo" title={`Redo (${shortcutLabel('Z', isMac, true)})`} disabled={!canRedo} onClick={() => void controller.redo()} />
          <div className="mx-1 h-5 w-px bg-line" aria-hidden="true" />
        </div>
        <SearchBar />
        <div className={`min-w-2 flex-1 ${compact}`} />
        {/* Phones pinch to zoom; the control doesn't fit there. */}
        <div className="flex items-center max-sm:hidden">
          <ZoomControl />
          <div className="mx-1 h-5 w-px bg-line" aria-hidden="true" />
        </div>
        <ToolButton
          icon="info"
          label="Workbook information"
          title={`Workbook information (${isMac ? '⌥⌘I' : 'Ctrl+Alt+I'})`}
          active={infoOpen}
          aria-pressed={infoOpen}
          disabled={!hasWorkbook}
          className={compact}
          onClick={() => controller.setInfoOpen(!infoOpen)}
        />
      </div>
      {hasWorkbook && <FormatBar />}
      <div className="flex h-8 items-center gap-2 border-t border-line/70 px-3">
        <NameBox />
        <div className="h-4 w-px bg-line" aria-hidden="true" />
        <FormulaBar />
      </div>
    </header>
  );
}
