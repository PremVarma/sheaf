import { useRef, useState } from 'react';
import { useController, usePlatform, useViewer } from '../../state/AppContext';
import type { CommandId } from '../../state/commands';
import { Icon } from '../Icon';
import { Menu, type MenuItem } from '../Menu/Menu';
import { SearchBar } from '../SearchBar/SearchBar';
import { shortcutLabel, ToolButton } from '../ui';
import { FormatBar, FormatSheet } from './FormatBar';
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
  const searching = useViewer((s) => s.search.query !== '' || s.replaceOpen);
  // Phones open search and the formatting bar on demand, on lines of their own;
  // larger screens always show them.
  const [searchOpen, setSearchOpen] = useState(false);
  const [formatOpen, setFormatOpen] = useState(false);
  const showSearch = searchOpen || searching;
  const moreRef = useRef<HTMLDivElement>(null);
  const [moreOpen, setMoreOpen] = useState(false);
  const run = (id: CommandId) => () => controller.runCommand(id);
  // Phones have no menu bar: the commands only found there.
  const moreItems: MenuItem[] = [
    { label: 'Save As…', icon: <Icon name="save" size={15} />, onSelect: run('file.saveAs') },
    { label: 'Close Workbook', icon: <Icon name="close" size={15} />, onSelect: run('file.close') },
    { kind: 'separator' },
    { label: 'Freeze Top Row', onSelect: run('view.freezeTopRow') },
    { label: 'Freeze First Column', onSelect: run('view.freezeFirstColumn') },
    { label: 'Freeze at Selection', onSelect: run('view.freezeAtSelection') },
    { label: 'Unfreeze Panes', onSelect: run('view.unfreeze') },
    { label: 'Show or Hide Gridlines', onSelect: run('view.toggleGridlines') },
    { kind: 'separator' },
    { label: 'Zoom In', icon: <Icon name="plus" size={15} />, onSelect: run('view.zoomIn') },
    { label: 'Zoom Out', icon: <Icon name="minus" size={15} />, onSelect: run('view.zoomOut') },
    { label: 'Actual Size', onSelect: run('view.zoomReset') },
    { kind: 'separator' },
    { label: 'Workbook Information', icon: <Icon name="info" size={15} />, checked: infoOpen, onSelect: run('view.toggleInfo') },
  ];

  return (
    <header className="shrink-0 border-b border-line bg-toolbar">
      <div className="flex min-h-11 flex-wrap items-center gap-1 px-3">
        <div className="flex items-center gap-1">
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
        <SearchBar className={showSearch ? 'max-sm:order-last max-sm:basis-full max-sm:pb-1.5' : 'max-sm:hidden'} />
        <div className="min-w-2 flex-1" />
        {/* Phones pinch to zoom; the control doesn't fit there. */}
        <div className="flex items-center max-sm:hidden">
          <ZoomControl />
          <div className="mx-1 h-5 w-px bg-line" aria-hidden="true" />
        </div>
        <ToolButton
          icon="search"
          label="Search"
          active={showSearch}
          aria-pressed={showSearch}
          disabled={!hasWorkbook}
          className="sm:hidden"
          onClick={() => {
            if (showSearch) {
              controller.clearSearch();
              setSearchOpen(false);
            } else {
              setSearchOpen(true);
              controller.runCommand('edit.find');
            }
          }}
        />
        <ToolButton
          label="Formatting"
          title="Show formatting tools"
          active={formatOpen}
          aria-pressed={formatOpen}
          disabled={!hasWorkbook}
          className="sm:hidden"
          onClick={() => setFormatOpen((open) => !open)}
        >
          <span className="text-[13px] font-semibold">Aa</span>
        </ToolButton>
        <ToolButton
          icon="info"
          label="Workbook information"
          title={`Workbook information (${isMac ? '⌥⌘I' : 'Ctrl+Alt+I'})`}
          active={infoOpen}
          aria-pressed={infoOpen}
          disabled={!hasWorkbook}
          className="max-sm:hidden"
          onClick={() => controller.setInfoOpen(!infoOpen)}
        />
        <div ref={moreRef} className="sm:hidden">
          <ToolButton icon="more" label="More" aria-haspopup="menu" aria-expanded={moreOpen} disabled={!hasWorkbook} onClick={() => setMoreOpen((open) => !open)} />
          {moreOpen && (
            <Menu
              label="More"
              placement={{ anchor: moreRef.current!, side: 'below', align: 'end' }}
              ignore={[moreRef.current]}
              autoFocus={false}
              onClose={() => setMoreOpen(false)}
              items={moreItems}
            />
          )}
        </div>
      </div>
      {hasWorkbook && <FormatBar className="max-sm:hidden" />}
      {hasWorkbook && formatOpen && <FormatSheet onClose={() => setFormatOpen(false)} />}
      <div className="flex h-8 items-center gap-2 border-t border-line/70 px-3 pointer-coarse:h-10">
        <NameBox />
        <div className="h-4 w-px bg-line" aria-hidden="true" />
        <FormulaBar />
      </div>
    </header>
  );
}
