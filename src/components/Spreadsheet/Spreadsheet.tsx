import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CellRange, CellRef } from '../../models/workbook';
import { useController, usePlatform, useViewer } from '../../state/AppContext';
import type { CommitMove } from '../../state/controller';
import type { CellEditState, Selection } from '../../state/viewerState';
import { isWholeColumns, isWholeRows, normalizeRange } from '../../utils/cellAddress';
import { Icon } from '../Icon';
import { Menu, type MenuItem } from '../Menu/Menu';
import { SpreadsheetGrid, type GridHitKind } from '../SpreadsheetGrid/SpreadsheetGrid';

const SHEET_KIND_LABELS = { chart: 'chart sheet', macro: 'macro sheet', dialog: 'dialog sheet' } as const;

/** Connects the active worksheet and its view state to the grid. */
export function Spreadsheet() {
  const controller = useController();
  const { isMac } = usePlatform();
  const workbook = useViewer((s) => s.workbook);
  const activeSheet = useViewer((s) => s.activeSheet);
  const view = useViewer((s) => s.sheetViews[s.activeSheet]);
  const zoom = useViewer((s) => s.zoom);
  const results = useViewer((s) => s.search.results);
  const current = useViewer((s) => s.search.current);
  const focusRequest = useViewer((s) => s.focusRequest);
  const editing = useViewer((s) => (s.editing && s.editing.sheet === s.activeSheet ? s.editing : null));
  const clipboardMark = useViewer((s) => s.clipboard);
  const gridRef = useRef<HTMLDivElement>(null);
  const [menu, setMenu] = useState<{ kind: GridHitKind; x: number; y: number } | null>(null);

  useEffect(() => {
    if (focusRequest?.target === 'grid') gridRef.current?.focus({ preventScroll: true });
  }, [focusRequest]);

  const sheet = workbook?.sheets[activeSheet];
  const match = current >= 0 ? results?.matches[current] : undefined;
  const matchRow = match && match.sheet === activeSheet ? match.row : -1;
  const matchCol = match && match.sheet === activeSheet ? match.col : -1;
  const currentMatch = useMemo<CellRef | null>(() => (matchRow >= 0 ? { row: matchRow, col: matchCol } : null), [matchRow, matchCol]);
  const clipboard = useMemo(
    () => (clipboardMark && sheet && clipboardMark.sheetId === sheet.id ? { range: clipboardMark.range, cut: clipboardMark.cut } : null),
    [clipboardMark, sheet],
  );

  const onSelectionChange = useCallback((selection: Selection, reveal?: boolean | 'focus') => controller.setSelection(selection, reveal), [controller]);
  const onColumnResize = useCallback((col: number, size: number | null) => controller.setColumnSize(col, size), [controller]);
  const onRowResize = useCallback((row: number, size: number | null) => controller.setRowSize(row, size), [controller]);
  const onColumnSizes = useCallback((sizes: ReadonlyMap<number, number | null>) => void controller.resizeColumns(sizes), [controller]);
  const onRowSize = useCallback((row: number, size: number | null) => void controller.resizeRows(row, size), [controller]);
  const onZoomChange = useCallback((value: number) => controller.setZoom(value), [controller]);
  const onCopy = useCallback(() => void controller.copySelection(), [controller]);
  const onSelectAll = useCallback(() => controller.selectAll(), [controller]);
  const onStartEdit = useCallback((mode: CellEditState['mode'], initial?: string) => controller.startEdit(mode, initial), [controller]);
  const onDraftChange = useCallback((text: string) => controller.setDraft(text, 'cell'), [controller]);
  const onCommitEdit = useCallback((move: CommitMove) => void controller.commitEdit(move), [controller]);
  const onCancelEdit = useCallback(() => controller.cancelEdit(), [controller]);
  const onClear = useCallback(() => void controller.clearSelection(), [controller]);
  const onPaste = useCallback(() => void controller.pasteFromClipboard(), [controller]);
  const onCut = useCallback(() => void controller.cutSelection(), [controller]);
  const onFill = useCallback((range: CellRange) => void controller.fillTo(range), [controller]);
  const onAutoFill = useCallback(() => void controller.autoFill(), [controller]);
  const onContextMenu = useCallback((kind: GridHitKind, x: number, y: number) => setMenu({ kind, x, y }), []);
  const sheetId = sheet?.id ?? -1;
  const onScrollSave = useCallback(
    (scroll: { left: number; top: number }, revealHandled: number) => controller.saveScroll(sheetId, scroll, revealHandled),
    [controller, sheetId],
  );

  if (!workbook || !sheet || !view) return null;

  if (sheet.kind !== 'worksheet') {
    return (
      <div className="flex h-full items-center justify-center bg-app p-8 text-center text-[13px] text-muted">
        “{sheet.name}” is a {SHEET_KIND_LABELS[sheet.kind]}, which Sheaf can’t display.
      </div>
    );
  }

  const menuItems = (kind: GridHitKind): MenuItem[] => {
    const range = normalizeRange(view.selection.anchor, view.selection.focus);
    const rows = isWholeRows(range) ? range.r1 - range.r0 + 1 : 0;
    const cols = isWholeColumns(range) ? range.c1 - range.c0 + 1 : 0;
    const plural = (n: number, word: string) => (n > 1 ? `${n} ${word}s` : word);
    const cmd = isMac ? '⌘' : 'Ctrl+';
    const merged = sheet.merges.some((m) => m.r0 <= range.r1 && m.r1 >= range.r0 && m.c0 <= range.c1 && m.c1 >= range.c0);
    const clipboardItems: MenuItem[] = [
      { label: 'Cut', shortcut: `${cmd}X`, icon: <span className="text-[13px]">✂︎</span>, onSelect: () => void controller.cutSelection() },
      { label: 'Copy', shortcut: `${cmd}C`, onSelect: () => void controller.copySelection() },
      { label: 'Paste', shortcut: `${cmd}V`, onSelect: () => void controller.pasteFromClipboard() },
      { label: 'Paste Values', shortcut: isMac ? '⇧⌘V' : 'Ctrl+Shift+V', onSelect: () => void controller.pasteFromClipboard('values') },
      { label: 'Paste Formatting', onSelect: () => void controller.pasteFromClipboard('formats') },
      { kind: 'separator' },
    ];
    if (kind === 'row') {
      const n = Math.max(1, rows);
      return [
        ...clipboardItems,
        { label: `Insert ${plural(n, 'Row')} Above`, icon: <Icon name="rowInsert" size={14} />, onSelect: () => void controller.insertRows('above') },
        { label: `Insert ${plural(n, 'Row')} Below`, onSelect: () => void controller.insertRows('below') },
        { label: `Delete ${plural(n, 'Row')}`, icon: <Icon name="rowDelete" size={14} />, onSelect: () => void controller.deleteRows() },
        { kind: 'separator' },
        { label: 'Hide', onSelect: () => void controller.setRowsHidden(true) },
        { label: 'Unhide', onSelect: () => void controller.setRowsHidden(false) },
        { label: 'Clear Contents', onSelect: () => void controller.clearSelection() },
        { label: 'Clear Formats', onSelect: () => void controller.clearFormats() },
      ];
    }
    if (kind === 'column') {
      const n = Math.max(1, cols);
      return [
        ...clipboardItems,
        { label: `Insert ${plural(n, 'Column')} Left`, icon: <Icon name="columnInsert" size={14} />, onSelect: () => void controller.insertColumns('left') },
        { label: `Insert ${plural(n, 'Column')} Right`, onSelect: () => void controller.insertColumns('right') },
        { label: `Delete ${plural(n, 'Column')}`, icon: <Icon name="columnDelete" size={14} />, onSelect: () => void controller.deleteColumns() },
        { kind: 'separator' },
        { label: 'Sort A → Z', icon: <Icon name="sortAsc" size={14} />, onSelect: () => void controller.sort(true) },
        { label: 'Sort Z → A', icon: <Icon name="sortDesc" size={14} />, onSelect: () => void controller.sort(false) },
        { kind: 'separator' },
        { label: 'Hide', onSelect: () => void controller.setColumnsHidden(true) },
        { label: 'Unhide', onSelect: () => void controller.setColumnsHidden(false) },
        { label: 'Clear Contents', onSelect: () => void controller.clearSelection() },
        { label: 'Clear Formats', onSelect: () => void controller.clearFormats() },
      ];
    }
    return [
      ...clipboardItems,
      {
        kind: 'submenu',
        label: 'Insert',
        icon: <Icon name="rowInsert" size={14} />,
        items: [
          { label: 'Rows Above', onSelect: () => void controller.insertRows('above') },
          { label: 'Rows Below', onSelect: () => void controller.insertRows('below') },
          { label: 'Columns Left', onSelect: () => void controller.insertColumns('left') },
          { label: 'Columns Right', onSelect: () => void controller.insertColumns('right') },
        ],
      },
      {
        kind: 'submenu',
        label: 'Delete',
        icon: <Icon name="rowDelete" size={14} />,
        items: [
          { label: range.r0 === range.r1 ? 'Row' : 'Rows', onSelect: () => void controller.deleteRows() },
          { label: range.c0 === range.c1 ? 'Column' : 'Columns', onSelect: () => void controller.deleteColumns() },
        ],
      },
      { label: 'Clear Contents', shortcut: 'Delete', onSelect: () => void controller.clearSelection() },
      { label: 'Clear Formats', onSelect: () => void controller.clearFormats() },
      { kind: 'separator' },
      { label: 'Sort A → Z', icon: <Icon name="sortAsc" size={14} />, onSelect: () => void controller.sort(true) },
      { label: 'Sort Z → A', icon: <Icon name="sortDesc" size={14} />, onSelect: () => void controller.sort(false) },
      { kind: 'separator' },
      merged
        ? { label: 'Unmerge Cells', icon: <Icon name="merge" size={14} />, onSelect: () => void controller.unmerge() }
        : { label: 'Merge & Center', icon: <Icon name="merge" size={14} />, disabled: range.r0 === range.r1 && range.c0 === range.c1, onSelect: () => void controller.merge('center') },
      { label: 'Fill Down', shortcut: `${cmd}D`, onSelect: () => void controller.fill('down') },
      { label: 'Fill Right', shortcut: `${cmd}R`, onSelect: () => void controller.fill('right') },
    ];
  };

  return (
    <>
      <SpreadsheetGrid
        key={sheet.id}
        ref={gridRef}
        sheet={sheet}
        styles={workbook.styles}
        view={view}
        zoom={zoom}
        matchFlags={results?.flags.get(activeSheet) ?? null}
        currentMatch={currentMatch}
        isMac={isMac}
        onSelectionChange={onSelectionChange}
        onColumnResize={onColumnResize}
        onRowResize={onRowResize}
        onColumnSizes={onColumnSizes}
        onRowSize={onRowSize}
        onZoomChange={onZoomChange}
        onCopy={onCopy}
        onSelectAll={onSelectAll}
        onScrollSave={onScrollSave}
        editing={editing}
        onStartEdit={onStartEdit}
        onDraftChange={onDraftChange}
        onCommitEdit={onCommitEdit}
        onCancelEdit={onCancelEdit}
        onClear={onClear}
        onPaste={onPaste}
        onCut={onCut}
        onFill={onFill}
        onAutoFill={onAutoFill}
        onContextMenu={onContextMenu}
        clipboard={clipboard}
      />
      {menu && (
        <Menu
          label="Cell actions"
          items={menuItems(menu.kind)}
          placement={{ x: menu.x, y: menu.y }}
          onClose={() => {
            setMenu(null);
            controller.focusGrid();
          }}
        />
      )}
    </>
  );
}
