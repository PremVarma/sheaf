import { useEffect } from 'react';
import { useKeyboardShortcuts } from '../../hooks/useKeyboardShortcuts';
import { useController, usePlatform, useViewer } from '../../state/AppContext';
import { EmptyState } from '../EmptyState/EmptyState';
import { DropOverlay, ErrorDialog, LoadingOverlay, PromptDialog } from '../Overlays/Overlays';
import { SheetTabs } from '../SheetTabs/SheetTabs';
import { Spreadsheet } from '../Spreadsheet/Spreadsheet';
import { StatusBar } from '../StatusBar/StatusBar';
import { Toolbar } from '../Toolbar/Toolbar';
import { WorkbookInfo } from '../WorkbookInfo/WorkbookInfo';

/**
 * Window layout:
 *   toolbar (file · undo · search · zoom · info), formatting bar, formula bar
 *   grid | info panel
 *   sheet tabs
 *   status bar
 */
export function AppShell() {
  const controller = useController();
  const platform = usePlatform();
  const hasWorkbook = useViewer((s) => s.workbook !== null);
  const infoOpen = useViewer((s) => s.infoOpen);

  useKeyboardShortcuts(controller, platform.isMac);
  useEffect(() => controller.attach(), [controller]);
  useEffect(() => {
    platform.appReady();
  }, [platform]);

  return (
    <div className="relative flex h-full flex-col bg-app text-fg">
      <Toolbar />
      <main className="relative flex min-h-0 flex-1">
        <div className="relative min-w-0 flex-1">{hasWorkbook ? <Spreadsheet /> : <EmptyState />}</div>
        {hasWorkbook && infoOpen && <WorkbookInfo />}
      </main>
      <SheetTabs />
      <StatusBar />
      <LoadingOverlay />
      <ErrorDialog />
      <PromptDialog />
      <DropOverlay />
    </div>
  );
}
