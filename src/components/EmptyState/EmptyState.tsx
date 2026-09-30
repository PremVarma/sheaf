import { useController, usePlatform, useViewer } from '../../state/AppContext';
import { baseName, dirName } from '../../utils/format';
import { Icon } from '../Icon';
import { shortcutLabel } from '../ui';

/** Shown when no workbook is open. */
export function EmptyState() {
  const controller = useController();
  const { isMac, kind } = usePlatform();
  const recentFiles = useViewer((s) => s.recentFiles);

  return (
    <div className="flex h-full items-center justify-center overflow-auto bg-app p-8">
      <div className="flex w-full max-w-md flex-col items-center">
        <div className="flex w-full flex-col items-center rounded-2xl border-2 border-dashed border-line-strong bg-surface/60 px-10 py-12 text-center">
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-accent-soft text-accent">
            <Icon name="sheet" size={32} strokeWidth={1.5} />
          </div>
          <h1 className="text-[17px] font-semibold text-fg">Open an Excel file</h1>
          <p className="mt-1.5 text-[13px] text-muted">Drag and drop a workbook here</p>
          <button
            type="button"
            className="mt-6 inline-flex h-8 items-center gap-2 rounded-lg bg-accent px-4 text-[13px] font-medium text-accent-fg shadow-sm transition hover:brightness-110 active:brightness-95"
            onClick={() => void controller.openDialog()}
          >
            <Icon name="open" size={15} />
            Open File
          </button>
          <p className="mt-4 text-[11.5px] text-muted">
            .xlsx, .xlsm, .xls, .csv and .tsv · {shortcutLabel('O', isMac)}
          </p>
        </div>

        {kind === 'desktop' && recentFiles.length > 0 && (
          <div className="mt-6 w-full">
            <div className="mb-1.5 flex items-center justify-between px-1">
              <h2 className="text-[11px] font-semibold uppercase tracking-[0.06em] text-muted">Recent</h2>
              <button type="button" className="text-[11.5px] text-muted hover:text-fg" onClick={() => controller.runCommand('recent.clear')}>
                Clear
              </button>
            </div>
            <ul className="overflow-hidden rounded-xl border border-line bg-surface">
              {recentFiles.map((path) => (
                <li key={path} className="border-b border-line last:border-b-0">
                  <button
                    type="button"
                    className="flex w-full items-center gap-3 px-3 py-2 text-left hover:bg-hover"
                    title={path}
                    onClick={() => void controller.openPath(path)}
                  >
                    <Icon name="file" size={16} className="shrink-0 text-accent" />
                    <span className="min-w-0 flex-1">
                      <span className="block truncate text-[12.5px] text-fg">{baseName(path)}</span>
                      <span className="block truncate text-[11px] text-muted">{dirName(path)}</span>
                    </span>
                  </button>
                </li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
}
