import appIcon from '../../../src-tauri/app-icon.svg';
import { useController, usePlatform, useViewer } from '../../state/AppContext';
import { baseName, dirName } from '../../utils/format';
import { Icon, type IconName } from '../Icon';
import { shortcutLabel } from '../ui';

/** Start screen: a blank workbook or a file to open, and recent files on desktop. */
export function EmptyState() {
  const controller = useController();
  const { isMac, kind, mobile } = usePlatform();
  const recentFiles = useViewer((s) => s.recentFiles);
  const showRecent = kind === 'desktop' && !mobile && recentFiles.length > 0;

  return (
    <div className="h-full overflow-auto bg-app">
      <div className="mx-auto flex min-h-full w-full max-w-xl flex-col justify-center px-6 py-10">
        <div className="flex flex-col items-center text-center">
          <img src={appIcon} alt="" width={80} height={80} className="h-20 w-20" draggable={false} />
          <h1 className="mt-2 text-[24px] font-semibold tracking-tight text-fg">Sheaf</h1>
          <p className="mt-1 text-[13px] text-muted">Open, edit and save spreadsheets.</p>
        </div>

        <div className="mt-8 grid gap-3 sm:grid-cols-2">
          <StartAction
            icon="open"
            title="Open File"
            detail=".xlsx, .xlsm, .xls, .csv and .tsv"
            shortcut={mobile ? undefined : shortcutLabel('O', isMac)}
            primary
            onClick={() => void controller.openDialog()}
          />
          <StartAction
            icon="newFile"
            title="Blank workbook"
            detail="Start from an empty sheet"
            shortcut={mobile ? undefined : shortcutLabel('N', isMac)}
            onClick={() => void controller.newWorkbook()}
          />
        </div>

        {showRecent && (
          <section className="mt-8" aria-labelledby="recent-heading">
            <div className="mb-2 flex items-center justify-between px-1">
              <h2 id="recent-heading" className="text-[11px] font-semibold uppercase tracking-[0.06em] text-muted">
                Recent
              </h2>
              <button type="button" className="text-[11.5px] text-muted hover:text-fg" onClick={() => controller.runCommand('recent.clear')}>
                Clear
              </button>
            </div>
            <ul className="overflow-hidden rounded-xl border border-line bg-surface shadow-sm">
              {recentFiles.map((path) => (
                <li key={path} className="border-b border-line last:border-b-0">
                  <button
                    type="button"
                    className="flex w-full items-center gap-3 px-3.5 py-2.5 text-left hover:bg-hover"
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
          </section>
        )}

        {!mobile && <p className="mt-8 text-center text-[11.5px] text-muted">You can also drop a spreadsheet anywhere in this window.</p>}
      </div>
    </div>
  );
}

function StartAction({
  icon,
  title,
  detail,
  shortcut,
  primary = false,
  onClick,
}: {
  icon: IconName;
  title: string;
  detail: string;
  shortcut?: string;
  primary?: boolean;
  onClick(): void;
}) {
  const detailId = `start-${icon}`;
  return (
    <button
      type="button"
      aria-label={title}
      aria-describedby={detailId}
      className="flex items-start gap-3 rounded-xl border border-line bg-surface p-4 text-left shadow-sm transition hover:border-accent/40 hover:shadow-md active:scale-[0.99]"
      onClick={onClick}
    >
      <span className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-lg ${primary ? 'bg-accent text-accent-fg' : 'bg-accent-soft text-accent'}`}>
        <Icon name={icon} size={20} />
      </span>
      <span className="min-w-0 flex-1">
        <span className="flex items-center justify-between gap-2">
          <span className="text-[13.5px] font-semibold text-fg">{title}</span>
          {shortcut && <kbd className="rounded border border-line px-1.5 py-px font-sans text-[10.5px] text-muted pointer-coarse:hidden">{shortcut}</kbd>}
        </span>
        <span id={detailId} className="mt-0.5 block text-[12px] text-muted">
          {detail}
        </span>
      </span>
    </button>
  );
}
