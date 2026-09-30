import { useDeferredValue, useEffect, useMemo, useState } from 'react';
import { computeRangeStats, clampToUsedRange } from '../../services/workbook/sheetService';
import { useViewer } from '../../state/AppContext';
import { normalizeRange } from '../../utils/cellAddress';
import { formatCount, formatStatistic, pluralize } from '../../utils/format';
import { Icon } from '../Icon';

const MESSAGE_MS = 4000;

export function StatusBar() {
  const workbook = useViewer((s) => s.workbook);
  const activeSheet = useViewer((s) => s.activeSheet);
  const selection = useViewer((s) => s.sheetViews[s.activeSheet]?.selection);
  const loading = useViewer((s) => s.loading);
  const message = useViewer((s) => s.statusMessage);
  const dirty = useViewer((s) => s.document.dirty);
  const saving = useViewer((s) => s.document.saving);
  const staleFormulas = useViewer((s) => s.document.staleFormulas);
  const [visibleMessage, setVisibleMessage] = useState<string | null>(null);

  useEffect(() => {
    if (!message) return;
    setVisibleMessage(message.text);
    const timer = setTimeout(() => setVisibleMessage(null), MESSAGE_MS);
    return () => clearTimeout(timer);
  }, [message]);

  const sheet = workbook?.sheets[activeSheet];
  const deferredSelection = useDeferredValue(selection);
  const stats = useMemo(() => {
    if (!sheet || !deferredSelection) return null;
    const range = normalizeRange(deferredSelection.anchor, deferredSelection.focus);
    if (range.r0 === range.r1 && range.c0 === range.c1) return null;
    return computeRangeStats(sheet, clampToUsedRange(range, sheet));
  }, [sheet, deferredSelection]);

  const status = loading ? 'Opening…' : saving ? 'Saving…' : (visibleMessage ?? 'Ready');

  return (
    <footer className="flex h-6 shrink-0 items-center gap-3 border-t border-line bg-toolbar px-3 text-[11.5px] text-muted" role="status">
      <span className={visibleMessage ? 'text-fg' : undefined} data-testid="status-text">
        {status}
      </span>
      {sheet && (
        <>
          <Divider />
          <span data-testid="sheet-dimensions">
            {sheet.rowCount === 0 ? 'Empty sheet' : `${pluralize(sheet.rowCount, 'row')} × ${pluralize(sheet.columnCount, 'column')}`}
          </span>
        </>
      )}
      {stats && stats.count > 0 && (
        <>
          <Divider />
          <span className="tabular-nums" data-testid="selection-stats">
            Count: {formatCount(stats.count)}
            {stats.numericCount > 0 && (
              <>
                {'  ·  '}Sum: {formatStatistic(stats.sum)}
                {'  ·  '}Average: {formatStatistic(stats.average ?? 0)}
              </>
            )}
          </span>
        </>
      )}
      <div className="flex-1" />
      {workbook && workbook.warnings.length > 0 && (
        <span className="flex items-center gap-1 text-warning" title={workbook.warnings.join('\n')}>
          <Icon name="warning" size={12} />
          Partially loaded
        </span>
      )}
      {staleFormulas > 0 && (
        <span
          className="flex items-center gap-1 text-warning"
          title="These formulas use functions Sheaf can’t calculate. They keep their last saved value, and Excel recalculates them when the file is opened."
          data-testid="stale-formulas"
        >
          <Icon name="warning" size={12} />
          {pluralize(staleFormulas, 'formula')} not recalculated
        </span>
      )}
      {workbook && (
        <span className={`flex items-center gap-1 ${dirty ? 'text-fg' : ''}`} data-testid="document-state">
          {dirty ? (
            <>
              <span className="h-1.5 w-1.5 rounded-full bg-accent" aria-hidden="true" />
              Unsaved changes
            </>
          ) : (
            'Saved'
          )}
        </span>
      )}
    </footer>
  );
}

function Divider() {
  return <span className="h-3 w-px bg-line" aria-hidden="true" />;
}
