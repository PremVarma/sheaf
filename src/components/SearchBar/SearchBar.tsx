import { useEffect, useRef, useState } from 'react';
import { MAX_SEARCH_MATCHES } from '../../services/search/searchService';
import { useController, useViewer } from '../../state/AppContext';
import { formatCount } from '../../utils/format';
import { Icon } from '../Icon';
import { ToolButton } from '../ui';

/**
 * Workbook search: type to highlight matches, Enter / Shift+Enter to step
 * through them, Esc to clear.
 */
export function SearchBar({ className = '' }: { className?: string }) {
  const controller = useController();
  const enabled = useViewer((s) => s.workbook !== null);
  const query = useViewer((s) => s.search.query);
  const results = useViewer((s) => s.search.results);
  const current = useViewer((s) => s.search.current);
  const scope = useViewer((s) => s.search.scope);
  const matchCase = useViewer((s) => s.search.matchCase);
  const wholeCell = useViewer((s) => s.search.wholeCell);
  const focusRequest = useViewer((s) => s.focusRequest);
  const replaceOpen = useViewer((s) => s.replaceOpen);
  const inputRef = useRef<HTMLInputElement>(null);
  const replaceRef = useRef<HTMLInputElement>(null);
  const [replacement, setReplacement] = useState('');

  useEffect(() => {
    if (focusRequest?.target === 'search') {
      inputRef.current?.focus();
      inputRef.current?.select();
    } else if (focusRequest?.target === 'replace') {
      replaceRef.current?.focus();
      replaceRef.current?.select();
    }
  }, [focusRequest]);

  const total = results?.total ?? 0;
  const capped = results !== null && results.total > MAX_SEARCH_MATCHES;
  let status = '';
  if (query && enabled) {
    if (!results) status = 'Searching…';
    else if (total === 0) status = 'No matches';
    else if (current >= 0) status = `${formatCount(current + 1)} of ${formatCount(results.matches.length)}${capped ? '+' : ''}`;
    else status = `${formatCount(total)} ${total === 1 ? 'match' : 'matches'}`;
  }

  return (
    <div className={`flex min-w-0 items-center gap-1 max-sm:flex-wrap max-sm:gap-y-1.5 ${className}`} role="search">
      <div className="relative flex h-7 w-56 min-w-36 shrink items-center max-sm:w-auto max-sm:min-w-0 max-sm:flex-1 pointer-coarse:h-9 rounded-md border border-control-line bg-control focus-within:border-accent focus-within:ring-2 focus-within:ring-accent/25">
        <Icon name="search" size={14} className="pointer-events-none absolute left-2 text-muted" />
        <input
          ref={inputRef}
          type="text"
          className="h-full w-full min-w-0 bg-transparent pl-7 pr-7 text-[12.5px] text-fg outline-none placeholder:text-muted disabled:opacity-50"
          placeholder="Search…"
          value={query}
          disabled={!enabled}
          spellCheck={false}
          autoComplete="off"
          aria-label="Search"
          onChange={(event) => controller.setSearchQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              controller.findNext(event.shiftKey ? -1 : 1);
            } else if (event.key === 'Escape') {
              event.preventDefault();
              controller.clearSearch();
              controller.focusGrid();
            }
          }}
        />
        {query && (
          <button
            type="button"
            className="absolute right-1.5 flex h-4 w-4 items-center justify-center rounded-full bg-muted/40 text-surface hover:bg-muted/70"
            aria-label="Clear search"
            title="Clear search (Esc)"
            onClick={() => {
              controller.clearSearch();
              inputRef.current?.focus();
            }}
          >
            <Icon name="close" size={10} strokeWidth={3} />
          </button>
        )}
      </div>

      {query && enabled && (
        <>
          <span className="min-w-16 whitespace-nowrap px-1 text-center text-[12px] tabular-nums text-muted" aria-live="polite" data-testid="search-status">
            {status}
          </span>
          <ToolButton icon="chevronUp" label="Previous match" title="Previous match (⇧↩)" disabled={total === 0} onClick={() => controller.findNext(-1)} />
          <ToolButton icon="chevronDown" label="Next match" title="Next match (↩)" disabled={total === 0} onClick={() => controller.findNext(1)} />
          <div className="ml-1 flex h-7 items-center rounded-md border border-control-line p-0.5 text-[11.5px] max-sm:hidden" role="group" aria-label="Search in">
            {(['sheet', 'workbook'] as const).map((value) => (
              <button
                key={value}
                type="button"
                aria-pressed={scope === value}
                className={`h-full rounded px-2 ${scope === value ? 'bg-accent text-accent-fg' : 'text-muted hover:text-fg'}`}
                onClick={() => controller.setSearchOptions({ scope: value })}
              >
                {value === 'sheet' ? 'Sheet' : 'Workbook'}
              </button>
            ))}
          </div>
          <ToolButton
            label="Match case"
            active={matchCase}
            aria-pressed={matchCase}
            className="text-[12px] font-semibold max-sm:hidden"
            onClick={() => controller.setSearchOptions({ matchCase: !matchCase })}
          >
            Aa
          </ToolButton>
          <ToolButton
            label="Match entire cell contents"
            active={wholeCell}
            aria-pressed={wholeCell}
            className="text-[11px] font-semibold max-sm:hidden"
            onClick={() => controller.setSearchOptions({ wholeCell: !wholeCell })}
          >
            [ab]
          </ToolButton>
        </>
      )}
      {enabled && (
        <ToolButton
          label="Replace"
          title="Find and replace"
          active={replaceOpen}
          aria-pressed={replaceOpen}
          className="text-[12px]"
          onClick={() => {
            controller.setReplaceOpen(!replaceOpen);
            if (!replaceOpen) (query ? replaceRef : inputRef).current?.focus();
          }}
        >
          <span className="text-[13px]">⇄</span>
        </ToolButton>
      )}
      {enabled && replaceOpen && (
        // On phones the replace controls get a line of their own.
        <div className="flex min-w-0 items-center gap-1 max-sm:basis-full">
          <input
            ref={replaceRef}
            type="text"
            className="h-7 w-36 min-w-24 shrink max-sm:w-auto max-sm:flex-1 pointer-coarse:h-9 rounded-md border border-control-line bg-control px-2 text-[12.5px] text-fg outline-none placeholder:text-muted focus:border-accent focus:ring-2 focus:ring-accent/25"
            placeholder="Replace with…"
            value={replacement}
            spellCheck={false}
            autoComplete="off"
            aria-label="Replace with"
            onChange={(event) => setReplacement(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                void controller.replaceCurrent(replacement);
              } else if (event.key === 'Escape') {
                event.preventDefault();
                controller.clearSearch();
                controller.focusGrid();
              }
            }}
          />
          <button
            type="button"
            disabled={total === 0}
            className="h-7 shrink-0 pointer-coarse:h-9 rounded-md border border-control-line bg-control px-2.5 text-[12px] text-fg hover:bg-hover disabled:opacity-40"
            onClick={() => void controller.replaceCurrent(replacement)}
          >
            Replace
          </button>
          <button
            type="button"
            disabled={total === 0}
            className="h-7 shrink-0 pointer-coarse:h-9 rounded-md border border-control-line bg-control px-2.5 text-[12px] text-fg hover:bg-hover disabled:opacity-40"
            onClick={() => void controller.replaceAll(replacement)}
          >
            Replace All
          </button>
        </div>
      )}
    </div>
  );
}
