import type { ReactNode } from 'react';
import { FORMAT_LABELS } from '../../services/workbook/formatDetection';
import { useController, useViewer } from '../../state/AppContext';
import { formatBytes, formatCount, formatDateTime } from '../../utils/format';
import { ToolButton } from '../ui';

/** Side panel with the essentials about the open file and workbook. */
export function WorkbookInfo() {
  const controller = useController();
  const workbook = useViewer((s) => s.workbook);
  const activeSheet = useViewer((s) => s.activeSheet);
  if (!workbook) return null;

  const sheet = workbook.sheets[activeSheet];
  const hidden = workbook.sheets.filter((s) => s.visibility !== 'visible').length;
  const { properties } = workbook;

  return (
    <aside className="flex w-64 shrink-0 flex-col overflow-y-auto border-l border-line bg-surface" aria-label="Workbook information">
      <div className="flex h-9 items-center justify-between border-b border-line pl-4 pr-1.5">
        <h2 className="text-[12.5px] font-semibold text-fg">Information</h2>
        <ToolButton icon="close" label="Close information panel" onClick={() => controller.setInfoOpen(false)} />
      </div>
      <Section title="File">
        <Row label="Name" value={workbook.fileName} />
        <Row label="Type" value={FORMAT_LABELS[workbook.format]} />
        <Row label="Size" value={formatBytes(workbook.fileSize)} />
      </Section>
      <Section title="Workbook">
        <Row label="Sheets" value={`${formatCount(workbook.sheets.length)}${hidden ? ` (${hidden} hidden)` : ''}`} />
        <Row label="Active sheet" value={sheet?.name ?? '—'} />
        <Row label="Rows" value={formatCount(sheet?.rowCount ?? 0)} />
        <Row label="Columns" value={formatCount(sheet?.columnCount ?? 0)} />
      </Section>
      {(properties.title || properties.author || properties.lastModifiedBy || properties.created || properties.modified) && (
        <Section title="Properties">
          {properties.title && <Row label="Title" value={properties.title} />}
          {properties.author && <Row label="Author" value={properties.author} />}
          {properties.lastModifiedBy && <Row label="Last saved by" value={properties.lastModifiedBy} />}
          {properties.created && <Row label="Created" value={formatDateTime(properties.created)} />}
          {properties.modified && <Row label="Modified" value={formatDateTime(properties.modified)} />}
        </Section>
      )}
      {workbook.warnings.length > 0 && (
        <Section title="Notes">
          {workbook.warnings.map((warning) => (
            <p key={warning} className="text-[12px] leading-snug text-warning">
              {warning}
            </p>
          ))}
        </Section>
      )}
    </aside>
  );
}

function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="border-b border-line px-4 py-3">
      <h3 className="mb-2 text-[10.5px] font-semibold uppercase tracking-[0.06em] text-muted">{title}</h3>
      <dl className="grid grid-cols-[88px_1fr] gap-x-3 gap-y-1.5">{children}</dl>
    </section>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <>
      <dt className="text-[12px] text-muted">{label}</dt>
      <dd className="min-w-0 break-words text-[12px] text-fg" title={value}>
        {value}
      </dd>
    </>
  );
}
