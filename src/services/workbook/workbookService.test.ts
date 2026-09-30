import { describe, expect, it, vi } from 'vitest';
import { buildWorkbookFile, sourceOf } from '../../test/fixtures';
import { WorkbookError } from './errors';
import { SIZE_LIMITS } from './formatDetection';
import { loadWorkbook } from './workbookService';

describe('loadWorkbook', () => {
  it('reads, then parses, reporting each stage', async () => {
    const stages: string[] = [];
    const { workbook, bytes } = await loadWorkbook(sourceOf('a.xlsx', buildWorkbookFile({ S: [['x']] })), {
      onStage: (stage) => stages.push(stage),
    });
    expect(stages).toEqual(['reading', 'parsing']);
    expect(workbook.sheets[0].name).toBe('S');
    expect(bytes.length).toBeGreaterThan(0);
  });

  it('rejects unsupported files without reading them', async () => {
    const read = vi.fn(async () => new Uint8Array([1]));
    await expect(loadWorkbook({ name: 'slides.pptx', read })).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' });
    expect(read).not.toHaveBeenCalled();
  });

  it('rejects files over the size limit before reading them', async () => {
    const read = vi.fn(async () => new Uint8Array([1]));
    const size = SIZE_LIMITS.workbook + 1;
    const error = await loadWorkbook({ name: 'huge.xlsx', size, read }).catch((e: WorkbookError) => e);
    expect(error).toMatchObject({ code: 'TOO_LARGE', details: { size, limit: SIZE_LIMITS.workbook } });
    expect(read).not.toHaveBeenCalled();
  });

  it('can be cancelled', async () => {
    const abort = new AbortController();
    const source = { name: 'a.xlsx', read: async () => (abort.abort(), buildWorkbookFile({ S: [['x']] })) };
    await expect(loadWorkbook(source, { signal: abort.signal })).rejects.toMatchObject({ code: 'CANCELLED' });
  });

  it('passes read failures through with their code', async () => {
    const source = { name: 'gone.xlsx', read: async () => Promise.reject(new WorkbookError('NOT_FOUND')) };
    await expect(loadWorkbook(source)).rejects.toMatchObject({ code: 'NOT_FOUND' });
  });
});
