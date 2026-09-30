import type { WorkbookModel } from '../../models/workbook';
import { deserializeError, serializeError, WorkbookError, type SerializedWorkbookError } from './errors';
import { isOpenableFileName, sizeLimitFor } from './formatDetection';

/** A file the user asked to open, independent of where it comes from. */
export interface WorkbookSource {
  name: string;
  /** Absolute path on disk (desktop app only). */
  path?: string;
  /** Size in bytes when known before reading. */
  size?: number;
  read(): Promise<Uint8Array>;
}

export type LoadStage = 'reading' | 'parsing';

/** A parsed workbook plus the file's bytes (kept so edits can be saved into the original package). */
export interface LoadedWorkbook {
  workbook: WorkbookModel;
  bytes: Uint8Array;
}

export type ParseFunction = (bytes: Uint8Array, fileName: string, signal?: AbortSignal) => Promise<LoadedWorkbook>;

export interface LoadOptions {
  signal?: AbortSignal;
  onStage?: (stage: LoadStage) => void;
  /** Override the parser (tests); defaults to a background worker. */
  parse?: ParseFunction;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) throw new WorkbookError('CANCELLED');
}

/** File loading → validation → background parsing → normalized model. */
export async function loadWorkbook(source: WorkbookSource, options: LoadOptions = {}): Promise<LoadedWorkbook> {
  const { signal, onStage } = options;
  if (!isOpenableFileName(source.name)) throw new WorkbookError('UNSUPPORTED_FORMAT', source.name);

  const limit = sizeLimitFor(source.name);
  if (source.size !== undefined && source.size > limit) {
    throw new WorkbookError('TOO_LARGE', undefined, { size: source.size, limit });
  }

  throwIfAborted(signal);
  onStage?.('reading');
  const bytes = await source.read();
  throwIfAborted(signal);
  if (bytes.length === 0) throw new WorkbookError('EMPTY_FILE');
  if (bytes.length > limit) throw new WorkbookError('TOO_LARGE', undefined, { size: bytes.length, limit });

  onStage?.('parsing');
  const loaded = await (options.parse ?? parseInBackground)(bytes, source.name, signal);
  throwIfAborted(signal);
  return loaded;
}

type WorkerResponse = { ok: true; workbook: WorkbookModel; original: ArrayBuffer } | { ok: false; error: SerializedWorkbookError };

/**
 * Parses in a dedicated worker so the UI stays responsive, then terminates it
 * to release SheetJS's memory. Falls back to the main thread where workers
 * aren't available (unit tests).
 */
export const parseInBackground: ParseFunction = (bytes, fileName, signal) => {
  if (typeof Worker === 'undefined') return parseOnMainThread(bytes, fileName, signal);

  return new Promise<LoadedWorkbook>((resolve, reject) => {
    const worker = new Worker(new URL('./parser.worker.ts', import.meta.url), { type: 'module', name: 'workbook-parser' });
    const finish = () => {
      worker.terminate();
      signal?.removeEventListener('abort', onAbort);
    };
    const onAbort = () => {
      finish();
      reject(new WorkbookError('CANCELLED'));
    };
    signal?.addEventListener('abort', onAbort, { once: true });

    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      finish();
      if (event.data.ok) resolve({ workbook: event.data.workbook, bytes: new Uint8Array(event.data.original) });
      else reject(deserializeError(event.data.error));
    };
    worker.onerror = (event) => {
      event.preventDefault();
      finish();
      reject(new WorkbookError(/memory/i.test(event.message ?? '') ? 'OUT_OF_MEMORY' : 'UNKNOWN', event.message));
    };
    worker.onmessageerror = () => {
      finish();
      reject(new WorkbookError('UNKNOWN', 'The parser result could not be transferred'));
    };

    const buffer =
      bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength
        ? (bytes.buffer as ArrayBuffer)
        : (bytes.slice().buffer as ArrayBuffer);
    worker.postMessage({ buffer, fileName }, [buffer]);
  });
};

async function parseOnMainThread(bytes: Uint8Array, fileName: string, signal?: AbortSignal): Promise<LoadedWorkbook> {
  await new Promise((resolve) => setTimeout(resolve, 0));
  throwIfAborted(signal);
  const { parseWorkbookBytes } = await import('./parser');
  try {
    return { workbook: parseWorkbookBytes(bytes, fileName), bytes };
  } catch (error) {
    throw error instanceof WorkbookError ? error : deserializeError(serializeError(error));
  }
}
