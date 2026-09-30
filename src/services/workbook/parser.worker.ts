import { serializeError } from './errors';
import { collectTransferables, parseWorkbookBytes } from './parser';

export interface ParseRequest {
  buffer: ArrayBuffer;
  fileName: string;
}

// Typed locally so this file doesn't need the WebWorker lib (which conflicts with DOM).
const scope = self as unknown as {
  onmessage: ((event: MessageEvent<ParseRequest>) => void) | null;
  postMessage(message: unknown, transfer: Transferable[]): void;
};

scope.onmessage = (event) => {
  const { buffer, fileName } = event.data;
  try {
    const workbook = parseWorkbookBytes(new Uint8Array(buffer), fileName);
    // The file bytes go back too (zero-copy), so edits can be saved into the original package.
    scope.postMessage({ ok: true, workbook, original: buffer }, [...collectTransferables(workbook), buffer]);
  } catch (error) {
    scope.postMessage({ ok: false, error: serializeError(error) }, []);
  }
};
