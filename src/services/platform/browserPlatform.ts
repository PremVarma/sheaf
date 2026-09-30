import { WorkbookError } from '../workbook/errors';
import type { WorkbookSource } from '../workbook/workbookService';
import { WORKBOOK_FILE_EXTENSIONS, type Platform, type PlatformEvents } from './types';

let beforeUnload: ((event: BeforeUnloadEvent) => void) | null = null;

/**
 * Browser implementation, used by `npm run dev` without Tauri and by tests.
 * Files come from <input type="file"> and HTML5 drag and drop.
 */

export function sourceFromFile(file: File): WorkbookSource {
  return {
    name: file.name,
    size: file.size,
    read: async () => {
      try {
        return new Uint8Array(await file.arrayBuffer());
      } catch (error) {
        throw new WorkbookError('READ_FAILED', error instanceof Error ? error.message : String(error));
      }
    },
  };
}

/** Lets the user pick a workbook with <input type="file">. */
export function pickWithFileInput(): Promise<WorkbookSource | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = WORKBOOK_FILE_EXTENSIONS.map((ext) => `.${ext}`).join(',');
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      resolve(file ? sourceFromFile(file) : null);
    });
    input.addEventListener('cancel', () => resolve(null));
    input.click();
  });
}

function hasFiles(event: DragEvent): boolean {
  return Array.from(event.dataTransfer?.types ?? []).includes('Files');
}

async function copyWithTextarea(text: string): Promise<void> {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();
  const ok = document.execCommand('copy');
  textarea.remove();
  if (!ok) throw new Error('Copy command was rejected');
}

export function createBrowserPlatform(): Platform {
  const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);

  return {
    kind: 'browser',
    isMac,

    pickWorkbookFile: pickWithFileInput,

    sourceFromPath(path) {
      return {
        name: path,
        read: () => Promise.reject(new WorkbookError('READ_FAILED', 'Paths can only be opened in the desktop app')),
      };
    },

    setWindowTitle(title) {
      document.title = title;
    },

    async writeClipboardText(text) {
      if (navigator.clipboard?.writeText) {
        try {
          await navigator.clipboard.writeText(text);
          return;
        } catch {
          // Fall through: some contexts reject the async API.
        }
      }
      await copyWithTextarea(text);
    },

    subscribe(events: PlatformEvents) {
      let depth = 0;
      const onDragEnter = (event: DragEvent) => {
        if (!hasFiles(event)) return;
        event.preventDefault();
        if (depth++ === 0) events.onDragStateChange?.(true);
      };
      const onDragOver = (event: DragEvent) => {
        if (!hasFiles(event)) return;
        event.preventDefault();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
      };
      const onDragLeave = (event: DragEvent) => {
        if (!hasFiles(event)) return;
        if (--depth <= 0) {
          depth = 0;
          events.onDragStateChange?.(false);
        }
      };
      const onDrop = (event: DragEvent) => {
        const files = Array.from(event.dataTransfer?.files ?? []);
        if (files.length === 0) return;
        event.preventDefault();
        depth = 0;
        events.onDragStateChange?.(false);
        events.onFilesDropped?.(files.map(sourceFromFile));
      };
      window.addEventListener('dragenter', onDragEnter);
      window.addEventListener('dragover', onDragOver);
      window.addEventListener('dragleave', onDragLeave);
      window.addEventListener('drop', onDrop);
      return () => {
        window.removeEventListener('dragenter', onDragEnter);
        window.removeEventListener('dragover', onDragOver);
        window.removeEventListener('dragleave', onDragLeave);
        window.removeEventListener('drop', onDrop);
      };
    },

    takeStartupFiles: async () => [],
    setRecentFiles() {},
    appReady() {},

    // The browser build can't pick a destination: the file is downloaded under the chosen name.
    saveDialog: async ({ defaultName }) => defaultName,

    async writeFile(path, bytes) {
      const url = URL.createObjectURL(new Blob([bytes as BlobPart]));
      const link = document.createElement('a');
      link.href = url;
      link.download = path.split(/[\\/]/).pop() ?? path;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 10_000);
    },

    async confirmUnsavedChanges(name) {
      return window.confirm(`Discard unsaved changes to “${name}”?`) ? 'discard' : 'cancel';
    },

    readClipboardText: () => navigator.clipboard.readText(),

    setDocumentEdited(edited) {
      if (beforeUnload) window.removeEventListener('beforeunload', beforeUnload);
      beforeUnload = edited ? (event) => event.preventDefault() : null;
      if (beforeUnload) window.addEventListener('beforeunload', beforeUnload);
    },

    quit() {},
  };
}
