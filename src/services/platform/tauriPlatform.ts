import { invoke } from '@tauri-apps/api/core';
import { listen, type UnlistenFn } from '@tauri-apps/api/event';
import { getCurrentWebview } from '@tauri-apps/api/webview';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { readText, writeText } from '@tauri-apps/plugin-clipboard-manager';
import { message, open, save } from '@tauri-apps/plugin-dialog';
import { baseName } from '../../utils/format';
import { WorkbookError, type WorkbookErrorCode } from '../workbook/errors';
import type { WorkbookSource } from '../workbook/workbookService';
import { SAVE_TYPE_LABELS, WORKBOOK_FILE_EXTENSIONS, type Platform, type PlatformEvents } from './types';

/** Error payload returned by the `read_workbook_file` Rust command. */
interface NativeFileError {
  code: WorkbookErrorCode;
  message?: string;
  size?: number;
  limit?: number;
}

function toWorkbookError(error: unknown): WorkbookError {
  if (error && typeof error === 'object' && 'code' in error) {
    const native = error as NativeFileError;
    return new WorkbookError(native.code, native.message, { size: native.size, limit: native.limit });
  }
  return new WorkbookError('READ_FAILED', String(error));
}

function sourceFromPath(path: string): WorkbookSource {
  return {
    name: baseName(path),
    path,
    read: async () => {
      try {
        // Returned as a raw ArrayBuffer (tauri::ipc::Response), not JSON.
        const buffer = await invoke<ArrayBuffer>('read_workbook_file', { path });
        return new Uint8Array(buffer);
      } catch (error) {
        throw toWorkbookError(error);
      }
    },
  };
}

export function createTauriPlatform(): Platform {
  const isMac = /Mac/.test(navigator.userAgent);

  return {
    kind: 'desktop',
    isMac,

    async pickWorkbookFile() {
      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          { name: 'Spreadsheets', extensions: WORKBOOK_FILE_EXTENSIONS },
          { name: 'All Files', extensions: ['*'] },
        ],
      });
      return typeof selected === 'string' ? sourceFromPath(selected) : null;
    },

    sourceFromPath,

    setWindowTitle(title) {
      void getCurrentWindow().setTitle(title);
    },

    writeClipboardText: (text) => writeText(text),

    subscribe(events: PlatformEvents) {
      const pending: Promise<UnlistenFn>[] = [
        getCurrentWebview().onDragDropEvent((event) => {
          const payload = event.payload;
          if (payload.type === 'enter' || payload.type === 'over') {
            events.onDragStateChange?.(true);
          } else if (payload.type === 'leave') {
            events.onDragStateChange?.(false);
          } else if (payload.type === 'drop') {
            events.onDragStateChange?.(false);
            if (payload.paths.length > 0) events.onFilesDropped?.(payload.paths.map(sourceFromPath));
          }
        }),
        listen<string[]>('open-files', (event) => events.onOpenPaths?.(event.payload)),
        listen<string>('menu-command', (event) => events.onMenuCommand?.(event.payload)),
        listen('close-requested', () => events.onCloseRequested?.()),
      ];
      return () => {
        for (const unlisten of pending) void unlisten.then((fn) => fn());
      };
    },

    takeStartupFiles: () => invoke<string[]>('take_startup_files'),

    setRecentFiles(paths) {
      void invoke('set_recent_files', { paths });
    },

    appReady() {
      void getCurrentWindow().show();
    },

    saveDialog: ({ defaultName, types }) =>
      save({
        defaultPath: defaultName,
        filters: types.map((type) => ({ name: SAVE_TYPE_LABELS[type], extensions: [type] })),
      }),

    async writeFile(path, bytes) {
      try {
        // Raw body (no JSON encoding); the path travels in a header.
        await invoke('write_workbook_file', bytes, { headers: { 'x-path': encodeURIComponent(path) } });
      } catch (error) {
        throw toWorkbookError(error);
      }
    },

    async confirmUnsavedChanges(name) {
      const choice = await message(`Do you want to save the changes you made to “${name}”?\n\nYour changes will be lost if you don’t save them.`, {
        title: 'Sheaf',
        kind: 'warning',
        buttons: { yes: 'Save', no: 'Don’t Save', cancel: 'Cancel' },
      });
      if (choice === 'Save' || choice === 'Yes') return 'save';
      if (choice === 'Don’t Save' || choice === 'No') return 'discard';
      return 'cancel';
    },

    readClipboardText: () => readText(),

    setDocumentEdited(edited) {
      void invoke('set_document_edited', { edited });
    },

    quit() {
      void invoke('quit_app');
    },
  };
}
