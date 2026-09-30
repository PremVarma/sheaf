import type { WorkbookSource } from '../workbook/workbookService';

export interface PlatformEvents {
  /** A file drag entered (true) or left (false) the window. */
  onDragStateChange?: (active: boolean) => void;
  onFilesDropped?: (files: WorkbookSource[]) => void;
  /** Files opened from the OS: command line, Finder "Open With", Open Recent. */
  onOpenPaths?: (paths: string[]) => void;
  /** A native menu item was chosen. */
  onMenuCommand?: (command: string) => void;
  /** The user tried to close the window or quit while there are unsaved changes. */
  onCloseRequested?: () => void;
}

export type SaveFileType = 'xlsx' | 'xlsm' | 'csv' | 'tsv';
export type UnsavedChoice = 'save' | 'discard' | 'cancel';

/** Everything the UI needs from the host environment. */
export interface Platform {
  readonly kind: 'desktop' | 'browser';
  readonly isMac: boolean;
  /** Phone or tablet (iOS, Android): no drag and drop, native menus or recent files. */
  readonly mobile: boolean;
  /** Shows a file picker; resolves null when cancelled. */
  pickWorkbookFile(): Promise<WorkbookSource | null>;
  /** A source for a path on disk (desktop only). */
  sourceFromPath(path: string): WorkbookSource;
  setWindowTitle(title: string): void;
  writeClipboardText(text: string): Promise<void>;
  subscribe(events: PlatformEvents): () => void;
  /** Files passed on the command line / at launch. Returns each file once. */
  takeStartupFiles(): Promise<string[]>;
  /** Updates the native "Open Recent" menu. */
  setRecentFiles(paths: string[]): void;
  /** Called once the first frame has rendered (desktop shows its window then). */
  appReady(): void;
  /** Save dialog; resolves the chosen path (a content:// URI on Android), or null when cancelled. */
  saveDialog(options: { defaultName: string; types: SaveFileType[] }): Promise<string | null>;
  /**
   * Where files are saved when there is no save dialog (iOS: the app's Documents folder,
   * which the Files app shows). Save As then asks for a name instead.
   */
  documentsDir?(): Promise<string>;
  /** Whether a file exists (used with `documentsDir`). */
  fileExists?(path: string): Promise<boolean>;
  /** Writes a file atomically (desktop) or downloads it (browser). */
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
  /** Asks what to do with unsaved changes. */
  confirmUnsavedChanges(documentName: string): Promise<UnsavedChoice>;
  readClipboardText(): Promise<string>;
  /** Tells the host whether there are unsaved changes (close/quit are intercepted while true). */
  setDocumentEdited(edited: boolean): void;
  /** Closes the app without asking again. */
  quit(): void;
}

export const SAVE_TYPE_LABELS: Record<SaveFileType, string> = {
  xlsx: 'Excel Workbook',
  xlsm: 'Excel Macro-Enabled Workbook',
  csv: 'CSV (Comma-separated)',
  tsv: 'Tab-separated text',
};

export const WORKBOOK_FILE_EXTENSIONS = ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'];
