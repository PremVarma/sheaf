import * as XLSX from 'xlsx';
import { CellKind, type WorkbookModel, type WorksheetModel } from '../models/workbook';
import type { Platform, UnsavedChoice } from '../services/platform/types';
import { SheetBuilder, type SheetLayout } from '../services/workbook/sheetBuilder';
import type { WorkbookSource } from '../services/workbook/workbookService';

type CellInput = string | number | boolean | null | XLSX.CellObject;

/** Builds a workbook with SheetJS and returns the file bytes. */
export function buildWorkbookFile(
  sheets: Record<string, CellInput[][]>,
  bookType: 'xlsx' | 'xls' | 'csv' = 'xlsx',
  configure?: (workbook: XLSX.WorkBook) => void,
): Uint8Array {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  configure?.(workbook);
  if (bookType === 'csv') {
    const first = workbook.Sheets[workbook.SheetNames[0]];
    return new TextEncoder().encode(XLSX.utils.sheet_to_csv(first));
  }
  return new Uint8Array(XLSX.write(workbook, { type: 'array', bookType }) as ArrayBuffer);
}

/** Rewrites XML parts inside an .xlsx package (to add styles, panes, …). */
export function patchXlsx(bytes: Uint8Array, patches: Record<string, (xml: string) => string>): Uint8Array {
  const container = XLSX.CFB.read(bytes, { type: 'array' });
  for (const [path, patch] of Object.entries(patches)) {
    const entry = XLSX.CFB.find(container, `/${path}`);
    if (!entry) throw new Error(`Missing part ${path}`);
    const xml = new TextDecoder().decode(entry.content as Uint8Array);
    const next = new TextEncoder().encode(patch(xml));
    entry.content = next as unknown as typeof entry.content;
    entry.size = next.length;
  }
  return new Uint8Array(XLSX.CFB.write(container, { fileType: 'zip', type: 'array' }) as number[]);
}

/** A worksheet built directly in the compact model (fast, for grid tests). */
export function createSheet(
  rows: number,
  cols: number,
  value: (row: number, col: number) => string | number | null = (r, c) => `R${r + 1}C${c + 1}`,
  layout: SheetLayout = {},
  name = 'Sheet1',
): WorksheetModel {
  const builder = new SheetBuilder({ name, index: 0 });
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const v = value(r, c);
      if (v === null || v === '') continue;
      if (typeof v === 'number') builder.add(r, c, CellKind.Number, v, String(v));
      else builder.add(r, c, CellKind.String, 0, v);
    }
  }
  return builder.finish(layout);
}

export function workbookOf(...sheets: WorksheetModel[]): WorkbookModel {
  return {
    fileName: 'test.xlsx',
    fileSize: 1024,
    format: 'xlsx',
    properties: {},
    sheets: sheets.map((sheet, index) => ({ ...sheet, id: index, index })),
    styles: [{}],
    activeSheetIndex: 0,
    date1904: false,
    warnings: [],
  };
}

export function sourceOf(name: string, bytes: Uint8Array, path?: string): WorkbookSource {
  return { name, path, size: bytes.length, read: async () => bytes };
}

export interface MockPlatform extends Platform {
  clipboard: string[];
  title: string;
  recent: string[];
  nextFile: WorkbookSource | null;
  /** Files written by Save, in order. */
  saved: { path: string; bytes: Uint8Array }[];
  /** Path the save dialog returns (null = cancelled); defaults to /Users/me/<name>. */
  nextSavePath: string | null | undefined;
  unsavedChoice: UnsavedChoice;
  unsavedPrompts: number;
  clipboardText: string;
  edited: boolean;
  quitCalled: boolean;
}

/** In-memory platform: the file picker returns `nextFile`, the clipboard is an array. */
export function createMockPlatform(): MockPlatform {
  const platform: MockPlatform = {
    kind: 'desktop',
    isMac: true,
    clipboard: [],
    title: '',
    recent: [],
    nextFile: null,
    saved: [],
    nextSavePath: undefined,
    unsavedChoice: 'discard',
    unsavedPrompts: 0,
    clipboardText: '',
    edited: false,
    quitCalled: false,
    pickWorkbookFile: async () => platform.nextFile,
    sourceFromPath: (path) => ({ name: path.split('/').pop() ?? path, path, read: async () => new Uint8Array() }),
    setWindowTitle: (title) => {
      platform.title = title;
    },
    writeClipboardText: async (text) => {
      platform.clipboard.push(text);
    },
    subscribe: () => () => {},
    takeStartupFiles: async () => [],
    setRecentFiles: (paths) => {
      platform.recent = paths;
    },
    appReady: () => {},
    saveDialog: async ({ defaultName }) => (platform.nextSavePath === undefined ? `/Users/me/${defaultName}` : platform.nextSavePath),
    writeFile: async (path, bytes) => {
      platform.saved.push({ path, bytes });
    },
    confirmUnsavedChanges: async () => {
      platform.unsavedPrompts++;
      return platform.unsavedChoice;
    },
    readClipboardText: async () => platform.clipboardText,
    setDocumentEdited: (edited) => {
      platform.edited = edited;
    },
    quit: () => {
      platform.quitCalled = true;
    },
  };
  return platform;
}
