import { formatBytes } from '../../utils/format';

export type WorkbookErrorCode =
  | 'UNSUPPORTED_FORMAT'
  | 'NOT_A_WORKBOOK'
  | 'INVALID_FILE'
  | 'CORRUPTED'
  | 'PASSWORD_PROTECTED'
  | 'EMPTY_FILE'
  | 'NO_SHEETS'
  | 'TOO_LARGE'
  | 'OUT_OF_MEMORY'
  | 'NOT_FOUND'
  | 'PERMISSION_DENIED'
  | 'READ_FAILED'
  | 'CANCELLED'
  | 'UNKNOWN';

export interface WorkbookErrorDetails {
  /** File size in bytes (TOO_LARGE). */
  size?: number;
  /** Size limit in bytes (TOO_LARGE). */
  limit?: number;
}

/**
 * A failure with a stable code. `message` is technical detail for logs only;
 * the UI always goes through `toUserFacingError`.
 */
export class WorkbookError extends Error {
  readonly code: WorkbookErrorCode;
  readonly details: WorkbookErrorDetails;

  constructor(code: WorkbookErrorCode, detail?: string, details: WorkbookErrorDetails = {}) {
    super(detail ?? code);
    this.name = 'WorkbookError';
    this.code = code;
    this.details = details;
  }
}

/** Plain-object form used to cross the worker boundary. */
export interface SerializedWorkbookError {
  code: WorkbookErrorCode;
  message: string;
  details?: WorkbookErrorDetails;
}

export function serializeError(error: unknown): SerializedWorkbookError {
  if (error instanceof WorkbookError) {
    return { code: error.code, message: error.message, details: error.details };
  }
  if (isOutOfMemory(error)) return { code: 'OUT_OF_MEMORY', message: String((error as Error).message) };
  return { code: 'UNKNOWN', message: error instanceof Error ? error.message : String(error) };
}

export function deserializeError(error: SerializedWorkbookError): WorkbookError {
  return new WorkbookError(error.code, error.message, error.details);
}

export function isOutOfMemory(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  return (
    error instanceof RangeError &&
    /out of memory|invalid (array|string|typed array) length|allocation failed|array buffer allocation/i.test(error.message)
  );
}

export function isCancelled(error: unknown): boolean {
  return error instanceof WorkbookError && error.code === 'CANCELLED';
}

export interface UserFacingError {
  code: WorkbookErrorCode;
  title: string;
  message: string;
}

const GENERIC_MESSAGE =
  'The file may be corrupted, password protected, or use an unsupported Excel feature.';

/** Friendly message for a failed save. */
export function toSaveError(error: unknown, fileName: string): UserFacingError {
  const code = error instanceof WorkbookError ? error.code : 'UNKNOWN';
  const messages: Partial<Record<WorkbookErrorCode, string>> = {
    PERMISSION_DENIED: 'Sheaf doesn’t have permission to save in this location. Choose a different folder.',
    NOT_FOUND: 'The folder couldn’t be found. It may have been moved or disconnected.',
    UNSUPPORTED_FORMAT: 'Files can only be saved as .xlsx, .xlsm, .csv or .tsv.',
  };
  return {
    code,
    title: `Couldn’t save “${fileName}”`,
    message: messages[code] ?? 'The file couldn’t be saved. Check that the location is available and has free space, then try again.',
  };
}

/** Converts any failure into a short, friendly message. Never exposes stack traces. */
export function toUserFacingError(error: unknown, fileName?: string): UserFacingError {
  const workbookError =
    error instanceof WorkbookError
      ? error
      : isOutOfMemory(error)
        ? new WorkbookError('OUT_OF_MEMORY')
        : new WorkbookError('UNKNOWN');
  const title = fileName ? `Unable to open “${fileName}”` : 'Unable to open this workbook';
  const { size, limit } = workbookError.details;

  const messages: Record<WorkbookErrorCode, string> = {
    UNSUPPORTED_FORMAT:
      'This type of file isn’t supported. Sheaf opens .xlsx, .xlsm, .xls, .csv and .tsv files.',
    NOT_A_WORKBOOK:
      'The file isn’t an Excel workbook, even though its name says it is. It may have been saved by a different application.',
    INVALID_FILE:
      'The file doesn’t appear to be a valid spreadsheet. It may be damaged or saved in a different format.',
    CORRUPTED: GENERIC_MESSAGE,
    PASSWORD_PROTECTED:
      'The workbook is password protected. Remove the password in Excel, save a copy, and try again.',
    EMPTY_FILE: 'The file is empty (0 bytes).',
    NO_SHEETS: 'The workbook doesn’t contain any worksheets.',
    TOO_LARGE:
      size !== undefined && limit !== undefined
        ? `The file is ${formatBytes(size)}, which is more than the ${formatBytes(limit)} Sheaf can open.`
        : 'The file is too large to open.',
    OUT_OF_MEMORY:
      'There isn’t enough memory to open this workbook. Try closing other applications or splitting the file.',
    NOT_FOUND: 'The file couldn’t be found. It may have been moved, renamed, or deleted.',
    PERMISSION_DENIED: 'Sheaf doesn’t have permission to read this file.',
    READ_FAILED: 'The file couldn’t be read. Check that it’s still available and try again.',
    CANCELLED: 'Opening the file was cancelled.',
    UNKNOWN: GENERIC_MESSAGE,
  };

  return { code: workbookError.code, title, message: messages[workbookError.code] };
}
