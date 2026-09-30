/** Recently opened file paths (desktop only), newest first, kept in local storage. */

const STORAGE_KEY = 'sheaf.recent-files';
export const MAX_RECENT_FILES = 10;

export function loadRecentFiles(): string[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed.filter((p): p is string => typeof p === 'string').slice(0, MAX_RECENT_FILES) : [];
  } catch {
    return [];
  }
}

function save(paths: string[]): string[] {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(paths));
  } catch {
    // Storage can be unavailable; the list just won't persist.
  }
  return paths;
}

export function addRecentFile(paths: string[], path: string): string[] {
  return save([path, ...paths.filter((p) => p !== path)].slice(0, MAX_RECENT_FILES));
}

export function removeRecentFile(paths: string[], path: string): string[] {
  return save(paths.filter((p) => p !== path));
}

export function clearRecentFiles(): string[] {
  return save([]);
}
