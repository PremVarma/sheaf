const integerFormat = new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 });
const decimalFormat = new Intl.NumberFormat(undefined, { maximumFractionDigits: 10 });

/** 1245 → "1,245" (locale aware). */
export function formatCount(value: number): string {
  return integerFormat.format(value);
}

/** Numbers for the status bar summary: grouping, up to 10 decimals, no float noise. */
export function formatStatistic(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  return decimalFormat.format(Number(value.toPrecision(15)));
}

/** 2_516_582 → "2.4 MB". */
export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} ${bytes === 1 ? 'byte' : 'bytes'}`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = bytes / 1024;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value >= 100 ? Math.round(value) : value.toFixed(1).replace(/\.0$/, '')} ${units[unit]}`;
}

export function formatDateTime(iso: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return iso;
  return new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(date);
}

export function pluralize(count: number, singular: string, plural = `${singular}s`): string {
  return `${formatCount(count)} ${count === 1 ? singular : plural}`;
}

/** Last path segment for both POSIX and Windows paths. */
export function baseName(path: string): string {
  const parts = path.split(/[\\/]/);
  return parts[parts.length - 1] || path;
}

/** Directory part of a path, for showing where a recent file lives. */
export function dirName(path: string): string {
  const index = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));
  return index > 0 ? path.slice(0, index) : '';
}
