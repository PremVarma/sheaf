import type { IndexMap } from '../../models/workbook';

/**
 * Monotone index maps (see IndexMap): current row/column → row/column before a
 * change. Used for a sheet's history since it was opened (SheetOrigin) and for a
 * single insert/delete when updating formulas.
 */

export const IDENTITY: IndexMap = { starts: [0], bases: [0] };

export function isIdentity(map: IndexMap): boolean {
  return map.starts.length === 1 && map.bases[0] === 0;
}

interface Segment {
  start: number;
  end: number;
  base: number;
}

function segments(map: IndexMap): Segment[] {
  return map.starts.map((start, k) => ({ start, end: map.starts[k + 1] ?? Infinity, base: map.bases[k] }));
}

/** Joins adjacent segments that continue each other; drops empty ones. */
function build(list: Segment[]): IndexMap {
  const starts: number[] = [];
  const bases: number[] = [];
  for (const seg of list.sort((a, b) => a.start - b.start)) {
    if (seg.end <= seg.start) continue;
    const k = starts.length - 1;
    if (k >= 0) {
      const prevLength = seg.start - starts[k];
      const continues = bases[k] < 0 ? seg.base < 0 : seg.base === bases[k] + prevLength;
      if (continues) continue;
    }
    starts.push(seg.start);
    bases.push(seg.base);
  }
  if (starts.length === 0 || starts[0] !== 0) return IDENTITY;
  return { starts, bases };
}

/** Last segment starting at or before `i`. */
function segmentAt(map: IndexMap, i: number): number {
  let lo = 0;
  let hi = map.starts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >>> 1;
    if (map.starts[mid] <= i) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** The earlier index of current index `i`, or -1 when it was inserted since. */
export function toBase(map: IndexMap, i: number): number {
  const k = segmentAt(map, i);
  const base = map.bases[k];
  return base < 0 ? -1 : base + (i - map.starts[k]);
}

/** The current index of earlier index `b`, or -1 when it was deleted. */
export function fromBase(map: IndexMap, b: number): number {
  for (let k = 0; k < map.starts.length; k++) {
    const base = map.bases[k];
    if (base < 0 || b < base) continue;
    const length = (map.starts[k + 1] ?? Infinity) - map.starts[k];
    if (b < base + length) return map.starts[k] + (b - base);
  }
  return -1;
}

/**
 * Where an earlier range [b0, b1] is now: its surviving entries, or null when all
 * of them were deleted (or pushed past `limit`). Entries inserted inside the range
 * become part of it, as in Excel.
 */
export function mapRange(map: IndexMap, b0: number, b1: number, limit: number): [number, number] | null {
  let first = -1;
  let last = -1;
  for (let k = 0; k < map.starts.length; k++) {
    const base = map.bases[k];
    if (base < 0) continue;
    const length = (map.starts[k + 1] ?? Infinity) - map.starts[k];
    const lo = Math.max(b0, base);
    const hi = Math.min(b1, base + length - 1);
    if (lo > hi) continue;
    if (first < 0) first = map.starts[k] + (lo - base);
    last = map.starts[k] + (hi - base);
  }
  if (first < 0 || first >= limit) return null;
  return [first, Math.min(last, limit - 1)];
}

/** After inserting `count` entries at current index `at`. */
export function insertAt(map: IndexMap, at: number, count: number): IndexMap {
  const out: Segment[] = [{ start: at, end: at + count, base: -1 }];
  for (const seg of segments(map)) {
    if (seg.end <= at) out.push(seg);
    else if (seg.start >= at) out.push({ start: seg.start + count, end: seg.end + count, base: seg.base });
    else {
      out.push({ start: seg.start, end: at, base: seg.base });
      out.push({ start: at + count, end: seg.end + count, base: seg.base < 0 ? -1 : seg.base + (at - seg.start) });
    }
  }
  return build(out);
}

/** After deleting current indexes [at, at + count). */
export function removeAt(map: IndexMap, at: number, count: number): IndexMap {
  const end = at + count;
  const out: Segment[] = [];
  for (const seg of segments(map)) {
    if (seg.start < at) out.push({ start: seg.start, end: Math.min(seg.end, at), base: seg.base });
    if (seg.end > end) {
      const from = Math.max(seg.start, end);
      out.push({ start: from - count, end: seg.end - count, base: seg.base < 0 ? -1 : seg.base + (from - seg.start) });
    }
  }
  return build(out);
}

/** The map for one insertion or deletion (current → before it). */
export function changeMap(kind: 'insert' | 'delete', at: number, count: number): IndexMap {
  return kind === 'insert' ? insertAt(IDENTITY, at, count) : removeAt(IDENTITY, at, count);
}
