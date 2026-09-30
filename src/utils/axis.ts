import type { SizeOverrides } from '../models/workbook';

/**
 * Positions along one grid axis (rows or columns) where most entries share a
 * default size and a sparse set of entries have custom sizes (0 = hidden).
 *
 * All lookups are O(log k) in the number of custom sizes, so a sheet with a
 * million rows costs nothing until rows are actually resized. Instances are
 * immutable; `withSize` returns a new axis.
 */
export class Axis {
  readonly count: number;
  readonly defaultSize: number;
  readonly total: number;
  private readonly idx: Int32Array;
  private readonly sizes: Float64Array;
  /** prefix[k] = sum over j < k of (sizes[j] - defaultSize). Length idx.length + 1. */
  private readonly prefix: Float64Array;

  private constructor(count: number, defaultSize: number, idx: Int32Array, sizes: Float64Array) {
    this.count = count;
    this.defaultSize = defaultSize;
    this.idx = idx;
    this.sizes = sizes;
    this.prefix = new Float64Array(idx.length + 1);
    for (let k = 0; k < idx.length; k++) this.prefix[k + 1] = this.prefix[k] + (sizes[k] - defaultSize);
    this.total = count * defaultSize + this.prefix[idx.length];
  }

  /**
   * @param base    sizes from the file (sorted by index)
   * @param overrides user adjustments; these win over `base`
   */
  static create(
    count: number,
    defaultSize: number,
    base?: SizeOverrides,
    overrides?: ReadonlyMap<number, number>,
  ): Axis {
    const baseIdx = base?.index ?? new Uint32Array(0);
    const baseSize = base?.size ?? new Float32Array(0);
    const extra = overrides && overrides.size > 0 ? [...overrides.keys()].sort((a, b) => a - b) : [];

    const idx: number[] = [];
    const sizes: number[] = [];
    const push = (i: number, size: number) => {
      if (i < 0 || i >= count || size === defaultSize) return;
      idx.push(i);
      sizes.push(size);
    };

    // Merge the two sorted lists; user overrides replace file sizes at the same index.
    let a = 0;
    let b = 0;
    while (a < baseIdx.length || b < extra.length) {
      const ia = a < baseIdx.length ? baseIdx[a] : Infinity;
      const ib = b < extra.length ? extra[b] : Infinity;
      if (ib <= ia) {
        push(ib, overrides!.get(ib)!);
        b++;
        if (ia === ib) a++;
      } else {
        push(ia, baseSize[a]);
        a++;
      }
    }
    return new Axis(count, defaultSize, Int32Array.from(idx), Float64Array.from(sizes));
  }

  /** Number of custom-size entries strictly before `i`. */
  private lowerBound(i: number): number {
    let lo = 0;
    let hi = this.idx.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (this.idx[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  sizeOf(i: number): number {
    const k = this.lowerBound(i);
    return k < this.idx.length && this.idx[k] === i ? this.sizes[k] : this.defaultSize;
  }

  /** Start offset of entry `i`; `offsetOf(count)` equals `total`. */
  offsetOf(i: number): number {
    if (i <= 0) return 0;
    if (i >= this.count) return this.total;
    return i * this.defaultSize + this.prefix[this.lowerBound(i)];
  }

  /** Index of the entry containing `offset`, clamped to [0, count - 1]. */
  indexAt(offset: number): number {
    if (this.count === 0) return 0;
    if (offset <= 0) return 0;
    if (offset >= this.total) return this.count - 1;
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >>> 1;
      if (this.offsetOf(mid) <= offset) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  isHidden(i: number): boolean {
    return this.sizeOf(i) === 0;
  }

  /** Nearest non-hidden index from `i` in direction `step` (inclusive), or -1. */
  nextVisible(i: number, step: 1 | -1): number {
    for (let j = i; j >= 0 && j < this.count; j += step) {
      if (this.sizeOf(j) > 0) return j;
    }
    return -1;
  }

  /** Same axis with a different entry count (e.g. when the display extent grows). */
  withCount(count: number): Axis {
    if (count === this.count) return this;
    const keep = this.lowerBound(count);
    return new Axis(count, this.defaultSize, this.idx.slice(0, keep), this.sizes.slice(0, keep));
  }
}

/** Integer pixel position of an axis offset at a zoom level (keeps gridlines crisp). */
export function scaled(offset: number, zoom: number): number {
  return Math.round(offset * zoom);
}
