import { describe, expect, it } from 'vitest';
import { Axis } from './axis';

describe('Axis', () => {
  const sizes = { index: Uint32Array.from([2, 5]), size: Float32Array.from([50, 0]) };
  const axis = Axis.create(10, 20, sizes);

  it('computes offsets with custom and hidden sizes', () => {
    expect(axis.offsetOf(0)).toBe(0);
    expect(axis.offsetOf(2)).toBe(40);
    expect(axis.offsetOf(3)).toBe(90);
    expect(axis.offsetOf(5)).toBe(130);
    expect(axis.offsetOf(6)).toBe(130); // row 5 is hidden
    expect(axis.total).toBe(8 * 20 + 50);
    expect(axis.sizeOf(2)).toBe(50);
    expect(axis.isHidden(5)).toBe(true);
  });

  it('finds the entry at an offset', () => {
    expect(axis.indexAt(-5)).toBe(0);
    expect(axis.indexAt(39)).toBe(1);
    expect(axis.indexAt(40)).toBe(2);
    expect(axis.indexAt(89)).toBe(2);
    expect(axis.indexAt(130)).toBe(6); // skips the hidden entry
    expect(axis.indexAt(10_000)).toBe(9);
  });

  it('layers user sizes over file sizes', () => {
    const resized = Axis.create(10, 20, sizes, new Map([[2, 20], [0, 100]]));
    expect(resized.sizeOf(0)).toBe(100);
    expect(resized.sizeOf(2)).toBe(20);
    expect(resized.total).toBe(100 + 7 * 20 + 20);
  });

  it('stays fast with a million rows', () => {
    const big = Axis.create(1_048_576, 20);
    expect(big.offsetOf(1_000_000)).toBe(20_000_000);
    expect(big.indexAt(20_000_010)).toBe(1_000_000);
    expect(big.nextVisible(3, 1)).toBe(3);
  });
});
