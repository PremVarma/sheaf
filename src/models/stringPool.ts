/**
 * A sheet's display strings packed into one string plus offsets.
 *
 * A million separate strings are slow to copy out of the parser worker and cost
 * an object header each; one packed string transfers in milliseconds, and
 * reading an entry is a cheap substring.
 */
export interface StringPool {
  data: string;
  /** Start of entry i; entry i ends where entry i + 1 starts. Length = count + 1. */
  offsets: Uint32Array;
}

export function createStringPool(strings: readonly string[]): StringPool {
  const offsets = new Uint32Array(strings.length + 1);
  let total = 0;
  for (let i = 0; i < strings.length; i++) {
    offsets[i] = total;
    total += strings[i].length;
  }
  offsets[strings.length] = total;
  return { data: strings.join(''), offsets };
}

export function stringAt(pool: StringPool, index: number): string {
  return pool.data.slice(pool.offsets[index], pool.offsets[index + 1]);
}

export function stringCount(pool: StringPool): number {
  return pool.offsets.length - 1;
}
