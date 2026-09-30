import type { SVGProps } from 'react';

const PATHS = {
  open: 'M3 6.5V17a1.5 1.5 0 0 0 1.5 1.5h15A1.5 1.5 0 0 0 21 17V9a1.5 1.5 0 0 0-1.5-1.5h-7L10.8 5.3a1.5 1.5 0 0 0-1.2-.6H4.5A1.5 1.5 0 0 0 3 6.2Z',
  search: 'M10.5 17a6.5 6.5 0 1 0 0-13 6.5 6.5 0 0 0 0 13Zm4.7-1.8L20 20',
  close: 'M6 6l12 12M18 6 6 18',
  chevronLeft: 'M14.5 6 8.5 12l6 6',
  chevronRight: 'm9.5 6 6 6-6 6',
  chevronUp: 'm6 14.5 6-6 6 6',
  chevronDown: 'm6 9.5 6 6 6-6',
  plus: 'M12 5v14M5 12h14',
  minus: 'M5 12h14',
  info: 'M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18Zm0-5v-5m0-3.2v-.1',
  lock: 'M7 10.5V8a5 5 0 0 1 10 0v2.5M6.5 10.5h11A1.5 1.5 0 0 1 19 12v7a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 5 19v-7a1.5 1.5 0 0 1 1.5-1.5Z',
  sheet: 'M4.5 4h15A1.5 1.5 0 0 1 21 5.5v13a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18.5v-13A1.5 1.5 0 0 1 4.5 4ZM3 9h18M3 14.5h18M9 9v11',
  list: 'M8 6.5h12M8 12h12M8 17.5h12M4 6.5h.01M4 12h.01M4 17.5h.01',
  warning: 'M12 3.8 21.5 20H2.5L12 3.8Zm0 6.2v4.5m0 2.7v.1',
  clock: 'M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18Zm0-13.5V12l3 2',
  file: 'M14 3.5H7A1.5 1.5 0 0 0 5.5 5v14A1.5 1.5 0 0 0 7 20.5h10a1.5 1.5 0 0 0 1.5-1.5V8L14 3.5Zm0 0V8h4.5',
  eyeOff: 'M3 3l18 18M10.6 5.1A9.7 9.7 0 0 1 12 5c5 0 8.5 4.3 9.5 7-.4 1-1.2 2.4-2.4 3.7M6.6 6.6C4.6 7.9 3.2 9.9 2.5 12c1 2.7 4.5 7 9.5 7 1.9 0 3.6-.6 5-1.5M9.9 9.9a3 3 0 0 0 4.2 4.2',
  drop: 'M12 4v11m0 0-4-4m4 4 4-4M5 19.5h14',
  newFile: 'M14 3.5H7A1.5 1.5 0 0 0 5.5 5v14A1.5 1.5 0 0 0 7 20.5h10a1.5 1.5 0 0 0 1.5-1.5V8L14 3.5Zm0 0V8h4.5M12 11v6m-3-3h6',
  save: 'M5 3.5h11.5L20.5 7.5V19a1.5 1.5 0 0 1-1.5 1.5H5A1.5 1.5 0 0 1 3.5 19V5A1.5 1.5 0 0 1 5 3.5Zm3 0v5h7v-5M7.5 20.5v-6h9v6',
  undo: 'M9 14 4 9l5-5M4 9h10.5a5.5 5.5 0 0 1 0 11H11',
  redo: 'm15 14 5-5-5-5m5 5H9.5a5.5 5.5 0 0 0 0 11H13',
  fill: 'M4.5 11.5 11 5l6.5 6.5L11 18l-6.5-6.5Zm0 0h13M8.5 3.5 11 6M20 14.5c.9 1.2 1.4 2.1 1.4 2.8a1.4 1.4 0 0 1-2.8 0c0-.7.5-1.6 1.4-2.8Z',
  borders: 'M4.5 4.5h15v15h-15zM4.5 12h15M12 4.5v15',
  alignLeft: 'M4 6h16M4 10h10M4 14h16M4 18h10',
  alignCenter: 'M4 6h16M7 10h10M4 14h16M7 18h10',
  alignRight: 'M4 6h16M10 10h10M4 14h16M10 18h10',
  alignTop: 'M4 4.5h16M12 20V9m-4 4 4-4 4 4',
  alignMiddle: 'M4 12h16M12 3.5v5m-3-2.5 3 2.5 3-2.5M12 20.5v-5m-3 2.5 3-2.5 3 2.5',
  alignBottom: 'M4 19.5h16M12 4v11m-4-4 4 4 4-4',
  wrap: 'M4 6h16M4 12h12.5a3 3 0 0 1 0 6H12m0 0 2-2m-2 2 2 2M4 18h5',
  merge: 'M3.5 5.5h17v13h-17zM7 12h10m-7.5-2.5L7 12l2.5 2.5m5-5L17 12l-2.5 2.5',
  sortAsc: 'M7 4v16m0 0-3-3m3 3 3-3M13.5 6h6.5M13.5 12h4.5M13.5 18h2.5',
  sortDesc: 'M7 4v16m0 0-3-3m3 3 3-3M13.5 6h2.5M13.5 12h4.5M13.5 18h6.5',
  eraser: 'M8.5 18.5h11M5.9 14.9l6.6-6.6a2 2 0 0 1 2.8 0l2.4 2.4a2 2 0 0 1 0 2.8l-5.2 5.2h-4.2l-2.4-2.4a1 1 0 0 1 0-1.4ZM9.5 11.5l5 5',
  indentMore: 'M4 5h16M11 10h9M11 14h9M4 19h16M4 9.5 7 12l-3 2.5',
  indentLess: 'M4 5h16M11 10h9M11 14h9M4 19h16M7 9.5 4 12l3 2.5',
  rowInsert: 'M3.5 9.5h17v9h-17zM3.5 14h17M12 2.5v5m-2.5-2.5h5',
  columnInsert: 'M9.5 3.5h9v17h-9zM14 3.5v17M2.5 12h5M5 9.5v5',
  rowDelete: 'M3.5 9.5h17v9h-17zM3.5 14h17M9.5 5h5',
  columnDelete: 'M9.5 3.5h9v17h-9zM14 3.5v17M3 12h5',
  table: 'M4.5 4.5h15v15h-15zM4.5 9.5h15M4.5 14.5h15M9.5 9.5v10',
  more: 'M6 12h.01M12 12h.01M18 12h.01',
} as const;

export type IconName = keyof typeof PATHS;

export function Icon({ name, size = 16, ...rest }: { name: IconName; size?: number } & SVGProps<SVGSVGElement>) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.8}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
      {...rest}
    >
      <path d={PATHS[name]} />
    </svg>
  );
}
