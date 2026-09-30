import type { ButtonHTMLAttributes, ReactNode } from 'react';
import { Icon, type IconName } from './Icon';

/** Toolbar-style button with an optional icon. */
export function ToolButton({
  icon,
  label,
  children,
  active = false,
  className = '',
  ...rest
}: { icon?: IconName; label: string; active?: boolean; children?: ReactNode } & ButtonHTMLAttributes<HTMLButtonElement>) {
  return (
    <button
      type="button"
      aria-label={label}
      title={rest.title ?? label}
      aria-pressed={rest['aria-pressed']}
      className={`inline-flex h-7 min-w-7 shrink-0 items-center justify-center gap-1.5 rounded-md px-1.5 text-fg transition-colors hover:bg-hover active:bg-pressed disabled:pointer-events-none disabled:opacity-40 ${
        active ? 'bg-accent-soft text-accent' : ''
      } ${className}`}
      {...rest}
    >
      {icon && <Icon name={icon} />}
      {children}
    </button>
  );
}

/** "⌘O" on macOS, "Ctrl+O" elsewhere. */
export function shortcutLabel(key: string, isMac: boolean, shift = false): string {
  if (isMac) return `${shift ? '⇧' : ''}⌘${key}`;
  return `Ctrl+${shift ? 'Shift+' : ''}${key}`;
}
