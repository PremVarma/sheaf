import { createBrowserPlatform } from './browserPlatform';
import { createTauriPlatform } from './tauriPlatform';
import type { Platform } from './types';

export type { Platform, PlatformEvents } from './types';

export function isTauri(): boolean {
  return typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;
}

export function createPlatform(): Platform {
  return isTauri() ? createTauriPlatform() : createBrowserPlatform();
}
