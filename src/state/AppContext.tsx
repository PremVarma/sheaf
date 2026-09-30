import { createContext, useContext, useSyncExternalStore, type ReactNode } from 'react';
import type { Platform } from '../services/platform/types';
import { loadRecentFiles } from '../services/recentFiles';
import type { ParseFunction } from '../services/workbook/workbookService';
import { ViewerController } from './controller';
import { createStore, type Store } from './store';
import { createInitialState, viewerReducer, type ViewerAction, type ViewerState } from './viewerState';

export interface AppServices {
  store: Store<ViewerState, ViewerAction>;
  controller: ViewerController;
  platform: Platform;
}

export function createAppServices(platform: Platform, options: { parse?: ParseFunction } = {}): AppServices {
  const recent = platform.kind === 'desktop' ? loadRecentFiles() : [];
  const store = createStore(viewerReducer, createInitialState(recent));
  return { store, platform, controller: new ViewerController(store, platform, options.parse) };
}

const AppContext = createContext<AppServices | null>(null);

export function AppProvider({ services, children }: { services: AppServices; children: ReactNode }) {
  return <AppContext.Provider value={services}>{children}</AppContext.Provider>;
}

function useServices(): AppServices {
  const services = useContext(AppContext);
  if (!services) throw new Error('AppProvider is missing');
  return services;
}

/** Subscribes to a slice of viewer state. The selector must return stable references. */
export function useViewer<T>(selector: (state: ViewerState) => T): T {
  const { store } = useServices();
  return useSyncExternalStore(
    store.subscribe,
    () => selector(store.getState()),
    () => selector(store.getState()),
  );
}

export function useController(): ViewerController {
  return useServices().controller;
}

export function usePlatform(): Platform {
  return useServices().platform;
}
