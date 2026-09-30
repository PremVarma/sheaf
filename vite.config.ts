import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

// Tauri sets these while running `tauri dev` / `tauri build`.
const host = process.env.TAURI_DEV_HOST;
const targetsWindows = process.env.TAURI_ENV_PLATFORM === 'windows';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    host: host || false,
    hmr: host ? { protocol: 'ws', host, port: 1421 } : undefined,
    watch: { ignored: ['**/src-tauri/**'] },
  },
  envPrefix: ['VITE_', 'TAURI_ENV_'],
  build: {
    // WebView2 on Windows is evergreen Chromium; macOS uses the system WKWebView.
    target: targetsWindows ? 'chrome111' : 'safari16',
    sourcemap: Boolean(process.env.TAURI_ENV_DEBUG),
    // The parser worker bundles SheetJS (~1 MB), which is expected.
    chunkSizeWarningLimit: 2500,
  },
  worker: { format: 'es' },
  test: {
    environment: 'jsdom',
    setupFiles: ['./src/test/setup.ts'],
    css: false,
    restoreMocks: true,
  },
});
