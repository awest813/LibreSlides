import { defineConfig } from 'vite';
import { VitePWA } from 'vite-plugin-pwa';
import { execSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

// Shown in Help › About and in bug-report details (src/diagnostics.js).
const version = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')).version;
let commit = 'unknown';
try {
  commit = execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
} catch { /* not a git checkout */ }

// `base: './'` keeps every asset URL relative, so the build in `dist/` can be
// served from any static host or sub-path (GitHub Pages, Netlify, S3, a USB stick…).
export default defineConfig({
  base: './',
  define: {
    __APP_VERSION__: JSON.stringify(version),
    __BUILD_COMMIT__: JSON.stringify(commit),
    __BUILD_DATE__: JSON.stringify(new Date().toISOString().slice(0, 10)),
  },
  build: {
    target: 'es2022',
    sourcemap: false,
    cssCodeSplit: true,
    chunkSizeWarningLimit: 1200,
  },
  plugins: [
    VitePWA({
      // A new version waits until the user agrees to reload (see main.js),
      // instead of never applying while any LibreSlides window is open.
      registerType: 'prompt',
      injectRegister: false,
      includeAssets: ['favicon.svg', 'icons/*.png'],
      manifest: {
        id: './',
        name: 'LibreSlides',
        short_name: 'LibreSlides',
        description: 'Make and present slides in your browser, offline. Opens and saves PowerPoint files.',
        theme_color: '#c43e1c',
        background_color: '#f3f2f1',
        display: 'standalone',
        start_url: './',
        scope: './',
        categories: ['productivity', 'education', 'business'],
        icons: [
          { src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' },
          { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png' },
          { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
        ],
        // Jump lists / long-press menu on the app icon.
        shortcuts: [
          { name: 'New presentation', short_name: 'New', url: './#/new', icons: [{ src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' }] },
        ],
        // Files opened from the file manager go to an open LibreSlides window instead of a new one.
        launch_handler: { client_mode: ['focus-existing', 'auto'] },
        // "Open with LibreSlides" in the file manager once the app is installed.
        file_handlers: [
          {
            action: './',
            icons: [{ src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' }],
            accept: {
              'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
              'application/vnd.libreslides+json': ['.lslides'],
            },
          },
        ],
      },
      workbox: {
        // WOFF too: the bundled fonts are embedded in PowerPoint files from those (src/io/bundled-fonts.js).
        globPatterns: ['**/*.{js,css,html,svg,png,woff,woff2}'],
        maximumFileSizeToCacheInBytes: 6 * 1024 * 1024,
        navigateFallback: 'index.html',
        cleanupOutdatedCaches: true,
      },
    }),
  ],
  test: {
    environment: 'jsdom',
    // One VM context per file instead of a new worker with its own jsdom: about a third faster.
    pool: 'vmThreads',
    include: ['tests/unit/**/*.test.js'],
  },
});
