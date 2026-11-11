import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';
import { legacyRuntime } from '../../tools/vite/legacy-runtime.mjs';

export default defineConfig({
  publicDir: false,
  plugins: [react({ fastRefresh: false }), legacyRuntime(['quire', 'ledger', 'lectern'])],
  server: { fs: { allow: [resolve(__dirname, '../..')] } },
  resolve: {
    alias: {
      '@vibeoffice/core': resolve(__dirname, '../../commonx/src/index.ts'),
      '@vibeoffice/wordx': resolve(__dirname, '../../wordx/src/index.ts'),
      '@vibeoffice/excelx': resolve(__dirname, '../../excelx/src/index.ts'),
      '@vibeoffice/pptx': resolve(__dirname, '../../pptx/src/index.ts'),
    },
  },
});
