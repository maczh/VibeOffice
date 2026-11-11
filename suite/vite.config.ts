import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { legacyRuntime } from '../tools/vite/legacy-runtime.mjs';
import { reactPwa } from '../tools/vite/react-pwa.mjs';
import { reactCsp } from '../tools/vite/react-csp.mjs';

export default defineConfig({ publicDir: false, plugins: [react(), reactCsp(), legacyRuntime(['quire', 'ledger', 'lectern']), reactPwa()], build: { outDir: 'dist' } });
