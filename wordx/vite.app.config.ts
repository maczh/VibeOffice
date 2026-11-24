import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { legacyRuntime } from '../tools/vite/legacy-runtime.mjs';
import { reactCsp } from '../tools/vite/react-csp.mjs';

export default defineConfig({ publicDir: false, plugins: [react(), reactCsp(), legacyRuntime('wordx')], build: { outDir: 'standalone', emptyOutDir: true } });
