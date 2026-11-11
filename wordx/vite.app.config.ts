import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { legacyRuntime } from '../tools/vite/legacy-runtime.mjs';

export default defineConfig({ publicDir: false, plugins: [react(), legacyRuntime('quire')], build: { outDir: 'standalone' } });
