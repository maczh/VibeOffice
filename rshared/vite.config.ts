import { defineConfig } from 'vite'
import { resolve } from 'node:path'
export default defineConfig({ build: { lib: { entry: resolve(import.meta.dirname, 'src/index.ts'), formats: ['es'], fileName: 'index' }, rollupOptions: { external: ['react', 'react-dom', 'react/jsx-runtime'] }, cssCodeSplit: false } })
