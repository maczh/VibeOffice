import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'node:path'

const r = (p: string) => resolve(__dirname, p)

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: [
      { find: '@voffice/shared/styles', replacement: r('../voffice-shared/src/styles/index.css') },
      { find: '@voffice/shared', replacement: r('../voffice-shared/src') },
      { find: '@', replacement: r('src') },
    ],
  },
  server: { port: 5173, strictPort: false },
  build: {
    lib: { entry: { index: r('src/index.ts'), engine: r('src/engine.ts') }, formats: ['es'], fileName: (_f, name) => name + '.js' },
    rollupOptions: { external: ['vue'] },
    cssCodeSplit: false,
  },
})
