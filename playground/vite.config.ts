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
      { find: '@voffice/vword', replacement: r('../vword/src') },
      { find: '@voffice/vexcel', replacement: r('../vexcel/src') },
      { find: '@voffice/vpptx', replacement: r('../vpptx/src') },
    ],
  },
  server: { port: 5170 },
})
