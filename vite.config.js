import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

const buildId = Date.now()

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    rollupOptions: {
      output: {
        entryFileNames: `assets/app-${buildId}.js`,
        chunkFileNames: `assets/chunk-${buildId}-[hash].js`,
        assetFileNames: `assets/asset-${buildId}-[hash][extname]`,
      },
    },
  },
})
