import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

const buildId = Date.now()

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    // Augmenter la limite d'avertissement chunk (notre App.jsx est gros intentionnellement)
    chunkSizeWarningLimit: 1500,
    rollupOptions: {
      output: {
        entryFileNames: `assets/app-${buildId}.js`,
        chunkFileNames: `assets/chunk-${buildId}-[hash].js`,
        assetFileNames: `assets/asset-${buildId}-[hash][extname]`,
        // Séparer les grosses dépendances vendor du code app
        manualChunks(id) {
          // Supabase Realtime — chargé séparément
          if (id.includes('@supabase/realtime-js')) {
            return 'vendor-supabase';
          }
          // SheetJS (xlsx) — très lourd, utilisé rarement (import Excel)
          if (id.includes('xlsx') || id.includes('sheetjs')) {
            return 'vendor-xlsx';
          }
          // React lui-même
          if (id.includes('node_modules/react/') || id.includes('node_modules/react-dom/')) {
            return 'vendor-react';
          }
        }
      }
    }
  },
  // Optimisation des dépendances en dev
  optimizeDeps: {
    include: ['react', 'react-dom', '@supabase/realtime-js', 'xlsx']
  }
})
