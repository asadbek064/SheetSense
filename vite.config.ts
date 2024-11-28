import { defineConfig } from 'vite';
import path from 'path';
import { configDefaults } from 'vitest/config';

export default defineConfig({
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  },
  build: {
    lib: {
      entry: path.resolve(__dirname, 'src/index.ts'),
      name: 'SheetSense',
      formats: ['iife', 'cjs'],
      fileName: (format) => {
        if (format === 'iife') return 'sheetsense.browser.js';
        return 'sheetsense.node.js';
      }
    },
    sourcemap: true,
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      // Only mark dependencies as external for CJS build
      external: (id) => {
        if (id === 'zod' || id === 'xlsx') {
          // Check the current format being built
          const currentFormat = process.env.FORMAT;
          return currentFormat === 'cjs';
        }
        return false;
      },
      // For IIFE build, define how external dependencies should be found in browser
      output: {
        globals: {
          zod: 'zod',
          xlsx: 'XLSX'
        }
      }
    }
  },
  test: {
    globals: true,
    environment: 'node',
    exclude: [...configDefaults.exclude, 'e2e/*'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html']
    }
  }
});