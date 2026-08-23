import { defineConfig } from 'vitest/config';
import path from 'path';

export default defineConfig({
  test: {
    environment: 'node',
    include: ['src/**/__tests__/**/*.test.ts'],
    globals: false,

    // xlsx & jszip terdiri dari ribuan file kecil. esbuild membundelnya
    // jadi beberapa file saja — waktu import turun drastis di mesin yang
    // pembacaan node_modules-nya lambat. exceljs sengaja TIDAK dibundel:
    // stream internalnya rusak kalau di-bundle esbuild.
    deps: {
      optimizer: {
        ssr: {
          enabled: true,
          include: ['xlsx', 'jszip'],
        },
      },
    },
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
});
