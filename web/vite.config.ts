import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/enyuauto/',
  server: {
    proxy: {
      // Public enyu.acccn.org search RSS (the WP REST API is locked).
      '/enyu-search': {
        target: 'https://enyu.acccn.org',
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/enyu-search/, ''),
      },
    },
  },
})
