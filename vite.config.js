import { defineConfig } from 'vite';

export default defineConfig({
  server: {
    port: 5173,
    open: true,
    proxy: {
      '/api': 'https://ai-job-applicator.vercel.app',
      '/auth': 'https://ai-job-applicator.vercel.app'
    }
  }
});
