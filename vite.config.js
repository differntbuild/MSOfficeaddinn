import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { getHttpsServerOptions } from 'office-addin-dev-certs';
import path from 'path';

export default defineConfig(async ({ command, mode }) => {
  const httpsOptions = await getHttpsServerOptions();

  return {
    plugins: [react()],
    root: '.',
    publicDir: 'assets',
    server: {
      port: 3000,
      https: httpsOptions,
    },
    build: {
       outDir: 'dist',
       rollupOptions: {
         input: {
           taskpane: 'index.html',
         }
       }
    }
  };
});
