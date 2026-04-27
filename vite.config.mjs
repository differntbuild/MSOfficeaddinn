import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import * as officeAddin from 'office-addin-dev-certs';

export default defineConfig(async ({ command, mode }) => {
  const httpsOptions = await officeAddin.getHttpsServerOptions();

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
