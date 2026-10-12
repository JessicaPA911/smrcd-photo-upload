import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests/browser',
  use: { baseURL: 'http://127.0.0.1:5173/smrcd-photo-upload/', channel: 'msedge' },
  webServer: {
    command: 'npm run dev -- --host 127.0.0.1',
    url: 'http://127.0.0.1:5173/smrcd-photo-upload/',
    reuseExistingServer: true
  }
});
