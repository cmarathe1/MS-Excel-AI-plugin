import { existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDb } from './db.js';
import { startServer } from './server.js';

const here = dirname(fileURLToPath(import.meta.url));

async function main(): Promise<void> {
  const db = openDb();
  const addinDist = join(here, '..', '..', 'addin', 'dist');
  const sidecar = await startServer({
    db,
    port: process.env.EXCELAI_PORT ? Number(process.env.EXCELAI_PORT) : 8923,
    staticDir: existsSync(addinDist) ? addinDist : undefined,
  });

  /* eslint-disable no-console */
  console.log(`Excel AI sidecar listening on http://127.0.0.1:${sidecar.port}`);
  if (sidecar.auth.currentPairingCode) {
    console.log(`Pairing code (enter it in the add-in once): ${sidecar.auth.currentPairingCode}`);
  }
  console.log('Press Ctrl+C to stop.');

  const shutdown = (): void => {
    void sidecar.close().finally(() => process.exit(0));
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}

main().catch((e) => {
  console.error('Sidecar failed to start:', e);
  process.exit(1);
});
