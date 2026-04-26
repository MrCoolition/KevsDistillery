import { createRequire } from 'node:module';

import { loadDotEnv } from './env.mjs';

loadDotEnv();

const require = createRequire(import.meta.url);
const { databaseStatus, ensureSchema, hasDatabase, REQUIRED_TABLES } = require('../api/_lib/db.js');

if (!hasDatabase()) {
  console.error('Set DATABASE_URL, POSTGRES_URL, or NEON_DATABASE_URL before running the migration.');
  process.exit(1);
}

await ensureSchema();
const status = await databaseStatus();

if (!status.ready) {
  console.error(`Distillery schema is missing tables: ${status.missingTables.join(', ')}`);
  process.exit(1);
}

console.log(`Applied Uncle Kev's Distillery Neon schema: ${status.schema}.${REQUIRED_TABLES.length} tables ready.`);
